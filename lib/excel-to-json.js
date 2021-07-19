const XLSX = require('xlsx')

const REF = '!ref'
const MERGES = '!merges'
const CHAR_CODE_COL = 'A'.charCodeAt(0)

function parseCellAddress(s) {
  let i = 0, c = 0, r = 0, code = 0
  while (i < s.length) {
    code = s.charCodeAt(i)
    if (code < CHAR_CODE_COL) break
    i++
    c = c * 26 + (code - CHAR_CODE_COL)
  }
  r = parseInt(s.slice(i), 10)
  return { c, r }
}

function formatCellAddress(c, r) {
  let codes = []
  while (c >= 26) {
    codes.unshift(CHAR_CODE_COL + c % 26)
    c = Math.floor(c / 26)
  }
  codes.unshift(CHAR_CODE_COL + c % 26)
  return String.fromCharCode(...codes) + r
}

function parseRange(s) {
  const [s, e] = s.split(':', 2)
  return {
    s: parseCellAddress(s),
    e: parseCellAddress(e),
  }
}

function formatRange(range) {
  const { s, e } = range
  return formatCellAddress(s) + ':' + formatCellAddress(e)
}

function cloneRange(range) {
  return {
    s: Object.assign({}, range.s),
    e: Object.assign({}, range.e),
  }
}

function computeRange(sheet, { defaultRange, throwError }) {
  if (sheet[REF]) {
    return parseRange(sheet[REF])
  }

  if (defaultRange) {
    if (typeof defaultRange === 'function') return defaultRange(sheet)
    if (typeof defaultRange === 'object' && defaultRange.s && defaultRange.e) return cloneRange(defaultRange)
    const err = new Error('unknown defaultRange: ' + defaultRange + '(' + (typeof defaultRange) + ')')
    if (throwError) throw err
    console.error('excel-to-json:', err)
  }

  let min, max
  Object.keys(sheet).filter(it => !it.startsWith('!')).forEach(it => {
    if (!min) min = it
    else if (it < min) min = it
    if (!max) max = it
    else if (it > max) max = it
  })

  if (!min) return null
  return { s: parseCellAddress(min), e: parseCellAddress(max) }
}

function cellContent(cell) {
  return cell.w || cell.z || cell.v
}

function cellLessOrEq(a, b) {
  return a.c <= b.c && a.r <= b.r
}

function cellEq(a, b) {
  return a.c === b.c && a.r === b.r
}

const header = {
  'product_name': {
    addr: '',
    c: 0,
    r: 0,
    dataIndex: 0,
    range: {},
    cell: {},
  },
  'strategy': {
    alias: '策略',
    nested: {
      'label_type': {
        index: 2,
      },
      'label_operator': {
        index: 3,
      }
    },
  }
}

const headerMap = {
  '1': {
    path: ['product_name'],
    parent: null,
    name: 'product_name',
    level: 0,
    meta: {
      index: 1,
      alias: ''
    }
  },
  2: {
    path: ['strategy', 'label_type'],
    parent: 'strategy',
    name: 'label_type',
    level: 1,
    meta: {}
  }
}

function sheetHeader(sheet, k0, k1, range, keyFunc = identity) {
  let cur = Object.assign({}, range.s)
  const header = {}
  let cell, key
  while (cur[k0] <= range.e[k0]) {
    if (!(cell = sheet[formatCellAddress(cur.c, cur.r)])) {
      break
    }
    merge = mergedCellRange(sheet[MERGES], cur.c, cur.r)
    lenDelta += merge.e[k0] - merge.s[k0] + 1
    len += lenDelta
    key = keyFunc(cellContent(cell), cell, k0, sheet)
    header[key] = {
      addr: formatRange(merge),
      c: cur.c,
      r: cur.r,
      range: merge,
      cell,
    }
    if (lenDelta > 1) {
      header[key].nested = sheetHeader(sheet, k0, k1, { s: merge.s, e: { [k0]: merge.e[k0], [k1]: range.e[k1] } }, keyFunc)
    } else {
      header[key].dataIndex = cur[k0]
    }
  }
  return header
}

function identity(x) {
  return x
}

function sheetColHeader(sheet, range, { headerKey } = {}) {
  return sheetHeader(sheet, 'c', 'r', range, headerKey)
}

function sheetRowHeader(sheet, range, { headerKey } = {}) {
  return sheetHeader(sheet, 'r', 'c', range, headerKey)
}

function mergedCellRange(merges, c, r) {
  const merge = merges.find(it => it.s.c === c && it.s.r === r)
  if (merge == null) return { s: { c, r }, e: { c, r } }
  return { s: Object.assign({}, merge.s), e: Object.assign({}, merge.e) }
}

function sheetHeaders(sheet, { range, headerType, ...options } = {}) {
  if (!range) {
    range = computeRange(sheet, options)
    if (range == null) return { col: undefined, row: undefined }
  } else if (range) {
    range = parseRange(range)
  }
  const { s: { r: START_ROW, c: START_COL }, e: { r: END_ROW, c: END_COL } } = range
  const merges = sheet[MERGES]
  let r = START_ROW, c = START_COL

  const { s: s0, e: e0 } = mergedCellRange(merges, c, r)
  let col, row

  let cell
  if (cell = sheet[formatCellAddress(c, r)] && cellContent(cell) == null) {
    if (headerType === 'both') return { col, row }
    if (headerType == null) headerType = 'col'
  }

  if (['both', 'col'].includes(headerType)) {
    col = sheetColHeader(sheet, { s: { c: Math.min(e0.c + 1, START_COL), r: s0.r }, e: range.e })
  }
  if (['both', 'row'].includes(headerType)) {
    row = sheetRowHeader(sheet, { s: { c: s0.c, r: Math.min(e0.r + 1, START_ROW) }, e: range.e })
  }
  return { col, row, range }
}

function compileHeader(header) {
  const ret = {}
  const visit = (t, name, level = 0, path = [], metas = []) => {
    if (t.nested) {
      Object.keys(t.nested).forEach(p => {
        visit(t.nested[p], p, level + 1, path.concat(p), [t, ...metas])
      })
    } else {
      const meta = t
      ret[t.dataIndex] = {
        name,
        dataIndex: t.dataIndex,
        level,
        path,
        parent: path[path.length - 2] || null,
        meta: t,
        metas,
      }
    }
  }
  visit(header)
  return ret
}

function headerRange(header) {
  let sc, sr, ec, er
  for (const k in header) {
    const { c, r } = header[k]
    if (sc == null || c < sc) sc = c
    if (sr == null || r < sr) sr = r
    if (ec == null || c > ec) ec = c
    if (er == null || r > er) er = r
  }
  return { s: { c: sc, r: sr }, e: { c: ec, r: er } }
}

function minMaxDataIndex(header) {
  let min, max
  for (const k in header) {
    const n = header[k].dataIndex
    if (min == null || n < min) min = n
    if (max == null || n > max) max = n
  }
  return [min, max]
}

function inRange(a, b) {
  return a.s.c >= b.s.c && a.s.r >= b.s.r && a.e.c <= b.e.c && a.e.r <= b.e.r
}

function computeDataRange(range, colHeader, rowHeader, { throwError }) {
  if (!colHeader && !rowHeader) return cloneRange(range)
  if (colHeader && rowHeader) {
    const [sc, ec] = minMaxDataIndex(colHeader)
    const [sr, er] = minMaxDataIndex(rowHeader)
    return { s: { c: sc, r: sr }, e: { c: ec, r: er } }
  }
  const r = headerRange(colHeader || rowHeader)
  if (!inRange(r, range)) {
    if (throwError) throw new Error('col header not in range')
    return null
  }
  if (colHeader) {
    return { s: { c: r.s.c, r: r.e.r + 1 }, e: { c: r.e.c, r: range.e.r } }
  } else {
    return { s: { c: r.s.c+1, r: r.e.r }, e: { c: range.e.c, r: r.e.r } }
  }
}

function normalizeHeader(sheet, header, k, addrMandatory, { throwError }) {
  for (const p in header) {
    const v = header[p]
    if (!v.addr && addrMandatory) {
      if (throwError) throw new Error('invalid header: addr is required since not both col and row headers being provided')
      return null
    }
    if (!v.addr && v.dataIndex == null) {
      if (throwError) throw new Error('invalid header: require at least one of addr and dataIndex')
      return null
    }
    if (v.addr) {
      const range = parseRange(v.addr)
      v.range = range
      v.c = range.s.c
      v.r = range.s.r
      v.cell = sheet[formatCellAddress(v.c, v.r)]
    }
    if (!v.dataIndex && v[k]) {
      v.dataIndex = v[k]
    }
  }
  return header
}

function sheetToJSON(sheet, { headers, dataRange, ...options } = {}) {
  if (!headers) {
    headers = sheetHeaders(sheet, options)
  } else {
    const addrMandatory = !(headers.col && headers.row)
    for (const k of ['col', 'row']) {
      if (!k in headers) continue
      headers[k] = normalizeHeader(sheet, headers[k], addrMandatory, k.charAt(0), options)
      if (headers[k] == null || Object.keys(headers[k]).length === 0) return []
    }
  }
  let colHeader, rowHeader
  if (headers.col) {
    colHeader = compileHeader(headers.col)
  }
  if (headers.row) {
    rowHeader = compileHeader(headers.row)
  }

  if (!dataRange) {
    dataRange = computeDataRange(
      headers.range || computeRange(sheet, options), colHeader, rowHeader, options
    )
    if (!dataRange) return []
  } else {
    dataRange = parseRange(dataRange)
  }

  const ret = []
  for (let r = dataRange.s.r; r < dataRange.e.r; r++) {
    for (let c = dataRange.s.c; c < dataRange.e.c; c++) {
    }
  }
}
