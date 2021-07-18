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

function computeRange(sheet) {

}

function cellContent(cell) {
  return cell.w || cell.z || cell.v
}

function sheetColHeader(sheet, range) {

}

function sheetRowHeader(sheet, range) { }

function mergedCellRange(merges, c, r) { }

function sheetHeader(sheet, range) {
  if (!range) {
    range = computeRange(sheet)
  } else if (range) {
    range = parseRange(range)
  }
  const { s: { r: START_ROW, c: START_COL }, e: { r: END_ROW, c: END_COL } } = range
  const merges = sheet[MERGES]
  let r = START_ROW, c = START_COL
  let cell
  if (cell = sheet[formatCellAddress(c, r)] && cellContent(cell)) {
    // Pure col or row header
    return
  }

  const { s: s0, e: e0 } = mergedCellRange(merges, c, r)
  return {
    col: sheetColHeader(sheet, { s: { c: Math.min(s0.c + 1, START_COL), r: s0.r }, e: { c: END_COL, r: e0.r } }),
    row: sheetRowHeader(sheet, { s: { c: s0.c, r: Math.min(s0.r + 1, START_ROW) }, e: { c: e0.c, r: END_ROW } }),
  }
}
