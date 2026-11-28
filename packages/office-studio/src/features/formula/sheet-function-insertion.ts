import type { CellRange, SpreadsheetCell, SpreadsheetSheet } from '../../domain/model.ts'
import { DEFAULT_SHEET_COLUMN_COUNT, DEFAULT_SHEET_ROW_COUNT } from '../../domain/model.ts'
import { normalizeSheetCellValueType } from '../../domain/sheet-cell-value.ts'
import { assertSheetCommandRanges, sheetRangeContainsCell } from '../sheet-workbench/sheet-range-command.ts'
import { translateFormulaReferences } from './formula-references.ts'

export interface SheetFunctionSelection { range: CellRange; primary?: { row: number; column: number } }
export interface SheetFunctionInsertion { range: CellRange; primary: { row: number; column: number }; formula: string }
export interface SheetFunctionWrite { row: number; column: number; formula: string }
export interface SheetFunctionInsertionPlan {
  list: SheetFunctionInsertion[]
  listOfRangeHasNumber: SheetFunctionInsertion[]
  edit?: SheetFunctionInsertion
}
export const MAX_FUNCTION_INSERTION_CELLS = 1_000_000

export function sheetInsertionFunctionName(value: unknown): string {
  if (typeof value !== 'string' || !/^[A-Za-z_][A-Za-z0-9_.]{0,127}$/.test(value)) throw new TypeError('函数名称无效。')
  return value.toUpperCase()
}

/** Match v1.6.0 isNumberCell: rich-document text has priority over scalar data. */
export function sheetFunctionNumberCell(cell: SpreadsheetCell | undefined): boolean {
  if (!cell) return false
  const publicBody = cell.metadata?.publicBody
  if (publicBody !== undefined && publicBody !== null) {
    if (typeof publicBody !== 'object' || Array.isArray(publicBody) || !('body' in publicBody)) return false
    const body = publicBody.body
    if (!body || typeof body !== 'object' || !('dataStream' in body) || typeof body.dataStream !== 'string') return false
    const text = body.dataStream.replace(/\r\n$/, '').trim()
    return text !== '' && !Number.isNaN(Number(text))
  }
  if (cell.richText?.length) {
    const text = cell.richText.map(run => typeof run.text === 'string' ? run.text : '').join('').replace(/\r\n$/, '').trim()
    return text !== '' && !Number.isNaN(Number(text))
  }
  const type = normalizeSheetCellValueType(cell.formula ? cell.cachedDataType : cell.dataType)
  return type === 'number' || type !== 'text' && type !== 'boolean' && type !== 'error' && typeof cell.value === 'number'
}

function columnName(column: number): string {
  let result = ''
  for (let value = column + 1; value > 0; value = Math.floor((value - 1) / 26)) result = String.fromCharCode(65 + (value - 1) % 26) + result
  return result
}
export function sheetFunctionRangeReference(range: CellRange): string {
  const first = `${columnName(range.startCol)}${range.startRow + 1}`, last = `${columnName(range.endCol)}${range.endRow + 1}`
  return first === last ? first : `${first}:${last}`
}

/** Sparse indexes avoid scanning every blank between a primary and the sheet edge. */
export function planSheetFunctionInsertion(sheet: SpreadsheetSheet, selections: readonly SheetFunctionSelection[], value: string): SheetFunctionInsertionPlan {
  const name = sheetInsertionFunctionName(value), rows = sheet.rowCount ?? DEFAULT_SHEET_ROW_COUNT, columns = sheet.columnCount ?? DEFAULT_SHEET_COLUMN_COUNT
  const numeric = Object.values(sheet.cells).filter(sheetFunctionNumberCell)
  const byColumn = new Map<number, number[]>(), byRow = new Map<number, number[]>()
  for (const cell of numeric) {
    const row = byRow.get(cell.row) ?? []; row.push(cell.col); byRow.set(cell.row, row)
    const column = byColumn.get(cell.col) ?? []; column.push(cell.row); byColumn.set(cell.col, column)
  }
  for (const positions of [...byColumn.values(), ...byRow.values()]) positions.sort((left, right) => left - right)
  const start = (positions: number[] | undefined, endpoint: number): number => {
    if (!positions?.length) return endpoint
    let index = positions.length - 1
    while (index >= 0 && positions[index]! >= endpoint) index--
    if (index < 0) return endpoint
    let first = positions[index]!
    while (index > 0 && positions[index - 1] === first - 1) first = positions[--index]!
    return first
  }
  const reference = (row: number, column: number): string => {
    const firstRow = start(byColumn.get(column), row)
    if (firstRow !== row) return sheetFunctionRangeReference({ startRow: firstRow, endRow: row - 1, startCol: column, endCol: column })
    const firstColumn = start(byRow.get(row), column)
    return firstColumn !== column ? sheetFunctionRangeReference({ startRow: row, endRow: row, startCol: firstColumn, endCol: column - 1 }) : ''
  }
  const plan: SheetFunctionInsertionPlan = { list: [], listOfRangeHasNumber: [] }
  for (const selection of selections) {
    const range = selection.range
    assertSheetCommandRanges([range])
    if (range.endRow >= rows || range.endCol >= columns) throw new RangeError('函数插入选区超出工作表范围。')
    const primary = selection.primary && sheetRangeContainsCell(range, selection.primary.row, selection.primary.column)
      ? { ...selection.primary } : { row: range.startRow, column: range.startCol }
    const hasNumber = numeric.some(cell => sheetRangeContainsCell(range, cell.row, cell.col))
    const single = range.startRow === range.endRow && range.startCol === range.endCol
    const multipleAxes = range.startRow !== range.endRow && range.startCol !== range.endCol
    if (selections.length === 1 && (single || multipleAxes && !hasNumber)) {
      plan.edit = { range: { ...range }, primary, formula: `=${name}(${reference(primary.row, primary.column)}` }; break
    }
    if (!hasNumber) {
      const ref = reference(primary.row, primary.column)
      if (!ref) { plan.edit = { range: { ...range }, primary, formula: `=${name}(` }; break }
      plan.list.push({ range: { ...range }, primary, formula: `=${name}(${ref})` }); continue
    }
    if (range.startRow === range.endRow) {
      let column = range.endCol
      while (column < columns - 1 && sheet.cells[`${range.startRow}:${column}`] !== undefined) column++
      const endCol = column === range.endCol ? range.endCol - 1 : range.endCol
      plan.listOfRangeHasNumber.push({ range: { ...range }, primary: { row: range.startRow, column },
        formula: `=${name}(${sheetFunctionRangeReference({ ...range, endCol })})` })
    } else {
      let destination = range.endRow
      for (let column = range.startCol; column <= range.endCol; column++) {
        let row = range.endRow
        while (row < rows - 1 && sheet.cells[`${row}:${column}`] !== undefined) row++
        destination = Math.max(destination, row)
      }
      const endRow = destination === range.endRow ? range.endRow - 1 : range.endRow
      for (let column = range.startCol; column <= range.endCol; column++) plan.listOfRangeHasNumber.push({
        range: { ...range }, primary: { row: destination, column },
        formula: `=${name}(${sheetFunctionRangeReference({ ...range, endRow, startCol: column, endCol: column })})`,
      })
    }
  }
  return plan
}

/** Expand shared formulas once, preserving relative/absolute references and last-write order. */
export function expandSheetFunctionInsertions(list: readonly SheetFunctionInsertion[], numericList: readonly SheetFunctionInsertion[],
  maximum = MAX_FUNCTION_INSERTION_CELLS): SheetFunctionWrite[] {
  const writes = new Map<string, SheetFunctionWrite>(); let planned = 0
  const put = (row: number, column: number, formula: string): void => {
    assertSheetCommandRanges([{ startRow: row, endRow: row, startCol: column, endCol: column }])
    writes.set(`${row}:${column}`, { row, column, formula })
  }
  for (const item of list) {
    assertSheetCommandRanges([item.range])
    planned += (item.range.endRow - item.range.startRow + 1) * (item.range.endCol - item.range.startCol + 1)
    if (planned > maximum) throw new RangeError('函数插入超过当前任务的单元格预算。')
    put(item.primary.row, item.primary.column, item.formula)
    for (let row = item.range.startRow; row <= item.range.endRow; row++) for (let column = item.range.startCol; column <= item.range.endCol; column++) {
      if (row !== item.primary.row || column !== item.primary.column) put(row, column,
        translateFormulaReferences(item.formula, row - item.primary.row, column - item.primary.column))
    }
  }
  if (planned + numericList.length > maximum) throw new RangeError('函数插入超过当前任务的单元格预算。')
  for (const item of numericList) put(item.primary.row, item.primary.column, item.formula)
  return [...writes.values()]
}
