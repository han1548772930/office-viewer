import type { CellRange, CellScalar, SheetFilterCriteria, SheetSortCriterion, SpreadsheetSheet, WorkbookModel } from '../../domain/model'
import { cloneModel } from '../../domain/serialization'
import { parseSheetAddress } from '../sheet-workbench/sheet-workbench'
import { readCellValue, setAutoFilter, setFilterCriteria } from './sort-filter-engine'

export const CONFIGURE_SHEET_FILTER_COMMAND = 'sheet.command.configure-filter'
export interface ConfigureSheetFilterInput {
  unitId: string
  sheetId: string
  range: CellRange | null
  columns: Record<number, SheetFilterCriteria>
  expectedRevision?: number
}
export interface DataPanelField { column: number; label: string }
export interface DataPanelValue { key: string; label: string; value: CellScalar }
export interface PanelSortKey { column: number; direction: 'ascending' | 'descending' }

function columnName(column: number): string {
  let value = column + 1, result = ''
  while (value > 0) { value--; result = String.fromCharCode(65 + value % 26) + result; value = Math.floor(value / 26) }
  return result
}
export function dataPanelAddress(range: CellRange): string {
  return `${columnName(range.startCol)}${range.startRow + 1}:${columnName(range.endCol)}${range.endRow + 1}`
}
export function parseDataPanelRange(sheet: SpreadsheetSheet, address: string): CellRange {
  if (!/^\$?[A-Z]+\$?[1-9]\d*(?::\$?[A-Z]+\$?[1-9]\d*)?$/i.test(address.trim())) throw new RangeError('请输入当前工作表的数据区域，例如 A1:C20。')
  const parsed = parseSheetAddress(address, { rowCount: sheet.rowCount ?? 1_048_576, columnCount: sheet.columnCount ?? 16_384 })
  if (!parsed || parsed.range.endRow <= parsed.range.startRow) throw new RangeError('请选择至少两行的数据区域，筛选区域需包含标题行。')
  return { startRow: parsed.range.startRow, endRow: parsed.range.endRow, startCol: parsed.range.startColumn, endCol: parsed.range.endColumn }
}
export function readDataPanelFields(sheet: SpreadsheetSheet, range: CellRange, hasHeader = true): DataPanelField[] {
  return Array.from({ length: range.endCol - range.startCol + 1 }, (_, offset) => {
    const column = range.startCol + offset, name = columnName(column)
    const value = hasHeader ? sheet.cells[`${range.startRow}:${column}`]?.value : undefined
    return { column, label: value == null || value === '' ? `列 ${name}` : `${name} · ${String(value)}` }
  })
}
export function dataPanelValueKey(value: CellScalar): string { return JSON.stringify([typeof value, value]) }

/** Read sparse cells without materializing the full rectangle or coercing typed values. */
export function readDataPanelValues(sheet: SpreadsheetSheet, range: CellRange, column: number): DataPanelValue[] {
  if (!Number.isInteger(column) || column < range.startCol || column > range.endCol) throw new RangeError('筛选字段必须位于数据区域内。')
  const values = new Map<string, DataPanelValue>()
  const rows = new Set<number>()
  const add = (value: CellScalar) => {
    const key = dataPanelValueKey(value)
    if (!values.has(key)) values.set(key, { key, value, label: value === null ? '（空白）' : value === '' ? '（空文本）' : typeof value === 'number' ? `${value}（数值）` : typeof value === 'boolean' ? (value ? 'TRUE（逻辑值）' : 'FALSE（逻辑值）') : value })
  }
  for (const cell of Object.values(sheet.cells)) {
    if (cell.col !== column || cell.row <= range.startRow || cell.row > range.endRow) continue
    rows.add(cell.row)
    const value = readCellValue(cell)
    if (value === null || typeof value === 'string' || typeof value === 'boolean' || typeof value === 'number' && Number.isFinite(value)) add(value)
  }
  if (rows.size < range.endRow - range.startRow) add(null)
  return [...values.values()]
}
export function parseDataPanelScalar(value: string, type: 'text' | 'number' | 'boolean'): CellScalar {
  if (type === 'text') return value
  if (type === 'boolean') {
    if (/^true$/i.test(value.trim())) return true
    if (/^false$/i.test(value.trim())) return false
    throw new TypeError('逻辑值请输入 TRUE 或 FALSE。')
  }
  if (!value.trim() || !Number.isFinite(Number(value))) throw new TypeError('请输入有效数值。日期条件请使用工作簿的日期序号。')
  return Number(value)
}
export function prepareDataPanelSort(range: CellRange, keys: readonly PanelSortKey[]): SheetSortCriterion[] {
  if (!keys.length) throw new RangeError('请至少添加一个排序条件。')
  const seen = new Set<number>()
  return keys.map((key) => {
    if (!Number.isInteger(key.column) || key.column < range.startCol || key.column > range.endCol || seen.has(key.column)) throw new RangeError('排序字段必须位于数据区域内，且不能重复。')
    if (key.direction !== 'ascending' && key.direction !== 'descending') throw new TypeError('排序方向无效。')
    seen.add(key.column)
    return { ...key }
  })
}

/** Preflight on a detached filter plan, then publish it as one model mutation. */
export function configureFilterFromPanel(workbook: WorkbookModel, input: ConfigureSheetFilterInput): void {
  if (input.unitId !== workbook.id) throw new Error('筛选目标工作簿不匹配。')
  const sheet = workbook.sheets.find((candidate) => candidate.id === input.sheetId)
  if (!sheet) throw new Error('筛选目标工作表已关闭。')
  const plan: SpreadsheetSheet = { ...sheet, autoFilter: sheet.autoFilter ? cloneModel(sheet.autoFilter) : undefined }
  setAutoFilter(plan, input.range)
  if (input.range !== null) {
    plan.autoFilter!.columns = {}
    for (const [column, criteria] of Object.entries(input.columns)) setFilterCriteria(plan, Number(column), criteria)
  }
  sheet.autoFilter = plan.autoFilter
}
