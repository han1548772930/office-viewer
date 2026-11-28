import type { CellRange, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import { normalizeSheetCommandRange, type SheetCommandRange, type SheetRangeCommandInput } from '../sheet-workbench/sheet-range-command.ts'
import { expandSheetContinuousRange } from '../sheet-workbench/sheet-continuous-range.ts'
import { parseSheetAddress } from '../sheet-workbench/sheet-workbench.ts'
import { createSheetTable, isSheetTableName } from './sheet-tables.ts'

export const SHEET_TABLE_COMMANDS = { openSelector: 'sheet.operation.open-table-selector', add: 'sheet.command.add-table' } as const

export interface SheetTableCommandContext {
  expectedRevision?: number
  expectedHistoryGeneration?: number
  expectedUserId?: string
  host?: 'workbench'
}
export interface AddSheetTableCommandInput extends SheetTableCommandContext {
  unitId?: string
  sheetId?: string
  subUnitId?: string
  range: SheetCommandRange | string
  id?: string
  name?: string
}
export type OpenSheetTableSelectorInput = SheetRangeCommandInput & SheetTableCommandContext
export interface SheetTableSelectorRequest {
  requestId: string
  unitId: string
  sheetId: string
  range: CellRange
  expectedRevision: number
  expectedHistoryGeneration: number
  expectedUserId: string
}

export function assertSheetTableCommandContext(input: unknown): asserts input is SheetTableCommandContext & { unitId?: string; sheetId?: string; subUnitId?: string } {
  if (!input || typeof input !== 'object' || Array.isArray(input)) throw new TypeError('表格命令参数无效。')
  for (const key of ['unitId', 'sheetId', 'subUnitId', 'expectedUserId'] as const) {
    const id = key in input ? input[key] : undefined
    if (id !== undefined && (typeof id !== 'string' || !id.trim())) throw new TypeError('表格命令身份无效。')
  }
  if ('sheetId' in input && 'subUnitId' in input && input.sheetId !== undefined && input.subUnitId !== undefined && input.sheetId !== input.subUnitId) throw new TypeError('表格目标工作表不一致。')
  if ('expectedRevision' in input && input.expectedRevision !== undefined && (typeof input.expectedRevision !== 'number' || !Number.isSafeInteger(input.expectedRevision) || input.expectedRevision < 0)) throw new TypeError('表格工作簿版本无效。')
  if ('expectedHistoryGeneration' in input && input.expectedHistoryGeneration !== undefined && (typeof input.expectedHistoryGeneration !== 'number' || !Number.isSafeInteger(input.expectedHistoryGeneration) || input.expectedHistoryGeneration <= 0)) throw new TypeError('表格历史版本无效。')
  if ('host' in input && input.host !== undefined && input.host !== 'workbench') throw new TypeError('表格工作台宿主无效。')
}

export function sheetTableSelectorRange(sheet: SpreadsheetSheet, range?: CellRange): CellRange {
  const selected = range ?? { startRow: 0, endRow: 0, startCol: 0, endCol: 0 }
  return selected.startRow === selected.endRow && selected.startCol === selected.endCol
    ? expandSheetContinuousRange(sheet, selected) : { ...selected }
}

export function sheetTableCreationRange(sheet: SpreadsheetSheet, range: SheetCommandRange | string): CellRange {
  if (typeof range !== 'string') return normalizeSheetCommandRange(range)
  const parsed = parseSheetAddress(range, { rowCount: sheet.rowCount ?? 1_048_576, columnCount: sheet.columnCount ?? 16_384 })
  if (!parsed) throw new RangeError('请输入当前工作表内的一个有效范围。')
  return { startRow: parsed.range.startRow, endRow: parsed.range.endRow, startCol: parsed.range.startColumn, endCol: parsed.range.endColumn }
}

export function createSheetTableCommandCandidate(workbook: WorkbookModel, sheetId: string, input: AddSheetTableCommandInput,
  labels: { tablePrefix: string; columnPrefix: string }): WorkbookModel {
  const sheet = workbook.sheets.find(item => item.id === sheetId)
  if (!sheet) throw new RangeError('找不到表格的目标工作表。')
  if (input.id !== undefined && (typeof input.id !== 'string' || !input.id.trim())) throw new TypeError('表格 ID 无效。')
  if (input.name !== undefined && typeof input.name !== 'string') throw new TypeError('表格名称无效。')
  if (input.id !== undefined && workbook.tables?.some(table => table.id === input.id)) throw new RangeError('表格 ID 已存在。')
  const range = sheetTableCreationRange(sheet, input.range)
  if (range.endRow >= (sheet.rowCount ?? 1_048_576) || range.endCol >= (sheet.columnCount ?? 16_384)) throw new RangeError('表格范围超出当前工作表。')
  const names = new Set((workbook.tables ?? []).map(table => table.name.toLocaleLowerCase()))
  for (const key of ['definedNames', 'localDefinedNames']) {
    const entries = workbook.metadata?.[key]
    if (!Array.isArray(entries)) continue
    for (const entry of entries) {
      if (entry && typeof entry === 'object' && 'name' in entry && typeof entry.name === 'string') names.add(entry.name.toLocaleLowerCase())
    }
  }
  const requested = input.name?.trim()
  const validName = requested && requested.length <= 255 && isSheetTableName(requested) && !names.has(requested.toLocaleLowerCase())
    && !parseSheetAddress(requested, { rowCount: 1_048_576, columnCount: 16_384 }) && !/^R\d+C\d+$/i.test(requested)
  const prefix = labels.tablePrefix
  const firstIndex = (workbook.tables?.length ?? 0) + 1
  const index = [...names].reduce((maximum, name) => {
    if (!name.startsWith(prefix.toLocaleLowerCase())) return maximum
    const suffix = Number(name.slice(prefix.length))
    return Number.isSafeInteger(suffix) && suffix >= maximum ? suffix + 1 : maximum
  }, firstIndex)
  return createSheetTable(workbook, { sheetId, range, id: input.id, name: validName ? requested : `${prefix}${index}`, columnPrefix: labels.columnPrefix })
}
