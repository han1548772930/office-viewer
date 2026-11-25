import type { FontStyle, WorkbookModel } from '../../domain/model.ts'
import { cellKey, toExcelCellView } from '../../services/excelViewAdapter.ts'
import { canEditProtectedRanges, canEditWorksheetProtection, canViewWorksheetProtection } from '../sheet-protection/range-protection.ts'
import { getSheetCellEditConflict } from './sheet-cell-edit-session.ts'
import { editSheetCellInlineText, sheetCellInputText } from './sheet-cell-edit.ts'
import { sheetCellInputCommandInput, sheetCellInputConflictMessage,
  type SheetCellInputCommandInput, type SheetCellInputDefaults, type SheetCellInputPlan } from './sheet-cell-input-command.ts'
import { fontFormatActionForCommand } from './sheet-font-format-command.ts'
import { isSheetCellTextBoundary, type SheetCellInputEditTrace } from './sheet-cell-input-trace.ts'
import { readSheetCellInlineFormats, replaySheetCellInlineFormats, sheetInlineFontChangeForAction,
  type SheetCellInlineFormatEdit, type SheetCellInlineFormatReplay } from './sheet-cell-inline-format.ts'

export const FORMAT_SHEET_CELL_INPUT_OPERATION = 'sheet.operation.format-cell-input'
export const FORMAT_SHEET_CELL_RICH_TEXT_COMMAND = 'sheet.command.format-cell-rich-text'
export type SheetCellInlineStylePatch = { [K in keyof FontStyle]?: FontStyle[K] | null }
export interface SheetCellInlineFontFormatInput extends SheetCellInputCommandInput {
  commandId: string
  params?: unknown
  selection: { start: number; end: number }
}
export interface SheetCellInlineFontFormatResult {
  accepted: boolean
  edit?: SheetCellInlineFormatEdit
  formats: SheetCellInlineFormatEdit[]
  replay?: SheetCellInlineFormatReplay
  message?: string
}
export interface SheetCellInlineStyleCommandInput {
  unitId: string
  sheetId: string
  row: number
  column: number
  start: number
  end: number
  style: SheetCellInlineStylePatch
}

function object(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError('局部文字格式参数必须是对象。')
  return value as Record<string, unknown>
}

function target(workbook: WorkbookModel, input: SheetCellInputCommandInput, userId: string) {
  if (workbook.id !== undefined && workbook.id !== input.unitId) throw new Error('文字格式必须属于目标工作簿。')
  const sheet = workbook.sheets.find(item => item.id === input.sheetId)
  if (!sheet || input.row >= (sheet.rowCount ?? 1_048_576) || input.column >= (sheet.columnCount ?? 16_384)) throw new RangeError('文字格式目标不在当前工作表中。')
  const range = sheet.mergedRanges.find(range => input.row >= range.startRow && input.row <= range.endRow
    && input.column >= range.startCol && input.column <= range.endCol)
    ?? { startRow: input.row, endRow: input.row, startCol: input.column, endCol: input.column }
  if (!canViewWorksheetProtection(sheet, userId) || !canEditWorksheetProtection(sheet, userId)
    || !canEditProtectedRanges(sheet, [range], userId)) throw new Error('当前用户没有修改目标文字格式的权限。')
  const cell = sheet.cells[`${range.startRow}:${range.startCol}`]
  return { sheet, range, modelCell: cell, cell: cell ? toExcelCellView(cell) : undefined }
}

/** Same font command IDs and value rules apply while editing; the operation never publishes document content. */
export function planSheetCellInlineFontFormat(workbook: WorkbookModel, value: unknown,
  defaults: SheetCellInputDefaults): SheetCellInlineFontFormatResult {
  const source = object(value), input = sheetCellInputCommandInput(source)
  if (!input.session || !input.editTrace) throw new TypeError('编辑中文字格式必须绑定输入会话。')
  const { range, cell } = target(workbook, input, defaults.userId)
  const conflict = getSheetCellEditConflict(input.session, { documentId: workbook.id, sheetId: input.sheetId,
    key: cellKey(range.startRow, range.startCol), date1904: workbook.date1904 ?? false, cell })
  const formats = input.inlineFormats ?? []
  if (conflict) return { accepted: false, formats, message: sheetCellInputConflictMessage(conflict) }
  if (input.editTrace.source !== sheetCellInputText(cell, input.session.date1904, input.session.region)) {
    return { accepted: false, formats, message: '文字格式轨迹不属于当前源内容，草稿已保留。' }
  }
  const selection = object(source.selection)
  if (typeof selection.start !== 'number' || typeof selection.end !== 'number' || selection.start > selection.end
    || !isSheetCellTextBoundary(input.input, selection.start) || !isSheetCellTextBoundary(input.input, selection.end)) throw new RangeError('局部文字选区无效。')
  if (typeof source.commandId !== 'string') throw new TypeError('文字格式命令标识无效。')
  const action = fontFormatActionForCommand(source.commandId, source.params)
  const edited = editSheetCellInlineText(cell, input.input, input.session.date1904, input.session.region, input.session.currentYear)
  if (!action || action.kind === 'color' && action.field !== 'color' || edited.formula
    || typeof edited.rawValue !== 'string' && edited.rawValue !== null) return { accepted: false, formats, message: '此输入不支持局部文字格式。' }
  const replay = replaySheetCellInlineFormats(cell, input.editTrace, formats,
    { numberFormat: input.session.source.numberFormat, overwrite: input.session.mode === 'overwrite' })
  if (!replay || replay.cell.rawValue !== (edited.rawValue ?? '')) return { accepted: false, formats, message: '局部文字格式与输入范围不一致，草稿已保留。' }
  const prefix = input.input.length - String(edited.rawValue ?? '').length
  const rawSelection = { start: Math.max(0, selection.start - prefix), end: Math.max(0, selection.end - prefix) }
  const change = sheetInlineFontChangeForAction(action, replay.cell, rawSelection, replay.typing)
  if (!change) return { accepted: false, formats }
  const edit: SheetCellInlineFormatEdit = { at: input.editTrace.operations.length, input: input.input,
    start: selection.start, end: selection.end, ...change }
  const next = readSheetCellInlineFormats([...formats, edit], input.editTrace)
  const candidate = replaySheetCellInlineFormats(cell, input.editTrace, next,
    { numberFormat: input.session.source.numberFormat, overwrite: input.session.mode === 'overwrite' })
  return candidate ? { accepted: true, edit, formats: next, replay: candidate }
    : { accepted: false, formats, message: '无法应用局部文字格式，草稿已保留。' }
}

/** API ranges use raw-cell offsets; draft operations use the input bridge's offsets. */
export function planSheetCellInlineStyle(workbook: WorkbookModel, value: unknown, defaults: SheetCellInputDefaults): SheetCellInputPlan {
  const source = object(value)
  const location = sheetCellInputCommandInput({ ...source, input: '' })
  const { range, cell, modelCell } = target(workbook, location, defaults.userId)
  const text = String(cell?.rawValue ?? cell?.value ?? '')
  if (cell?.formula || typeof cell?.rawValue === 'number' || typeof cell?.rawValue === 'boolean') throw new TypeError('局部文字格式仅用于文本单元格。')
  if (typeof source.start !== 'number' || typeof source.end !== 'number' || source.start > source.end
    || !isSheetCellTextBoundary(text, source.start) || !isSheetCellTextBoundary(text, source.end)) throw new RangeError('局部文字选区超出单元格内容。')
  const style = object(source.style)
  if (Object.keys(style).some(key => !['family', 'size', 'bold', 'italic', 'underline', 'strike', 'color'].includes(key))) throw new TypeError('局部文字格式包含未知字体字段。')
  const trace: SheetCellInputEditTrace = { version: 1, sessionId: 'api-cell-inline-format', source: text, operations: [] }
  const formats = readSheetCellInlineFormats(Object.entries(style).filter(([, value]) => value !== undefined).map(([field, value]) => ({
    at: 0, input: text, start: source.start, end: source.end, field, value,
  })), trace)
  // This API changes styles only. Its raw offsets never reinterpret numbers, apostrophes or stored text types.
  const replay = replaySheetCellInlineFormats(cell, trace, formats, { numberFormat: '@' })
  if (!replay) throw new TypeError('文字格式与源富文本范围不一致。')
  const position = { row: range.startRow, column: range.startCol }
  if (!modelCell) return { position, result: { accepted: true, changed: false, issues: [] } }
  const candidate = { ...modelCell, richText: replay.cell.richText ?? undefined }
  return { position, cell: candidate, result: { accepted: true,
    changed: JSON.stringify(modelCell) !== JSON.stringify(candidate), issues: [] } }
}
