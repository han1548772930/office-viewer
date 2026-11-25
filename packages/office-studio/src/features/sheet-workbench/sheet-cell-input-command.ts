import type { SpreadsheetCell, WorkbookModel } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { normalizeSheetCellValueType } from '../../domain/sheet-cell-value.ts'
import { applyExcelCellView, cellKey, parseCellKey, sheetModelCellViewAt } from '../../services/excelViewAdapter.ts'
import { resolveSheetCellStyle, setSheetCellStyleProperties } from './sheet-style-inheritance.ts'
import { canEditProtectedRanges, canEditWorksheetProtection, canViewWorksheetProtection } from '../sheet-protection/range-protection.ts'
import { validateSheetDataValidationChanges, type SheetValidationIssue } from '../sheet-validation/validation-engine.ts'
import { editSheetCellInlineText, editSheetCellText, sheetCellInputText } from './sheet-cell-edit.ts'
import { getSheetCellEditConflict, type SheetCellEditConflict, type SheetCellEditSession } from './sheet-cell-edit-session.ts'
import { editSheetCellRichText, sheetCellRichTextMatches } from './sheet-cell-rich-text-edit.ts'
import { projectSheetCellInputEditTrace, readSheetCellInputEditTrace, replaySheetCellInputEditTrace,
  type SheetCellInputEditTrace } from './sheet-cell-input-trace.ts'
import { isSheetTextNumberFormat } from './sheet-number-format.ts'
import { readSheetCellInlineFormats, replaySheetCellInlineFormats, type SheetCellInlineFormatEdit } from './sheet-cell-inline-format.ts'

export const COMMIT_SHEET_CELL_INPUT_COMMAND = 'sheet.command.commit-cell-input'
export interface SheetCellInputCommandInput {
  unitId: string
  sheetId: string
  row: number
  column: number
  input: string
  session?: SheetCellEditSession
  editTrace?: SheetCellInputEditTrace
  inlineFormats?: SheetCellInlineFormatEdit[]
  region?: string
  date1904?: boolean
  currentYear?: number
}
export type SheetCellInputOptions = Pick<SheetCellInputCommandInput, 'session' | 'editTrace' | 'inlineFormats' | 'region' | 'date1904' | 'currentYear'>
export interface SheetCellInputCommandResult {
  accepted: boolean
  changed: boolean
  issues: SheetValidationIssue[]
  reason?: 'conflict' | 'validation' | 'rich-text-source' | 'edit-trace' | 'inline-format'
  conflict?: SheetCellEditConflict
  message?: string
}
export interface SheetCellInputPlan {
  position: { row: number; column: number }
  cell?: SpreadsheetCell
  result: SheetCellInputCommandResult
}
export interface SheetCellInputDefaults { region: string; currentYear: number; userId: string }

function record(value: unknown, label: string): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError(`${label}必须是对象。`)
  return value as Record<string, unknown>
}
function identifier(value: unknown, label: string): string {
  if (typeof value !== 'string' || !value.trim()) throw new TypeError(`${label}无效。`)
  return value
}
function coordinate(value: unknown, limit: number): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 0 || value >= limit) throw new RangeError('单元格坐标超出 Office 范围。')
  return value
}
function year(value: unknown): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 1 || value > 9999) throw new RangeError('输入年份无效。')
  return value
}
function region(value: unknown): string {
  const source = identifier(value, '输入地区')
  try { new Intl.NumberFormat(source) } catch { throw new TypeError('输入地区无效。') }
  return source
}
function readSession(value: unknown): SheetCellEditSession {
  const input = record(value, '单元格编辑会话'), source = record(input.source, '编辑源内容')
  const documentId = input.documentId === undefined ? undefined : identifier(input.documentId, '编辑文档标识')
  const sessionId = input.sessionId === undefined ? undefined : identifier(input.sessionId, '编辑会话标识')
  const address = parseCellKey(identifier(input.key, '编辑地址'))
  if (!address) throw new TypeError('编辑地址必须是单个单元格。')
  coordinate(address.row, 1_048_576); coordinate(address.col, 16_384)
  if (typeof input.date1904 !== 'boolean' || input.mode !== undefined && input.mode !== 'edit' && input.mode !== 'overwrite') throw new TypeError('编辑会话设置无效。')
  const formula = source.formula, scalar = source.value, type = source.type
  const normalizedType = typeof type === 'string' ? normalizeSheetCellValueType(type) : undefined
  if (formula !== null && typeof formula !== 'string' || scalar !== null && typeof scalar !== 'string' && typeof scalar !== 'boolean' && typeof scalar !== 'number'
    || typeof scalar === 'number' && !Number.isFinite(scalar) || typeof source.text !== 'string' || typeof type !== 'string'
    || normalizedType === undefined || normalizedType !== type || typeof source.forcedText !== 'boolean' || typeof source.numberFormat !== 'string') {
    throw new TypeError('编辑源内容无效。')
  }
  let richText: ReadonlyArray<Record<string, unknown>> | null | undefined
  if (source.richText !== undefined) {
    if (source.richText === null) richText = null
    else {
      if (!Array.isArray(source.richText) || source.richText.some(run => !run || typeof run !== 'object'
        || Array.isArray(run) || typeof run.text !== 'string')) throw new TypeError('编辑源富文本无效。')
      richText = source.richText.length ? cloneModel(source.richText) : null
    }
  }
  return { documentId, sheetId: identifier(input.sheetId, '编辑工作表标识'),
    ...(sessionId === undefined ? {} : { sessionId }),
    key: cellKey(address.row, address.col), date1904: input.date1904, region: region(input.region), currentYear: year(input.currentYear),
    ...(input.mode === undefined ? {} : { mode: input.mode }),
    source: { formula, value: scalar, text: source.text, type: normalizedType, forcedText: source.forcedText, numberFormat: source.numberFormat,
      ...(richText === undefined ? {} : { richText }) } }
}

/** UI and Facade use one runtime boundary; an open draft owns its interpretation settings. */
export function sheetCellInputCommandInput(value: unknown): SheetCellInputCommandInput {
  const source = record(value, '单元格输入命令参数')
  if (typeof source.input !== 'string') throw new TypeError('单元格输入必须是文本。')
  const sheetId = identifier(source.sheetId ?? source.subUnitId, '工作表标识')
  if (source.sheetId !== undefined && source.sheetId !== sheetId || source.subUnitId !== undefined && source.subUnitId !== sheetId) throw new TypeError('工作表标识不一致。')
  if (source.column !== undefined && source.col !== undefined && source.column !== source.col) throw new TypeError('单元格列坐标不一致。')
  const session = source.session === undefined ? undefined : readSession(source.session)
  const editTrace = source.editTrace === undefined ? undefined : readSheetCellInputEditTrace(source.editTrace)
  if (editTrace && (!session?.sessionId || editTrace.sessionId !== session.sessionId)) throw new TypeError('输入范围轨迹必须绑定当前编辑会话。')
  if (editTrace && replaySheetCellInputEditTrace(editTrace) !== source.input) throw new TypeError('输入范围轨迹与最终输入不一致。')
  if (source.inlineFormats !== undefined && (!editTrace || !session)) throw new TypeError('局部文字格式必须绑定当前编辑会话和输入轨迹。')
  const inlineFormats = source.inlineFormats === undefined ? undefined : readSheetCellInlineFormats(source.inlineFormats, editTrace!)
  const parsed: SheetCellInputCommandInput = { unitId: identifier(source.unitId, '工作簿标识'), sheetId,
    row: coordinate(source.row, 1_048_576), column: coordinate(source.column ?? source.col, 16_384), input: source.input,
    ...(session ? { session } : {}), ...(editTrace ? { editTrace } : {}), ...(inlineFormats ? { inlineFormats } : {}) }
  if (source.region !== undefined) parsed.region = region(source.region)
  if (source.currentYear !== undefined) parsed.currentYear = year(source.currentYear)
  if (source.date1904 !== undefined) {
    if (typeof source.date1904 !== 'boolean') throw new TypeError('输入日期系统无效。')
    parsed.date1904 = source.date1904
  }
  if (session && (parsed.region !== undefined && parsed.region !== session.region || parsed.currentYear !== undefined && parsed.currentYear !== session.currentYear
    || parsed.date1904 !== undefined && parsed.date1904 !== session.date1904)) throw new TypeError('输入设置与编辑会话不一致。')
  return parsed
}

export function sheetCellInputConflictMessage(conflict: SheetCellEditConflict): string {
  return conflict === 'date-system' ? '编辑期间日期系统已更改，草稿已保留。请取消后重新编辑。'
    : conflict === 'number-format' ? '编辑期间数字格式已更改，草稿已保留。请取消后重新编辑。'
      : conflict === 'rich-text' ? '编辑期间文字运行、链接或扩展已更改，草稿已保留。请取消后重新编辑。'
      : '编辑目标的内容已更改，草稿已保留。请取消后重新编辑。'
}

/** Prepare the canonical merged-cell edit without changing the authoritative workbook. */
export function planSheetCellInput(workbook: WorkbookModel, value: unknown, defaults: SheetCellInputDefaults): SheetCellInputPlan {
  const input = sheetCellInputCommandInput(value)
  if (workbook.id !== undefined && workbook.id !== input.unitId) throw new Error('输入必须属于目标工作簿。')
  const sheet = workbook.sheets.find(item => item.id === input.sheetId)
  if (!sheet) throw new Error('找不到输入的目标工作表。')
  if (input.row >= (sheet.rowCount ?? 1_048_576) || input.column >= (sheet.columnCount ?? 16_384)) throw new RangeError('输入超出工作表范围。')
  const merge = sheet.mergedRanges.find(range => input.row >= range.startRow && input.row <= range.endRow && input.column >= range.startCol && input.column <= range.endCol)
  const position = { row: merge?.startRow ?? input.row, column: merge?.startCol ?? input.column }
  const previous = sheet.cells[`${position.row}:${position.column}`], view = sheetModelCellViewAt(workbook, sheet, position.row, position.column)
  const date1904 = workbook.date1904 ?? false
  const interpretation = { region: input.session?.region ?? input.region ?? region(defaults.region),
    currentYear: input.session?.currentYear ?? input.currentYear ?? year(defaults.currentYear) }
  if (previous?.richText?.length && !previous.formula && input.session?.mode !== 'overwrite' && input.session
    && !sheetCellRichTextMatches(previous.richText, String(previous.value ?? ''))) {
    return { position, result: { accepted: false, changed: false, issues: [], reason: 'rich-text-source', message: '富文本运行与源内容不一致，草稿已保留。' } }
  }
  const conflict = input.session ? getSheetCellEditConflict(input.session, { documentId: workbook.id, sheetId: sheet.id,
    key: cellKey(position.row, position.column), date1904, cell: view }) : input.date1904 !== undefined && input.date1904 !== date1904 ? 'date-system' : undefined
  if (conflict) return { position, result: { accepted: false, changed: false, issues: [], reason: 'conflict', conflict, message: sheetCellInputConflictMessage(conflict) } }
  if (input.editTrace && input.editTrace.source !== sheetCellInputText(view, date1904, interpretation.region)) {
    return { position, result: { accepted: false, changed: false, issues: [], reason: 'edit-trace', message: '输入范围轨迹不属于当前源内容，草稿已保留。' } }
  }
  if (!canViewWorksheetProtection(sheet, defaults.userId)) throw new Error('当前用户没有查看此工作表的权限。')
  const overwritesRichText = previous?.richText?.length && (!input.session || input.session.mode === 'overwrite')
  const overwritesImage = previous?.image && (!input.session || input.session.mode === 'overwrite')
  if (!overwritesRichText && !overwritesImage && !input.editTrace?.operations.length && !input.inlineFormats?.length
    && sheetCellInputText(view, date1904, interpretation.region) === input.input) return { position, result: { accepted: true, changed: false, issues: [] } }
  if (!canEditWorksheetProtection(sheet, defaults.userId)) throw new Error('此工作表受保护，当前用户没有编辑权限。')
  const range = merge ?? { startRow: position.row, endRow: position.row, startCol: position.column, endCol: position.column }
  if (!canEditProtectedRanges(sheet, [range], defaults.userId)) throw new Error('目标单元格受保护，当前用户没有编辑权限。')
  const editor = input.inlineFormats?.length || previous?.richText?.length && input.session?.mode !== 'overwrite'
    ? editSheetCellInlineText : editSheetCellText
  const edited = editor(view, input.input, date1904, interpretation.region, interpretation.currentYear)
  if (input.inlineFormats?.length && input.editTrace && input.session) {
    const replay = replaySheetCellInlineFormats(view, input.editTrace, input.inlineFormats,
      { numberFormat: input.session.source.numberFormat, overwrite: input.session.mode === 'overwrite' })
    if (!replay || edited.formula || typeof edited.rawValue !== 'string' && edited.rawValue !== null
      || replay.cell.rawValue !== (edited.rawValue ?? '')) {
      return { position, result: { accepted: false, changed: false, issues: [], reason: 'inline-format',
        message: '局部文字格式无法应用到当前输入。请保持文本类型，或取消局部格式后提交；草稿已保留。' } }
    }
    edited.richText = replay.cell.richText ?? (previous?.richText?.length ? null : undefined)
  } else if (previous?.richText?.length) {
    if (input.session?.mode !== 'overwrite' && input.session && typeof edited.rawValue === 'string' && !edited.formula) {
      const operations = input.editTrace ? projectSheetCellInputEditTrace(input.editTrace, String(previous.value ?? ''), edited.rawValue,
        isSheetTextNumberFormat(input.session.source.numberFormat)) : undefined
      if (input.editTrace && !operations) return { position, result: { accepted: false, changed: false, issues: [], reason: 'edit-trace', message: '输入范围无法映射到当前富文本，草稿已保留。' } }
      const runs = editSheetCellRichText(previous.richText, String(previous.value ?? ''), edited.rawValue, operations)
      if (!runs) return { position, result: { accepted: false, changed: false, issues: [], reason: 'rich-text-source', message: '富文本运行与源内容不一致，草稿已保留。' } }
      edited.richText = runs
    } else edited.richText = null
  }
  const cell = cloneModel(applyExcelCellView({ row: position.row, col: position.column }, edited, previous,
    { date1904, previousDate1904: date1904, numberFormatLocale: interpretation.region }))
  if (edited.rawNumberFormat !== undefined && edited.rawNumberFormat !== view.rawNumberFormat) {
    setSheetCellStyleProperties(cell, { numberFormat: edited.rawNumberFormat })
  }
  cell.style = resolveSheetCellStyle(sheet, position.row, position.column, cell, workbook)
  const candidate = { ...workbook, sheets: workbook.sheets.map(item => item.id === sheet.id
    ? { ...sheet, cells: { ...sheet.cells, [cell.id]: cell } } : item) }
  const formattingOnly = input.inlineFormats?.length && previous && Object.is(previous.value, cell.value)
    && previous.formula === cell.formula && normalizeSheetCellValueType(previous.dataType) === normalizeSheetCellValueType(cell.dataType)
  const validation = formattingOnly ? { blocked: false, issues: [] } : validateSheetDataValidationChanges(candidate, sheet.id, [position])
  if (validation.blocked) {
    const issue = validation.issues.find(item => item.severity === 'stop')!
    return { position, result: { accepted: false, changed: false, reason: 'validation', issues: cloneModel(validation.issues), message: `${issue.title}：${issue.message}` } }
  }
  return { position, cell, result: { accepted: true, changed: JSON.stringify(previous) !== JSON.stringify(cell), issues: cloneModel(validation.issues) } }
}
