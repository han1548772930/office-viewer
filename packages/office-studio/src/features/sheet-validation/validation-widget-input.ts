import type { CellRange, CellScalar, SheetDataValidationRule, SpreadsheetCell, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { applyExcelCellView, cellKey, sheetModelCellViewAt, type ExcelCellView } from '../../services/excelViewAdapter.ts'
import { canEditProtectedRanges, canEditWorksheetProtection, canViewProtectedRanges, canViewWorksheetProtection } from '../sheet-protection/range-protection.ts'
import { getSheetCellEditConflict, resolveSheetDate1904 } from '../sheet-workbench/sheet-cell-edit-session.ts'
import { sheetCellInputCommandInput, sheetCellInputConflictMessage, type SheetCellInputCommandInput,
  type SheetCellInputDefaults, type SheetCellInputPlan } from '../sheet-workbench/sheet-cell-input-command.ts'
import { resolveSheetCellStyle, setSheetCellStyleProperties } from '../sheet-workbench/sheet-style-inheritance.ts'
import { dataValidationsAt, validateSheetDataValidationChanges } from './validation-engine.ts'
import { sheetValidationSourceReader } from './validation-value-read.ts'

export interface SheetValidationWidgetTargetInput extends Omit<SheetCellInputCommandInput, 'input' | 'editTrace' | 'inlineFormats'> {
  ruleId: string
  host?: 'api' | 'workbench'
  expectedRevision?: number
  expectedHistoryGeneration?: number
  expectedUserId?: string
  expectedRuleSignature?: string
}
export interface SheetValidationWidgetTarget {
  workbook: WorkbookModel
  sheet: SpreadsheetSheet
  rule: SheetDataValidationRule
  position: SheetCellInputPlan['position']
  range: CellRange
  view: ExcelCellView
  date1904: boolean
  region: string
  currentYear: number
  userId?: string
}

export function validationWidgetRecord(value: unknown, label: string): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError(`${label}必须是对象。`)
  return value as Record<string, unknown>
}
function text(value: unknown, label: string): string {
  if (typeof value !== 'string' || !value.trim()) throw new TypeError(`${label}无效。`)
  return value
}

/** Typed widgets reuse the canonical target/session boundary without parsing a label or serial as cell text. */
export function sheetValidationWidgetTargetInput(source: Record<string, unknown>): SheetValidationWidgetTargetInput {
  if (source.editTrace !== undefined || source.inlineFormats !== undefined) throw new TypeError('选择控件不能携带文字输入范围或局部格式。')
  const { input: _input, ...target } = sheetCellInputCommandInput({ unitId: source.unitId, sheetId: source.sheetId,
    subUnitId: source.subUnitId, row: source.row, column: source.column, col: source.col, input: '',
    session: source.session, region: source.region, currentYear: source.currentYear, date1904: source.date1904 })
  const parsed: SheetValidationWidgetTargetInput = { ...target, ruleId: text(source.ruleId, '数据验证标识') }
  const host = source.host, userId = source.expectedUserId, ruleSignature = source.expectedRuleSignature
  if (host !== undefined) {
    if (host !== 'api' && host !== 'workbench') throw new TypeError('选择控件宿主无效。')
    parsed.host = host
  }
  for (const key of ['expectedRevision', 'expectedHistoryGeneration'] as const) {
    const value = source[key]
    if (value !== undefined) {
      if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < (key === 'expectedRevision' ? 0 : 1)) throw new TypeError('选择控件内容或历史版本无效。')
      parsed[key] = value
    }
  }
  if (userId !== undefined) parsed.expectedUserId = text(userId, '选择控件用户')
  if (ruleSignature !== undefined) parsed.expectedRuleSignature = text(ruleSignature, '选择控件规则版本')
  return parsed
}

/** Typed validation controls edit the merged main cell and apply the same permissions and edit-session conflicts. */
export function resolveSheetValidationWidgetTarget(workbook: WorkbookModel, input: SheetValidationWidgetTargetInput,
  defaults: SheetCellInputDefaults): SheetValidationWidgetTarget | SheetCellInputPlan {
  if (workbook.id !== undefined && workbook.id !== input.unitId) throw new Error('选择控件必须属于目标工作簿。')
  const sheet = workbook.sheets.find(candidate => candidate.id === input.sheetId)
  if (!sheet) throw new Error('找不到选择控件的目标工作表。')
  if (input.row >= (sheet.rowCount ?? 1_048_576) || input.column >= (sheet.columnCount ?? 16_384)) throw new RangeError('选择控件超出工作表范围。')
  const merge = sheet.mergedRanges.find(range => input.row >= range.startRow && input.row <= range.endRow
    && input.column >= range.startCol && input.column <= range.endCol)
  const position = { row: merge?.startRow ?? input.row, column: merge?.startCol ?? input.column }
  const range = merge ?? { startRow: position.row, endRow: position.row, startCol: position.column, endCol: position.column }
  if (!canViewWorksheetProtection(sheet, defaults.userId) || !canViewProtectedRanges(sheet, [range], defaults.userId)) throw new Error('当前用户没有查看目标单元格的权限。')
  if (!canEditWorksheetProtection(sheet, defaults.userId) || !canEditProtectedRanges(sheet, [range], defaults.userId)) throw new Error('目标单元格受保护，当前用户没有编辑权限。')
  const date1904 = resolveSheetDate1904(undefined, workbook), view = sheetModelCellViewAt(workbook, sheet, position.row, position.column)
  const conflict = input.session ? getSheetCellEditConflict(input.session, { documentId: workbook.id, sheetId: sheet.id,
    key: cellKey(position.row, position.column), date1904, cell: view })
    : input.date1904 !== undefined && input.date1904 !== date1904 ? 'date-system' : undefined
  if (conflict) return { position, result: { accepted: false, changed: false, issues: [], reason: 'conflict', conflict,
    message: sheetCellInputConflictMessage(conflict) } }
  const rule = dataValidationsAt(sheet, position.row, position.column).find(candidate => candidate.id === input.ruleId)
  if (!rule || input.expectedRuleSignature !== undefined && JSON.stringify(rule) !== input.expectedRuleSignature) throw new Error('数据验证规则已更改，请重新打开选择控件。')
  return { workbook, sheet, rule, position, range, date1904, view,
    region: input.session?.region ?? input.region ?? defaults.region, currentYear: input.session?.currentYear ?? input.currentYear ?? defaults.currentYear,
    userId: defaults.userId }
}

/** Rule initialization and gestures share scalar mapping, formatting inheritance and cache invalidation. */
export function createSheetValidationScalarCell(target: Pick<SheetValidationWidgetTarget, 'workbook' | 'sheet' | 'position' | 'view' | 'date1904' | 'region'>,
  value: CellScalar, type: 'number' | 'text' | 'boolean', numberFormat?: string): SpreadsheetCell {
  const { workbook, sheet, position, view, date1904 } = target
  const previous = sheet.cells[`${position.row}:${position.column}`]
  const edited: ExcelCellView = { ...cloneModel(view), value: String(value ?? ''), rawValue: value, dataType: type,
    formula: undefined, cachedValue: undefined, displayValue: undefined, cachedDataType: null, richText: null, image: null,
    ...(numberFormat === undefined ? {} : { rawNumberFormat: numberFormat }) }
  const cell = cloneModel(applyExcelCellView({ row: position.row, col: position.column }, edited, previous,
    { date1904, previousDate1904: date1904, numberFormatLocale: target.region }))
  if (numberFormat !== undefined && numberFormat !== view.rawNumberFormat) setSheetCellStyleProperties(cell, { numberFormat })
  cell.style = resolveSheetCellStyle(sheet, position.row, position.column, cell, workbook)
  return cell
}

/** Canonical cell mapping preserves unrelated styles/metadata and clears replaced formulas, rich text and value caches. */
export function planSheetValidationWidgetValue(target: SheetValidationWidgetTarget, value: CellScalar,
  type: 'number' | 'text' | 'boolean', numberFormat?: string): SheetCellInputPlan {
  const { workbook, sheet, position } = target, previous = sheet.cells[`${position.row}:${position.column}`]
  const cell = createSheetValidationScalarCell(target, value, type, numberFormat)
  const candidate = { ...workbook, sheets: workbook.sheets.map(item => item.id === sheet.id
    ? { ...sheet, cells: { ...sheet.cells, [cell.id]: cell } } : item) }
  const validation = validateSheetDataValidationChanges(candidate, sheet.id, [position], sheetValidationSourceReader(target.userId), target.region)
  if (validation.blocked) {
    const issue = validation.issues.find(item => item.severity === 'stop')!
    return { position, result: { accepted: false, changed: false, reason: 'validation', issues: cloneModel(validation.issues),
      message: `${issue.title}：${issue.message}` } }
  }
  return { position, cell, result: { accepted: true, changed: JSON.stringify(previous) !== JSON.stringify(cell), issues: cloneModel(validation.issues) } }
}
