import type { CellRange, SheetDataValidationType, SheetDataValidationErrorStyle, WorkbookModel } from '../../domain/model.ts'
import { cellKey } from '../../services/excelViewAdapter.ts'
import { dataValidationsAt, validateSheetDataValidationChanges, type SheetValidationPosition } from './validation-engine.ts'
import { sheetValidationSourceReader } from './validation-value-read.ts'

export interface SheetValidationCellNotice { ruleId: string; title?: string; message?: string }
export interface SheetValidationCellIssue extends SheetValidationCellNotice { type: SheetDataValidationType; severity: SheetDataValidationErrorStyle }
export interface SheetValidationCellFeedback extends SheetValidationPosition {
  key: string; prompt?: SheetValidationCellNotice; issues: readonly SheetValidationCellIssue[]
}
export type SheetValidationFeedbackReader = (positions: readonly SheetValidationPosition[]) => ReadonlyMap<string, SheetValidationCellFeedback>

/** Feedback is derived for the requested cells only and never carries cell values, source ranges or whole rules into UI. */
export function sheetValidationCellFeedback(workbook: WorkbookModel, sheetId: string, positions: readonly SheetValidationPosition[],
  userId?: string, region = 'zh-CN'): Map<string, SheetValidationCellFeedback> {
  const sheet = workbook.sheets.find(candidate => candidate.id === sheetId), result = new Map<string, SheetValidationCellFeedback>()
  if (!sheet) return result
  const canRead = sheetValidationSourceReader(userId)
  for (const position of positions) {
    if (!Number.isSafeInteger(position.row) || !Number.isSafeInteger(position.column) || position.row < 0 || position.column < 0
      || position.row >= (sheet.rowCount ?? 1_048_576) || position.column >= (sheet.columnCount ?? 16_384)) continue
    const merge = sheet.mergedRanges.find(range => position.row >= range.startRow && position.row <= range.endRow
      && position.column >= range.startCol && position.column <= range.endCol)
    const row = merge?.startRow ?? position.row, column = merge?.startCol ?? position.column, key = cellKey(row, column)
    const range: CellRange = merge ?? { startRow: row, endRow: row, startCol: column, endCol: column }
    if (result.has(key) || canRead && !canRead(sheet, range)) continue
    const rules = dataValidationsAt(sheet, row, column)
    if (!rules.length) continue
    const promptRule = rules.find(rule => rule.showInputMessage && (rule.promptTitle || rule.prompt))
    result.set(key, { key, row, column, issues: [], ...(promptRule ? { prompt: { ruleId: promptRule.id,
      ...(promptRule.promptTitle ? { title: promptRule.promptTitle } : {}), ...(promptRule.prompt ? { message: promptRule.prompt } : {}) } } : {}) })
  }
  const validation = validateSheetDataValidationChanges(workbook, sheet.id, [...result.values()], canRead, region)
  const grouped = new Map<string, SheetValidationCellIssue[]>()
  for (const issue of validation.issues) {
    const key = cellKey(issue.row, issue.column), rule = dataValidationsAt(sheet, issue.row, issue.column).find(rule => rule.id === issue.ruleId)
    const notices = grouped.get(key) ?? []
    notices.push({ ruleId: issue.ruleId, type: issue.type, severity: issue.severity,
      ...(rule?.showErrorMessage !== false && rule?.errorTitle ? { title: rule.errorTitle } : {}),
      ...(rule?.showErrorMessage !== false && rule?.error ? { message: rule.error } : {}) })
    grouped.set(key, notices)
  }
  for (const [key, issues] of grouped) result.get(key)!.issues = issues
  return result
}
