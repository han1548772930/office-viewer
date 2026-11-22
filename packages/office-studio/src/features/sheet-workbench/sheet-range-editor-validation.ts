import type { SpreadsheetCell, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import { validateSheetDataValidationChanges } from '../sheet-validation/validation-engine.ts'

function contentIdentity(cell: SpreadsheetCell | undefined): string {
  return JSON.stringify([cell?.value ?? null, cell?.formula ?? null, cell?.richText ?? null])
}

/** Validate the complete candidate before any range-edit transaction is committed. */
export function assertSheetRangeChangedValuesValid(before: SpreadsheetSheet, after: WorkbookModel, sheetId: string): void {
  const sheet = after.sheets.find(candidate => candidate.id === sheetId)
  if (!sheet) throw new Error('找不到目标工作表。')
  if (!sheet.dataValidations?.length) return
  const positions = [...new Set([...Object.keys(before.cells), ...Object.keys(sheet.cells)])].flatMap(key => {
    const previous = before.cells[key], next = sheet.cells[key]
    if (contentIdentity(previous) === contentIdentity(next)) return []
    const cell = next ?? previous!
    return [{ row: cell.row, column: cell.col }]
  })
  const validation = validateSheetDataValidationChanges(after, sheetId, positions)
  if (validation.blocked) throw new Error(validation.issues.find(issue => issue.severity === 'stop')?.message ?? '内容未通过数据验证，数据没有修改。')
}
