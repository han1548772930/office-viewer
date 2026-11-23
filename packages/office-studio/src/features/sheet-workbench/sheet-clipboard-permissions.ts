import type { CellRange, WorkbookModel } from '../../domain/model.ts'
import { compactSheetClipboardRanges, sheetClipboardChangedSheetRanges } from './sheet-clipboard-changes.ts'
import { sheetClipboardWorkbookResourceChangedRanges } from './sheet-clipboard-workbook-resources.ts'
import { assertSheetRangesEditable } from './sheet-range-command.ts'

/** Validate the actual replacement against current permissions, including a recent-paste restoration. */
export function assertSheetClipboardWorkbookChangesEditable(
  current: WorkbookModel, candidate: WorkbookModel, userId: string,
): Record<string, CellRange[]> {
  if (current.id !== candidate.id || current.sheets.length !== candidate.sheets.length
    || current.sheets.some(sheet => !candidate.sheets.some(next => next.id === sheet.id))) {
    throw new Error('剪贴板候选与当前工作簿结构不一致。')
  }
  const changed = sheetClipboardWorkbookResourceChangedRanges(current, candidate)
  for (const before of current.sheets) {
    const after = candidate.sheets.find(sheet => sheet.id === before.id)!
    const ranges = compactSheetClipboardRanges([...(changed[before.id] ?? []), ...sheetClipboardChangedSheetRanges(before, after)])
    if (!ranges.length) { delete changed[before.id]; continue }
    assertSheetRangesEditable(before, ranges, userId)
    changed[before.id] = ranges
  }
  return changed
}
