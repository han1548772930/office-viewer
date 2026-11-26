import type { WorkbookModel } from '../../domain/model.ts'
import { toExcelWorkbookView } from '../../services/excelViewAdapter.ts'
import { createSheetSelectionAxisIndexes } from '../sheet-workbench/sheet-selection.ts'
import { SheetDrawingError } from '../drawing-groups/drawing-group-types.ts'

/** Reuse the Canvas view's extent and visible axes, including table filters and auto-height. */
export function sheetDrawingMoveExtent(workbook: WorkbookModel, sheetId: string): { width: number; height: number } {
  const view = toExcelWorkbookView(workbook).sheets.find(sheet => sheet.id === sheetId)
  if (!view) throw new SheetDrawingError('not-found', '绘图所在工作表已移除。')
  const axes = createSheetSelectionAxisIndexes(view)
  return { width: axes.columns.offset(axes.columns.count), height: axes.rows.offset(axes.rows.count) }
}
