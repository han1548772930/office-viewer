import type { CellRange, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import { GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH } from '../../rendering/spreadsheet/layout.ts'
import { createSheetSelectionAxisIndexes } from './sheet-selection.ts'
import { assertSheetCommandRanges } from './sheet-range-command.ts'

export interface SheetClipboardObjectBounds { left: number; top: number; width: number; height: number }

/** Objects and the Canvas use the same CSS-pixel axes, including filters, compact hides and collapsed outlines. */
export function sheetClipboardObjectAxes(workbook: WorkbookModel, sheet: SpreadsheetSheet) {
  const filtered = [...(sheet.autoFilter?.filteredRows ?? []), ...(workbook.tables ?? [])
    .filter(table => table.sheetId === sheet.id).flatMap(table => table.autoFilter?.filteredRows ?? [])]
  return createSheetSelectionAxisIndexes({ ...sheet, rowCount: 1_048_576, columnCount: 16_384 }, filtered)
}

export function sheetClipboardRangeBounds(workbook: WorkbookModel, sheet: SpreadsheetSheet, range: CellRange): SheetClipboardObjectBounds {
  assertSheetCommandRanges([range])
  const axes = sheetClipboardObjectAxes(workbook, sheet)
  return {
    left: GRID_HEADER_WIDTH + axes.columns.offset(range.startCol), top: GRID_HEADER_HEIGHT + axes.rows.offset(range.startRow),
    width: axes.columns.offset(range.endCol + 1) - axes.columns.offset(range.startCol),
    height: axes.rows.offset(range.endRow + 1) - axes.rows.offset(range.startRow),
  }
}

export function sheetClipboardBoundsContain(outer: SheetClipboardObjectBounds, inner: SheetClipboardObjectBounds): boolean {
  return inner.left >= outer.left && inner.top >= outer.top && inner.left + inner.width <= outer.left + outer.width
    && inner.top + inner.height <= outer.top + outer.height
}

export function sheetClipboardObjectRange(workbook: WorkbookModel, sheetId: string, bounds: SheetClipboardObjectBounds): CellRange {
  const sheet = workbook.sheets.find(item => item.id === sheetId)
  if (!sheet) throw new Error('剪贴板对象所在工作表已移除。')
  if (![bounds.left, bounds.top, bounds.width, bounds.height].every(Number.isFinite) || bounds.width < 0 || bounds.height < 0) {
    throw new RangeError('剪贴板对象位置或尺寸无效。')
  }
  const axes = sheetClipboardObjectAxes(workbook, sheet), left = Math.max(0, bounds.left - GRID_HEADER_WIDTH), top = Math.max(0, bounds.top - GRID_HEADER_HEIGHT)
  const right = Math.max(left, bounds.left + bounds.width - GRID_HEADER_WIDTH), bottom = Math.max(top, bounds.top + bounds.height - GRID_HEADER_HEIGHT)
  if (right > axes.columns.offset(16_384) || bottom > axes.rows.offset(1_048_576)) throw new RangeError('剪贴板对象超出工作表最大范围。')
  const interior = (end: number, start: number): number => end > start ? Math.max(start, end - Math.max(1e-7, Math.abs(end) * Number.EPSILON)) : start
  return { startRow: axes.rows.locate(top), endRow: axes.rows.locate(interior(bottom, top)),
    startCol: axes.columns.locate(left), endCol: axes.columns.locate(interior(right, left)) }
}

export function sheetClipboardBoundsDrawing(workbook: WorkbookModel, sheetId: string, bounds: SheetClipboardObjectBounds) {
  const sheet = workbook.sheets.find(item => item.id === sheetId)
  if (!sheet) throw new Error('剪贴板绘图目标工作表已移除。')
  sheetClipboardObjectRange(workbook, sheetId, bounds)
  const axes = sheetClipboardObjectAxes(workbook, sheet)
  const point = (left: number, top: number) => {
    const x = left - GRID_HEADER_WIDTH, y = top - GRID_HEADER_HEIGHT
    const column = axes.columns.locate(Math.max(0, x)), row = axes.rows.locate(Math.max(0, y))
    return { row, column, offsetX: x - axes.columns.offset(column), offsetY: y - axes.rows.offset(row) }
  }
  return { from: point(bounds.left, bounds.top), to: point(bounds.left + bounds.width, bounds.top + bounds.height) }
}
