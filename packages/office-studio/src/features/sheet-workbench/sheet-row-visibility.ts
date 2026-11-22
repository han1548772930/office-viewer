import { effectiveSheetRowSizes, sheetColumnWidth, sheetRowHeight, sheetAxisSpanAt, SparseAxisIndex, type SheetAxisDimensions } from '../../domain/sheet-axis.ts'

export interface SheetRowVisibility extends SheetAxisDimensions {
  rowCount: number
  hiddenRows?: readonly number[]
  filteredRows?: readonly number[]
  outlineHiddenRows?: readonly number[]
}

/** Keep manual and filter reasons distinct in the model; combine them only for viewing. */
export function hiddenSheetRows(sheet: SheetRowVisibility): Set<number> {
  const hidden = new Set([...(sheet.hiddenRows ?? []), ...(sheet.filteredRows ?? []), ...(sheet.outlineHiddenRows ?? [])])
  for (const key of new Set([...Object.keys(sheet.rowHeights ?? {}), ...Object.keys(sheet.rowAutoHeights ?? {})])) if (sheetRowHeight(sheet, Number(key)) <= 0) hidden.add(Number(key))
  return new Set([...hidden].filter(row => Number.isInteger(row) && row >= 0 && row < sheet.rowCount))
}

/** Constant zero-size spans stay compact; callers query them instead of expanding them into a Set. */
export function isSheetRowVisible(sheet: SheetRowVisibility, row: number, hidden: ReadonlySet<number> = hiddenSheetRows(sheet)): boolean {
  return Number.isInteger(row) && row >= 0 && row < sheet.rowCount && !hidden.has(row) && !sheetAxisSpanAt(sheet.rowHiddenSpans, row)?.value && sheetRowHeight(sheet, row) > 0
}

export function findVisibleSheetRow(sheet: SheetRowVisibility, start: number, direction: -1 | 1, fallback = true): number | undefined {
  if (!sheet.rowCount || !Number.isFinite(start)) return undefined
  const sizes = effectiveSheetRowSizes(sheet)
  const axis = new SparseAxisIndex(sheet.rowCount, sheet.defaultRowHeight ?? 26, sizes.sizes, [...hiddenSheetRows(sheet)], sizes.spans, sheet.rowHiddenSpans)
  const first = Math.max(0, Math.min(sheet.rowCount - 1, Math.trunc(start)))
  return axis.nextVisible(first, direction) ?? (fallback ? axis.nextVisible(first - direction, direction === 1 ? -1 : 1) : undefined)
}

export interface SheetColumnVisibility extends SheetAxisDimensions { columnCount: number; hiddenColumns?: readonly number[]; outlineHiddenColumns?: readonly number[] }
export function hiddenSheetColumns(sheet: SheetColumnVisibility): Set<number> {
  return new Set([...(sheet.hiddenColumns ?? []), ...(sheet.outlineHiddenColumns ?? [])].filter(column => Number.isInteger(column) && column >= 0 && column < sheet.columnCount))
}
export function isSheetColumnVisible(sheet: SheetColumnVisibility, column: number, hidden: ReadonlySet<number> = hiddenSheetColumns(sheet)): boolean {
  return Number.isInteger(column) && column >= 0 && column < sheet.columnCount && !hidden.has(column) && !sheetAxisSpanAt(sheet.columnHiddenSpans, column)?.value && sheetColumnWidth(sheet, column) > 0
}
export function findVisibleSheetColumn(sheet: SheetColumnVisibility, start: number, direction: -1 | 1, fallback = true): number | undefined {
  if (!sheet.columnCount || !Number.isFinite(start)) return undefined
  const sizes = sheet.defaultColumnWidth === undefined && !sheetAxisSpanAt(sheet.columnWidthSpans, 0) ? { 0: 152, ...sheet.columnWidths } : sheet.columnWidths
  const axis = new SparseAxisIndex(sheet.columnCount, sheet.defaultColumnWidth ?? 112, sizes, [...hiddenSheetColumns(sheet)], sheet.columnWidthSpans, sheet.columnHiddenSpans)
  const first = Math.max(0, Math.min(sheet.columnCount - 1, Math.trunc(start)))
  return axis.nextVisible(first, direction) ?? (fallback ? axis.nextVisible(first - direction, direction === 1 ? -1 : 1) : undefined)
}
