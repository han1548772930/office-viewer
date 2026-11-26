import { canonicalSheetAxisSpans, effectiveSheetRowSizes, sheetColumnWidth, sheetRowHeight, sheetAxisSpanAt, SparseAxisIndex, type SheetAxisDimensions, type SheetAxisSpan } from '../../domain/sheet-axis.ts'
import type { SheetOutlineState } from '../../domain/model.ts'

export interface SheetOutlineVisibility {
  outline?: SheetOutlineState
  outlineHiddenRows?: readonly number[]
  outlineHiddenColumns?: readonly number[]
  outlineHiddenRowSpans?: readonly SheetAxisSpan<boolean>[]
  outlineHiddenColumnSpans?: readonly SheetAxisSpan<boolean>[]
}

/** Outline visibility is independent of manual hide/show commands. */
export function sheetOutlineHiddenSpans(sheet: SheetOutlineVisibility, axis: 'row' | 'column'): SheetAxisSpan<boolean>[] {
  const derived = axis === 'row' ? sheet.outlineHiddenRowSpans : sheet.outlineHiddenColumnSpans
  return canonicalSheetAxisSpans([
    ...canonicalSheetAxisSpans(derived ?? []).filter(span => span.value),
    ...(sheet.outline?.groups ?? []).filter(group => group.axis === axis && group.collapsed).map(group => ({ start: group.start, end: group.end, value: true })),
  ])
}

/** Union independent visibility reasons after resolving each reason's own overrides. */
export function sheetAxisViewHiddenSpans(sheet: SheetAxisDimensions & SheetOutlineVisibility, axis: 'row' | 'column'): SheetAxisSpan<boolean>[] {
  const manual = axis === 'row' ? sheet.rowHiddenSpans : sheet.columnHiddenSpans
  return canonicalSheetAxisSpans([
    ...canonicalSheetAxisSpans(manual ?? []).filter(span => span.value),
    ...sheetOutlineHiddenSpans(sheet, axis),
  ])
}

export interface SheetRowVisibility extends SheetAxisDimensions, SheetOutlineVisibility {
  rowCount: number
  hiddenRows?: readonly number[]
  filteredRows?: readonly number[]
}

/** Keep manual and filter reasons distinct in the model; combine them only for viewing. */
export function hiddenSheetRows(sheet: SheetRowVisibility): Set<number> {
  const hidden = new Set([...(sheet.hiddenRows ?? []), ...(sheet.filteredRows ?? []), ...(sheet.outlineHiddenRows ?? [])])
  for (const key of new Set([...Object.keys(sheet.rowHeights ?? {}), ...Object.keys(sheet.rowAutoHeights ?? {})])) if (sheetRowHeight(sheet, Number(key)) <= 0) hidden.add(Number(key))
  return new Set([...hidden].filter(row => Number.isInteger(row) && row >= 0 && row < sheet.rowCount))
}

/** Constant zero-size spans stay compact; callers query them instead of expanding them into a Set. */
export function isSheetRowVisible(sheet: SheetRowVisibility, row: number, hidden: ReadonlySet<number> = hiddenSheetRows(sheet)): boolean {
  return Number.isInteger(row) && row >= 0 && row < sheet.rowCount && !hidden.has(row) && !sheetAxisSpanAt(sheet.rowHiddenSpans, row)?.value
    && !sheetAxisSpanAt(sheet.outlineHiddenRowSpans, row)?.value
    && !sheet.outline?.groups.some(group => group.axis === 'row' && group.collapsed && group.start <= row && row <= group.end) && sheetRowHeight(sheet, row) > 0
}

export function findVisibleSheetRow(sheet: SheetRowVisibility, start: number, direction: -1 | 1, fallback = true): number | undefined {
  if (!sheet.rowCount || !Number.isFinite(start)) return undefined
  const sizes = effectiveSheetRowSizes(sheet)
  const axis = new SparseAxisIndex(sheet.rowCount, sheet.defaultRowHeight ?? 26, sizes.sizes, [...hiddenSheetRows(sheet)], sizes.spans, sheetAxisViewHiddenSpans(sheet, 'row'))
  const first = Math.max(0, Math.min(sheet.rowCount - 1, Math.trunc(start)))
  return axis.nextVisible(first, direction) ?? (fallback ? axis.nextVisible(first - direction, direction === 1 ? -1 : 1) : undefined)
}

export interface SheetColumnVisibility extends SheetAxisDimensions, SheetOutlineVisibility { columnCount: number; hiddenColumns?: readonly number[] }
export function hiddenSheetColumns(sheet: SheetColumnVisibility): Set<number> {
  return new Set([...(sheet.hiddenColumns ?? []), ...(sheet.outlineHiddenColumns ?? [])].filter(column => Number.isInteger(column) && column >= 0 && column < sheet.columnCount))
}
export function isSheetColumnVisible(sheet: SheetColumnVisibility, column: number, hidden: ReadonlySet<number> = hiddenSheetColumns(sheet)): boolean {
  return Number.isInteger(column) && column >= 0 && column < sheet.columnCount && !hidden.has(column) && !sheetAxisSpanAt(sheet.columnHiddenSpans, column)?.value
    && !sheetAxisSpanAt(sheet.outlineHiddenColumnSpans, column)?.value
    && !sheet.outline?.groups.some(group => group.axis === 'column' && group.collapsed && group.start <= column && column <= group.end) && sheetColumnWidth(sheet, column) > 0
}
export function findVisibleSheetColumn(sheet: SheetColumnVisibility, start: number, direction: -1 | 1, fallback = true): number | undefined {
  if (!sheet.columnCount || !Number.isFinite(start)) return undefined
  const sizes = sheet.defaultColumnWidth === undefined && !sheetAxisSpanAt(sheet.columnWidthSpans, 0) ? { 0: 152, ...sheet.columnWidths } : sheet.columnWidths
  const axis = new SparseAxisIndex(sheet.columnCount, sheet.defaultColumnWidth ?? 112, sizes, [...hiddenSheetColumns(sheet)], sheet.columnWidthSpans, sheetAxisViewHiddenSpans(sheet, 'column'))
  const first = Math.max(0, Math.min(sheet.columnCount - 1, Math.trunc(start)))
  return axis.nextVisible(first, direction) ?? (fallback ? axis.nextVisible(first - direction, direction === 1 ? -1 : 1) : undefined)
}
