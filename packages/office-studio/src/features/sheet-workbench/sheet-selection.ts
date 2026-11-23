import type { CellRange, SheetOutlineState } from '../../domain/model.ts'
import { effectiveSheetRowSizes, sheetAxisSpanAt, SparseAxisIndex } from '../../domain/sheet-axis.ts'
import { assertSheetCommandRanges, sheetRangeContainsCell } from './sheet-range-command.ts'
import { hiddenSheetRows, hiddenSheetColumns, type SheetRowVisibility, type SheetColumnVisibility } from './sheet-row-visibility.ts'

export interface SheetSelectionRange extends CellRange { rangeType?: 0 | 1 | 2 | 3 }
export interface SheetSelectionCell { row: number; column: number }
export interface SheetSelectionState {
  ranges: SheetSelectionRange[]
  activeRangeIndex: number
  primary?: SheetSelectionCell
}
export interface SheetSelectionOptions { activeRangeIndex?: number; primary?: SheetSelectionCell }

/** Keep every selection in its own coordinates; overlapping ranges are legal and commands deduplicate their targets. */
export function normalizeSheetSelection(ranges: readonly SheetSelectionRange[], extent: { rowCount?: number; columnCount?: number },
  options: SheetSelectionOptions = {}): SheetSelectionState {
  if (!Array.isArray(ranges)) throw new TypeError('选区必须是范围数组。')
  const rows = extent.rowCount ?? 1_048_576, columns = extent.columnCount ?? 16_384
  if (!Number.isSafeInteger(rows) || rows < 1 || rows > 1_048_576 || !Number.isSafeInteger(columns) || columns < 1 || columns > 16_384) {
    throw new RangeError('选区的工作表行列总数无效。')
  }
  assertSheetCommandRanges(ranges)
  const next = ranges.map(range => {
    if (range.endRow >= rows || range.endCol >= columns) throw new RangeError('选区超出工作表。')
    if (range.rangeType !== undefined && ![0, 1, 2, 3].includes(range.rangeType)) throw new TypeError('选区类型无效。')
    return { startRow: range.startRow, endRow: range.endRow, startCol: range.startCol, endCol: range.endCol,
      ...(range.rangeType === undefined ? {} : { rangeType: range.rangeType }) }
  })
  if (!next.length) {
    if (options.primary !== undefined || options.activeRangeIndex !== undefined && options.activeRangeIndex !== -1) throw new RangeError('空选区不能指定活动单元格。')
    return { ranges: [], activeRangeIndex: -1 }
  }
  let activeRangeIndex = options.activeRangeIndex ?? next.length - 1
  if (!Number.isSafeInteger(activeRangeIndex) || activeRangeIndex < 0 || activeRangeIndex >= next.length) throw new RangeError('活动选区索引无效。')
  const selected = next[activeRangeIndex]!, primary = options.primary ?? { row: selected.startRow, column: selected.startCol }
  if (!Number.isSafeInteger(primary.row) || !Number.isSafeInteger(primary.column) || primary.row < 0 || primary.column < 0
    || primary.row >= rows || primary.column >= columns) throw new RangeError('活动单元格超出工作表。')
  if (!sheetRangeContainsCell(selected, primary.row, primary.column)) {
    if (options.activeRangeIndex !== undefined) throw new RangeError('活动单元格必须在活动选区内。')
    activeRangeIndex = -1
    for (let index = next.length - 1; index >= 0; index--) {
      if (sheetRangeContainsCell(next[index]!, primary.row, primary.column)) { activeRangeIndex = index; break }
    }
    if (activeRangeIndex < 0) throw new RangeError('活动单元格必须在选区内。')
  }
  return { ranges: next, activeRangeIndex, primary: { row: primary.row, column: primary.column } }
}

export function cloneSheetSelection(state: SheetSelectionState): SheetSelectionState {
  return { ranges: state.ranges.map(range => ({ ...range })), activeRangeIndex: state.activeRangeIndex,
    ...(state.primary ? { primary: { ...state.primary } } : {}) }
}

/** Structural snapshots can change sheet bounds before their saved selection is restored. */
export function reconcileSheetSelection(state: SheetSelectionState, extent: { rowCount?: number; columnCount?: number }): SheetSelectionState {
  const checked = normalizeSheetSelection(state.ranges, {}, state)
  const rows = extent.rowCount ?? 1_048_576, columns = extent.columnCount ?? 16_384
  if (!checked.ranges.length) return normalizeSheetSelection([], extent)
  const ranges = checked.ranges.map(range => ({ ...range,
    startRow: range.rangeType === 2 || range.rangeType === 3 ? 0 : Math.min(rows - 1, range.startRow),
    endRow: range.rangeType === 2 || range.rangeType === 3 ? rows - 1 : Math.min(rows - 1, range.endRow),
    startCol: range.rangeType === 1 || range.rangeType === 3 ? 0 : Math.min(columns - 1, range.startCol),
    endCol: range.rangeType === 1 || range.rangeType === 3 ? columns - 1 : Math.min(columns - 1, range.endCol),
  }))
  const active = ranges[checked.activeRangeIndex]!, primary = checked.primary
    ? { row: Math.max(active.startRow, Math.min(active.endRow, checked.primary.row)), column: Math.max(active.startCol, Math.min(active.endCol, checked.primary.column)) }
    : undefined
  return normalizeSheetSelection(ranges, extent, { activeRangeIndex: checked.activeRangeIndex, primary })
}

export interface SheetSelectionSheet extends SheetRowVisibility, SheetColumnVisibility { outline?: SheetOutlineState }

export function createSheetSelectionAxisIndexes(sheet: SheetSelectionSheet, filteredRows: readonly number[] = []) {
  const collapsed = sheet.outline?.groups.filter(group => group.collapsed) ?? []
  const rowSizes = effectiveSheetRowSizes(sheet)
  const rows = new SparseAxisIndex(sheet.rowCount, sheet.defaultRowHeight ?? 26, rowSizes.sizes,
    [...hiddenSheetRows({ ...sheet, filteredRows: [...(sheet.filteredRows ?? []), ...filteredRows] })], rowSizes.spans,
    [...(sheet.rowHiddenSpans ?? []), ...collapsed.filter(group => group.axis === 'row').map(group => ({ start: group.start, end: group.end, value: true }))])
  const columnSizes = sheet.defaultColumnWidth === undefined && !sheetAxisSpanAt(sheet.columnWidthSpans, 0)
    ? { 0: 152, ...sheet.columnWidths } : sheet.columnWidths
  const columns = new SparseAxisIndex(sheet.columnCount, sheet.defaultColumnWidth ?? 112, columnSizes,
    [...hiddenSheetColumns(sheet)], sheet.columnWidthSpans,
    [...(sheet.columnHiddenSpans ?? []), ...collapsed.filter(group => group.axis === 'column').map(group => ({ start: group.start, end: group.end, value: true }))])
  return { rows, columns }
}

/** Keep API and Canvas on the same visible primary while retaining the other selected rectangles. */
export function resolveVisibleSheetSelection(state: SheetSelectionState, sheet: SheetSelectionSheet,
  filteredRows: readonly number[] = []): SheetSelectionState {
  const next = reconcileSheetSelection(state, sheet)
  if (!next.primary) return next
  const { rows, columns } = createSheetSelectionAxisIndexes(sheet, filteredRows)
  const point = next.primary
  if (rows.size(point.row) > 0 && columns.size(point.column) > 0) return next
  const within = (axis: SparseAxisIndex, start: number, end: number, preferred: number): number | undefined => {
    const at = Math.max(start, Math.min(end, preferred)), forward = axis.nextVisible(at, 1)
    if (forward !== undefined && forward <= end) return forward
    const backward = axis.nextVisible(at, -1)
    return backward !== undefined && backward >= start ? backward : undefined
  }
  const candidates = [next.activeRangeIndex, ...next.ranges.map((_range, index) => index).reverse().filter(index => index !== next.activeRangeIndex)]
  for (const index of candidates) {
    const range = next.ranges[index]!, row = within(rows, range.startRow, range.endRow, point.row)
    const column = within(columns, range.startCol, range.endCol, point.column)
    if (row !== undefined && column !== undefined) return normalizeSheetSelection(next.ranges, sheet, { activeRangeIndex: index, primary: { row, column } })
  }
  const row = rows.nextVisible(point.row, 1) ?? rows.nextVisible(point.row, -1)
  const column = columns.nextVisible(point.column, 1) ?? columns.nextVisible(point.column, -1)
  if (row === undefined || column === undefined) return normalizeSheetSelection([], sheet)
  const active = next.ranges[next.activeRangeIndex]!
  next.ranges[next.activeRangeIndex] = active.rangeType === 1
    ? { startRow: row, endRow: row, startCol: 0, endCol: sheet.columnCount - 1, rangeType: 1 }
    : active.rangeType === 2
      ? { startRow: 0, endRow: sheet.rowCount - 1, startCol: column, endCol: column, rangeType: 2 }
      : { startRow: row, endRow: row, startCol: column, endCol: column, ...(active.rangeType === undefined ? {} : { rangeType: 0 }) }
  return normalizeSheetSelection(next.ranges, sheet, { activeRangeIndex: next.activeRangeIndex, primary: { row, column } })
}
