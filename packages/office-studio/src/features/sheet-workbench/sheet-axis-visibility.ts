import type { CellRange, SheetOutlineState, SpreadsheetSheet } from '../../domain/model.ts'
import { assignSheetAxisSpan, canonicalSheetAxisSpans, effectiveSheetRowSizes, sheetAxisSpanAt, SparseAxisIndex,
  type SheetAxisDimensions, type SheetAxisSpan } from '../../domain/sheet-axis.ts'
import { mergeSheetAxisIntervals, type SheetAxis, type SheetAxisInterval } from '../../domain/sheet-axis-structure.ts'
import type { SheetAxisDimensionChanges } from '../../domain/sheet-axis-transform.ts'
import { assertSheetCommandRanges, assertSheetRangesEditable, sheetRangeCommandInputRanges, type SheetRangeCommandInput } from './sheet-range-command.ts'
import { sheetAxisDimensionTargetRanges } from './sheet-axis-dimensions.ts'

export const SHEET_AXIS_VISIBILITY_COMMANDS = {
  hideRows: 'sheet.command.set-rows-hidden', hideColumns: 'sheet.command.set-col-hidden',
  showRows: 'sheet.command.set-specific-rows-visible', showColumns: 'sheet.command.set-col-visible-on-cols',
  hideRowsConfirm: 'sheet.command.hide-row-confirm', hideColumnsConfirm: 'sheet.command.hide-col-confirm',
  showSelectedRows: 'sheet.command.set-selected-rows-visible', showSelectedColumns: 'sheet.command.set-selected-cols-visible',
} as const
export const SHEET_AXIS_VISIBILITY_COMMAND_IDS = [...Object.values(SHEET_AXIS_VISIBILITY_COMMANDS),
  'sheet.command.hide-row', 'sheet.command.hide-column', 'sheet.command.show-hidden-rows', 'sheet.command.show-hidden-columns'] as const
export interface SheetAxisVisibilityCommandInput extends SheetRangeCommandInput { index?: number; count?: number; all?: boolean; expectedRevision?: number }
export interface SheetAxisVisibilityAction { axis: SheetAxis; visible: boolean; selectedOnly: boolean }
export interface SheetAxisVisibilityState extends SheetAxisDimensions {
  rowCount?: number; columnCount?: number; hiddenRows?: readonly number[]; hiddenColumns?: readonly number[]
  filteredRows?: readonly number[]; outlineHiddenRows?: readonly number[]; outlineHiddenColumns?: readonly number[]
  autoFilter?: { filteredRows: readonly number[] }; outline?: SheetOutlineState
}
export interface SheetAxisVisibilityChanges extends SheetAxisDimensionChanges { hiddenRows?: number[]; hiddenColumns?: number[] }
export interface SheetAxisVisibilityPlan { changed: boolean; changes: SheetAxisVisibilityChanges; ranges: CellRange[]; selections: CellRange[] }
export class SheetAxisVisibilityError extends Error {
  readonly code = 'SHEET_AXIS_HIDE_ALL'
  readonly axis: SheetAxis
  constructor(axis: SheetAxis) { super(axis === 'row' ? '不能隐藏全部可见行，请至少保留一行。' : '不能隐藏全部可见列，请至少保留一列。'); this.name = 'SheetAxisVisibilityError'; this.axis = axis }
}

export function sheetAxisVisibilityActionForCommand(id: string): SheetAxisVisibilityAction | undefined {
  if (!(SHEET_AXIS_VISIBILITY_COMMAND_IDS as readonly string[]).includes(id)) return undefined
  return { axis: /(?:-row|-rows)(?:-|$)/.test(id) ? 'row' : 'column', visible: !id.includes('.hide-') && !id.endsWith('-hidden'),
    selectedOnly: id === SHEET_AXIS_VISIBILITY_COMMANDS.showSelectedRows || id === SHEET_AXIS_VISIBILITY_COMMANDS.showSelectedColumns }
}
export function sheetAxisVisibilityCommandInput(value?: unknown): SheetAxisVisibilityCommandInput {
  if (value === undefined) return {}
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError('行列显示命令参数无效。')
  const input = value as SheetAxisVisibilityCommandInput
  const ranges = sheetRangeCommandInputRanges(input)
  for (const key of ['index', 'count', 'expectedRevision'] as const) {
    const number = input[key]
    if (number !== undefined && (!Number.isSafeInteger(number) || number < 0 || key === 'count' && number === 0)) throw new RangeError('行列显示命令的索引、数量或版本无效。')
  }
  if (input.all !== undefined && typeof input.all !== 'boolean') throw new TypeError('显示全部选项必须是布尔值。')
  return { ...input, ...(ranges === undefined ? {} : { ranges }) }
}
function counts(sheet: SheetAxisVisibilityState): { rows: number; columns: number } {
  const rows = sheet.rowCount ?? 1_048_576, columns = sheet.columnCount ?? 16_384
  if (!Number.isSafeInteger(rows) || rows <= 0 || rows > 1_048_576 || !Number.isSafeInteger(columns) || columns <= 0 || columns > 16_384) throw new RangeError('工作表行列总数无效。')
  return { rows, columns }
}
function axisRange(sheet: SheetAxisVisibilityState, axis: SheetAxis, interval: SheetAxisInterval): CellRange {
  const { rows, columns } = counts(sheet)
  return axis === 'row' ? { startRow: interval.start, endRow: interval.end, startCol: 0, endCol: columns - 1 }
    : { startRow: 0, endRow: rows - 1, startCol: interval.start, endCol: interval.end }
}
function intervals(ranges: readonly CellRange[], axis: SheetAxis): SheetAxisInterval[] {
  return mergeSheetAxisIntervals(ranges.map(range => axis === 'row' ? { start: range.startRow, end: range.endRow } : { start: range.startCol, end: range.endCol }))
}
export function sheetAxisVisibilityTargetRanges(sheet: SheetAxisVisibilityState, action: SheetAxisVisibilityAction, ranges: readonly CellRange[], input: SheetAxisVisibilityCommandInput = {}): CellRange[] {
  const { rows, columns } = counts(sheet), count = action.axis === 'row' ? rows : columns
  assertSheetCommandRanges(ranges)
  if (input.index !== undefined && (!Number.isSafeInteger(input.index) || input.index < 0 || input.index >= count)
    || input.count !== undefined && (!Number.isSafeInteger(input.count) || input.count <= 0)) throw new RangeError('行列显示目标的索引或数量无效。')
  if (input.all) return [axisRange(sheet, action.axis, { start: 0, end: count - 1 })]
  if (input.index !== undefined && input.count !== undefined) {
    if (input.index + input.count > count) throw new RangeError('目标行列超出工作表。')
    return [axisRange(sheet, action.axis, { start: input.index, end: input.index + input.count - 1 })]
  }
  const source = action.selectedOnly && input.index === undefined ? ranges.filter(range => action.axis === 'row'
    ? range.startCol === 0 && range.endCol >= columns - 1 : range.startRow === 0 && range.endRow >= rows - 1) : ranges
  return intervals(sheetAxisDimensionTargetRanges(sheet, action.axis, source, input.index), action.axis).map(interval => axisRange(sheet, action.axis, interval))
}
function manualHidden(sheet: SheetAxisVisibilityState, axis: SheetAxis): SheetAxisSpan<boolean>[] {
  const { rows, columns } = counts(sheet), count = axis === 'row' ? rows : columns
  const spans = axis === 'row' ? sheet.rowHiddenSpans : sheet.columnHiddenSpans, points = axis === 'row' ? sheet.hiddenRows : sheet.hiddenColumns
  if ((points ?? []).some(index => !Number.isSafeInteger(index) || index < 0 || index >= count)
    || (spans ?? []).some(span => span.end >= count || typeof span.value !== 'boolean')) throw new RangeError('工作表隐藏行列记录无效。')
  return canonicalSheetAxisSpans([...(spans ?? []), ...(points ?? []).map(index => ({ start: index, end: index, value: true }))]).filter(span => span.value)
}
function sizeState(sheet: SheetAxisVisibilityState, axis: SheetAxis) {
  if (axis === 'row') return { ...effectiveSheetRowSizes(sheet), defaultSize: sheet.defaultRowHeight ?? 26 }
  const sizes = sheet.defaultColumnWidth === undefined && !sheetAxisSpanAt(sheet.columnWidthSpans, 0) ? { 0: 152, ...sheet.columnWidths } : sheet.columnWidths ?? {}
  return { sizes, spans: sheet.columnWidthSpans ?? [], defaultSize: sheet.defaultColumnWidth ?? 112 }
}
function axisIndex(sheet: SheetAxisVisibilityState, axis: SheetAxis, hidden: readonly SheetAxisSpan<boolean>[], filteredRows: readonly number[]) {
  const { rows, columns } = counts(sheet), size = sizeState(sheet, axis)
  const collapsed = sheet.outline?.groups.filter(group => group.axis === axis && group.collapsed).map(group => ({ start: group.start, end: group.end, value: true })) ?? []
  return new SparseAxisIndex(axis === 'row' ? rows : columns, size.defaultSize, size.sizes,
    axis === 'row' ? [...(sheet.filteredRows ?? []), ...(sheet.autoFilter?.filteredRows ?? []), ...(sheet.outlineHiddenRows ?? []), ...filteredRows] : sheet.outlineHiddenColumns,
    size.spans, canonicalSheetAxisSpans([...hidden, ...collapsed]))
}
function intersection(first: readonly SheetAxisInterval[], second: readonly SheetAxisInterval[]): SheetAxisInterval[] {
  const result: SheetAxisInterval[] = []; let left = 0, right = 0
  while (left < first.length && right < second.length) {
    const a = first[left]!, b = second[right]!, start = Math.max(a.start, b.start), end = Math.min(a.end, b.end)
    if (start <= end) result.push({ start, end })
    if (a.end < b.end) left += 1; else right += 1
  }
  return result
}
function subtraction(source: readonly SheetAxisInterval[], removed: readonly SheetAxisInterval[]): SheetAxisInterval[] {
  const result: SheetAxisInterval[] = []; let first = 0
  for (const interval of source) {
    let start = interval.start
    while (first < removed.length && removed[first]!.end < start) first += 1
    for (let index = first; index < removed.length && removed[index]!.start <= interval.end; index++) {
      const gap = removed[index]!
      if (gap.start > start) result.push({ start, end: gap.start - 1 })
      start = Math.max(start, gap.end + 1)
      if (start > interval.end) break
    }
    if (start <= interval.end) result.push({ start, end: interval.end })
  }
  return result
}
function zeroSizeIntervals(sheet: SheetAxisVisibilityState, axis: SheetAxis): SheetAxisInterval[] {
  const { rows, columns } = counts(sheet), state = sizeState(sheet, axis)
  return canonicalSheetAxisSpans([{ start: 0, end: (axis === 'row' ? rows : columns) - 1, value: state.defaultSize }, ...state.spans,
    ...Object.entries(state.sizes).map(([index, value]) => ({ start: Number(index), end: Number(index), value }))]).filter(span => span.value === 0)
}
export function sheetAxisVisibilityTargetHasHidden(sheet: SheetAxisVisibilityState, axis: SheetAxis, ranges: readonly CellRange[]): boolean {
  if (!ranges.length) return false
  const targets = intervals(ranges, axis)
  return intersection(targets, manualHidden(sheet, axis)).length > 0 || intersection(targets, zeroSizeIntervals(sheet, axis)).length > 0
}
function clearPoints(values: Readonly<Record<number, number>> | undefined, ranges: readonly SheetAxisInterval[]): Record<number, number> {
  const selected = ranges.map(interval => ({ ...interval, value: true }))
  return Object.fromEntries(Object.entries(values ?? {}).filter(([index]) => !sheetAxisSpanAt(selected, Number(index))))
}
function restoreZeroSizes(sheet: SpreadsheetSheet, axis: SheetAxis, ranges: readonly SheetAxisInterval[], changes: SheetAxisVisibilityChanges): void {
  if (!ranges.length) return
  if (axis === 'column') {
    changes.columnWidths = clearPoints(sheet.columnWidths, ranges)
    let spans = sheet.columnWidthSpans
    for (const range of ranges) spans = assignSheetAxisSpan(spans, range.start, range.end, (sheet.defaultColumnWidth ?? 112) || 112)
    changes.columnWidthSpans = spans
  } else {
    changes.rowHeights = clearPoints(sheet.rowHeights, ranges)
    changes.rowAutoHeights = clearPoints(sheet.rowAutoHeights, ranges)
    let spans = sheet.rowHeightSpans, modes = sheet.rowAutoHeightSpans
    for (const range of ranges) {
      spans = assignSheetAxisSpan(spans, range.start, range.end, (sheet.defaultRowHeight ?? 26) || 26)
      modes = assignSheetAxisSpan(modes, range.start, range.end, false)
    }
    changes.rowHeightSpans = spans; changes.rowAutoHeightSpans = modes
  }
}

/** Plan all model/selection changes before commit; large contiguous hides never allocate one entry per axis. */
export function planSheetAxisVisibilityCommand(sheet: SpreadsheetSheet, ranges: readonly CellRange[], action: SheetAxisVisibilityAction,
  input: SheetAxisVisibilityCommandInput, userId: string, filteredRows: readonly number[] = []): SheetAxisVisibilityPlan {
  const targets = sheetAxisVisibilityTargetRanges(sheet, action, ranges, input)
  if (!targets.length) return { changed: false, changes: {}, ranges: [], selections: [] }
  assertSheetRangesEditable(sheet, targets, userId)
  const previous = manualHidden(sheet, action.axis), selected = intervals(targets, action.axis)
  let next = previous
  for (const range of selected) next = assignSheetAxisSpan(next, range.start, range.end, action.visible ? undefined : true)
  const changes: SheetAxisVisibilityChanges = {}, hidesChanged = JSON.stringify(previous) !== JSON.stringify(next)
  if (hidesChanged) {
    if (action.axis === 'row') { changes.hiddenRows = []; changes.rowHiddenSpans = next }
    else { changes.hiddenColumns = []; changes.columnHiddenSpans = next }
  }
  // A hidden flag and a zero size are separate source properties. Restore the
  // size only for axes that were manually visible before this explicit show.
  const zeroSizes = action.visible ? subtraction(intersection(selected, zeroSizeIntervals(sheet, action.axis)), previous) : []
  restoreZeroSizes(sheet, action.axis, zeroSizes, changes)
  const draft = { ...sheet, ...changes }, index = axisIndex(draft, action.axis, next, filteredRows)
  if (!action.visible && index.nextVisible(0, 1) === undefined) throw new SheetAxisVisibilityError(action.axis)
  const selections = action.visible ? targets : selected.flatMap(range => {
    const before = range.start > 0 ? index.nextVisible(range.start - 1, -1) : undefined
    const nextIndex = before ?? index.nextVisible(range.end + 1, 1)
    return nextIndex === undefined ? [] : [axisRange(sheet, action.axis, { start: nextIndex, end: nextIndex })]
  })
  return { changed: hidesChanged || zeroSizes.length > 0, changes, ranges: targets, selections: selections.filter((range, position) =>
    !selections.slice(0, position).some(previous => previous.startRow === range.startRow && previous.startCol === range.startCol && previous.endRow === range.endRow && previous.endCol === range.endCol)) }
}
