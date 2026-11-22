import type { CellRange } from '../../domain/model.ts'
import { sheetAxisMaximum, sheetRangeAxisInterval, type SheetAxis, type SheetAxisInterval } from '../../domain/sheet-axis-structure.ts'

interface SheetHeaderExtent { rowCount: number; columnCount: number }
export interface SheetAxisHeaderMoveSource { axis: SheetAxis; interval: SheetAxisInterval; range: CellRange }

export function sheetAxisHeaderRange(extent: SheetHeaderExtent, axis: SheetAxis, first: number, last = first): CellRange {
  const count = axis === 'row' ? extent.rowCount : extent.columnCount
  if (!Number.isSafeInteger(first) || !Number.isSafeInteger(last) || first < 0 || last < 0 || first >= count || last >= count) throw new RangeError('行列头选区超出工作表。')
  const start = Math.min(first, last), end = Math.max(first, last)
  return axis === 'row' ? { startRow: start, endRow: end, startCol: 0, endCol: extent.columnCount - 1 }
    : { startRow: 0, endRow: extent.rowCount - 1, startCol: start, endCol: end }
}

/** The first whole-axis selection containing the pressed header is the move source. */
export function sheetAxisHeaderMoveSource(extent: SheetHeaderExtent, ranges: readonly CellRange[], axis: SheetAxis, index: number): SheetAxisHeaderMoveSource | undefined {
  const wholeSheet = (range: CellRange) => range.startRow === 0 && range.endRow === extent.rowCount - 1
    && range.startCol === 0 && range.endCol === extent.columnCount - 1
  const range = ranges.find(candidate => !wholeSheet(candidate) && (axis === 'row'
    ? candidate.startCol === 0 && candidate.endCol === extent.columnCount - 1 && index >= candidate.startRow && index <= candidate.endRow
    : candidate.startRow === 0 && candidate.endRow === extent.rowCount - 1 && index >= candidate.startCol && index <= candidate.endCol))
  return range ? { axis, interval: sheetRangeAxisInterval(range, axis), range: { ...range } } : undefined
}

/** Destination uses coordinates before source removal, as in the fixed-version move command. */
export function sheetAxisHeaderDestination(source: SheetAxisInterval, hoverIndex: number): number | undefined {
  if (!Number.isSafeInteger(hoverIndex) || hoverIndex < 0) return undefined
  return hoverIndex < source.start ? hoverIndex : hoverIndex > source.end ? hoverIndex + 1 : undefined
}

export function sheetAxisHeaderAffectedRange(extent: SheetHeaderExtent, source: SheetAxisHeaderMoveSource, destination: number, copy: boolean): CellRange {
  const start = Math.min(source.interval.start, destination)
  const end = copy ? sheetAxisMaximum(source.axis) - 1 : Math.max(source.interval.end, destination - 1)
  return source.axis === 'row' ? { startRow: start, endRow: end, startCol: 0, endCol: extent.columnCount - 1 }
    : { startRow: 0, endRow: extent.rowCount - 1, startCol: start, endCol: end }
}
