import type { CellRange } from './model.ts'
import { assertSheetAxisEdit, sheetAxisMaximum, sheetRangeAxisInterval, sheetRangeWithAxisInterval, transformSheetAxisRange } from './sheet-axis-structure.ts'

/** A structure edit affects one perpendicular band, unlike inserting a whole axis. */
export interface SheetCellShiftEdit {
  kind: 'insert' | 'delete'
  axis: 'row' | 'column'
  start: number
  count: number
  bandStart: number
  bandEnd: number
}
export function assertSheetCellShiftEdit(edit: SheetCellShiftEdit): void {
  if (!edit || typeof edit !== 'object') throw new TypeError('单元格移位参数无效。')
  assertSheetAxisEdit(edit)
  const bandMaximum = sheetAxisMaximum(edit.axis === 'row' ? 'column' : 'row')
  if (edit.kind !== 'insert' && edit.kind !== 'delete' || !Number.isSafeInteger(edit.bandStart) || !Number.isSafeInteger(edit.bandEnd)
    || edit.bandStart < 0 || edit.bandEnd < edit.bandStart || edit.bandEnd >= bandMaximum
    || edit.start + edit.count > sheetAxisMaximum(edit.axis)) throw new RangeError('单元格移位带或尺寸无效。')
}
export function sheetCellShiftAffectedRange(edit: SheetCellShiftEdit): CellRange {
  assertSheetCellShiftEdit(edit)
  return edit.axis === 'row'
    ? { startRow: edit.start, endRow: 1_048_575, startCol: edit.bandStart, endCol: edit.bandEnd }
    : { startRow: edit.bandStart, endRow: edit.bandEnd, startCol: edit.start, endCol: 16_383 }
}
export function sheetCellShiftContainsBand(range: CellRange, edit: SheetCellShiftEdit): boolean {
  const band = sheetRangeAxisInterval(range, edit.axis === 'row' ? 'column' : 'row')
  return band.start >= edit.bandStart && band.end <= edit.bandEnd
}
export function mapSheetCellShiftPosition(row: number, col: number, edit: SheetCellShiftEdit): { row: number; col: number } | undefined {
  const band = edit.axis === 'row' ? col : row, axis = edit.axis === 'row' ? row : col
  if (band < edit.bandStart || band > edit.bandEnd || axis < edit.start) return { row, col }
  if (edit.kind === 'delete' && axis < edit.start + edit.count) return undefined
  const next = axis + (edit.kind === 'insert' ? edit.count : -edit.count)
  if (next >= sheetAxisMaximum(edit.axis)) throw new RangeError('移位会把已有内容移出 Office 行列上限。')
  return edit.axis === 'row' ? { row: next, col } : { row, col: next }
}
export function inverseSheetCellShiftPosition(row: number, col: number, edit: SheetCellShiftEdit): { row: number; col: number } | undefined {
  const band = edit.axis === 'row' ? col : row, axis = edit.axis === 'row' ? row : col
  if (band < edit.bandStart || band > edit.bandEnd || axis < edit.start) return { row, col }
  if (edit.kind === 'insert' && axis < edit.start + edit.count) return undefined
  const next = axis + (edit.kind === 'insert' ? -edit.count : edit.count)
  if (next >= sheetAxisMaximum(edit.axis)) return undefined
  return edit.axis === 'row' ? { row: next, col } : { row, col: next }
}
/** At most three disjoint rectangles; existing membership is never clipped at the Office boundary. */
export function transformSheetCellShiftRange(range: CellRange, edit: SheetCellShiftEdit, expandInsertion = true): CellRange[] {
  assertSheetCellShiftEdit(edit)
  const bandAxis = edit.axis === 'row' ? 'column' : 'row', band = sheetRangeAxisInterval(range, bandAxis), axis = sheetRangeAxisInterval(range, edit.axis)
  if (band.end < edit.bandStart || band.start > edit.bandEnd || axis.end < edit.start) return [{ ...range }]
  if (edit.kind === 'insert' && axis.end + edit.count >= sheetAxisMaximum(edit.axis)) throw new RangeError('移位会把已有范围移出 Office 行列上限。')
  const result: CellRange[] = []
  if (band.start < edit.bandStart) result.push(sheetRangeWithAxisInterval(range, bandAxis, { start: band.start, end: edit.bandStart - 1 }))
  if (band.end > edit.bandEnd) result.push(sheetRangeWithAxisInterval(range, bandAxis, { start: edit.bandEnd + 1, end: band.end }))
  const selected = sheetRangeWithAxisInterval(range, bandAxis, { start: Math.max(band.start, edit.bandStart), end: Math.min(band.end, edit.bandEnd) })
  result.push(...transformSheetAxisRange(selected, edit, expandInsertion))
  return result
}
