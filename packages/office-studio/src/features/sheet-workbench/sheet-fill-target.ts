import type { CellRange } from '../../domain/model.ts'
import { assertSheetCommandRanges, sheetRangesEqual } from './sheet-range-command.ts'

export type SheetFillDirection = 'up' | 'down' | 'left' | 'right'
export interface SheetFillTarget {
  direction: SheetFillDirection
  selection: CellRange
  destination: CellRange
  clear: boolean
}

/** The command, Facade and pointer preview share one physical range contract. */
export function resolveSheetFillTarget(source: CellRange, target: CellRange, allowClear = true): SheetFillTarget | undefined {
  assertSheetCommandRanges([source, target])
  if (sheetRangesEqual(source, target)) return undefined
  const sameColumns = source.startCol === target.startCol && source.endCol === target.endCol
  const sameRows = source.startRow === target.startRow && source.endRow === target.endRow
  if (allowClear && target.startRow === source.startRow && target.startCol === source.startCol) {
    if (sameColumns && target.endRow < source.endRow) return { direction: 'down', selection: { ...target }, clear: true,
      destination: { ...source, startRow: target.endRow + 1 } }
    if (sameRows && target.endCol < source.endCol) return { direction: 'right', selection: { ...target }, clear: true,
      destination: { ...source, startCol: target.endCol + 1 } }
  }
  if (sameColumns && source.startRow === target.startRow && target.endRow > source.endRow) return {
    direction: 'down', selection: { ...target }, destination: { ...target, startRow: source.endRow + 1 }, clear: false }
  if (sameColumns && source.endRow === target.endRow && target.startRow < source.startRow) return {
    direction: 'up', selection: { ...target }, destination: { ...target, endRow: source.startRow - 1 }, clear: false }
  if (sameRows && source.startCol === target.startCol && target.endCol > source.endCol) return {
    direction: 'right', selection: { ...target }, destination: { ...target, startCol: source.endCol + 1 }, clear: false }
  if (sameRows && source.endCol === target.endCol && target.startCol < source.startCol) return {
    direction: 'left', selection: { ...target }, destination: { ...target, endCol: source.startCol - 1 }, clear: false }
  throw new RangeError('自动填充的目标必须沿一个方向扩展源区域。')
}

/** A corner drag chooses one axis; dragging inside the source clears its trailing cells. */
export function sheetFillPointerTarget(source: CellRange, row: number, column: number): CellRange {
  if (!Number.isSafeInteger(row) || !Number.isSafeInteger(column) || row < 0 || column < 0) throw new RangeError('填充位置无效。')
  const rowDelta = row < source.startRow ? source.startRow - row : Math.max(0, row - source.endRow)
  const columnDelta = column < source.startCol ? source.startCol - column : Math.max(0, column - source.endCol)
  if (rowDelta || columnDelta) return rowDelta >= columnDelta
    ? { ...source, startRow: Math.min(source.startRow, row), endRow: Math.max(source.endRow, row) }
    : { ...source, startCol: Math.min(source.startCol, column), endCol: Math.max(source.endCol, column) }
  const rows = source.endRow - row, columns = source.endCol - column
  return rows >= columns && source.endRow > source.startRow ? { ...source, endRow: row }
    : source.endCol > source.startCol ? { ...source, endCol: column } : { ...source }
}

export function sheetCopyFillRanges(range: CellRange, direction: 'down' | 'right'): { sourceRange: CellRange; targetRange: CellRange } | undefined {
  assertSheetCommandRanges([range])
  if (direction === 'down') {
    if (range.startRow === range.endRow && range.startRow === 0) return undefined
    const startRow = range.startRow === range.endRow ? range.startRow - 1 : range.startRow
    return { sourceRange: { ...range, startRow, endRow: startRow }, targetRange: { ...range, startRow } }
  }
  if (range.startCol === range.endCol && range.startCol === 0) return undefined
  const startCol = range.startCol === range.endCol ? range.startCol - 1 : range.startCol
  return { sourceRange: { ...range, startCol, endCol: startCol }, targetRange: { ...range, startCol } }
}
