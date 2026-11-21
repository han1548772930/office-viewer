import type { GridRange } from './layout.ts'

export const SHEET_SELECTION_STROKE_WIDTH = 1
export const SHEET_SELECTION_HANDLE_SIZE = 6

export function containsCell(ranges: readonly GridRange[], row: number, column: number): boolean {
  return ranges.some(range => row >= range.startRow && row <= range.endRow && column >= range.startColumn && column <= range.endColumn)
}

/** Returns disjoint rectangles, so whole-sheet selections remain constant sized. */
export function subtractRange(source: GridRange, cut: GridRange): GridRange[] {
  const top = Math.max(source.startRow, cut.startRow), bottom = Math.min(source.endRow, cut.endRow)
  const left = Math.max(source.startColumn, cut.startColumn), right = Math.min(source.endColumn, cut.endColumn)
  if (top > bottom || left > right) return [source]
  const pieces: GridRange[] = []
  if (source.startRow < top) pieces.push({ ...source, endRow: top - 1 })
  if (bottom < source.endRow) pieces.push({ ...source, startRow: bottom + 1 })
  if (source.startColumn < left) pieces.push({ startRow: top, endRow: bottom, startColumn: source.startColumn, endColumn: left - 1 })
  if (right < source.endColumn) pieces.push({ startRow: top, endRow: bottom, startColumn: right + 1, endColumn: source.endColumn })
  return pieces
}

export function addRange(ranges: readonly GridRange[], range: GridRange): GridRange[] {
  let additions = [range]
  for (const existing of ranges) additions = additions.flatMap(addition => subtractRange(addition, existing))
  return [...ranges, ...additions]
}

export function selectionCount(ranges: readonly GridRange[]): number {
  return ranges.reduce((total, range) => total + (range.endRow - range.startRow + 1) * (range.endColumn - range.startColumn + 1), 0)
}

export function selectionBounds(ranges: readonly GridRange[]): GridRange | undefined {
  if (!ranges.length) return undefined
  return ranges.reduce((bounds, range) => ({ startRow: Math.min(bounds.startRow, range.startRow), startColumn: Math.min(bounds.startColumn, range.startColumn), endRow: Math.max(bounds.endRow, range.endRow), endColumn: Math.max(bounds.endColumn, range.endColumn) }))
}
