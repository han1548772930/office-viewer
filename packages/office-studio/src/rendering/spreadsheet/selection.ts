import type { GridRange, GridViewport, SpreadsheetGeometry } from './layout.ts'
import type { RenderRect } from '../contracts.ts'

export const SHEET_SELECTION_STROKE_WIDTH = 1
export const SHEET_SELECTION_HANDLE_SIZE = 6
const SHEET_SELECTION_HANDLE_HIT_SIZE = 14

export function fillPreviewRange(
  source: GridRange,
  target: { row: number; column: number },
): GridRange | undefined {
  if (target.row >= source.startRow && target.row <= source.endRow
      && target.column >= source.startColumn && target.column <= source.endColumn) return undefined
  return {
    startRow: Math.min(source.startRow, target.row),
    endRow: Math.max(source.endRow, target.row),
    startColumn: Math.min(source.startColumn, target.column),
    endColumn: Math.max(source.endColumn, target.column),
  }
}

export function selectionFillHandleHitRect(
  geometry: SpreadsheetGeometry,
  ranges: readonly GridRange[],
  viewport: GridViewport,
): RenderRect | undefined {
  const range = ranges.at(-1)
  if (!range) return undefined
  const endpoint = geometry.cellRect(range.endRow, range.endColumn, viewport, false)
  const size = SHEET_SELECTION_HANDLE_HIT_SIZE
  return { x: endpoint.x + endpoint.width - size / 2, y: endpoint.y + endpoint.height - size / 2, width: size, height: size }
}

export function hitSelectionFillHandle(
  geometry: SpreadsheetGeometry,
  ranges: readonly GridRange[],
  viewport: GridViewport,
  x: number,
  y: number,
): GridRange | undefined {
  const range = ranges.at(-1)
  const rect = selectionFillHandleHitRect(geometry, ranges, viewport)
  if (!range || !rect || x < rect.x || x >= rect.x + rect.width || y < rect.y || y >= rect.y + rect.height) return undefined
  return range
}

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
