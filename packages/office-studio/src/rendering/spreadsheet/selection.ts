import type { GridRange, GridViewport, SpreadsheetGeometry } from './layout.ts'
import { GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH } from './layout.ts'
import type { LineCommand, RenderRect } from '../contracts.ts'
import { resolveSheetFillTarget, sheetFillPointerTarget } from '../../features/sheet-workbench/sheet-fill-target.ts'

export const SHEET_SELECTION_STROKE_WIDTH = 1
export const SHEET_SELECTION_HANDLE_SIZE = 6
const SHEET_SELECTION_HANDLE_HIT_SIZE = 14

/** Source outlines use the same visible panes as selection and emit only visible dash segments, even for whole-sheet ranges. */
export function sheetRangeDashedOutlineCommands(
  geometry: SpreadsheetGeometry,
  range: GridRange,
  viewport: GridViewport,
  color = '#246bfd',
): LineCommand[] {
  const commands: LineCommand[] = []
  const width = SHEET_SELECTION_STROKE_WIDTH, inset = width / 2, dash = 4, period = 7
  const addEdge = (start: number, end: number, fixed: number, horizontal: boolean, clip: RenderRect) => {
    const low = horizontal ? clip.x : clip.y
    const high = low + (horizontal ? clip.width : clip.height)
    const crossLow = horizontal ? clip.y : clip.x
    const crossHigh = crossLow + (horizontal ? clip.height : clip.width)
    if (fixed < crossLow - inset || fixed > crossHigh + inset || end <= low || start >= high) return
    const first = Math.max(0, Math.floor((low - start) / period))
    const limit = Math.min(end, high)
    for (let offset = start + first * period; offset < limit; offset += period) {
      const from = Math.max(low, offset), to = Math.min(limit, offset + dash)
      if (to <= from) continue
      commands.push(horizontal
        ? { kind: 'line', x1: from, y1: fixed, x2: to, y2: fixed, color, width, clip }
        : { kind: 'line', x1: fixed, y1: from, x2: fixed, y2: to, color, width, clip })
    }
  }
  for (const pane of geometry.panes(viewport)) {
    if (pane.clip.width <= 0 || pane.clip.height <= 0 || !pane.rows.length || !pane.columns.length) continue
    if (!pane.rows.some(row => row >= range.startRow && row <= range.endRow)
      || !pane.columns.some(column => column >= range.startColumn && column <= range.endColumn)) continue
    const frozenRowPane = pane.rows[0]! < geometry.frozenRows
    const frozenColumnPane = pane.columns[0]! < geometry.frozenColumns
    const originX = GRID_HEADER_WIDTH - (frozenColumnPane ? 0 : viewport.scrollLeft)
    const originY = GRID_HEADER_HEIGHT - (frozenRowPane ? 0 : viewport.scrollTop)
    const left = originX + geometry.columns.offset(range.startColumn) * viewport.zoom + inset
    const right = originX + geometry.columns.offset(range.endColumn + 1) * viewport.zoom + inset
    const top = originY + geometry.rows.offset(range.startRow) * viewport.zoom + inset
    const bottom = originY + geometry.rows.offset(range.endRow + 1) * viewport.zoom + inset
    if ((range.startRow < geometry.frozenRows) === frozenRowPane) addEdge(left, right, top, true, pane.clip)
    if ((range.endRow < geometry.frozenRows) === frozenRowPane) addEdge(left, right, bottom, true, pane.clip)
    if ((range.startColumn < geometry.frozenColumns) === frozenColumnPane) addEdge(top, bottom, left, false, pane.clip)
    if ((range.endColumn < geometry.frozenColumns) === frozenColumnPane) addEdge(top, bottom, right, false, pane.clip)
  }
  return commands
}

export function fillPreviewRange(
  source: GridRange,
  target: { row: number; column: number },
): GridRange | undefined {
  const range = { startRow: source.startRow, endRow: source.endRow, startCol: source.startColumn, endCol: source.endColumn }
  const selection = sheetFillPointerTarget(range, target.row, target.column)
  if (!resolveSheetFillTarget(range, selection)) return undefined
  return { startRow: selection.startRow, endRow: selection.endRow, startColumn: selection.startCol, endColumn: selection.endCol }
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
  // A geometric fragment no longer carries the source's whole-axis selection intent.
  if (source.startRow < top) pieces.push({ startRow: source.startRow, endRow: top - 1, startColumn: source.startColumn, endColumn: source.endColumn })
  if (bottom < source.endRow) pieces.push({ startRow: bottom + 1, endRow: source.endRow, startColumn: source.startColumn, endColumn: source.endColumn })
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
  const disjoint = ranges.reduce<GridRange[]>((result, range) => addRange(result, range), [])
  return disjoint.reduce((total, range) => total + (range.endRow - range.startRow + 1) * (range.endColumn - range.startColumn + 1), 0)
}

export function selectionBounds(ranges: readonly GridRange[]): GridRange | undefined {
  if (!ranges.length) return undefined
  return ranges.reduce((bounds, range) => ({ startRow: Math.min(bounds.startRow, range.startRow), startColumn: Math.min(bounds.startColumn, range.startColumn), endRow: Math.max(bounds.endRow, range.endRow), endColumn: Math.max(bounds.endColumn, range.endColumn) }))
}
