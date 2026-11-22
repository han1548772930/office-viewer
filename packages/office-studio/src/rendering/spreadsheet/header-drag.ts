import type { DrawCommand, RenderRect } from '../contracts.ts'
import { GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH, type GridViewport, type SpreadsheetGeometry } from './layout.ts'
import type { SheetAxis, SheetAxisInterval } from '../../domain/sheet-axis-structure.ts'

export interface SheetHeaderDragPreview {
  axis: SheetAxis
  source: SheetAxisInterval
  destination?: number
  allowed: boolean
}
export interface SheetHeaderDragColors { source: string; target: string; invalid: string; preview: string }

/** Header hit coordinates ignore merges and use the same frozen axis transform as the grid. */
export function sheetHeaderAxisIndexAt(geometry: SpreadsheetGeometry, viewport: GridViewport, axis: SheetAxis, coordinate: number): number | undefined {
  if (!Number.isFinite(coordinate) || !Number.isFinite(viewport.zoom) || viewport.zoom <= 0) return undefined
  const index = axis === 'row' ? geometry.rows : geometry.columns
  const frozen = axis === 'row' ? geometry.frozenRows : geometry.frozenColumns
  const header = axis === 'row' ? GRID_HEADER_HEIGHT : GRID_HEADER_WIDTH
  const scroll = axis === 'row' ? viewport.scrollTop : viewport.scrollLeft
  const limit = axis === 'row' ? viewport.height : viewport.width
  if (limit <= header) return undefined
  const pixel = Math.max(header, Math.min(limit - 1, coordinate))
  const offset = (pixel - header + (pixel >= header + index.offset(frozen) * viewport.zoom ? scroll : 0)) / viewport.zoom
  const total = index.offset(index.count)
  if (!index.count || total <= 0 || offset < 0) return undefined
  if (offset >= total) return index.nextVisible(index.count - 1, -1)
  const located = index.locate(offset)
  return index.size(located) > 0 ? located : undefined
}

/** Preview commands are CSS coordinates appended after the header scene, on both rendering backends. */
export function sheetHeaderDragCommands(geometry: SpreadsheetGeometry, viewport: GridViewport, preview: SheetHeaderDragPreview | undefined,
  colors: SheetHeaderDragColors): DrawCommand[] {
  if (!preview || viewport.width <= GRID_HEADER_WIDTH || viewport.height <= GRID_HEADER_HEIGHT) return []
  const { axis, source, destination } = preview, index = axis === 'row' ? geometry.rows : geometry.columns
  const frozen = axis === 'row' ? geometry.frozenRows : geometry.frozenColumns
  const header = axis === 'row' ? GRID_HEADER_HEIGHT : GRID_HEADER_WIDTH
  const scroll = axis === 'row' ? viewport.scrollTop : viewport.scrollLeft
  const extent = axis === 'row' ? viewport.height : viewport.width
  const frozenEnd = Math.min(extent, header + index.offset(frozen) * viewport.zoom)
  const commands: DrawCommand[] = []
  const band = (first: number, last: number, origin: number, low: number, high: number) => {
    if (first > last || high <= low) return
    const start = origin + index.offset(first) * viewport.zoom, end = origin + index.offset(last + 1) * viewport.zoom
    const clip: RenderRect = axis === 'row' ? { x: 0, y: low, width: viewport.width, height: high - low }
      : { x: low, y: 0, width: high - low, height: viewport.height }
    commands.push(axis === 'row' ? { kind: 'rect', x: 0, y: start, width: viewport.width, height: end - start, fill: colors.source, clip }
      : { kind: 'rect', x: start, y: 0, width: end - start, height: viewport.height, fill: colors.source, clip })
  }
  band(source.start, Math.min(source.end, frozen - 1), header, header, frozenEnd)
  band(Math.max(source.start, frozen), source.end, header - scroll, frozenEnd, extent)
  if (destination === undefined) return commands
  const fixed = frozen > 0 && destination <= frozen, start = header + index.offset(destination) * viewport.zoom - (fixed ? 0 : scroll)
  const low = fixed ? header : frozenEnd, high = fixed ? frozenEnd : extent
  if (start < low - 2 || start > high + 2 || high <= low) return commands
  const clip: RenderRect = axis === 'row' ? { x: 0, y: low, width: viewport.width, height: high - low }
    : { x: low, y: 0, width: high - low, height: viewport.height }
  const size = (index.offset(source.end + 1) - index.offset(source.start)) * viewport.zoom
  commands.push(axis === 'row'
    ? { kind: 'rect', x: 0, y: start, width: viewport.width, height: size, fill: colors.preview, clip }
    : { kind: 'rect', x: start, y: 0, width: size, height: viewport.height, fill: colors.preview, clip })
  commands.push(axis === 'row' ? { kind: 'line', x1: 0, y1: start, x2: viewport.width, y2: start,
    color: preview.allowed ? colors.target : colors.invalid, width: 4, clip }
    : { kind: 'line', x1: start, y1: 0, x2: start, y2: viewport.height,
      color: preview.allowed ? colors.target : colors.invalid, width: 4, clip })
  return commands
}
