import type { RenderRect } from '../contracts'
import { intersectRect } from '../geometry.ts'
import { GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH, parseAddress, type GridViewport, type SpreadsheetGeometry } from './layout.ts'

export interface CellEditorLayout {
  rect: RenderRect
  clip: RenderRect
  clipPath: string
}

/** Keep the input bridge at the cell's coordinates, clipped to its frozen pane. */
export function createCellEditorLayout(geometry: SpreadsheetGeometry, address: string, viewport: GridViewport): CellEditorLayout | undefined {
  const point = parseAddress(address)
  if (!point || point.row >= geometry.rows.count || point.column >= geometry.columns.count) return undefined
  const merge = geometry.mergeAt(point.row, point.column)
  const row = merge?.startRow ?? point.row, column = merge?.startColumn ?? point.column
  const rect = geometry.cellRect(row, column, viewport)
  const frozenRight = Math.min(viewport.width, GRID_HEADER_WIDTH + geometry.columns.offset(geometry.frozenColumns) * viewport.zoom)
  const frozenBottom = Math.min(viewport.height, GRID_HEADER_HEIGHT + geometry.rows.offset(geometry.frozenRows) * viewport.zoom)
  const x = column < geometry.frozenColumns ? GRID_HEADER_WIDTH : frozenRight
  const y = row < geometry.frozenRows ? GRID_HEADER_HEIGHT : frozenBottom
  const right = column < geometry.frozenColumns ? frozenRight : viewport.width
  const bottom = row < geometry.frozenRows ? frozenBottom : viewport.height
  const clip = intersectRect(rect, { x, y, width: Math.max(0, right - x), height: Math.max(0, bottom - y) })
  const clipPath = clip.width <= 0 || clip.height <= 0
    ? 'inset(50%)'
    : `inset(${Math.max(0, clip.y - rect.y)}px ${Math.max(0, rect.x + rect.width - clip.x - clip.width)}px ${Math.max(0, rect.y + rect.height - clip.y - clip.height)}px ${Math.max(0, clip.x - rect.x)}px)`
  return { rect, clip, clipPath }
}
