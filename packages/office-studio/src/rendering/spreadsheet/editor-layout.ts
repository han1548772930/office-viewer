import type { RenderRect } from '../contracts'
import { intersectRect } from '../geometry.ts'
import { GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH, parseAddress, type GridViewport, type SpreadsheetGeometry } from './layout.ts'
import { SHEET_SELECTION_HANDLE_SIZE, SHEET_SELECTION_STROKE_WIDTH } from './selection.ts'

export interface CellEditorLayout {
  rect: RenderRect
  fillHandle: RenderRect
  clip: RenderRect
  clipPath: string
  fillHandleClipPath: string
}

export interface CellEditorLayoutOptions { contentHeight?: number }

/** Keep the input bridge at the cell's coordinates, clipped to its frozen pane. */
export function createCellEditorLayout(geometry: SpreadsheetGeometry, address: string, viewport: GridViewport, options: CellEditorLayoutOptions = {}): CellEditorLayout | undefined {
  const point = parseAddress(address)
  if (!point || point.row >= geometry.rows.count || point.column >= geometry.columns.count) return undefined
  const merge = geometry.mergeAt(point.row, point.column)
  const row = merge?.startRow ?? point.row, column = merge?.startColumn ?? point.column
  const cellRect = geometry.cellRect(row, column, viewport)
  const frozenRight = Math.min(viewport.width, GRID_HEADER_WIDTH + geometry.columns.offset(geometry.frozenColumns) * viewport.zoom)
  const frozenBottom = Math.min(viewport.height, GRID_HEADER_HEIGHT + geometry.rows.offset(geometry.frozenRows) * viewport.zoom)
  const x = column < geometry.frozenColumns ? GRID_HEADER_WIDTH : frozenRight
  const y = row < geometry.frozenRows ? GRID_HEADER_HEIGHT : frozenBottom
  const right = column < geometry.frozenColumns ? frozenRight : viewport.width
  const bottom = row < geometry.frozenRows ? frozenBottom : viewport.height
  const pane = { x, y, width: Math.max(0, right - x), height: Math.max(0, bottom - y) }
  const halfStroke = SHEET_SELECTION_STROKE_WIDTH / 2
  const contentHeight = Number.isFinite(options.contentHeight) ? Math.max(cellRect.height, options.contentHeight!) : cellRect.height
  const height = Math.max(cellRect.height + SHEET_SELECTION_STROKE_WIDTH,
    Math.min(contentHeight + SHEET_SELECTION_STROKE_WIDTH, bottom - cellRect.y + halfStroke))
  const rect = {
    x: cellRect.x - halfStroke,
    y: cellRect.y - halfStroke,
    width: cellRect.width + SHEET_SELECTION_STROKE_WIDTH,
    height,
  }
  const clip = intersectRect(rect, pane)
  const handleSize = SHEET_SELECTION_HANDLE_SIZE
  const fillHandle = {
    x: cellRect.x + cellRect.width - handleSize / 2,
    y: cellRect.y + cellRect.height - handleSize / 2,
    width: handleSize,
    height: handleSize,
  }
  const fillHandleClip = intersectRect(fillHandle, pane)
  return {
    rect,
    fillHandle,
    clip,
    clipPath: clipPathFor(rect, clip),
    fillHandleClipPath: clipPathFor(fillHandle, fillHandleClip),
  }
}

function clipPathFor(rect: RenderRect, visible: RenderRect): string {
  if (visible.width <= 0 || visible.height <= 0) return 'inset(50%)'
  return `inset(${Math.max(0, visible.y - rect.y)}px ${Math.max(0, rect.x + rect.width - visible.x - visible.width)}px ${Math.max(0, rect.y + rect.height - visible.y - visible.height)}px ${Math.max(0, visible.x - rect.x)}px)`
}
