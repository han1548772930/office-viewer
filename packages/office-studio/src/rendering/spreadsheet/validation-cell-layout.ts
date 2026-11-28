import type { RenderRect } from '../contracts.ts'
import { intersectRect } from '../geometry.ts'
import { parseAddress, type GridViewport, type SpreadsheetGeometry } from './layout.ts'

/** Cell notices and controls use a visible fragment of the same merged/frozen layout. */
export function sheetValidationCellVisibleRect(geometry: SpreadsheetGeometry, viewport: GridViewport, key?: string, minimumSize = 0): RenderRect | undefined {
  const position = key && parseAddress(key)
  if (!position || !Number.isFinite(viewport.zoom) || viewport.zoom <= 0) return undefined
  const merge = geometry.mergeAt(position.row, position.column)
  const range = merge ?? { startRow: position.row, endRow: position.row, startColumn: position.column, endColumn: position.column }
  const candidates = geometry.panes(viewport).flatMap(pane => {
    const rows = pane.rows.filter(row => row >= range.startRow && row <= range.endRow)
    const columns = pane.columns.filter(column => column >= range.startColumn && column <= range.endColumn)
    if (!rows.length || !columns.length) return []
    const frozenRows = pane.rows[0]! < geometry.frozenRows, frozenColumns = pane.columns[0]! < geometry.frozenColumns
    const x = geometry.headerWidth + geometry.columns.offset(range.startColumn) * viewport.zoom - (frozenColumns ? 0 : viewport.scrollLeft)
    const y = geometry.headerHeight + geometry.rows.offset(range.startRow) * viewport.zoom - (frozenRows ? 0 : viewport.scrollTop)
    const rect = intersectRect(pane.clip, { x, y,
      width: (geometry.columns.offset(range.endColumn + 1) - geometry.columns.offset(range.startColumn)) * viewport.zoom,
      height: (geometry.rows.offset(range.endRow + 1) - geometry.rows.offset(range.startRow)) * viewport.zoom })
    return rect.width > 0 && rect.height > 0 && rect.width >= minimumSize && rect.height >= minimumSize
      ? [{ rect, lastColumn: columns.at(-1)!, firstRow: rows[0]! }] : []
  }).sort((left, right) => right.lastColumn - left.lastColumn || left.firstRow - right.firstRow)
  return candidates[0]?.rect
}

/** A single Ant trigger follows the visible merged cell through the same frozen-pane transforms as Canvas. */
export function sheetValidationCellControlRect(geometry: SpreadsheetGeometry, viewport: GridViewport, key?: string): RenderRect | undefined {
  const visible = sheetValidationCellVisibleRect(geometry, viewport, key, 12)
  if (!visible) return undefined
  const width = Math.min(20, visible.width - 2), height = Math.min(20, visible.height - 2)
  return { x: visible.x + visible.width - width - 1, y: visible.y + (visible.height - height) / 2, width, height }
}
