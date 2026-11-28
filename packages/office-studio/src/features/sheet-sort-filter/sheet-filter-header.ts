import type { SheetAutoFilter } from '../../domain/model.ts'
import type { RenderRect } from '../../rendering/contracts.ts'
import { intersectRect } from '../../rendering/geometry.ts'
import type { GridViewport, SpreadsheetGeometry } from '../../rendering/spreadsheet/layout.ts'
import { sheetViewCellAt } from '../../services/excelViewAdapter.ts'

export const SHEET_FILTER_ICON_SIZE = 16
export const SHEET_FILTER_ICON_PADDING = 1
export interface SheetFilterHeaderControl {
  key: string
  row: number
  column: number
  active: boolean
  disabled: boolean
  rect: RenderRect
  clipPath: string
}

/** Only visible filter-heading cells receive controls; frozen panes share the same cell transform. */
export function sheetFilterHeaderControls(geometry: SpreadsheetGeometry, viewport: GridViewport, filter: SheetAutoFilter | undefined,
  canOpen: (column: number) => boolean): SheetFilterHeaderControl[] {
  if (!filter) return []
  const row = filter.range.startRow, controls: SheetFilterHeaderControl[] = []
  for (const [paneIndex, pane] of geometry.panes(viewport).entries()) {
    if (!pane.rows.includes(row)) continue
    for (const column of pane.columns) {
      if (column < filter.range.startCol || column > filter.range.endCol) continue
      const cell = geometry.cellRect(row, column, viewport, false), size = SHEET_FILTER_ICON_SIZE * viewport.zoom,
        padding = SHEET_FILTER_ICON_PADDING * viewport.zoom
      if (cell.width <= padding || cell.height <= padding) continue
      const vertical = sheetViewCellAt(geometry.sheet, row, column)?.vertical ?? 'bottom'
      const rect = { x: cell.x + cell.width - size - padding,
        y: vertical === 'top' ? cell.y + padding : vertical === 'middle' ? cell.y + Math.max(0, (cell.height - size) / 2)
          : cell.y + cell.height - size - padding, width: size, height: size }
      const clip = intersectRect(intersectRect(rect, cell), pane.clip)
      if (clip.width <= 0 || clip.height <= 0) continue
      controls.push({ key: `${paneIndex}:${column}`, row, column, active: filter.columns[column] !== undefined, disabled: !canOpen(column), rect,
        clipPath: `inset(${Math.max(0, clip.y - rect.y)}px ${Math.max(0, rect.x + rect.width - clip.x - clip.width)}px ${Math.max(0, rect.y + rect.height - clip.y - clip.height)}px ${Math.max(0, clip.x - rect.x)}px)` })
    }
  }
  return controls
}
