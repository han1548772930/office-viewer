import type { ExcelCellView } from '../../services/excelViewAdapter.ts'
import { sheetViewCellAt } from '../../services/excelViewAdapter.ts'
import type { RenderRect } from '../contracts.ts'
import { intersectRect } from '../geometry.ts'
import { cellAddress, type GridViewport, type SpreadsheetGeometry } from './layout.ts'

export interface SheetCheckboxCellPosition { key: string; row: number; column: number; rect: RenderRect; clip: RenderRect; clips: readonly RenderRect[] }

/** The widget follows the same merged-cell rectangle, clipping and visible sparse axes as Canvas. */
export function sheetCheckboxRect(cell: ExcelCellView | undefined, bounds: RenderRect, zoom: number): RenderRect {
  const fontSize = Number.isFinite(cell?.fontSize) && cell!.fontSize! > 0 ? cell!.fontSize! : 10
  const size = fontSize * 1.6 * zoom, margin = 6 * zoom
  const left = cell?.align === 'left' ? margin : cell?.align === 'right' ? bounds.width - margin - size : (bounds.width - size) / 2
  const top = cell?.vertical === 'top' ? 0 : cell?.vertical === 'bottom' ? bounds.height - size : (bounds.height - size) / 2
  return { x: bounds.x + left, y: bounds.y + top, width: size, height: size }
}

/** There is one Ant Checkbox per visible main cell, even if a merged range crosses multiple panes. */
export function sheetCheckboxCellPositions(geometry: SpreadsheetGeometry, viewport: GridViewport,
  renderable: (row: number, column: number) => boolean): SheetCheckboxCellPosition[] {
  if (!Number.isFinite(viewport.zoom) || viewport.zoom <= 0) return []
  const positions = new Map<string, { key: string; row: number; column: number; rect: RenderRect; clips: RenderRect[] }>()
  const supportedCells = new Map<string, boolean>()
  for (const pane of geometry.panes(viewport)) {
    if (pane.clip.width <= 0 || pane.clip.height <= 0) continue
    const anchors = new Map<string, { row: number; column: number }>()
    for (const row of pane.rows) for (const column of pane.columns) {
      const merge = geometry.mergeAt(row, column), mainRow = merge?.startRow ?? row, mainColumn = merge?.startColumn ?? column
      anchors.set(cellAddress(mainRow, mainColumn), { row: mainRow, column: mainColumn })
    }
    for (const [key, main] of anchors) {
      let supported = supportedCells.get(key)
      if (supported === undefined) { supported = renderable(main.row, main.column); supportedCells.set(key, supported) }
      if (!supported) continue
      const bounds = geometry.cellRect(main.row, main.column, viewport), clip = intersectRect(bounds, pane.clip)
      const rect = sheetCheckboxRect(sheetViewCellAt(geometry.sheet, main.row, main.column), bounds, viewport.zoom)
      if (clip.width <= 0 || clip.height <= 0) continue
      const previous = positions.get(key)
      if (previous) previous.clips.push(clip)
      else positions.set(key, { key, ...main, rect, clips: [clip] })
    }
  }
  return [...positions.values()].flatMap(position => {
    if (!position.clips.some(clip => { const visible = intersectRect(position.rect, clip); return visible.width > 0 && visible.height > 0 })) return []
    const x = Math.min(...position.clips.map(clip => clip.x)), y = Math.min(...position.clips.map(clip => clip.y))
    const right = Math.max(...position.clips.map(clip => clip.x + clip.width)), bottom = Math.max(...position.clips.map(clip => clip.y + clip.height))
    return [{ ...position, clip: { x, y, width: right - x, height: bottom - y } }]
  })
}
