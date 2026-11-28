import type { SheetDataValidationRule } from '../../domain/model.ts'
import type { ExcelCellView } from '../../services/excelViewAdapter.ts'
import { sheetCellTextAlignment } from '../../features/sheet-workbench/sheet-cell-alignment.ts'
import { layoutSheetValidationTagCell, SHEET_VALIDATION_TAG_PADDING, type SheetValidationTagOption } from '../../features/sheet-validation/validation-tag-layout.ts'
import type { DrawingAffineMatrix } from '../../domain/drawing-affine.ts'
import type { DrawCommand, RenderRect } from '../contracts.ts'
import { intersectRect } from '../geometry.ts'
import { renderRectPoints } from '../affine-command-geometry.ts'
import { measureSpreadsheetText } from './text-metrics.ts'

export interface SheetValidationTagTheme { fill: string; text: string }
function roundedTag(rect: RenderRect, radius: number, color: string, clip: RenderRect): DrawCommand {
  const { x, y, width, height } = rect, r = Math.min(radius, width / 2, height / 2)
  return { kind: 'polygon', fill: color, clip, points: [
    { x: x + r, y }, { x: x + width - r, y }, { x: x + width, y: y + r }, { x: x + width, y: y + height - r },
    { x: x + width - r, y: y + height }, { x: x + r, y: y + height }, { x, y: y + height - r }, { x, y: y + r },
  ] }
}

/** Local list values use the same CSV reader as validation and typed commits. Painting never rewrites the cell. */
export function createSheetValidationTagCommands(rule: SheetDataValidationRule | undefined, cell: ExcelCellView | undefined,
  rect: RenderRect, paneClip: RenderRect, zoom: number, theme: SheetValidationTagTheme,
  measure: (font: string, text: string) => number = (font, text) => measureSpreadsheetText(text, font),
  options?: readonly SheetValidationTagOption[]): DrawCommand[] | undefined {
  if (!rule || rule.type !== 'listMultiple' && rule.type !== 'list' || rule.renderMode === 'text') return undefined
  if (!Number.isFinite(zoom) || zoom <= 0) return []
  const layout = layoutSheetValidationTagCell(rule, cell, { width: rect.width, height: rect.height, zoom,
    measure: (text, font) => measure(font, text), options })
  if (!layout) return undefined
  const clip = intersectRect(paneClip, { x: rect.x + layout.leftInset, y: rect.y + layout.verticalInset,
    width: layout.availableWidth, height: Math.max(0, rect.height - layout.verticalInset * 2) })
  if (clip.width <= 0 || clip.height <= 0 || layout.scale <= 0 || !layout.lines.length) return []
  const align = sheetCellTextAlignment(cell), radians = layout.rotation * Math.PI / 180
  const matrix: DrawingAffineMatrix = { a: Math.cos(radians) * layout.scale, b: Math.sin(radians) * layout.scale,
    c: -Math.sin(radians) * layout.scale, d: Math.cos(radians) * layout.scale, e: 0, f: 0 }
  const corners = renderRectPoints({ x: 0, y: 0, width: layout.width, height: layout.height }, matrix)
  const minimumX = Math.min(...corners.map(point => point.x)), minimumY = Math.min(...corners.map(point => point.y))
  const left = rect.x + layout.leftInset + (align === 'center' ? (layout.availableWidth - layout.size.width) / 2
    : align === 'right' ? layout.availableWidth - layout.size.width : 0)
  const top = cell?.vertical === 'top' ? rect.y + layout.verticalInset
    : cell?.vertical === 'bottom' ? rect.y + rect.height - layout.verticalInset - layout.size.height : rect.y + (rect.height - layout.size.height) / 2
  matrix.e = left - minimumX; matrix.f = top - minimumY
  const commands: DrawCommand[] = []
  const visibleRect = (rect: RenderRect): boolean => {
    const points = renderRectPoints(rect, matrix), x = Math.min(...points.map(point => point.x)), y = Math.min(...points.map(point => point.y))
    const visible = intersectRect({ x, y, width: Math.max(...points.map(point => point.x)) - x,
      height: Math.max(...points.map(point => point.y)) - y }, clip)
    return visible.width > 0 && visible.height > 0
  }
  for (const line of layout.lines) {
    for (const tag of line.items) {
      const tagRect = { x: tag.x, y: tag.y, width: tag.width, height: tag.height }
      if (!visibleRect(tagRect)) continue
      commands.push({ ...roundedTag(tagRect, SHEET_VALIDATION_TAG_PADDING.radius * zoom, tag.color ?? theme.fill, clip), matrix })
      const localClip = { x: tag.x + layout.paddingX, y: tag.y + layout.paddingY,
        width: Math.max(0, tag.width - layout.paddingX * 2), height: Math.max(0, tag.height - layout.paddingY * 2) }
      for (const fragment of tag.fragments) {
        if (!fragment.text) continue
        const x = tag.x + fragment.x, y = tag.y + fragment.y, color = cell?.color ?? theme.text
        const start = fragment.align === 'center' ? x - fragment.width / 2 : x
        if (!visibleRect({ x: start, y: y - layout.font.fontSize / 2, width: fragment.width, height: layout.font.fontSize })) continue
        commands.push({ kind: 'text', x, y, text: fragment.text, font: layout.font.font, color, baseline: 'middle',
          align: fragment.align, matrix, localClip, clip })
        for (const offset of [cell?.underline ? layout.font.fontSize * 0.35 : undefined, cell?.strike ? 0 : undefined]) {
          if (offset !== undefined) commands.push({ kind: 'line', x1: start, x2: start + fragment.width,
            y1: y + offset, y2: y + offset, color, width: zoom, matrix, localClip, clip })
        }
      }
    }
  }
  return commands
}
