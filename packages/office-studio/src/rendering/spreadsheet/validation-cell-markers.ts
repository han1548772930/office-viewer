import type { DrawCommand, RenderRect } from '../contracts.ts'
import { intersectRect } from '../geometry.ts'

/** The validation corner mark is editor content. Replaceable popup controls remain Ant components. */
export function createSheetValidationInvalidMarker(rect: RenderRect, clip: RenderRect, zoom: number, color: string): DrawCommand[] {
  const size = Math.min(6 * zoom, rect.width, rect.height)
  if (!Number.isFinite(size) || size <= 0) return []
  const marker = { x: rect.x + rect.width - size, y: rect.y, width: size, height: size }
  const visible = intersectRect(marker, clip)
  return visible.width > 0 && visible.height > 0 ? [{ kind: 'polygon', fill: color, clip,
    points: [{ x: marker.x, y: marker.y }, { x: marker.x + size, y: marker.y }, { x: marker.x + size, y: marker.y + size }] }] : []
}
