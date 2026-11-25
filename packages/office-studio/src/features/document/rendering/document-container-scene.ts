import type { DrawCommand, RenderRect } from '../../../rendering/contracts'
import { parseColor } from '../../../rendering/geometry'
import type { LayoutContainerDecoration } from '../layout/document-layout-types'

function roundedPoints(rect: RenderRect, radius: number): Array<{ x: number; y: number }> {
  const r = Math.min(Math.max(0, radius), rect.width / 2, rect.height / 2)
  if (!r) return [{ x: rect.x, y: rect.y }, { x: rect.x + rect.width, y: rect.y }, { x: rect.x + rect.width, y: rect.y + rect.height }, { x: rect.x, y: rect.y + rect.height }]
  const centers = [{ x: rect.x + rect.width - r, y: rect.y + r }, { x: rect.x + rect.width - r, y: rect.y + rect.height - r },
    { x: rect.x + r, y: rect.y + rect.height - r }, { x: rect.x + r, y: rect.y + r }]
  return centers.flatMap((center, corner) => Array.from({ length: 7 }, (_, index) => {
    const angle = (corner - 1) * Math.PI / 2 + index * Math.PI / 12
    return { x: center.x + Math.cos(angle) * r, y: center.y + Math.sin(angle) * r }
  }))
}

/** Rounded/dashed decoration uses primitives supported by both rendering backends. */
export function appendDocumentContainerDecoration(commands: DrawCommand[], container: LayoutContainerDecoration, rectangle: RenderRect, clip: RenderRect, scale: number): void {
  if (container.kind === 'columns') return
  if (container.kind === 'quote') {
    commands.push({ kind: 'line', x1: rectangle.x + 3 * scale, y1: rectangle.y, x2: rectangle.x + 3 * scale,
      y2: rectangle.y + rectangle.height, color: '#8796aa', width: 3 * scale, clip })
    return
  }
  const config = container.callout, points = roundedPoints(rectangle, (config?.borderRadius ?? 4) * scale)
  commands.push({ kind: 'polygon', points, fill: config?.backgroundColor ?? '#f3f5f8', clip })
  const width = (config?.borderWidth ?? 1) * scale, color = parseColor(config?.borderColor ?? '#dce3ec')
  const stroke = `rgba(${Math.round(color[0]! * 255)},${Math.round(color[1]! * 255)},${Math.round(color[2]! * 255)},${(color[3] ?? 1) * (config?.borderOpacity ?? 1)})`
  if (width > 0 && (config?.borderOpacity ?? 1) > 0) {
    const dash = config?.borderStyle === 'dashed' ? Math.max(3 * scale, width * 4) : config?.borderStyle === 'dotted' ? Math.max(scale, width) : undefined
    let travelled = 0
    for (const [index, begin] of points.entries()) {
      const end = points[(index + 1) % points.length]!, dx = end.x - begin.x, dy = end.y - begin.y, length = Math.hypot(dx, dy)
      if (!dash) commands.push({ kind: 'line', x1: begin.x, y1: begin.y, x2: end.x, y2: end.y, color: stroke, width, clip })
      else for (let offset = 0; offset < length;) {
        const phase = (travelled + offset) % (dash * 2), step = Math.min(length - offset, phase < dash ? dash - phase : dash * 2 - phase)
        if (phase < dash && step > 1e-6) commands.push({ kind: 'line', x1: begin.x + dx * offset / length, y1: begin.y + dy * offset / length,
          x2: begin.x + dx * (offset + step) / length, y2: begin.y + dy * (offset + step) / length, color: stroke, width, clip })
        offset += Math.max(step, 1e-6)
      }
      travelled += length
    }
  }
  if (config?.showIcon && config.icon && container.first) commands.push({ kind: 'text', x: rectangle.x + config.paddingLeft * scale,
    y: rectangle.y + config.paddingTop * scale, text: config.icon, font: `${config.iconSize * scale}px "Segoe UI Emoji", Arial`,
    color: '#384b63', baseline: 'top', clip })
}
