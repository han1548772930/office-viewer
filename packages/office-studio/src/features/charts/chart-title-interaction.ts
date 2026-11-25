import type { ChartTitleManualLayout } from '../../domain/model'
import type { DrawCommand, RenderRect } from '../../rendering/contracts'
import type { ChartVector } from './chart-geometry'
import type { ChartTitleKind, ChartTitlePaintLayout } from './chart-titles'

export interface ChartMeasuredTitleText {
  layout: ChartTitlePaintLayout
  polygon: ChartVector[]
  /** Leading offset measured by ZRender, in chart coordinates. */
  leading?: ChartVector
}
export interface ChartCanvasTitleTarget {
  kind: ChartTitleKind
  anchor: ChartVector
  polygons: ChartVector[][]
  clip: RenderRect
  bounds: RenderRect
}

/** Read only our authored graphic metadata; titles are never inferred from
 * the displayed string or a chart/series label with the same text. */
export function chartTitleGraphicLayouts(elements: unknown): Map<string, ChartTitlePaintLayout> {
  const result = new Map<string, ChartTitlePaintLayout>()
  const record = (value: unknown): Record<string, unknown> | undefined => value && typeof value === 'object' && !Array.isArray(value) ? value as Record<string, unknown> : undefined
  const visit = (value: unknown): void => {
    if (!Array.isArray(value)) return
    for (const item of value) {
      const element = record(item)
      if (!element) continue
      visit(element.children)
      const layout = record(record(element.info)?.chartTitle), anchor = record(layout?.anchor), clip = record(layout?.clip)
      if (typeof element.id !== 'string' || !layout || !anchor || !clip || (layout.kind !== 'title' && layout.kind !== 'subtitle')
        || typeof layout.lineIndex !== 'number' || !Number.isSafeInteger(layout.lineIndex) || layout.lineIndex < 0
        || typeof layout.manual !== 'boolean' || typeof layout.vertical !== 'boolean'
        || (layout.alignment !== 'left' && layout.alignment !== 'center' && layout.alignment !== 'right')
        || !finite(anchor.x) || !finite(anchor.y) || !finite(clip.x) || !finite(clip.y) || !finite(clip.width) || !finite(clip.height)
        || clip.width <= 0 || clip.height <= 0) continue
      result.set(element.id, { kind: layout.kind, lineIndex: layout.lineIndex, manual: layout.manual, vertical: layout.vertical, alignment: layout.alignment,
        anchor: { x: anchor.x, y: anchor.y }, clip: { x: clip.x, y: clip.y, width: clip.width, height: clip.height } })
    }
  }
  visit(elements)
  return result
}

function finite(value: unknown): value is number { return typeof value === 'number' && Number.isFinite(value) }

/** Clip real transformed text bounds before hit testing or outlining. A title
 * outside its chart frame cannot steal input from a neighbouring chart. */
export function clipChartTitlePolygon(polygon: readonly ChartVector[], clip: RenderRect): ChartVector[] {
  if (!polygon.length || !polygon.every(point => finite(point.x) && finite(point.y))) return []
  let output = polygon.map(point => ({ ...point }))
  for (const [axis, boundary, lower] of [['x', clip.x, true], ['x', clip.x + clip.width, false], ['y', clip.y, true], ['y', clip.y + clip.height, false]] as const) {
    const input = output; output = []
    for (let index = 0; index < input.length; index++) {
      const from = input[index]!, to = input[(index + 1) % input.length]!
      const fromInside = lower ? from[axis] >= boundary : from[axis] <= boundary, toInside = lower ? to[axis] >= boundary : to[axis] <= boundary
      if (fromInside) output.push(from)
      if (fromInside !== toInside) {
        const fraction = (boundary - from[axis]) / (to[axis] - from[axis])
        output.push({ x: from.x + (to.x - from.x) * fraction, y: from.y + (to.y - from.y) * fraction })
      }
    }
  }
  return output
}

export function createChartTitleTargets(texts: readonly ChartMeasuredTitleText[]): ChartCanvasTitleTarget[] {
  const targets: ChartCanvasTitleTarget[] = []
  for (const kind of ['title', 'subtitle'] as const) {
    const lines = texts.filter(text => text.layout.kind === kind).sort((a, b) => a.layout.lineIndex - b.layout.lineIndex)
    const first = lines[0]
    if (!first) continue
    const polygons = lines.map(line => clipChartTitlePolygon(line.polygon, line.layout.clip)).filter(polygon => polygon.length >= 3)
    let left = Infinity, top = Infinity, right = -Infinity, bottom = -Infinity
    for (const polygon of polygons) for (const point of polygon) {
      left = Math.min(left, point.x); top = Math.min(top, point.y); right = Math.max(right, point.x); bottom = Math.max(bottom, point.y)
    }
    if (!Number.isFinite(left)) continue
    const bounds = { x: left, y: top, width: right - left, height: bottom - top }
    if (bounds.width <= 0 || bounds.height <= 0) continue
    // Fixed C5 renders manually positioned horizontal text left-aligned. The
    // real leading edge anchors the first drag instead of estimating a width.
    const shift = !first.layout.manual && !first.layout.vertical && first.layout.alignment !== 'left' ? first.leading : undefined
    targets.push({ kind, polygons, bounds, clip: { ...first.layout.clip }, anchor: {
      x: first.layout.anchor.x + (shift?.x ?? 0), y: first.layout.anchor.y + (shift?.y ?? 0),
    } })
  }
  return targets
}

function insidePolygon(polygon: readonly ChartVector[], point: ChartVector): boolean {
  let inside = false
  for (let index = 0; index < polygon.length; index++) {
    const a = polygon[index]!, b = polygon[(index + 1) % polygon.length]!
    const cross = (point.x - a.x) * (b.y - a.y) - (point.y - a.y) * (b.x - a.x)
    if (Math.abs(cross) < 1e-7 && point.x >= Math.min(a.x, b.x) - 1e-7 && point.x <= Math.max(a.x, b.x) + 1e-7
      && point.y >= Math.min(a.y, b.y) - 1e-7 && point.y <= Math.max(a.y, b.y) + 1e-7) return true
    if ((a.y > point.y) !== (b.y > point.y) && point.x < (b.x - a.x) * (point.y - a.y) / (b.y - a.y) + a.x) inside = !inside
  }
  return inside
}

export function hitChartTitle(targets: readonly ChartCanvasTitleTarget[], point?: ChartVector): ChartCanvasTitleTarget | undefined {
  if (!point || !finite(point.x) || !finite(point.y)) return undefined
  for (let index = targets.length - 1; index >= 0; index--) {
    const target = targets[index]!
    if (target.polygons.some(polygon => insidePolygon(polygon, point))) return target
  }
  return undefined
}

export function chartTitleOutlineCommands(target: ChartCanvasTitleTarget, position: ChartVector, color: string, width = 1): DrawCommand[] {
  const clip = { ...target.clip, x: target.clip.x + position.x, y: target.clip.y + position.y }
  return target.polygons.flatMap(polygon => polygon.map((point, index) => {
    const next = polygon[(index + 1) % polygon.length]!
    return { kind: 'line' as const, x1: position.x + point.x, y1: position.y + point.y,
      x2: position.x + next.x, y2: position.y + next.y, color, width, clip }
  }))
}

export function chartTitleDragLayout(target: ChartCanvasTitleTarget, size: { width: number; height: number }, delta: ChartVector): ChartTitleManualLayout {
  if (![size.width, size.height, target.anchor.x, target.anchor.y, delta.x, delta.y].every(Number.isFinite) || size.width <= 0 || size.height <= 0) throw new RangeError('标题拖动需要有效尺寸和坐标。')
  return { left: Math.min(1, Math.max(0, (target.anchor.x + delta.x) / size.width)), top: Math.min(1, Math.max(0, (target.anchor.y + delta.y) / size.height)) }
}

export interface ChartTitleAuthorContext { owner: object; workbookId: string; sheetId: string; chartId: string; revision: number; snapshot: string }
export function chartTitleAuthorContextMatches(original: ChartTitleAuthorContext, current?: ChartTitleAuthorContext): boolean {
  return !!current && original.owner === current.owner && original.workbookId === current.workbookId && original.sheetId === current.sheetId
    && original.chartId === current.chartId && original.revision === current.revision && original.snapshot === current.snapshot
}
