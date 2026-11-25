import type { ChartColor, ChartLabelOptions, ChartSunburstCenterLabelOptions, ChartSunburstLabelOptions, ChartSunburstOptions } from '../../domain/model'
import type { RenderRect } from '../../rendering/contracts'
import type { ChartHierarchyItem, ChartHierarchyNode } from './chart-data'
import { chartHierarchyItems } from './chart-data'
import { chartColorRepresentative } from './chart-color'
import { chartLabelContentMask } from './chart-labels'
import { finiteChartValue } from './chart-math'
import { definedChartStyleFields } from './chart-options'
import { chartTreemapValueText } from './chart-treemap'
import type { ChartVector } from './chart-geometry'

export const CHART_SUNBURST_RESIDUAL_PREFIX = '__univer_sunburst_residual__:'
export const CHART_SUNBURST_LABEL_MIN_SHARE = 0.04
export const CHART_SUNBURST_INNER_RADIUS_RATIO = 0.18
export const CHART_SUNBURST_OUTER_RADIUS_RATIO = 0.88

export interface ChartSunburstRenderNode extends ChartHierarchyItem {
  id: string; name: string; value: number; formattedValue?: string; semanticParentId?: string; isSunburstResidual?: boolean
  children?: ChartSunburstRenderNode[]
}
interface SunburstStyleOwner { options?: { appearance?: { font?: ChartLabelOptions }; sunburst?: ChartSunburstOptions } }

export function resolveChartSunburstLabel(chart: SunburstStyleOwner, inherited?: ChartLabelOptions): ChartSunburstLabelOptions & { visible: boolean; fontSize: number } {
  const authored = chart.options?.sunburst?.label
  const label = { ...definedChartStyleFields(chart.options?.appearance?.font), ...definedChartStyleFields(inherited), ...definedChartStyleFields(authored) }
  // Shared series labels can retain positions for other types. Only the fixed
  // Sunburst placement is active; omitted placement lets the renderer choose.
  const position = label.position === 'inside' || label.position === 'outside' ? label.position : undefined
  return { ...label, position, visible: label.visible ?? true, contentType: label.contentType ?? 'category', fontSize: label.fontSize ?? 12 }
}

export function resolveChartSunburstCenterLabel(chart: SunburstStyleOwner): ChartSunburstCenterLabelOptions & { visible: boolean; fontSize: number } {
  const authored = chart.options?.sunburst?.centerLabel
  return { ...definedChartStyleFields(chart.options?.appearance?.font), ...definedChartStyleFields(authored?.font), ...definedChartStyleFields(authored),
    visible: authored?.visible ?? false, fontSize: authored?.fontSize ?? authored?.font?.fontSize ?? chart.options?.appearance?.font?.fontSize ?? 16 }
}

/** Fixed qOe keeps a signed parent's semantic value and appends only positive
 * parent-owned facts as unnamed residual children. It does not replace a
 * parent's value with its rendered children sum, as Treemap does. */
export function chartSunburstRenderNodes(nodes: ChartHierarchyNode[]): ChartSunburstRenderNode[] {
  const items = chartHierarchyItems(nodes), rendered = new Map<ChartHierarchyNode, ChartSunburstRenderNode>()
  for (let index = items.length - 1; index >= 0; index--) {
    const item = items[index]!, node = item.node, children = (node.children ?? []).map(child => rendered.get(child)!).filter(Boolean)
    if (children.length && node.selfValue > 0) children.push({ ...item, depth: item.depth + 1, id: CHART_SUNBURST_RESIDUAL_PREFIX + node.id,
      name: '', value: node.selfValue, formattedValue: node.formattedSelfValue, semanticParentId: node.id, isSunburstResidual: true })
    rendered.set(node, { ...item, id: node.id, name: node.name, value: node.value, formattedValue: node.formattedValue,
      ...(children.length ? { children } : {}) })
  }
  return nodes.map(node => rendered.get(node)!)
}

export function chartSunburstTotal(nodes: readonly ChartHierarchyNode[]): number {
  return nodes.reduce((sum, node) => finiteChartValue(sum + node.value, '旭日图层级总量'), 0)
}

/** Fixed V6 lightens six-digit hex by 18% per depth, capped at 54%. */
export function chartSunburstShade(color: ChartColor, depth: number): string {
  const solid = chartColorRepresentative(color), match = /^#([0-9a-f]{6})$/i.exec(solid)
  if (!match || depth <= 0) return solid
  const value = Number.parseInt(match[1]!, 16), ratio = Math.min(0.18 * depth, 0.54)
  const channel = (shift: number): string => Math.round((value >>> shift & 255) + (255 - (value >>> shift & 255)) * ratio).toString(16).padStart(2, '0')
  return `#${channel(16)}${channel(8)}${channel(0)}`
}

export function chartSunburstLabelText(node: ChartSunburstRenderNode, label: ChartSunburstLabelOptions, date1904 = false): string {
  if (node.isSunburstResidual) return ''
  const content: string[] = [], mask = chartLabelContentMask(label.contentType ?? 'category')
  if (mask & 2) content.push(node.name)
  if (mask & 8) content.push(chartTreemapValueText(node.value, node.formattedValue, label.format, date1904))
  return content.join('\n')
}
export function chartSunburstLabelVisible(node: ChartSunburstRenderNode, total: number, label: ChartSunburstLabelOptions): boolean {
  return label.visible !== false && !node.isSunburstResidual && (total <= 0 || node.value / total >= CHART_SUNBURST_LABEL_MIN_SHARE)
}
export function chartSunburstCenterText(nodes: readonly ChartHierarchyNode[], label: ChartSunburstCenterLabelOptions, date1904 = false): string {
  const total = chartSunburstTotal(nodes)
  return chartTreemapValueText(total, String(total), label.format, date1904)
}

export interface ChartSunburstLabelSector { start: number; end: number; inner: number; outer: number; depth: number }
export interface ChartSunburstLabelAnchor extends ChartVector { rotation: number; align: 'left' | 'center' | 'right' }
export function chartSunburstLabelFits(rect: Pick<RenderRect, 'width' | 'height'> | undefined,
  labelRect: Pick<RenderRect, 'width' | 'height'> | undefined): boolean {
  return Boolean(rect && labelRect && [rect.width, rect.height, labelRect.width, labelRect.height].every(value => Number.isFinite(value) && value >= 0)
    && labelRect.width <= rect.width && labelRect.height <= rect.height)
}
function normalizedRadians(value: number): number { const result = value % (Math.PI * 2); return result < 0 ? result + Math.PI * 2 : result }

/** Mirrors the locked SunburstPiece placement and upright-flip contract. Scene
 * rotation is clockwise degrees; zrender text rotation is counterclockwise. */
export function chartSunburstLabelAnchor(center: ChartVector, sector: ChartSunburstLabelSector, position?: ChartSunburstLabelOptions['position']): ChartSunburstLabelAnchor {
  const mid = (sector.start + sector.end) / 2, tangential = sector.depth === 0, normal = normalizedRadians(tangential ? Math.PI / 2 - mid : mid)
  const flip = normal > Math.PI / 2 + 1e-8 && normal < Math.PI * 1.5
  const radius = position === 'outside' ? sector.outer + 5 : (sector.inner + sector.outer) / 2
  let rotation = -(tangential ? Math.PI / 2 - mid : -mid) * 180 / Math.PI - (flip ? 180 : 0)
  rotation = (rotation % 360 + 540) % 360 - 180
  return { x: center.x + radius * Math.cos(mid), y: center.y + radius * Math.sin(mid), rotation,
    align: position === 'outside' ? flip ? 'right' : 'left' : 'center' }
}

export function chartSunburstSectorRect(center: ChartVector, sector: ChartSunburstLabelSector): RenderRect {
  if (sector.end - sector.start >= Math.PI * 2) return { x: center.x - sector.outer, y: center.y - sector.outer, width: sector.outer * 2, height: sector.outer * 2 }
  const start = normalizedRadians(sector.start), end = start + Math.max(0, sector.end - sector.start)
  const angles = [start, end], step = Math.PI / 2
  for (let quadrant = Math.ceil(start / step); quadrant * step < end; quadrant++) angles.push(quadrant * step)
  let left = Infinity, right = -Infinity, top = Infinity, bottom = -Infinity
  for (const radius of [sector.inner, sector.outer]) for (const angle of angles) {
    const x = center.x + radius * Math.cos(angle), y = center.y + radius * Math.sin(angle)
    left = Math.min(left, x); right = Math.max(right, x); top = Math.min(top, y); bottom = Math.max(bottom, y)
  }
  return { x: left, y: top, width: right - left, height: bottom - top }
}
