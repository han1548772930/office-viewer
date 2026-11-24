import type { ChartColor, ChartLabelOptions, ChartTreemapLabelOptions, ChartTreemapOptions } from '../../domain/model'
import type { RenderRect } from '../../rendering/contracts'
import type { ChartHierarchyItem, ChartHierarchyNode } from './chart-data'
import { chartHierarchyItems } from './chart-data'
import { chartColorChannels, chartColorRepresentative, chartSolidColorOpacity, isChartLinearGradient } from './chart-color'
import { definedChartStyleFields } from './chart-options'
import { chartLabelContentMask } from './chart-labels'
import { finiteChartValue } from './chart-math'
import { formatSheetValue } from '../sheet-workbench/sheet-number-format'
import type { ChartAutomaticGradient } from './chart-auto-gradient'

export const CHART_TREEMAP_DEFAULT_FONT_SIZE = 12
export const CHART_TREEMAP_RESIDUAL_PREFIX = '__univer_treemap_residual__:'
export interface ChartTreemapRenderNode extends ChartHierarchyItem {
  id: string; name: string; value: number; semanticParentId?: string; isTreemapResidual?: boolean; children?: ChartTreemapRenderNode[]
}

export function resolveChartTreemapLabel(chart: { options?: { appearance?: { font?: ChartLabelOptions }; treemap?: ChartTreemapOptions } },
  inherited?: ChartLabelOptions): ChartTreemapLabelOptions & { visible: boolean; fontSize: number } {
  const font = chart.options?.appearance?.font, authored = chart.options?.treemap?.label
  const label = { ...definedChartStyleFields(font), ...definedChartStyleFields(inherited), ...definedChartStyleFields(authored) }
  return { ...label, contentType: label.contentType ?? 'category', visible: label.visible ?? true, fontSize: label.fontSize ?? CHART_TREEMAP_DEFAULT_FONT_SIZE }
}

export function chartTreemapParentLayout(options?: ChartTreemapOptions): NonNullable<ChartTreemapOptions['parentLabelLayout']> {
  return options?.parentLabelLayout ?? 'banner'
}

/** Overlapping text keeps its own line box, while child layout reserves no banner. */
export function chartTreemapHeaderHeight(layout: NonNullable<ChartTreemapOptions['parentLabelLayout']>, fontSize: number, padding = 8): number {
  return layout === 'none' ? 0 : layout === 'banner' ? fontSize + padding : fontSize
}

/** Semantic signed totals stay untouched. A positive parent-owned fact gets
 * an unnamed residual leaf; containers use their actual rendered child sum. */
export function chartTreemapRenderNodes(nodes: ChartHierarchyNode[]): ChartTreemapRenderNode[] {
  const items = chartHierarchyItems(nodes), rendered = new Map<ChartHierarchyNode, ChartTreemapRenderNode>()
  for (let index = items.length - 1; index >= 0; index--) {
    const item = items[index]!, node = item.node, id = node.id ?? JSON.stringify(item.path.map(name => `string:${JSON.stringify(name)}`))
    const children = (node.children ?? []).map(child => rendered.get(child)!).filter(Boolean)
    const selfValue = node.selfValue ?? node.value - (node.children ?? []).reduce((sum, child) => sum + child.value, 0)
    if (children.length && selfValue > 0) children.push({ ...item, id: CHART_TREEMAP_RESIDUAL_PREFIX + id, name: '', value: selfValue,
      depth: item.depth + 1, semanticParentId: id, isTreemapResidual: true })
    const value = children.length ? children.reduce((sum, child) => finiteChartValue(sum + child.value, '树图绘制子节点合计'), 0) : node.value
    children.sort((left, right) => right.value - left.value)
    rendered.set(node, { ...item, id, name: node.name, value, ...(children.length ? { children } : {}) })
  }
  return nodes.map(node => rendered.get(node)!).sort((left, right) => right.value - left.value)
}

/** Fixed H6 lightens #rrggbb by 12% per depth, capped at 48%. */
export function chartTreemapShade(color: ChartColor, depth: number): string {
  const solid = chartColorRepresentative(color), match = /^#([0-9a-f]{6})$/i.exec(solid)
  if (!match || depth <= 0) return solid
  const value = Number.parseInt(match[1]!, 16), ratio = Math.min(0.12 * depth, 0.48)
  const channel = (shift: number): string => Math.round((value >>> shift & 255) + (255 - (value >>> shift & 255)) * ratio).toString(16).padStart(2, '0')
  return `#${channel(16)}${channel(8)}${channel(0)}`
}

export function chartTreemapTextColor(fill: ChartColor, authored?: string, fallback = '#64748b'): string {
  if (authored) return authored
  const match = /^#([0-9a-f]{6})$/i.exec(chartColorRepresentative(fill))
  if (!match) return fallback
  const value = Number.parseInt(match[1]!, 16)
  return 0.2126 * (value >>> 16 & 255) + 0.7152 * (value >>> 8 & 255) + 0.0722 * (value & 255) > 150 ? '#000000' : '#ffffff'
}

export function chartTreemapValueText(value: number, source: string | undefined, format?: string, date1904 = false): string {
  if (format) {
    try { const formatted = formatSheetValue(value, format, date1904); if (formatted !== undefined) return formatted } catch { /* Keep source display on an unsupported mask. */ }
  }
  return source ?? String(value)
}

export function chartTreemapLabelText(node: ChartTreemapRenderNode, label: ChartTreemapLabelOptions, date1904 = false): string {
  const mask = chartLabelContentMask(label.contentType ?? 'category'), content: string[] = []
  if (!node.isTreemapResidual && mask & 2) content.push(node.node.name)
  if (mask & 8) content.push(chartTreemapValueText(node.isTreemapResidual ? node.value : node.node.value,
    node.isTreemapResidual ? node.node.formattedSelfValue : node.node.formattedValue, label.format, date1904))
  return content.join('\n')
}

/** Fixed wp represents label patterns by their mean color; the authored pattern
 * remains independent from text color and is preserved in the model. */
export function chartTreemapLabelBackground(background: ChartTreemapLabelOptions['backgroundColor']): ChartColor | undefined {
  if (background === undefined || typeof background === 'string' || isChartLinearGradient(background)) return background
  const foreground = chartColorChannels(background.foregroundColor), base = chartColorChannels(background.backgroundColor)
  const percent = /^pct(\d+)$/.exec(background.preset), ratio = percent ? Math.min(100, Number(percent[1])) / 100 : 0.5
  const channels = base.map((channel, index) => channel + (foreground[index]! - channel) * ratio)
  return `rgba(${Math.round(channels[0]! * 255)}, ${Math.round(channels[1]! * 255)}, ${Math.round(channels[2]! * 255)}, ${channels[3]})`
}

export function chartTreemapBackgroundPaint(background: ChartColor | undefined, bounds: RenderRect): ChartAutomaticGradient | undefined {
  if (!isChartLinearGradient(background)) return undefined
  return { authored: true, start: { ...background.start }, end: { ...background.end }, bounds: { ...bounds },
    stops: background.stops.map(stop => ({ offset: stop.offset, color: chartSolidColorOpacity(stop.color, stop.opacity ?? 1) })).sort((a, b) => a.offset - b.offset) }
}

export function chartTreemapLabelFits(rect: Pick<RenderRect, 'width' | 'height'> | undefined,
  labelRect: Pick<RenderRect, 'width' | 'height'> | undefined): boolean {
  return !rect || !labelRect || labelRect.width <= rect.width && labelRect.height <= rect.height
}

export interface ChartTreemapTextBlock { lines: string[]; width: number; height: number; lineHeight: number }

/** Complete content can wrap, as fixed overflow=break does, but it never
 * truncates or compresses into a smaller maxWidth. Metrics come from the renderer. */
export function chartTreemapTextBlock(text: string, rectangle: Pick<RenderRect, 'width' | 'height'>, font: string, lineHeight: number,
  measure: (text: string, font: string) => number, graphemes: (text: string) => string[] = Array.from): ChartTreemapTextBlock | undefined {
  if (!text || rectangle.width <= 0 || rectangle.height < lineHeight) return undefined
  const lines: string[] = []
  for (const sourceLine of text.split('\n')) {
    let current = ''
    for (const glyph of graphemes(sourceLine)) {
      if (measure(glyph, font) > rectangle.width) return undefined
      if (current && measure(current + glyph, font) > rectangle.width) {
        lines.push(current); current = ''
        if ((lines.length + 1) * lineHeight > rectangle.height) return undefined
      }
      current += glyph
    }
    lines.push(current)
    if (lines.length * lineHeight > rectangle.height) return undefined
  }
  const width = lines.reduce((maximum, line) => Math.max(maximum, measure(line, font)), 0), height = lines.length * lineHeight
  return chartTreemapLabelFits(rectangle, { width, height }) ? { lines, width, height, lineHeight } : undefined
}
