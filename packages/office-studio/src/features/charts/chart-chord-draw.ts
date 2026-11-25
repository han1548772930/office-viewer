import type { ChartFontOptions, ChartLabelOptions, ChartLineOptions, ChartModel, ChartSeriesFillPattern, ChartSeriesOptions } from '../../domain/model'
import type { DrawCommand, RenderRect, TextCommand } from '../../rendering/contracts'
import { ComputeError } from '../../compute/errors'
import type { ChartData } from './chart-engine'
import type { ChartAutomaticGradient } from './chart-auto-gradient'
import type { ChartChordLayout, ChartChordLinkLayout, ChartChordNodeLayout } from './chart-chord-layout'
import type { ChartRenderTheme } from './chart-theme'
import { chartChordLegacyLabel, resolveChartChordOptions } from './chart-chord'
import { resolveChartChordColor } from './chart-chord-color'
import { chartPointStyle } from './chart-derived'
import { chartColorRepresentative } from './chart-color'
import { chartColorOpacity, chartFont, chartLabelText, chartPolar, chartPolygon } from './chart-geometry'
import { chartFillPatternPaint } from './chart-fill-pattern'
import { chartSeriesStyle, definedChartStyleFields } from './chart-options'
import { drawChartSector } from './chart-sector-draw'
import { CHART_GRADIENT_COMMAND_BUDGET } from './chart-gradient-geometry'

export interface ChartChordNodePaint {
  color: string; opacity: number; border: ChartLineOptions; pattern?: ChartSeriesFillPattern
  label: ChartLabelOptions; font: ChartFontOptions; rotation: number; distance: number; text: string
}
export interface ChartChordSceneMark { kind: 'node' | 'link'; index: number; commands: DrawCommand[] }
export interface ChartChordScene {
  layout: ChartChordLayout; nodes: ChartChordNodePaint[]; marks: ChartChordSceneMark[]; commands: DrawCommand[]
  focus: 'adjacency' | 'self' | 'none'; weightMode: 'value' | 'count'; valueFieldName: string
}

/** Resolve authors once. Hover uses this retained paint/geometry and never
 * reads source cells, formats numbers or invokes a second layout engine. */
export function createChartChordScene(data: ChartData, chart: ChartModel, layout: ChartChordLayout,
  theme: ChartRenderTheme = {}, palette: readonly string[], labelBounds: RenderRect = layout.plot): ChartChordScene {
  const source = data.series[0], textColor = chart.options?.appearance?.font?.color ?? theme.textColor ?? '#334155'
  const options = resolveChartChordOptions(chart.options?.chord, chart.options?.appearance?.font, textColor)
  const colorTheme = { ...theme, backgroundColor: chart.options?.appearance?.background ?? theme.backgroundColor }
  const scale = layout.nodes.reduce((maximum, node) => Math.max(maximum, node.value), 0) || 1
  const total = layout.nodes.reduce((sum, node) => sum + node.value / scale, 0)
  const nodes: ChartChordNodePaint[] = layout.nodes.map(node => {
    const style: ChartSeriesOptions = source ? chartPointStyle(chartSeriesStyle(chart, source.field, node.index, palette, source.sourceName ?? source.name), node.index, chart.type) : {}
    const explicit = chart.options?.chord?.node, ownLabel = chart.options?.chord?.label
    const font: ChartFontOptions = { ...definedChartStyleFields(chart.options?.appearance?.font), ...definedChartStyleFields(style.label),
      fontSize: ownLabel?.fontSize ?? style.label?.fontSize ?? options.label.fontSize, bold: ownLabel?.bold ?? style.label?.bold ?? options.label.bold,
      italic: ownLabel?.italic ?? style.label?.italic ?? options.label.italic, family: ownLabel?.family ?? style.label?.family ?? options.label.family,
      color: resolveChartChordColor(ownLabel?.color ?? style.label?.color ?? options.label.color, colorTheme) }
    const label: ChartLabelOptions = { contentType: 'category', ...definedChartStyleFields(style.label), ...chartChordLegacyLabel(chart.options?.chord),
      ...font, visible: ownLabel?.visible ?? style.label?.visible ?? true, position: ownLabel?.position ?? options.label.position }
    const borderColor = resolveChartChordColor(explicit?.borderColor ?? style.border?.color ?? options.node.borderColor, colorTheme)
    return { color: resolveChartChordColor(explicit?.color ?? chartColorRepresentative(style.color ?? palette[node.index % palette.length] ?? '#5470c6'), colorTheme),
      opacity: explicit?.opacity ?? style.fillOpacity ?? options.node.opacity,
      border: { color: borderColor, width: Math.min(explicit?.borderWidth ?? style.border?.width ?? options.node.borderWidth, Math.max(layout.plot.width, layout.plot.height) * 2),
        opacity: style.border?.opacity ?? 1, visible: explicit?.borderWidth === undefined ? style.border?.visible : true,
        dashType: explicit?.borderType ?? style.border?.dashType ?? options.node.borderType },
      pattern: style.fillPattern ? { ...style.fillPattern } : undefined, label, font,
      rotation: options.label.rotation % 360, distance: options.label.distance,
      text: chartLabelText(label, node.value, node.name, source?.name ?? '', total ? (node.value / scale) / total * 100 : 0) }
  })
  const marks: ChartChordSceneMark[] = [], commands: DrawCommand[] = []
  const append = (mark: ChartChordSceneMark): void => {
    if (commands.length + mark.commands.length > CHART_GRADIENT_COMMAND_BUDGET) throw new ComputeError('COMPUTE_MEMORY_LIMIT', '完整和弦场景超出绘制指令预算。')
    marks.push(mark); for (const command of mark.commands) commands.push(command)
  }
  for (const link of layout.links) {
    const items: DrawCommand[] = [], sourcePaint = nodes[link.sourceIndex]!, targetPaint = nodes[link.targetIndex]!
    if (layout.innerRadius > 0 && Math.abs(link.sourceEnd - link.sourceStart) > 0 && Math.abs(link.targetEnd - link.targetStart) > 0) {
      const selected = options.ribbon.color === 'target' ? targetPaint : sourcePaint
      const fill = chartColorOpacity(selected.color, options.ribbon.opacity)
      const gradient = options.ribbon.color === 'gradient' ? chartChordRibbonGradient(layout, link, sourcePaint.color, targetPaint.color, options.ribbon.opacity) : undefined
      const border: ChartLineOptions = { color: fill, width: Math.min(options.ribbon.width, Math.max(layout.plot.width, layout.plot.height) * 2),
        dashType: options.ribbon.type, visible: options.ribbon.width > 0, opacity: 1 }
      // Texture is a retained local series extension. Public Chord's three
      // ribbon color modes stay independent of node.opacity.
      chartPolygon(items, link.band, fill, border, layout.plot, selected.pattern ? chartFillPatternPaint(selected.pattern, options.ribbon.opacity) : undefined, gradient)
    }
    append({ kind: 'link', index: link.index, commands: items })
  }
  const labels: ChartChordSceneMark[] = []
  for (const node of layout.nodes) {
    const style = nodes[node.index]!, items: DrawCommand[] = [], text: DrawCommand[] = []
    if (layout.outerRadius > layout.innerRadius && Math.abs(node.end - node.start) > 0) {
      drawChartSector(items, layout.center, layout.innerRadius, layout.outerRadius, node.start, node.end,
        chartColorOpacity(style.color, style.opacity), style.border, layout.plot, chartFillPatternPaint(style.pattern, style.opacity))
      if (style.label.visible !== false && (style.font.fontSize ?? 12) > 0 && style.text) {
        const anchor = chartChordLabelAnchor(node, layout, style.label.position, style.distance, labelBounds)
        if (anchor && anchor.maxWidth > 0) text.push({ kind: 'text', ...anchor, text: style.text, font: chartFont(style.font, 12),
          color: style.font.color!, rotation: style.rotation })
      }
    }
    append({ kind: 'node', index: node.index, commands: items }); labels.push({ kind: 'node', index: node.index, commands: text })
  }
  for (const mark of labels) append(mark)
  return { layout, nodes, marks, commands, focus: options.emphasis.focus,
    weightMode: data.chord?.weightMode ?? 'value', valueFieldName: data.chord?.valueFieldName ?? source?.name ?? '' }
}

/** Fixed ChordEdge gradient uses averages of the two arc endpoint positions,
 * in one global coordinate space, rather than a different gradient per triangle. */
export function chartChordRibbonGradient(layout: ChartChordLayout, link: ChartChordLinkLayout, sourceColor: string, targetColor: string, opacity: number): ChartAutomaticGradient {
  const sourceFirst = chartPolar(layout.center, layout.innerRadius, link.sourceStart), sourceLast = chartPolar(layout.center, layout.innerRadius, link.sourceEnd)
  const targetFirst = chartPolar(layout.center, layout.innerRadius, link.targetStart), targetLast = chartPolar(layout.center, layout.innerRadius, link.targetEnd)
  const point = (first: { x: number; y: number }, last: { x: number; y: number }) => ({
    x: ((first.x + last.x) / 2 - layout.plot.x) / Math.max(1, layout.plot.width),
    y: ((first.y + last.y) / 2 - layout.plot.y) / Math.max(1, layout.plot.height),
  })
  return { start: point(sourceFirst, sourceLast), end: point(targetFirst, targetLast), bounds: { ...layout.plot },
    stops: [{ offset: 0, color: chartColorOpacity(sourceColor, opacity) }, { offset: 1, color: chartColorOpacity(targetColor, opacity) }] }
}

export function chartChordLabelAnchor(node: ChartChordNodeLayout, layout: ChartChordLayout, position: ChartLabelOptions['position'], distance: number,
  bounds: RenderRect): (Pick<TextCommand, 'x' | 'y' | 'align' | 'baseline' | 'clip'> & { maxWidth: number }) | undefined {
  const outside = position !== 'inside' && position !== 'center', radius = outside ? layout.outerRadius + distance : (layout.outerRadius + layout.innerRadius) / 2
  const angle = (node.start + node.end) / 2, point = chartPolar(layout.center, radius, angle)
  if (![point.x, point.y].every(Number.isFinite) || point.x < bounds.x || point.x > bounds.x + bounds.width || point.y < bounds.y || point.y > bounds.y + bounds.height) return undefined
  const dx = point.x - layout.center.x, dy = point.y - layout.center.y
  const align = outside ? dx >= 0 ? 'left' : 'right' : 'center', baseline = outside ? dy >= 0 ? 'top' : 'bottom' : 'middle'
  const maxWidth = align === 'left' ? bounds.x + bounds.width - point.x : align === 'right' ? point.x - bounds.x : 2 * Math.min(point.x - bounds.x, bounds.x + bounds.width - point.x)
  return { ...point, align, baseline, maxWidth: Math.max(0, maxWidth), clip: { ...bounds } }
}
