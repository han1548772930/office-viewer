import type { ChartFontOptions, ChartLabelOptions, ChartLineOptions, ChartModel, ChartSeriesFillPattern } from '../../domain/model'
import type { DrawCommand, RenderRect, TextCommand } from '../../rendering/contracts'
import type { ChartData } from './chart-engine'
import { chartPointStyle } from './chart-derived'
import { chartColorRepresentative } from './chart-color'
import { chartColorOpacity, chartFont, chartLabelText, chartPolygon, chartRect } from './chart-geometry'
import { chartFillPatternPaint } from './chart-fill-pattern'
import { chartSeriesStyle, definedChartStyleFields } from './chart-options'
import { moveChartSankeyNode, type ChartSankeyLayout, type ChartSankeyNodeLayout } from './chart-sankey-layout'
import type { ChartVector } from './chart-geometry'

export interface ChartSankeyNodePaint {
  color: string
  opacity: number
  border?: ChartLineOptions
  pattern?: ChartSeriesFillPattern
  label: ChartLabelOptions
  font: ChartFontOptions
  text: string
}
export interface ChartSankeyPaint {
  nodes: ChartSankeyNodePaint[]
  textColor: string
  labelBounds: RenderRect
}
export interface ChartSankeySceneMark {
  kind: 'node' | 'link'
  index: number
  commands: DrawCommand[]
}
/** Runtime geometry/paint only. No coordinates are written into chart options. */
export interface ChartSankeyScene {
  layout: ChartSankeyLayout
  paint: ChartSankeyPaint
  marks: ChartSankeySceneMark[]
  commands: DrawCommand[]
}

/** Fixed Gje source-node palette, inherited font and name labels. Authored
 * masks/positions/fonts are local extensions, resolved before any interaction. */
export function createChartSankeyScene(data: ChartData, chart: ChartModel, layout: ChartSankeyLayout,
  textColor: string, palette: readonly string[], labelBounds: RenderRect = layout.plot): ChartSankeyScene {
  const source = data.series[0]
  if (!source) return renderChartSankeyScene(layout, { nodes: [], textColor, labelBounds: { ...labelBounds } })
  const percentages = chartSankeyNodePercentages(layout)
  const nodes = layout.nodes.map(node => {
    const style = chartPointStyle(chartSeriesStyle(chart, source.field, node.index, palette, source.sourceName ?? source.name), node.index, chart.type)
    const label: ChartLabelOptions = { visible: true, contentType: 'category', position: 'auto', ...definedChartStyleFields(style.label) }
    const font = { ...definedChartStyleFields(chart.options?.appearance?.font), ...definedChartStyleFields(label) }
    return { color: chartColorRepresentative(style.color ?? palette[node.index % palette.length] ?? '#5470c6'),
      opacity: style.fillOpacity ?? 1, border: style.border ? { ...style.border } : undefined,
      pattern: style.fillPattern ? { ...style.fillPattern } : undefined, label, font,
      text: chartLabelText(label, node.value, node.name, style.name ?? source.name, percentages[node.index]) }
  })
  return renderChartSankeyScene(layout, { nodes, textColor, labelBounds: { ...labelBounds } })
}

/** Cached author paint is reused: dragging does not reread source cells,
 * derive data, discover SCCs or enter the content command/history pipeline. */
export function moveChartSankeyScene(scene: ChartSankeyScene, nodeIndex: number, position: ChartVector): ChartSankeyScene {
  return renderChartSankeyScene(moveChartSankeyNode(scene.layout, nodeIndex, position), scene.paint)
}

export function renderChartSankeyScene(layout: ChartSankeyLayout, paint: ChartSankeyPaint): ChartSankeyScene {
  const marks: ChartSankeySceneMark[] = [], commands: DrawCommand[] = []
  for (const link of layout.links) {
    const node = paint.nodes[link.sourceIndex], items: DrawCommand[] = []
    if (node && link.width > 0) chartPolygon(items, link.band, chartColorOpacity(node.color, 0.2 * node.opacity), undefined,
      layout.plot, chartFillPatternPaint(node.pattern, 0.2 * node.opacity))
    marks.push({ kind: 'link', index: link.index, commands: items })
    for (const command of items) commands.push(command)
  }
  const nodeMarks: ChartSankeySceneMark[] = [], labels: ChartSankeySceneMark[] = []
  for (const node of layout.nodes) {
    const style = paint.nodes[node.index], items: DrawCommand[] = [], text: DrawCommand[] = []
    if (style && node.width > 0 && node.height > 0) {
      chartRect(items, node, chartColorOpacity(style.color, style.opacity), style.border, layout.plot,
        chartFillPatternPaint(style.pattern, style.opacity))
      if (style.label.visible !== false && style.text) {
        const anchor = chartSankeyLabelAnchor(style.label.position, node, layout, style.font.fontSize ?? 10, paint.labelBounds)
        if (anchor.maxWidth > 0) text.push({ kind: 'text', ...anchor, text: style.text, font: chartFont(style.font, 10),
          color: style.label.color ?? style.font.color ?? paint.textColor })
      }
    }
    nodeMarks.push({ kind: 'node', index: node.index, commands: items })
    labels.push({ kind: 'node', index: node.index, commands: text })
  }
  // All labels remain above all rectangles and bands, including after hover.
  for (const mark of [...nodeMarks, ...labels]) { marks.push(mark); for (const command of mark.commands) commands.push(command) }
  return { layout, paint, marks, commands }
}

export type ChartSankeyLabelAnchor = Pick<TextCommand, 'x' | 'y' | 'align' | 'baseline' | 'clip'> & { maxWidth: number }
export function chartSankeyLabelAnchor(position: ChartLabelOptions['position'], node: ChartSankeyNodeLayout,
  layout: ChartSankeyLayout, fontSize = 10, bounds: RenderRect = layout.plot): ChartSankeyLabelAnchor {
  const gap = 4, center = { x: node.x + node.width / 2, y: node.y + node.height / 2 }
  const endColumn = node.x + node.width >= layout.nodeArea.x + layout.nodeArea.width - 1e-6
  const current = position === undefined || position === 'auto' || position === 'outside' ? endColumn ? 'left' : 'right' : position
  const anchor: ChartSankeyLabelAnchor = { ...center, align: 'center', baseline: 'middle', maxWidth: 0, clip: { ...bounds } }
  if (current === 'inside' || current === 'center' || current === 'insideBase' || current === 'insideEnd') {
    const inset = Math.min(gap, node.width / 2)
    anchor.clip = { x: node.x, y: node.y, width: node.width, height: node.height }
    anchor.maxWidth = Math.max(0, node.width - inset * 2)
    if (current === 'insideBase') { anchor.x = node.x + inset; anchor.align = 'left' }
    else if (current === 'insideEnd') { anchor.x = node.x + node.width - inset; anchor.align = 'right' }
  } else if (current === 'left') {
    anchor.x = node.x - gap; anchor.align = 'right'; anchor.maxWidth = Math.max(0, anchor.x - bounds.x)
  } else if (current === 'right') {
    anchor.x = node.x + node.width + gap; anchor.align = 'left'; anchor.maxWidth = Math.max(0, bounds.x + bounds.width - anchor.x)
  } else {
    anchor.y = current === 'top' ? node.y - gap - fontSize / 2 : node.y + node.height + gap + fontSize / 2
    anchor.maxWidth = Math.max(0, Math.min(anchor.x - bounds.x, bounds.x + bounds.width - anchor.x) * 2)
  }
  return anchor
}

/** A local percentage label denotes the node's share of the largest stage's
 * total throughput. Normalization avoids summing transit nodes twice or
 * overflowing disconnected finite flows. Fixed Gje only emits names. */
export function chartSankeyNodePercentages(layout: ChartSankeyLayout): number[] {
  const scale = layout.nodes.reduce((maximum, node) => Math.max(maximum, node.value), 0) || 1
  const totals = new Map<number, number>()
  for (const node of layout.nodes) totals.set(node.layer, (totals.get(node.layer) ?? 0) + node.value / scale)
  let total = 0
  for (const value of totals.values()) total = Math.max(total, value)
  return layout.nodes.map(node => total ? (node.value / scale) / total * 100 : 0)
}
