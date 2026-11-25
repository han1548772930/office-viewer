import type { ChartFontOptions, ChartLabelOptions, ChartLineOptions, ChartModel, ChartPointShape, ChartSeriesFillPattern, ChartSeriesOptions } from '../../domain/model'
import type { DrawCommand, RenderRect, TextCommand } from '../../rendering/contracts'
import { ComputeError } from '../../compute/errors'
import type { ChartData } from './chart-engine'
import type { ChartRelationGraphLayout, ChartRelationNodeLayout } from './chart-relation-layout'
import type { ChartRenderTheme } from './chart-theme'
import { CHART_RELATION_DEFAULTS } from './chart-relation'
import { chartPointStyle } from './chart-derived'
import { chartColorRepresentative } from './chart-color'
import { chartBezier, chartColorOpacity, chartFont, chartLabelText, chartLine, chartSymbol, formatChartNumber, type ChartVector } from './chart-geometry'
import { chartMarkerPoints } from './chart-markers'
import { chartFillPatternPaint } from './chart-fill-pattern'
import { chartSeriesStyle, definedChartStyleFields } from './chart-options'
import { CHART_GRADIENT_COMMAND_BUDGET } from './chart-gradient-geometry'

export interface ChartRelationNodePaint {
  color: string; opacity: number; border: ChartLineOptions; shape: ChartPointShape; visible: boolean
  pattern?: ChartSeriesFillPattern; label: ChartLabelOptions; font: ChartFontOptions; text: string; formattedValue: string
}
export interface ChartRelationSceneMark { kind: 'node' | 'link'; index: number; commands: DrawCommand[] }
export interface ChartRelationScene {
  layout: ChartRelationGraphLayout; nodes: ChartRelationNodePaint[]; marks: ChartRelationSceneMark[]; commands: DrawCommand[]
  emphasisEnabled: boolean; valueFieldName: string; linkFormattedValues: string[]
}

/** Oje/HOe's category palette, source-colored curved edges, name labels and
 * explicit node symbol all use the retained Worker coordinate set. Local
 * series/point paint extensions pass through the same command scene. */
export function createChartRelationScene(data: ChartData, chart: ChartModel, layout: ChartRelationGraphLayout,
  theme: ChartRenderTheme = {}, palette: readonly string[], labelBounds: RenderRect = layout.plot): ChartRelationScene {
  const source = data.series[0], textColor = chart.options?.appearance?.font?.color ?? theme.textColor ?? '#334155'
  let maximum = 0, total = 0
  for (const node of layout.nodes) maximum = Math.max(maximum, node.value)
  for (const node of layout.nodes) total += maximum ? node.value / maximum : 0
  const nodes: ChartRelationNodePaint[] = layout.nodes.map(node => {
    const style: ChartSeriesOptions = source ? chartPointStyle(chartSeriesStyle(chart, source.field, node.colorIndex, palette, source.sourceName ?? source.name), node.index, chart.type) : {}
    const label: ChartLabelOptions = { visible: true, contentType: 'category', position: 'right', ...definedChartStyleFields(style.label) }
    const font: ChartFontOptions = { ...definedChartStyleFields(chart.options?.appearance?.font), ...definedChartStyleFields(label) }
    return { color: chartColorRepresentative(style.point?.color ?? style.color ?? palette[node.colorIndex % palette.length] ?? '#5470c6'),
      opacity: style.fillOpacity ?? 1, border: { ...style.border },
      shape: chart.options?.relation?.nodeShape ?? style.point?.shape ?? CHART_RELATION_DEFAULTS.nodeShape,
      visible: style.point?.visible !== false, pattern: style.fillPattern ? { ...style.fillPattern } : undefined, label, font,
      text: chartLabelText(label, node.value, node.name, style.name ?? source?.name ?? '', total ? (node.value / maximum) / total * 100 : 0),
      formattedValue: formatChartNumber(node.value, label.format) }
  })
  const commands: DrawCommand[] = [], marks: ChartRelationSceneMark[] = [], linkFormattedValues: string[] = []
  const append = (mark: ChartRelationSceneMark): void => {
    if (commands.length + mark.commands.length > CHART_GRADIENT_COMMAND_BUDGET) throw new ComputeError('COMPUTE_MEMORY_LIMIT', '完整关系图场景超出绘制指令预算。')
    marks.push(mark); for (const command of mark.commands) commands.push(command)
  }
  for (const link of layout.links) {
    const first = layout.nodes[link.sourceIndex]!, second = layout.nodes[link.targetIndex]!, paint = nodes[link.sourceIndex]!
    const firstPaint = nodes[link.sourceIndex]!, secondPaint = nodes[link.targetIndex]!
    const points = link.selfLoop ? link.points : chartRelationCurve(first, second,
      firstPaint.visible ? chartMarkerPoints(first, first.symbolSize, firstPaint.shape) : [],
      secondPaint.visible ? chartMarkerPoints(second, second.symbolSize, secondPaint.shape) : [])
    const items: DrawCommand[] = [], edge: ChartLineOptions = { color: paint.color, opacity: 0.5 * paint.opacity, width: link.width }
    for (let index = 1; index < points.length; index++) chartLine(items, points[index - 1]!, points[index]!, edge, paint.color, layout.plot)
    linkFormattedValues.push(formatChartNumber(link.value, paint.label.format))
    append({ kind: 'link', index: link.index, commands: items })
  }
  const labels: ChartRelationSceneMark[] = []
  for (const node of layout.nodes) {
    const paint = nodes[node.index]!, items: DrawCommand[] = [], text: DrawCommand[] = []
    if (paint.visible) chartSymbol(items, node, node.symbolSize, chartColorOpacity(paint.color, paint.opacity), paint.shape,
      paint.border, layout.plot, chartFillPatternPaint(paint.pattern, paint.opacity))
    if (paint.label.visible !== false && paint.text) {
      const anchor = chartRelationLabelAnchor(node, layout, paint.label.position, labelBounds,
        chart.options?.relation?.layout === 'circular' && (chart.options?.relation?.circularLabelRotation ?? CHART_RELATION_DEFAULTS.circularLabelRotation))
      if (anchor.maxWidth > 0) text.push({ kind: 'text', ...anchor, text: paint.text, font: chartFont(paint.font, 12), color: paint.font.color ?? textColor })
    }
    append({ kind: 'node', index: node.index, commands: items }); labels.push({ kind: 'node', index: node.index, commands: text })
  }
  // Labels stay above all symbols/edges, including after adjacency emphasis.
  for (const mark of labels) append(mark)
  return { layout, nodes, marks, commands, linkFormattedValues,
    emphasisEnabled: chart.options?.relation?.emphasisEnabled ?? CHART_RELATION_DEFAULTS.emphasisEnabled,
    valueFieldName: source?.name ?? '权重' }
}

/** HOe uses .3 curveness. Trim both endpoints to the submitted marker polygons
 * so a hollow symbol's unpainted center does not inherit the edge stroke. */
export function chartRelationCurve(first: ChartVector, second: ChartVector,
  firstPolygon: readonly ChartVector[] = [], secondPolygon: readonly ChartVector[] = []): ChartVector[] {
  const dx = second.x - first.x, dy = second.y - first.y
  if (Math.hypot(dx, dy) < 1e-9) return []
  const control = { x: (first.x + second.x) / 2 + dy * 0.3, y: (first.y + second.y) / 2 - dx * 0.3 }
  const points = chartBezier(first, { x: first.x + (control.x - first.x) * 2 / 3, y: first.y + (control.y - first.y) * 2 / 3 },
    { x: second.x + (control.x - second.x) * 2 / 3, y: second.y + (control.y - second.y) * 2 / 3 }, second)
  const start = curveExit(points, firstPolygon, false), end = curveExit(points, secondPolygon, true)
  if (!start || !end || start.segment > end.segment) return []
  const result = [start.point]
  for (let index = start.segment + 1; index <= end.segment; index++) result.push(points[index]!)
  result.push(end.point)
  return result
}

function curveExit(points: readonly ChartVector[], polygon: readonly ChartVector[], reverse: boolean): { segment: number; point: ChartVector } | undefined {
  if (!polygon.length) return reverse ? { segment: points.length - 2, point: points.at(-1)! } : { segment: 0, point: points[0]! }
  for (let offset = 0; offset < points.length - 1; offset++) {
    const segment = reverse ? points.length - 2 - offset : offset
    const from = points[reverse ? segment + 1 : segment]!, to = points[reverse ? segment : segment + 1]!
    let crossing = Infinity
    for (let cursor = 0; cursor < polygon.length; cursor++) {
      const a = polygon[cursor]!, b = polygon[(cursor + 1) % polygon.length]!
      const dx = to.x - from.x, dy = to.y - from.y, ex = b.x - a.x, ey = b.y - a.y, cross = dx * ey - dy * ex
      if (Math.abs(cross) < 1e-12) continue
      const t = ((a.x - from.x) * ey - (a.y - from.y) * ex) / cross
      const u = ((a.x - from.x) * dy - (a.y - from.y) * dx) / cross
      if (t >= -1e-9 && t <= 1 + 1e-9 && u >= -1e-9 && u <= 1 + 1e-9) crossing = Math.min(crossing, Math.max(0, Math.min(1, t)))
    }
    if (crossing !== Infinity) return { segment, point: { x: from.x + (to.x - from.x) * crossing, y: from.y + (to.y - from.y) * crossing } }
  }
  return undefined
}

export function chartRelationLabelAnchor(node: ChartRelationNodeLayout, layout: ChartRelationGraphLayout,
  position: ChartLabelOptions['position'], bounds: RenderRect, rotateCircular = false): Pick<TextCommand, 'x' | 'y' | 'align' | 'baseline' | 'rotation' | 'clip'> & { maxWidth: number } {
  const radius = node.symbolSize / 2, gap = radius + 5, current = position ?? 'right'
  const anchor: Pick<TextCommand, 'x' | 'y' | 'align' | 'baseline' | 'rotation' | 'clip'> & { maxWidth: number } = {
    x: node.x, y: node.y, align: 'center', baseline: 'middle', clip: { ...bounds }, maxWidth: 0 }
  if (rotateCircular && ['right', 'auto', 'outside'].includes(current)) {
    const angle = Math.atan2(node.y - layout.center.y, node.x - layout.center.x), dx = Math.cos(angle), dy = Math.sin(angle)
    anchor.x += dx * gap; anchor.y += dy * gap
    const left = dx < 0
    anchor.align = left ? 'right' : 'left'; anchor.rotation = -(angle + (left ? Math.PI : 0)) * 180 / Math.PI
    const horizontal = Math.abs(dx) < 1e-9 ? Infinity : (dx > 0 ? bounds.x + bounds.width - anchor.x : anchor.x - bounds.x) / Math.abs(dx)
    const vertical = Math.abs(dy) < 1e-9 ? Infinity : (dy > 0 ? bounds.y + bounds.height - anchor.y : anchor.y - bounds.y) / Math.abs(dy)
    anchor.maxWidth = Math.max(0, Math.min(horizontal, vertical))
  } else if (current === 'inside' || current === 'center' || current === 'insideBase' || current === 'insideEnd') {
    anchor.maxWidth = Math.max(0, node.symbolSize - 4)
  } else if (current === 'left') {
    anchor.x -= gap; anchor.align = 'right'; anchor.maxWidth = Math.max(0, anchor.x - bounds.x)
  } else if (current === 'top' || current === 'bottom') {
    anchor.y += current === 'top' ? -gap : gap; anchor.baseline = current === 'top' ? 'bottom' : 'top'
    anchor.maxWidth = Math.max(0, 2 * Math.min(anchor.x - bounds.x, bounds.x + bounds.width - anchor.x))
  } else {
    anchor.x += gap; anchor.align = 'left'; anchor.maxWidth = Math.max(0, bounds.x + bounds.width - anchor.x)
  }
  if (anchor.x < bounds.x || anchor.x > bounds.x + bounds.width || anchor.y < bounds.y || anchor.y > bounds.y + bounds.height) anchor.maxWidth = 0
  return anchor
}

export interface ChartRelationViewport { scale: number; x: number; y: number }
/** View-only scale is confined to .4..2. Authored geometry, source records,
 * prepared Rust coordinates and content history remain with their owners. */
export function transformChartRelationScene(scene: ChartRelationScene, viewport: ChartRelationViewport): ChartRelationScene {
  const { scale, x, y } = viewport
  if (![scale, x, y].every(Number.isFinite) || scale < 0.4 || scale > 2) throw new RangeError('关系图视图缩放无效。')
  if (scale === 1 && x === 0 && y === 0) return scene
  const point = (item: ChartVector) => ({ x: item.x * scale + x, y: item.y * scale + y })
  const command = (item: DrawCommand): DrawCommand => {
    if (item.kind === 'line') { const first = point({ x: item.x1, y: item.y1 }), second = point({ x: item.x2, y: item.y2 }); return {
      ...item, x1: first.x, y1: first.y, x2: second.x, y2: second.y, width: (item.width ?? 1) * scale } }
    if (item.kind === 'polygon') return { ...item, points: item.points.map(point) }
    if (item.kind === 'rect') return { ...item, ...point(item), width: item.width * scale, height: item.height * scale,
      ...(item.strokeWidth === undefined ? {} : { strokeWidth: item.strokeWidth * scale }) }
    if (item.kind === 'text') return { ...item, ...point(item), maxWidth: item.maxWidth === undefined ? undefined : item.maxWidth * scale,
      font: item.font.replace(/([\d.]+)px /, (_match, size: string) => `${Number(size) * scale}px `) }
    return { ...item, ...point(item), width: item.width * scale, height: item.height * scale }
  }
  const marks = scene.marks.map(mark => ({ ...mark, commands: mark.commands.map(command) }))
  return { ...scene, marks, commands: marks.flatMap(mark => mark.commands), layout: { ...scene.layout,
    center: point(scene.layout.center), nodes: scene.layout.nodes.map(node => ({ ...node, ...point(node),
      symbolSize: node.symbolSize * scale, width: node.width * scale, height: node.height * scale })),
    links: scene.layout.links.map(link => ({ ...link, width: link.width * scale, points: link.points.map(point) })) } }
}
