import type { ChartAxisOptions, ChartAxisPointerOptions, ChartFontOptions, ChartModel } from '../../domain/model'
import type { DrawCommand, RenderRect } from '../../rendering/contracts'
import type { ChartData, ChartLayout } from './chart-engine'
import { resolveChartType } from './chart-catalog'
import { normalizeChartColor } from './chart-color'
import { chartDateAxisLabel } from './chart-date-axis'
import { chartFont, chartLine, chartPolar, formatChartNumber, type ChartVector } from './chart-geometry'
import { chartInterpolate } from './chart-math'
import { chartPalette, definedChartStyleFields } from './chart-options'
import { chartRadarGeometry } from './chart-radar'
import type { ChartRenderTheme } from './chart-theme'
import { createChartHeatmapPointerGeometry, chartHeatmapPointerCommands, type ChartHeatmapPointerGeometry } from './chart-heatmap-interaction'
import { hitChartHeatmapCell } from './chart-heatmap'

export interface ChartAxisPointerPresentation extends Required<ChartAxisPointerOptions> {
  font: ChartFontOptions
}

interface PointerCategory { position: number; index: number }
interface RadarPointerGeometry {
  center: ChartVector
  radius: number
  min: number
  max: number
  vertices: ChartVector[]
  ring: ChartVector[]
}

/** Cached view geometry. It is owned by the raster surface and never enters a snapshot. */
export interface ChartAxisPointerGeometry {
  chart: ChartModel
  data: ChartData
  layout: ChartLayout
  presentation: ChartAxisPointerPresentation
  categories: PointerCategory[]
  radar?: RadarPointerGeometry
  heatmap?: ChartHeatmapPointerGeometry
}

export type ChartAxisPointerRuntime = { mode: 'native' } | { mode: 'graphic'; geometry: ChartAxisPointerGeometry }

/** CSS viewport coordinates enter the surface once; renderer DPR is a separate transform. */
export function chartAxisPointerLocalPoint(point: ChartVector, position: ChartVector, view: { zoom: number; scrollLeft: number; scrollTop: number }): ChartVector {
  const scale = Math.max(0.1, view.zoom / 100)
  return { x: (point.x + view.scrollLeft) / scale - position.x, y: (point.y + view.scrollTop) / scale - position.y }
}

/** Fixed chart-ui v1.6.0 qi/Hi IndicatorLine capability; 3D excludes the entry. */
export function supportsChartAxisPointer(chart: Pick<ChartModel, 'type'>): boolean {
  const descriptor = resolveChartType(chart.type)
  return !descriptor.threeD && ['column', 'bar', 'line', 'area', 'combination', 'scatter', 'bubble', 'radar', 'waterfall', 'heatmap', 'boxplot', 'pareto'].includes(descriptor.family)
}

export function chartAxisPointerPresentation(chart: ChartModel, theme: ChartRenderTheme = {}): ChartAxisPointerPresentation {
  const authored = chart.options?.axisPointer
  return {
    indicatorLineColor: normalizeChartColor(authored?.indicatorLineColor ?? chart.options?.appearance?.font?.color ?? theme.textColor ?? '#334155'),
    indicatorLineType: authored?.indicatorLineType ?? 'solid',
    indicatorLabelColor: normalizeChartColor(authored?.indicatorLabelColor ?? theme.axisPointerLabelBackground ?? '#64748b'),
    indicatorLabelTextColor: normalizeChartColor(authored?.indicatorLabelTextColor ?? theme.axisPointerLabelColor ?? '#ffffff'),
    font: { fontSize: 10, ...definedChartStyleFields(chart.options?.appearance?.font) },
  }
}

/** Build once with the content frame; hover only queries cached positions and emits guides. */
export function chartAxisPointerGeometry(data: ChartData, chart: ChartModel, layout: ChartLayout, theme: ChartRenderTheme = {}): ChartAxisPointerGeometry {
  const categories: PointerCategory[] = []
  if (layout.xMin === undefined || layout.dateAxis) {
    const count = chart.type === 'heatmap' ? data.series.length : layout.categories.length
    for (let index = 0; index < count; index++) {
      const position = layout.dateAxis ? layout.categoryPositions?.[index]
        : chart.xAxis.reverse ? 1 - (index + 0.5) / count : (index + 0.5) / count
      if (position != null && Number.isFinite(position) && position >= 0 && position <= 1) categories.push({ position, index })
    }
    categories.sort((a, b) => a.position - b.position || a.index - b.index)
  }
  const geometry: ChartAxisPointerGeometry = { chart, data, layout, categories, presentation: chartAxisPointerPresentation(chart, theme) }
  if (chart.type === 'heatmap') geometry.heatmap = createChartHeatmapPointerGeometry(data, chart, layout, chartPalette(chart, theme.palette),
    chart.options?.appearance?.font?.color ?? theme.textColor ?? '#334155')
  if (chart.type === 'radar') {
    const radar = chartRadarGeometry(data, chart, layout.plot), vertices = data.categories.map((_category, index) => radar.vertex(index, 1))
    const ring = chart.options?.radar?.shape === 'circle'
      ? Array.from({ length: 36 }, (_value, index) => chartPolar(radar.center, radar.radius, index * Math.PI / 18)) : vertices
    geometry.radar = { center: radar.center, radius: radar.radius, min: radar.min, max: radar.max, vertices, ring }
  }
  return geometry
}

/** Indicator labels and lines use the same logical frame as the chart; no data preparation runs here. */
export function chartAxisPointerCommands(geometry: ChartAxisPointerGeometry, pointer?: ChartVector): DrawCommand[] {
  const { chart, layout } = geometry
  if (!pointer || !Number.isFinite(pointer.x) || !Number.isFinite(pointer.y) || !supportsChartAxisPointer(chart)
    || !geometry.data.categories.length || !geometry.data.series.length) return []
  const heatmapCommands = geometry.heatmap ? chartHeatmapPointerCommands(geometry.heatmap, pointer) : []
  if (!inside(pointer, layout.plot)) return heatmapCommands
  if (geometry.radar) return radarPointerCommands(geometry, pointer)
  const { plot } = layout, horizontal = Boolean(layout.horizontal), commands: DrawCommand[] = []
  const baseRatio = horizontal ? (pointer.y - plot.y) / plot.height : (pointer.x - plot.x) / plot.width
  let basePosition = baseRatio, baseText: string
  if (layout.xMin === undefined || layout.dateAxis) {
    const category = nearestCategory(geometry.categories, baseRatio)
    if (!category) return []
    basePosition = category.position
    const date = geometry.data.categoryDateValues?.[category.index]
    baseText = geometry.heatmap ? geometry.heatmap.fieldNames[category.index] ?? '' : layout.dateAxis && date != null
      ? chartDateAxisLabel(date, chart.xAxis, layout.date1904, (layout.dateSourceMax ?? 1) - (layout.dateSourceMin ?? 0)) : layout.categories[category.index] ?? ''
  } else baseText = formatChartNumber(axisValue(baseRatio, layout.xMin, layout.xMax!, chart.xAxis), chart.xAxis.format)
  const baseCoordinate = horizontal ? plot.y + basePosition * plot.height : plot.x + basePosition * plot.width
  guide(commands, horizontal ? { x: plot.x, y: baseCoordinate } : { x: baseCoordinate, y: plot.y },
    horizontal ? { x: plot.x + plot.width, y: baseCoordinate } : { x: baseCoordinate, y: plot.y + plot.height }, geometry)
  pointerLabel(commands, baseText, horizontal ? { x: plot.x, y: baseCoordinate } : { x: baseCoordinate, y: plot.y + plot.height }, horizontal ? 'left' : 'below', geometry)
  if (chart.type === 'heatmap') {
    const cell = hitChartHeatmapCell(geometry.data, chart, plot, pointer)!
    const count = geometry.data.categories.length, visibleRow = chart.yAxis.reverse ? cell.categoryIndex : count - cell.categoryIndex - 1
    const y = plot.y + (visibleRow + 0.5) * plot.height / count
    guide(commands, { x: plot.x, y }, { x: plot.x + plot.width, y }, geometry)
    pointerLabel(commands, geometry.data.categories[cell.categoryIndex] ?? '', { x: plot.x, y }, 'left', geometry)
    return [...commands, ...heatmapCommands]
  }
  const valueRatio = horizontal ? (pointer.x - plot.x) / plot.width : 1 - (pointer.y - plot.y) / plot.height
  const coordinate = horizontal ? pointer.x : pointer.y, value = axisValue(valueRatio, layout.min, layout.max, chart.yAxis)
  const suffix = resolveChartType(chart.type).percent && !chart.yAxis.format ? '%' : ''
  guide(commands, horizontal ? { x: coordinate, y: plot.y } : { x: plot.x, y: coordinate },
    horizontal ? { x: coordinate, y: plot.y + plot.height } : { x: plot.x + plot.width, y: coordinate }, geometry)
  pointerLabel(commands, `${formatChartNumber(value, chart.yAxis.format)}${suffix}`,
    horizontal ? { x: coordinate, y: plot.y + plot.height } : { x: plot.x, y: coordinate }, horizontal ? 'below' : 'left', geometry)
  if (layout.rightMin !== undefined && layout.rightMax !== undefined) {
    const axis = chart.options?.rightYAxis ?? { scaleType: 'linear' }, value = axisValue(valueRatio, layout.rightMin, layout.rightMax, axis)
    pointerLabel(commands, `${formatChartNumber(value, axis.format)}${chart.type === 'pareto' && !axis.format ? '%' : ''}`,
      horizontal ? { x: coordinate, y: plot.y } : { x: plot.x + plot.width, y: coordinate }, horizontal ? 'above' : 'right', geometry)
  }
  return commands
}

function radarPointerCommands(geometry: ChartAxisPointerGeometry, pointer: ChartVector): DrawCommand[] {
  const radar = geometry.radar!, { chart, layout, data } = geometry, count = radar.vertices.length
  if (count < 3 || radar.radius <= 0) return []
  const dx = pointer.x - radar.center.x, dy = pointer.y - radar.center.y, distance = Math.hypot(dx, dy)
  const step = Math.PI * 2 / count, angle = Math.atan2(dy, dx), logical = ((angle + Math.PI / 2) * (chart.xAxis.reverse ? -1 : 1) % (Math.PI * 2) + Math.PI * 2) % (Math.PI * 2)
  const boundary = chart.options?.radar?.shape === 'circle' ? radar.radius
    : radar.radius * Math.cos(step / 2) / Math.cos(logical % step - step / 2)
  if (distance > boundary + 1e-8) return []
  const index = Math.round(logical / step) % count, end = radar.vertices[index]!, ratio = Math.min(1, distance / radar.radius), commands: DrawCommand[] = []
  guide(commands, radar.center, end, geometry)
  const ring = radar.ring.map(point => ({ x: radar.center.x + (point.x - radar.center.x) * ratio, y: radar.center.y + (point.y - radar.center.y) * ratio }))
  if (ratio > 0) for (let vertex = 0; vertex < ring.length; vertex++) guide(commands, ring[vertex]!, ring[(vertex + 1) % ring.length]!, geometry)
  pointerLabel(commands, data.categories[index] ?? '', end, end.x < radar.center.x ? 'left' : 'right', geometry)
  const value = axisValue(ratio, radar.min, radar.max, chart.yAxis), label = chart.options?.radar?.axisLabel?.format ?? chart.yAxis.format
  pointerLabel(commands, formatChartNumber(value, label), { x: radar.center.x + (end.x - radar.center.x) * ratio, y: radar.center.y + (end.y - radar.center.y) * ratio }, 'right', geometry)
  return commands.map(command => command.clip ? command : { ...command, clip: layout.frame })
}

function nearestCategory(categories: readonly PointerCategory[], position: number): PointerCategory | undefined {
  const firstAt = (target: number): number => {
    let low = 0, high = categories.length
    while (low < high) { const middle = (low + high) >>> 1; if (categories[middle]!.position < target) low = middle + 1; else high = middle }
    return low
  }
  const low = firstAt(position)
  const after = categories[low], before = categories[low - 1]
  const nearest = !before ? after : !after || position - before.position <= after.position - position ? before : after
  // A calendar bucket can hold several original dates. Stable source-index
  // order resolves their shared coordinate from either side of the bucket.
  return nearest ? categories[firstAt(nearest.position)] : undefined
}

function axisValue(position: number, min: number, max: number, axis: ChartAxisOptions): number {
  const ratio = Math.max(0, Math.min(1, axis.reverse ? 1 - position : position))
  return axis.scaleType === 'log' ? ratio <= 0 ? min : ratio >= 1 ? max : Math.exp(chartInterpolate(Math.log(min), Math.log(max), ratio))
    : chartInterpolate(min, max, ratio)
}

function inside(point: ChartVector, rect: RenderRect): boolean {
  return point.x >= rect.x && point.x <= rect.x + rect.width && point.y >= rect.y && point.y <= rect.y + rect.height
}

function guide(commands: DrawCommand[], start: ChartVector, end: ChartVector, geometry: ChartAxisPointerGeometry): void {
  const pointer = geometry.presentation
  chartLine(commands, start, end, { color: pointer.indicatorLineColor, width: 1, dashType: pointer.indicatorLineType }, pointer.indicatorLineColor, geometry.layout.plot)
}

function pointerLabel(commands: DrawCommand[], text: string, anchor: ChartVector, placement: 'left' | 'right' | 'above' | 'below', geometry: ChartAxisPointerGeometry): void {
  if (!text) return
  const { frame } = geometry.layout, style = geometry.presentation, size = style.font.fontSize ?? 10
  const insetX = frame.width > 2 ? 1 : 0, insetY = frame.height > 2 ? 1 : 0
  const availableWidth = Math.max(1, frame.width - insetX * 2), availableHeight = Math.max(1, frame.height - insetY * 2)
  let width = 0, label = ''
  // Bound long source labels without allocating a character array for the entire category.
  for (const character of text) {
    const advance = size * ((character.codePointAt(0) ?? 0) > 255 ? 1 : 0.6)
    if (width + advance > availableWidth - 8) { label += '…'; width = availableWidth - 8; break }
    label += character; width += advance
  }
  width = Math.min(availableWidth, Math.max(12, width + 8))
  const height = Math.min(availableHeight, size + 6)
  let x = placement === 'left' ? anchor.x - width - 6 : placement === 'right' ? anchor.x + 6 : anchor.x - width / 2
  let y = placement === 'above' ? anchor.y - height - 6 : placement === 'below' ? anchor.y + 6 : anchor.y - height / 2
  x = Math.max(frame.x + insetX, Math.min(frame.x + frame.width - width - insetX, x)); y = Math.max(frame.y + insetY, Math.min(frame.y + frame.height - height - insetY, y))
  commands.push({ kind: 'rect', x, y, width, height, fill: style.indicatorLabelColor, clip: frame },
    { kind: 'text', x: x + width / 2, y: y + height / 2, text: label, font: chartFont(style.font, 10), color: style.indicatorLabelTextColor,
      align: 'center', baseline: 'middle', maxWidth: Math.max(1, width - 8), clip: frame })
}
