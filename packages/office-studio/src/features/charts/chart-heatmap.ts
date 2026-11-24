import type { ChartAxisOptions, ChartFontOptions, ChartLabelOptions, ChartModel } from '../../domain/model'
import type { RenderRect } from '../../rendering/contracts'
import type { ChartData } from './chart-engine'
import { chartColorChannels, normalizeChartColor } from './chart-color'
import { formatChartNumber } from './chart-geometry'
import { chartInterpolate, chartLinearRatio } from './chart-math'
import { chartSeriesStyle, definedChartStyleFields } from './chart-options'

export interface ChartHeatmapExtent { min: number; max: number }
export interface ChartHeatmapPiece {
  lower: number
  upper: number
  lowerClosed: boolean
  color: string
  text: string
}
export interface ChartHeatmapScale {
  type: 'continuous' | 'piecewise'
  extent: ChartHeatmapExtent
  palette: string[]
  channels: Array<ReturnType<typeof chartColorChannels>>
  pieces: ChartHeatmapPiece[]
}
export interface ChartHeatmapCellIdentity { seriesPosition: number; categoryIndex: number }
export interface ChartHeatmapLabel extends ChartFontOptions {
  visible: boolean
  position: NonNullable<ChartLabelOptions['position']>
  format?: string
}

/** Fixed Oae source semantics: fields/headers are X; source categories are Y. */
export function chartHeatmapFieldNames(data: ChartData, chart: Pick<ChartModel, 'type' | 'options'>): string[] {
  return data.series.map((series, position) => chartSeriesStyle(chart, series.field, position, undefined, series.sourceName ?? series.name).name ?? series.name)
}

/** Actual finite extent, including the fixed empty-source 0/0 default. No
 * full-array spread or implicit zero/one padding is used. ECharts sorts an
 * effective one-sided author/automatic pair; author fields remain untouched. */
export function chartHeatmapExtent(data: ChartData, chart: Pick<ChartModel, 'options'>): ChartHeatmapExtent {
  let min = Infinity, max = -Infinity
  for (const series of data.series) for (const value of series.values) {
    if (value === null || !Number.isFinite(value)) continue
    min = Math.min(min, value); max = Math.max(max, value)
  }
  if (min === Infinity) { min = 0; max = 0 }
  const authored = chart.options?.heatmap, first = authored?.min ?? min, last = authored?.max ?? max
  if (!Number.isFinite(first) || !Number.isFinite(last)) throw new RangeError('热力图色阶范围须为有限数值。')
  if (authored?.min !== undefined && authored.max !== undefined && authored.min > authored.max) throw new RangeError('热力图最小值不能大于最大值。')
  return { min: Math.min(first, last), max: Math.max(first, last) }
}

export function createChartHeatmapScale(extent: ChartHeatmapExtent, palette: readonly string[], type: ChartHeatmapScale['type'] = 'continuous'): ChartHeatmapScale {
  if (![extent.min, extent.max].every(Number.isFinite) || extent.min > extent.max || !palette.length) throw new RangeError('热力图色阶需要有效范围与非空调色板。')
  const colors = palette.map(normalizeChartColor), scale: ChartHeatmapScale = { type, extent: { ...extent }, palette: colors,
    channels: colors.map(chartColorChannels), pieces: [] }
  if (extent.min === extent.max) {
    scale.pieces.push({ lower: extent.min, upper: extent.max, lowerClosed: true, color: chartHeatmapRatioColor(scale, 0.5), text: String(extent.min) })
  } else {
    // Fixed splitNumber=8. Explicit binary64 bounds avoid the upstream decimal
    // step rounding that collapses subnormal or extreme finite extents.
    for (let index = 0; index < 8; index++) {
      const lower = chartInterpolate(extent.min, extent.max, index / 8), upper = chartInterpolate(extent.min, extent.max, (index + 1) / 8)
      if (upper <= lower) continue
      let first = formatChartNumber(lower), last = formatChartNumber(upper)
      if (first === last) { first = String(lower); last = String(upper) }
      scale.pieces.push({ lower, upper, lowerClosed: scale.pieces.length === 0, color: chartHeatmapRatioColor(scale, index / 7), text: `${first} – ${last}` })
    }
  }
  return scale
}

export function chartHeatmapValueRatio(value: number, extent: ChartHeatmapExtent): number {
  return extent.min === extent.max ? 0.5 : Math.max(0, Math.min(1, chartLinearRatio(value, extent.min, extent.max)))
}
export function chartHeatmapRatioColor(scale: Pick<ChartHeatmapScale, 'palette' | 'channels'>, ratio: number): string {
  const position = Math.max(0, Math.min(1, ratio)) * (scale.palette.length - 1), index = Math.floor(position), fraction = position - index
  const first = scale.channels[index]!, last = scale.channels[Math.min(index + 1, scale.channels.length - 1)]!
  const channels = first.map((channel, axis) => channel * (1 - fraction) + last[axis]! * fraction)
  return `rgba(${Math.round(channels[0]! * 255)}, ${Math.round(channels[1]! * 255)}, ${Math.round(channels[2]! * 255)}, ${channels[3]})`
}
/** Boundaries belong to the preceding piece, matching reformIntervals. */
export function chartHeatmapPieceIndex(scale: ChartHeatmapScale, value: number, clamp = false): number {
  if (!Number.isFinite(value)) return -1
  if (value < scale.extent.min) return clamp ? 0 : -1
  if (value > scale.extent.max) return clamp ? scale.pieces.length - 1 : -1
  for (let index = 0; index < scale.pieces.length; index++) if (value <= scale.pieces[index]!.upper) return index
  return -1
}
export function chartHeatmapScaleColor(scale: ChartHeatmapScale, value: number): string {
  return scale.type === 'piecewise' ? scale.pieces[chartHeatmapPieceIndex(scale, value, true)]!.color
    : chartHeatmapRatioColor(scale, chartHeatmapValueRatio(value, scale.extent))
}
/** Fixed continuous default range is unbounded at both ends; piecewise's
 * inactive out-of-range marks are transparent. Values stay in the source. */
export function chartHeatmapValueVisible(scale: ChartHeatmapScale, value: number): boolean {
  return Number.isFinite(value) && (scale.type === 'continuous' || chartHeatmapPieceIndex(scale, value) >= 0)
}

export function chartHeatmapCellRect(data: Pick<ChartData, 'categories' | 'series'>,
  chart: { xAxis?: Partial<ChartAxisOptions>; yAxis?: Partial<ChartAxisOptions> }, plot: RenderRect,
  seriesPosition: number, categoryIndex: number): RenderRect | undefined {
  const columns = data.series.length, rows = data.categories.length
  if (!Number.isInteger(seriesPosition) || !Number.isInteger(categoryIndex) || seriesPosition < 0 || seriesPosition >= columns || categoryIndex < 0 || categoryIndex >= rows) return undefined
  const width = plot.width / columns, height = plot.height / rows
  const column = chart.xAxis?.reverse ? columns - seriesPosition - 1 : seriesPosition
  const row = chart.yAxis?.reverse ? categoryIndex : rows - categoryIndex - 1
  return { x: plot.x + column * width, y: plot.y + row * height, width, height }
}

export function hitChartHeatmapCell(data: Pick<ChartData, 'categories' | 'series'>,
  chart: { xAxis?: Partial<ChartAxisOptions>; yAxis?: Partial<ChartAxisOptions> }, plot: RenderRect,
  point?: { x: number; y: number }): ChartHeatmapCellIdentity | undefined {
  if (!point || !Number.isFinite(point.x) || !Number.isFinite(point.y) || point.x < plot.x || point.x > plot.x + plot.width || point.y < plot.y || point.y > plot.y + plot.height
    || !data.series.length || !data.categories.length || plot.width <= 0 || plot.height <= 0) return undefined
  const column = Math.min(data.series.length - 1, Math.floor((point.x - plot.x) / plot.width * data.series.length))
  const row = Math.min(data.categories.length - 1, Math.floor((point.y - plot.y) / plot.height * data.categories.length))
  return { seriesPosition: chart.xAxis?.reverse ? data.series.length - column - 1 : column,
    categoryIndex: chart.yAxis?.reverse ? row : data.categories.length - row - 1 }
}

/** Only value-label semantics apply while Heatmap is active. Retained generic
 * series/point font and format overrides remain usable without emitting a
 * second label or consuming another chart type's content mask. */
export function resolveChartHeatmapLabel(chart: Pick<ChartModel, 'options'>, inherited?: ChartLabelOptions, color = '#334155'): ChartHeatmapLabel {
  const label = { ...definedChartStyleFields(inherited), ...definedChartStyleFields(chart.options?.heatmap?.label) }
  return { visible: label.visible ?? false, position: label.position ?? 'inside', format: label.format,
    family: label.family ?? chart.options?.appearance?.font?.family ?? 'Arial', fontSize: label.fontSize ?? 12,
    color: label.color ?? color, bold: label.bold, italic: label.italic }
}
export function chartHeatmapLabelText(value: number, formattedValue: string | null | undefined, label: Pick<ChartHeatmapLabel, 'format'>): string {
  return label.format !== undefined ? formatChartNumber(value, label.format) : formattedValue ?? formatChartNumber(value)
}
export function chartHeatmapLabelAnchor(rect: RenderRect, label: Pick<ChartHeatmapLabel, 'position' | 'fontSize'>): {
  x: number; y: number; align: 'left' | 'center' | 'right'; baseline: 'middle'; nativePosition: 'inside' | 'top' | 'bottom' | 'left' | 'right' | 'insideTop' | 'insideBottom'
} {
  const anchor = { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2, align: 'center' as 'left' | 'center' | 'right', baseline: 'middle' as const,
    nativePosition: 'inside' as 'inside' | 'top' | 'bottom' | 'left' | 'right' | 'insideTop' | 'insideBottom' }
  const gap = 5 + (label.fontSize ?? 12) / 2
  if (label.position === 'top' || label.position === 'auto' || label.position === 'outside') { anchor.y = rect.y - gap; anchor.nativePosition = 'top' }
  else if (label.position === 'bottom') { anchor.y = rect.y + rect.height + gap; anchor.nativePosition = 'bottom' }
  else if (label.position === 'left') { anchor.x = rect.x - 5; anchor.align = 'right'; anchor.nativePosition = 'left' }
  else if (label.position === 'right') { anchor.x = rect.x + rect.width + 5; anchor.align = 'left'; anchor.nativePosition = 'right' }
  else if (label.position === 'insideEnd') { anchor.y = rect.y + Math.min(gap, rect.height / 2); anchor.nativePosition = 'insideTop' }
  else if (label.position === 'insideBase') { anchor.y = rect.y + rect.height - Math.min(gap, rect.height / 2); anchor.nativePosition = 'insideBottom' }
  return anchor
}

export function chartHeatmapBoundaryText(value: number, unit?: string): string {
  const trimmed = unit?.trim()
  return `${String(value)}${trimmed ? ` ${trimmed}` : ''}`
}

/** Automatic legend space is shared with the plot frame; authored plot areas
 * keep their exact bounds and can overlap chrome, as with other chart types. */
export function chartHeatmapLegendBand(width: number, type: ChartHeatmapScale['type'], fontSize = 11): number {
  return type === 'continuous' ? Math.max(30, fontSize + 20) : Math.ceil(8 / Math.max(1, Math.min(8, Math.floor((width - 16) / 108)))) * (fontSize + 8) + 8
}
