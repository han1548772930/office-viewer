import type { ChartErrorBarOptions, ChartModel, ChartSeriesOptions, ChartTrendlineOptions, ChartType } from '../../domain/model'
import type { ChartData } from './chart-engine'
import type { ChartHierarchyItem } from './chart-data'
import { resolveChartType, supportsChartErrorDirection, supportsChartStatistics } from './chart-catalog'
import { chartComputeInputSteps } from './chart-input-preparation'
import { chartSeriesColorIsAuthored, chartSeriesStyle, definedChartStyleFields } from './chart-options'
import { validateChartNamedSeriesBindings } from './chart-series-selector'
import { chartShade, formatChartNumber } from './chart-geometry'
import { chartFiniteSum, chartInterpolate, chartLinearRatio, finiteChartValue } from './chart-math'
import { chartParetoDateValues, prepareChartDateData, type ChartDateAxisConfig } from './chart-date-axis'
import { chartColorChannels, chartColorRepresentative, isChartLinearGradient } from './chart-color'
import { supportsChartExplicitGradient } from './chart-gradient'
import {
  chartBoxplot, chartErrorIntervals, chartHistogram, chartPareto, chartTrendline, chartWaterfall,
  type BoxplotData, type ChartErrorInterval, type HistogramData, type ParetoData, type TrendlineData, type WaterfallPoint,
} from './chart-statistics'

export interface ChartDerivedData {
  data: ChartData
  histogram?: HistogramData
  boxes?: Array<{ field: number; seriesPosition: number; sourceName: string; name: string; statistics?: BoxplotData }>
  waterfalls?: WaterfallPoint[][]
  pareto?: ParetoData
  trends: Array<{ field: number; seriesPosition: number; statistics: TrendlineData; config: ChartTrendlineOptions }>
  errors: Array<{ field: number; seriesPosition: number; intervals: ChartErrorInterval[]; config: ChartErrorBarOptions }>
}

/** Worker result for the current source/config revision. The caller owns cancellation and stale-result rejection. */
export interface ChartComputedStatistics {
  /** Input after category aggregation and value normalization, before statistical transforms. */
  preparedData?: ChartData
  histogram?: HistogramData
  boxplots?: Array<{ field: number; seriesPosition: number; statistics?: BoxplotData }>
  waterfalls?: Array<{ field: number; seriesPosition: number; points: WaterfallPoint[] }>
  pareto?: ParetoData
  /** Source field IDs may repeat after Bubble grouping. Each result retains the
   * authored config slot and its actual preparedData.series position. */
  trendlines?: Array<{ configIndex: number; seriesPosition: number; statistics: TrendlineData }>
  errorBars?: Array<{ configIndex: number; seriesPosition: number; intervals: ChartErrorInterval[] }>
}

export function chartComputeInputs(source: ChartData, chart: ChartDateAxisConfig): ChartData {
  const steps = chartComputeInputSteps(source, chart)
  while (true) { const step = steps.next(); if (step.done) return step.value }
}

export function chartSeriesKind(chart: ChartModel, field: number, index: number, sourceName?: string): { type: 'column' | 'bar' | 'line' | 'area'; stacked: boolean } {
  const descriptor = resolveChartType(chart.type)
  if (descriptor.family === 'combination') {
    const type = chartSeriesStyle(chart, field, index, undefined, sourceName).type ?? (index === 0 ? 'column' : 'line')
    return { type: type.startsWith('column') ? 'column' : type.startsWith('area') ? 'area' : 'line', stacked: type.endsWith('Stacked') }
  }
  if (descriptor.family === 'histogram' || descriptor.family === 'pareto') return { type: 'column', stacked: false }
  return { type: descriptor.family === 'bar' ? 'bar' : descriptor.family === 'column' ? 'column' : descriptor.family === 'area' ? 'area' : 'line', stacked: descriptor.stacked ?? false }
}

/** The same derived values feed the bitmap renderer and the public Canvas command API. */
export function deriveChartData(source: ChartData, chart: ChartModel, computed?: ChartComputedStatistics): ChartDerivedData {
  let data = computed?.preparedData ?? chartComputeInputs(source, chart)
  validateChartNamedSeriesBindings(chart.options, data.series)
  const result: ChartDerivedData = { data, trends: [], errors: [] }
  if (!data.categories.length || !data.series.length) return result
  if (chart.type === 'histogram') {
    const series = data.series[0]!
    const histogram = computed?.histogram ?? chartHistogram(series.values.filter((value): value is number => value !== null), chart.options?.histogram)
    result.histogram = histogram
    data = { date1904: data.date1904, categories: histogram.bins.map((bin, index) => bin.kind === 'underflow' ? `≤ ${formatChartNumber(bin.upperBound!)}`
      : bin.kind === 'overflow' ? `> ${formatChartNumber(bin.lowerBound!)}` : `${index === 0 && chart.options?.histogram?.underflowThreshold === undefined ? '[' : '('}${formatChartNumber(bin.lowerBound!)}, ${formatChartNumber(bin.upperBound!)}]`),
    ...(chart.options?.mapping?.useDateAxis ? { categoryValues: histogram.bins.map(bin => bin.lowerBound === undefined ? bin.upperBound! : bin.upperBound === undefined ? bin.lowerBound : chartInterpolate(bin.lowerBound, bin.upperBound, 0.5)) } : {}),
    series: [{ ...series, values: histogram.bins.map(bin => bin.count) }] }
  } else if (chart.type === 'boxplot') {
    result.boxes = data.series.map((series, seriesPosition) => {
      const box = computed?.boxplots?.find(box => box.seriesPosition === seriesPosition)
      if (computed?.boxplots && !box) throw new RangeError('缺少与当前系列位置匹配的箱线统计结果。')
      const sourceName = series.sourceName ?? series.name
      return { field: series.field, seriesPosition, sourceName, name: chartSeriesStyle(chart, series.field, seriesPosition, undefined, sourceName).name ?? series.name, statistics: box ? box.statistics
        : chartBoxplot(series.values.filter((value): value is number => value !== null), chart.options?.boxplot) }
    })
    data = { date1904: data.date1904, categories: result.boxes.map(box => box.name), categoryValues: result.boxes.map(box => box.name), series: [{ field: data.series[0]!.field, name: '样本分布',
      sourceName: data.series[0]!.sourceName ?? data.series[0]!.name, values: result.boxes.map(box => box.statistics?.median ?? null) }] }
  } else if (chart.type === 'pareto') {
    const series = data.series[0]!
    result.pareto = computed?.pareto ?? chartPareto(data.categories, series.values, chart.options?.pareto?.includeZeroValues)
    let categoryValues: ChartData['categoryValues'] = result.pareto.categories
    if (chart.options?.mapping?.useDateAxis) {
      categoryValues = chartParetoDateValues(prepareChartDateData(data, chart), result.pareto.categories, chart.options?.pareto?.includeZeroValues)
    }
    data = { date1904: data.date1904, categories: result.pareto.categories, categoryValues, series: [{ ...series, values: result.pareto.values }] }
  } else if (chart.type === 'waterfall') result.waterfalls = data.series.map((series, seriesPosition) => {
    const waterfall = computed?.waterfalls?.find(result => result.seriesPosition === seriesPosition)
    if (computed?.waterfalls && !waterfall) throw new RangeError('缺少与当前系列位置匹配的瀑布统计结果。')
    return waterfall?.points ?? chartWaterfall(series, chart)
  })
  data = prepareChartDateData(data, chart)
  result.data = data
  for (const [configIndex, config] of (chart.options?.trendlines ?? []).entries()) {
    for (const [seriesPosition, series] of data.series.entries()) {
      if (series.field !== config.seriesIndex || !supportsChartStatistics(chart, series.field, seriesPosition, series.sourceName ?? series.name)) continue
      const trend = computed?.trendlines?.find(result => result.configIndex === configIndex && result.seriesPosition === seriesPosition)
      if (computed?.trendlines && !trend) throw new RangeError('缺少与当前配置和系列位置匹配的趋势线结果。')
      result.trends.push({ field: series.field, seriesPosition, statistics: trend?.statistics ?? chartTrendline(series, config), config })
    }
  }
  for (const [configIndex, config] of (chart.options?.errorBars ?? []).entries()) {
    for (const [seriesPosition, series] of data.series.entries()) {
      if (series.field !== config.seriesIndex || !supportsChartStatistics(chart, series.field, seriesPosition, series.sourceName ?? series.name) || !supportsChartErrorDirection(chart, config.direction)) continue
      const error = computed?.errorBars?.find(result => result.configIndex === configIndex && result.seriesPosition === seriesPosition)
      if (computed?.errorBars && !error) throw new RangeError('缺少与当前配置和系列位置匹配的误差线结果。')
      result.errors.push({ field: series.field, seriesPosition, intervals: error?.intervals ?? chartErrorIntervals(data, series, config), config })
    }
  }
  return result
}

export interface ChartPiePart { category: string; value: number; index: number }
export interface ChartCompoundPieData { primary: ChartPiePart[]; secondary: ChartPiePart[]; aggregateIndex: number }

export function chartCompoundPieData(data: ChartData, chart: ChartModel): ChartCompoundPieData {
  const source = data.series[0]!, parts = data.categories.map((category, index) => ({ category, index, value: Math.max(0, source.values[index] ?? 0) }))
  const split = chart.options?.pie?.secondaryPlot?.split ?? { type: 'auto' as const }
  const scale = parts.reduce((max, part) => Math.max(max, part.value), 0) || 1, total = parts.reduce((sum, part) => sum + part.value / scale, 0)
  const secondary = parts.filter((part, index) => split.type === 'custom' ? split.pointIndexes?.includes(index)
    : split.type === 'value' ? part.value < (split.lessThan ?? 0)
      : split.type === 'percentage' ? total > 0 && (part.value / scale) / total * 100 < (split.lessThan ?? 10)
        : index >= Math.max(1, parts.length - (split.count ?? Math.max(1, Math.min(3, Math.ceil(parts.length / 3))))))
  if (!secondary.length || secondary.length === parts.length) throw new RangeError('复合饼图须同时保留主图和副图的项目。请调整拆分规则。')
  const indices = new Set(secondary.map(part => part.index))
  const aggregateIndex = parts.length
  const primary = parts.filter(part => !indices.has(part.index))
  primary.push({ category: chart.options?.pie?.aggregateLabel ?? '其他', value: finiteChartValue(chartFiniteSum(secondary.map(part => part.value)), '复合饼图副图合计'), index: aggregateIndex })
  return { primary, secondary, aggregateIndex }
}

export function chartPointStyle(style: ChartSeriesOptions, categoryIndex: number, type?: ChartType): ChartSeriesOptions {
  const authored = style.dataPoints?.[categoryIndex]
  // A type switch retains authored data. Unsupported point gradients stop
  // overriding the active series; marker colors are always solid CSS colors.
  const point = type !== undefined && isChartLinearGradient(authored?.color) && !supportsChartExplicitGradient(type, 'dataPoint')
    ? { ...authored, color: undefined } : authored
  return { ...style, ...definedChartStyleFields(point), border: { ...style.border, ...definedChartStyleFields(point?.border) }, point: { ...style.point,
    ...(point?.color === undefined ? {} : { color: chartColorRepresentative(point.color) }),
    ...(point?.size === undefined ? {} : { size: point.size }), ...(point?.shape === undefined ? {} : { shape: point.shape }) } }
}

/** An authored marker size is an absolute diameter; source weights still determine
 * the default diameter and invalid/zero weights never become visible markers. */
export function chartBubblePointSize(style: ChartSeriesOptions, value: number, max: number, options: ChartModel['options'], available: number): number {
  if (!Number.isFinite(value) || value <= 0 || !Number.isFinite(available) || available <= 0) return 0
  if (style.point?.size !== undefined) return style.point.size
  const ratio = value / Math.max(1, max)
  return (options?.bubble?.scale ?? 1) * Math.max(0, Math.min(54, available * 0.2))
    * (options?.bubble?.sizeRepresents === 'width' ? ratio : Math.sqrt(ratio))
}

/** Both hierarchy adapters resolve authored styles using source DFS indices, before any layout sorting. */
export function chartHierarchyPointStyle(chart: ChartModel, field: number, item: Pick<ChartHierarchyItem, 'pointIndex' | 'rootIndex' | 'depth'>, palette: readonly string[], sourceName?: string): ChartSeriesOptions {
  const base = chartSeriesStyle(chart, field, item.rootIndex, palette, sourceName)
  if (!chartSeriesColorIsAuthored(chart, field, sourceName)) {
    const step = chart.type === 'sunburst' ? 0.08 : 0.1
    base.color = chartShade(base.color!, 1 - Math.min(0.4, item.depth * step))
  }
  return chartPointStyle(base, item.pointIndex, chart.type)
}

export function chartHeatmapExtent(data: ChartData, chart: ChartModel): { min: number; max: number } {
  let min = 0, max = 1
  for (const series of data.series) for (const value of series.values) if (value !== null && Number.isFinite(value)) { min = Math.min(min, value); max = Math.max(max, value) }
  return { min: chart.options?.heatmap?.min ?? min, max: chart.options?.heatmap?.max ?? max }
}

/** Authored series and point colors override the same deterministic color scale
 * in both adapters; the scale itself never sums or spreads the full data set. */
export function chartHeatmapPointStyle(chart: ChartModel, field: number, row: number, column: number, value: number, extent: { min: number; max: number }, palette: readonly string[], sourceName?: string): ChartSeriesOptions {
  const base = chartSeriesStyle(chart, field, row, palette, sourceName)
  if (!chartSeriesColorIsAuthored(chart, field, sourceName)) {
    const ratio = Math.max(0, Math.min(1, chartLinearRatio(value, extent.min, extent.max)))
    if (chart.options?.heatmap?.visualMapType === 'piecewise') base.color = palette[Math.min(palette.length - 1, Math.floor(ratio * palette.length))]!
    else {
      const position = ratio * (palette.length - 1), index = Math.floor(position), fraction = position - index
      const first = chartColorChannels(palette[index]!), last = chartColorChannels(palette[Math.min(index + 1, palette.length - 1)]!)
      const channels = first.map((channel, axis) => channel * (1 - fraction) + last[axis]! * fraction)
      base.color = `rgba(${Math.round(channels[0]! * 255)}, ${Math.round(channels[1]! * 255)}, ${Math.round(channels[2]! * 255)}, ${channels[3]})`
    }
  }
  return chartPointStyle(base, column, chart.type)
}

export function chartCandlestickPointStyle(chart: ChartModel, field: number, index: number, rising: boolean, palette: readonly string[], sourceName?: string): ChartSeriesOptions {
  const base = chartSeriesStyle(chart, field, 0, palette, sourceName)
  if (!chartSeriesColorIsAuthored(chart, field, sourceName)) base.color = rising ? chart.options?.candlestick?.risingColor ?? '#dc2626' : chart.options?.candlestick?.fallingColor ?? '#059669'
  return chartPointStyle(base, index, chart.type)
}
