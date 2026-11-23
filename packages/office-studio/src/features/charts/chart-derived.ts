import type { ChartErrorBarOptions, ChartModel, ChartSeriesOptions, ChartTrendlineOptions } from '../../domain/model'
import type { ChartData } from './chart-engine'
import { resolveChartType, supportsChartErrorDirection, supportsChartStatistics } from './chart-catalog'
import { normalizedChartValues } from './chart-data'
import { chartSeriesStyle } from './chart-options'
import { formatChartNumber } from './chart-geometry'
import { chartFiniteSum, finiteChartValue } from './chart-math'
import {
  chartBoxplot, chartErrorIntervals, chartHistogram, chartPareto, chartTrendline, chartWaterfall,
  type BoxplotData, type ChartErrorInterval, type HistogramData, type ParetoData, type TrendlineData, type WaterfallPoint,
} from './chart-statistics'

export interface ChartDerivedData {
  data: ChartData
  histogram?: HistogramData
  boxes?: Array<{ field: number; name: string; statistics?: BoxplotData }>
  waterfalls?: WaterfallPoint[][]
  pareto?: ParetoData
  trends: Array<{ field: number; statistics: TrendlineData; config: ChartTrendlineOptions }>
  errors: Array<{ field: number; intervals: ChartErrorInterval[]; config: ChartErrorBarOptions }>
}

/** Worker result for the current source/config revision. The caller owns cancellation and stale-result rejection. */
export interface ChartComputedStatistics {
  histogram?: HistogramData
  boxplots?: Array<{ field: number; statistics?: BoxplotData }>
  waterfalls?: Array<{ field: number; points: WaterfallPoint[] }>
  pareto?: ParetoData
  /** Entries keep the authored trendlines/errorBars array order. */
  trendlines?: TrendlineData[]
  errorBars?: ChartErrorInterval[][]
}

export function chartComputeInputs(source: ChartData, chart: ChartModel): ChartData {
  const descriptor = resolveChartType(chart.type)
  const categoryData = descriptor.dataShape === 'category' ? aggregateChartCategories(source, chart) : source
  const values = normalizedChartValues(categoryData, chart)
  return { ...categoryData, series: categoryData.series.map((series, index) => ({ ...series, values: values[index]! })) }
}

export function chartSeriesKind(chart: ChartModel, field: number, index: number): { type: 'column' | 'bar' | 'line' | 'area'; stacked: boolean } {
  const descriptor = resolveChartType(chart.type)
  if (descriptor.family === 'combination') {
    const type = chartSeriesStyle(chart, field, index).type ?? (index === 0 ? 'column' : 'line')
    return { type: type.startsWith('column') ? 'column' : type.startsWith('area') ? 'area' : 'line', stacked: type.endsWith('Stacked') }
  }
  if (descriptor.family === 'histogram' || descriptor.family === 'pareto') return { type: 'column', stacked: false }
  return { type: descriptor.family === 'bar' ? 'bar' : descriptor.family === 'column' ? 'column' : descriptor.family === 'area' ? 'area' : 'line', stacked: descriptor.stacked ?? false }
}

/** The same derived values feed the bitmap renderer and the public Canvas command API. */
export function deriveChartData(source: ChartData, chart: ChartModel, computed?: ChartComputedStatistics): ChartDerivedData {
  let data = chartComputeInputs(source, chart)
  const result: ChartDerivedData = { data, trends: [], errors: [] }
  if (chart.type === 'histogram') {
    const series = data.series[0]!
    const histogram = computed?.histogram ?? chartHistogram(series.values.filter((value): value is number => value !== null), chart.options?.histogram)
    result.histogram = histogram
    data = { categories: histogram.bins.map((bin, index) => bin.kind === 'underflow' ? `≤ ${formatChartNumber(bin.upperBound!)}`
      : bin.kind === 'overflow' ? `> ${formatChartNumber(bin.lowerBound!)}` : `${index === 0 && chart.options?.histogram?.underflowThreshold === undefined ? '[' : '('}${formatChartNumber(bin.lowerBound!)}, ${formatChartNumber(bin.upperBound!)}]`),
    series: [{ ...series, values: histogram.bins.map(bin => bin.count) }] }
  } else if (chart.type === 'boxplot') {
    result.boxes = data.series.map(series => ({ field: series.field, name: series.name, statistics: computed?.boxplots
      ? computed.boxplots.find(box => box.field === series.field)?.statistics
      : chartBoxplot(series.values.filter((value): value is number => value !== null), chart.options?.boxplot) }))
    data = { categories: result.boxes.map(box => box.name), series: [{ field: data.series[0]!.field, name: '样本分布', values: result.boxes.map(box => box.statistics?.median ?? null) }] }
  } else if (chart.type === 'pareto') {
    const series = data.series[0]!
    result.pareto = computed?.pareto ?? chartPareto(data.categories, series.values, chart.options?.pareto?.includeZeroValues)
    data = { categories: result.pareto.categories, series: [{ ...series, values: result.pareto.values }] }
  } else if (chart.type === 'waterfall') result.waterfalls = data.series.map(series => computed?.waterfalls?.find(result => result.field === series.field)?.points ?? chartWaterfall(series, chart))
  result.data = data
  for (const [index, config] of (chart.options?.trendlines ?? []).entries()) {
    const series = data.series.find(candidate => candidate.field === config.seriesIndex)
    if (series && supportsChartStatistics(chart, series.field, data.series.indexOf(series))) result.trends.push({ field: series.field, statistics: computed?.trendlines?.[index] ?? chartTrendline(series, config), config })
  }
  for (const [index, config] of (chart.options?.errorBars ?? []).entries()) {
    const series = data.series.find(candidate => candidate.field === config.seriesIndex)
    if (series && supportsChartStatistics(chart, series.field, data.series.indexOf(series)) && supportsChartErrorDirection(chart, config.direction)) result.errors.push({ field: series.field, intervals: computed?.errorBars?.[index] ?? chartErrorIntervals(data, series, config), config })
  }
  return result
}

function aggregateChartCategories(data: ChartData, chart: ChartModel): ChartData {
  const options = chart.options?.aggregation
  if (!options?.aggregate && !options?.topN) return data
  const groups = new Map<string, number[]>()
  data.categories.forEach((category, index) => {
    const key = options.aggregate ? category : `${index}\u0000${category}`
    const group = groups.get(key) ?? []; group.push(index); groups.set(key, group)
  })
  let entries = [...groups.values()].map(indexes => ({
    category: data.categories[indexes[0]!]!, indexes,
    values: data.series.map(series => { const values = indexes.map(index => series.values[index]).filter((value): value is number => value != null); return values.length ? chartFiniteSum(values, '图表分类合计') : null }),
  }))
  if (options.topN) {
    const top = options.topN
    let scale = 1
    for (const entry of entries) for (const value of entry.values) scale = Math.max(scale, Math.abs(value ?? 0))
    const weight = (values: Array<number | null>): number => values.reduce((sum, value) => sum + Math.abs(value ?? 0) / scale, 0)
    const sorted = [...entries].sort((a, b) => weight(b.values) - weight(a.values) || a.indexes[0]! - b.indexes[0]!)
    const retained = sorted.slice(0, top.count), remaining = sorted.slice(top.count)
    entries = retained
    if (top.includeOthers && remaining.length) entries.push({
      category: top.othersLabel ?? '其他', indexes: remaining.flatMap(entry => entry.indexes),
      values: data.series.map((_series, seriesIndex) => { const values = remaining.map(entry => entry.values[seriesIndex]).filter((value): value is number => value != null); return values.length ? chartFiniteSum(values, '图表其他分类合计') : null }),
    })
  }
  return { categories: entries.map(entry => entry.category), fields: data.fields,
    ...(data.records && entries.every(entry => entry.indexes.length === 1) ? { records: entries.map(entry => data.records![entry.indexes[0]!]!) } : {}),
    series: data.series.map((series, index) => ({ ...series, values: entries.map(entry => entry.values[index] ?? null) })) }
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

export function chartPointStyle(style: ChartSeriesOptions, categoryIndex: number): ChartSeriesOptions {
  const point = style.dataPoints?.[categoryIndex]
  return { ...style, ...point, border: { ...style.border, ...point?.border }, point: { ...style.point, ...(point?.size === undefined ? {} : { size: point.size }) } }
}
