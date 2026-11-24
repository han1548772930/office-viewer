import type { ChartModel } from '../../domain/model'
import type { ChartData } from './chart-engine'
import { resolveChartType } from './chart-catalog'
import { prepareChartDateSource } from './chart-date-axis'
import { chartSeriesKind } from './chart-derived'
import { chartSeriesStyle } from './chart-options'
import { validateChartNamedSeriesBindings } from './chart-series-selector'
import { CHART_PREPARATION_CHUNK, consumeChartPreparation, consumeChartPreparationSync,
  type ChartComputePreparationContext } from './chart-compute-lifecycle'

/** Source validation is shared by synchronous small-source calls and the
 * Studio's queued authoring lifecycle. Date scene admission is a later stage
 * for asynchronous edits, after the final statistical vectors exist. */
export function validateChartSourceData(chart: ChartModel, data: ChartData): void {
  consumeChartPreparationSync(chartSourceValidationSteps(chart, data, true))
}

export function validateChartSourceDataAsync(chart: ChartModel, data: ChartData,
  context: ChartComputePreparationContext = {}): Promise<void> {
  return consumeChartPreparation(chartSourceValidationSteps(chart, data, false), context)
}

export function* chartSourceValidationSteps(chart: ChartModel, source: ChartData, validateDates = false): Generator<void, void, void> {
  const data = validateDates ? prepareChartDateSource(source, chart) : source
  validateChartNamedSeriesBindings(chart.options, data.series)
  const type = chart.type, descriptor = resolveChartType(type), axis = chart.yAxis
  if (descriptor.stacked && axis.scaleType !== 'linear') throw new TypeError('堆积图须使用线性数值轴。')
  let work = 0
  for (let index = 0; index < data.series.length; index++) {
    const series = data.series[index]!, name = series.sourceName ?? series.name
    const style = chartSeriesStyle(chart, series.field, index, undefined, name)
    const valueAxis = style.rightAxis ? chart.options?.rightYAxis ?? { scaleType: 'linear' } : axis
    if (chartSeriesKind(chart, series.field, index, name).stacked && valueAxis.scaleType !== 'linear') throw new TypeError('堆积系列须使用线性数值轴。')
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  if (type === 'radar' && axis.scaleType !== 'linear') throw new TypeError('Radar charts require a linear value axis')
  if (data.visibleFields !== undefined && (!data.categories.length || !data.series.length)) return
  if (descriptor.dataShape === 'hierarchy' && data.hierarchy && !data.hierarchy.length && data.series.length === 1) return
  if (!data.categories.length || !data.series.length) throw new RangeError('Chart source must contain at least one category and value series')
  if (type === 'radar' && data.categories.length < 3) throw new RangeError('Radar charts require at least three categories')
  const paired = type === 'scatter' || type === 'bubble'
  let hasPair = false, pieFinite = false, piePositive = false, pieNegative = false, negativeSize = false, negativeWeight = false
  for (const series of data.series) {
    if (!series.values.length || series.values.length !== data.categories.length) throw new RangeError('Chart series must have one value per category')
    if (paired && (!series.xValues || series.xValues.length !== series.values.length)) throw new RangeError('Scatter series must have one X value per Y value')
    if (paired || descriptor.family === 'pie' || type === 'wordCloud') {
      for (let index = 0; index < series.values.length; index++) {
        const value = series.values[index]!
        if (paired && value !== null && series.xValues?.[index] != null) hasPair = true
        if (descriptor.family === 'pie' && value !== null && Number.isFinite(value)) {
          pieFinite = true; piePositive ||= value > 0; pieNegative ||= value < 0
        }
        if (type === 'wordCloud' && value !== null && value < 0) negativeWeight = true
        if (++work % CHART_PREPARATION_CHUNK === 0) yield
      }
    }
    if (type === 'bubble') for (const value of series.sizeValues ?? []) {
      if (value !== null && value < 0) negativeSize = true
      if (++work % CHART_PREPARATION_CHUNK === 0) yield
    }
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  if (paired && !hasPair) throw new RangeError('Scatter chart requires at least one numeric X/Y pair')
  if (descriptor.family === 'pie') {
    if (data.series.length !== 1) throw new RangeError(`${type} charts require exactly one value series`)
    if (!pieFinite || pieNegative || !piePositive) throw new RangeError(`${type} charts require non-negative values with a positive total`)
  }
  if (negativeSize) throw new RangeError('气泡大小不能为负数。')
  if (type === 'candlestick') {
    let complete = false
    for (const candle of data.candles ?? []) {
      if (candle) { complete = true; break }
      if (++work % CHART_PREPARATION_CHUNK === 0) yield
    }
    if (!complete) throw new RangeError('K 线须至少包含一条完整有效的 HLC 或 OHLC 数据。')
  }
  if (type === 'histogram' && data.series.length !== 1) throw new RangeError('直方图须选择一个数值样本字段。')
  if (descriptor.dataShape === 'hierarchy' && data.series.length !== 1) throw new RangeError('层级图须选择一个数值权重字段。')
  if (type === 'pareto' && data.series.length !== 1) throw new RangeError('帕累托图须选择一个数值权重字段。')
  if ((type === 'funnel' || type === 'wordCloud') && data.series.length !== 1) throw new RangeError('漏斗和词云图须选择一个权重字段。')
  if (negativeWeight) throw new RangeError('此图表的权重不能为负数。')
}
