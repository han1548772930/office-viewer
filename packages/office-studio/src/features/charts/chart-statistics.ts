import type { ChartErrorBarOptions, ChartModel, ChartSemanticOptions, ChartTrendlineOptions, ChartWaterfallStackType } from '../../domain/model'
import { ComputeError } from '../../compute/errors'
import type { ChartData, ChartSeriesData } from './chart-engine'
import { chartNumber } from './chart-data'
import { chartFiniteProduct, chartFiniteSum, chartInterpolate, chartScaledMoments, finiteChartValue } from './chart-math'
import { validateChartOptions } from './chart-options'
import { chartBoxplotStatisticOptions } from './chart-boxplot-options'
import { buildChartHistogram } from './chart-histogram'

export interface HistogramBin { kind: 'underflow' | 'regular' | 'overflow'; lowerBound?: number; upperBound?: number; count: number }
export interface HistogramData { validSampleCount: number; bins: HistogramBin[] }
export interface BoxplotData {
  min: number; q1: number; median: number; q3: number; max: number
  outliers: number[]
  /** Parallel identities in preparedData, including original empty sample slots. */
  outlierSourceIndices: number[]
}
export interface ParetoData { categories: string[]; values: number[]; cumulativePercent: number[] }
export interface WaterfallPoint { start: number; end: number; value: number; role: 'change' | 'total'; categoryIndex: number }
export const CHART_WATERFALL_GLOBAL_SERIES = 0xffff_ffff
export interface WaterfallLayoutPoint extends WaterfallPoint {
  seriesPosition: number
  /** A source row for real points; sourceCategoryCount for automatic subtotals. */
  dataPointIndex: number
  synthetic: boolean
}
export interface WaterfallLayoutConnector {
  from: { categoryIndex: number; value: number }
  to: { categoryIndex: number; value: number }
}
export interface WaterfallLayoutData {
  stackType: ChartWaterfallStackType
  useSubtotal: boolean
  sourceCategoryCount: number
  seriesFields: number[]
  points: WaterfallLayoutPoint[]
  connectors: WaterfallLayoutConnector[]
  cumulative: number
}
export interface TrendlineData { points: Array<[number, number]>; formula?: string; r2?: number }
/** Coefficients fit centered normalized Y to centered/spread normalized X.
 * Log/exp/power transforms are applied before these scale factors. */
export interface ChartTrendlineFit { coefficients: number[]; xScale: number; xMean: number; xSpread: number; yScale: number; yMean: number }
export interface ChartErrorInterval { categoryIndex: number; x: number; y: number; minus: number; plus: number }

/** Small-data reference contract. Worker/Rust must use the same boundaries and statistics. */
export function chartHistogram(samples: readonly number[], options: ChartSemanticOptions['histogram'] = {}): HistogramData {
  validateChartOptions({ histogram: options })
  return buildChartHistogram(samples, options)
}

export function chartBoxplot(samples: readonly (number | null)[], options: ChartSemanticOptions['boxplot'] = {}): BoxplotData | undefined {
  validateChartOptions({ boxplot: options })
  const sorted = samples.flatMap((value, sourceIndex) => value !== null && Number.isFinite(value) ? [{ value, sourceIndex }] : [])
    .sort((a, b) => a.value - b.value || a.sourceIndex - b.sourceIndex)
  if (!sorted.length) return undefined
  const effective = chartBoxplotStatisticOptions(options)
  // The fixed converter's empty halves yield NaN for a single sample, so it
  // has no drawable box. Keep the finite protocol instead of emitting NaN.
  if (effective.quartileMethod === 'medianOfHalves' && sorted.length < 2) return undefined
  const halfMedian = (start: number, end: number): number => {
    const length = end - start, middle = start + Math.floor(length / 2)
    return length % 2 ? sorted[middle]!.value : chartInterpolate(sorted[middle - 1]!.value, sorted[middle]!.value, 0.5)
  }
  const quartile = (p: number): number => {
    if (effective.quartileMethod === 'medianOfHalves') return p === 0.25 ? halfMedian(0, Math.floor(sorted.length / 2))
      : p === 0.75 ? halfMedian(Math.ceil(sorted.length / 2), sorted.length) : halfMedian(0, sorted.length)
    const position = effective.quartileMethod === 'exclusive' ? p * (sorted.length + 1) - 1 : p * (sorted.length - 1)
    const lower = Math.max(0, Math.min(sorted.length - 1, Math.floor(position))), upper = Math.max(0, Math.min(sorted.length - 1, Math.ceil(position)))
    return chartInterpolate(sorted[lower]!.value, sorted[upper]!.value, Math.max(0, Math.min(1, position - Math.floor(position))))
  }
  const q1 = quartile(0.25), q3 = quartile(0.75), median = quartile(0.5)
  if (effective.whiskerMethod === 'range') return { min: sorted[0]!.value, q1, median, q3, max: sorted.at(-1)!.value,
    outliers: [], outlierSourceIndices: [] }
  const scale = Math.max(Math.abs(q1), Math.abs(q3), Math.abs(sorted[0]!.value), Math.abs(sorted.at(-1)!.value)) || 1
  const distance = (q3 / scale - q1 / scale) * effective.whiskerMultiplier
  const isWithin = (value: number): boolean => value / scale >= q1 / scale - distance && value / scale <= q3 / scale + distance
  const within = sorted.filter(sample => isWithin(sample.value)), outliers = sorted.filter(sample => !isWithin(sample.value))
  return { min: within[0]?.value ?? q1, q1, median, q3, max: within.at(-1)?.value ?? q3,
    outliers: outliers.map(sample => sample.value), outlierSourceIndices: outliers.map(sample => sample.sourceIndex) }
}

export function chartPareto(categories: readonly string[], samples: readonly (number | null)[], includeZeroValues = false): ParetoData {
  const grouped = new Map<string, { value: number; index: number }>()
  samples.forEach((value, index) => {
    if (value === null || !Number.isFinite(value)) return
    const name = categories[index] ?? String(index + 1), group = grouped.get(name) ?? { value: 0, index }
    group.value = finiteChartValue(group.value + value, '帕累托分类合计'); grouped.set(name, group)
  })
  // Fixed Pareto builder contract aggregates signed rows before filtering.
  // includeZeroValues retains zero totals; negative totals are always excluded.
  const sorted = [...grouped].filter(([_name, group]) => group.value > 0 || includeZeroValues && group.value === 0).sort((a, b) => b[1].value - a[1].value || a[1].index - b[1].index)
  const scale = sorted[0]?.[1].value || 1
  let rawTotal = 0, scaledTotal = 0
  for (const [_name, group] of sorted) { rawTotal += group.value; scaledTotal += group.value / scale }
  // Fixed OOe truncates percentages to two decimal places. Preserve its raw
  // addition/division for finite totals: max-value normalization alone can
  // turn an exact 60% for [6, 4] into 59.99%. Scaling only handles overflow.
  const scaled = !Number.isFinite(rawTotal), total = scaled ? scaledTotal : rawTotal
  let cumulative = 0
  return { categories: sorted.map(([name]) => name), values: sorted.map(([_name, group]) => group.value),
    cumulativePercent: sorted.map(([_name, group], index) => {
      cumulative += scaled ? group.value / scale : group.value
      return total ? index === sorted.length - 1 ? 100 : Math.min(100, Math.floor(cumulative / total * 10_000) / 100) : 0
    }) }
}

/** Legacy single-series projection, never a second accumulation engine. */
export function chartWaterfall(series: ChartSeriesData, chart: ChartModel, computed?: WaterfallLayoutData): WaterfallPoint[] {
  validateChartOptions({ waterfall: chart.options?.waterfall })
  if (!computed) throw new ComputeError('COMPUTE_NOT_READY', '瀑布图布局需要先完成 Rust/WASM 计算。')
  const positions = computed.seriesFields.flatMap((field, position) => field === series.field ? [position] : [])
  if (positions.length !== 1 || computed.sourceCategoryCount !== series.values.length) throw new RangeError('瀑布图投影与当前来源系列不匹配。')
  return computed.points.filter(point => point.seriesPosition === positions[0] && !point.synthetic).map(point => ({
    start: point.start, end: point.end, value: point.value, role: point.role, categoryIndex: point.dataPointIndex,
  }))
}

export function chartTrendline(series: ChartSeriesData, options: ChartTrendlineOptions): TrendlineData {
  validateChartOptions({ trendlines: [options] })
  const pairs = series.values.flatMap((y, index): Array<[number, number]> => {
    const x = series.xValues === undefined ? index + 1 : series.xValues[index]
    if (y === null || x == null || !Number.isFinite(x) || !Number.isFinite(y)) return []
    if ((options.type === 'log' || options.type === 'pow') && x <= 0 || (options.type === 'exp' || options.type === 'pow') && y <= 0) return []
    return [[x, y]]
  })
  if (options.type === 'movingAvg') {
    const period = options.period ?? 2
    const points: Array<[number, number]> = []
    for (let index = period - 1; index < series.values.length; index++) {
      const values = series.values.slice(index - period + 1, index + 1)
      const x = series.xValues === undefined ? index + 1 : series.xValues[index]
      if (x == null || !Number.isFinite(x) || values.some(value => value === null || !Number.isFinite(value))
        || series.xValues?.slice(index - period + 1, index + 1).some(value => value == null || !Number.isFinite(value))) continue
      points.push([x, chartScaledMoments(values.map(value => value!)).mean])
    }
    return { points }
  }
  if (pairs.length < 2) return { points: [] }
  const transformX = (x: number): number => options.type === 'log' || options.type === 'pow' ? Math.log(x) : x
  const transformY = (y: number): number => options.type === 'exp' || options.type === 'pow' ? Math.log(y) : y
  const degree = options.type === 'poly' ? options.order ?? 2 : 1
  if (pairs.length <= degree) return { points: [] }
  const xs = pairs.map(([x]) => transformX(x)), ys = pairs.map(([_x, y]) => transformY(y))
  const xMoments = chartScaledMoments(xs), yMoments = chartScaledMoments(ys)
  let spread = 0
  for (const x of xs) spread = Math.max(spread, Math.abs(x / xMoments.scale - xMoments.scaledMean))
  if (!spread) return { points: [] }
  const coefficients = polynomialFit(xs.map(x => (x / xMoments.scale - xMoments.scaledMean) / spread), ys.map(y => y / yMoments.scale - yMoments.scaledMean), degree)
  if (!coefficients) return { points: [] }
  const predictScaled = (x: number): number => {
    const argument = (transformX(x) / xMoments.scale - xMoments.scaledMean) / spread
    return yMoments.scaledMean + coefficients.reduceRight((sum, coefficient) => sum * argument + coefficient, 0)
  }
  const exponential = options.type === 'exp' || options.type === 'pow'
  const predict = (x: number): number => exponential ? Math.exp(predictScaled(x) * yMoments.scale) : predictScaled(x) * yMoments.scale
  let minX = pairs[0]![0], maxX = minX
  for (const [x] of pairs) { minX = Math.min(minX, x); maxX = Math.max(maxX, x) }
  const points = Array.from({ length: 33 }, (_unused, index): [number, number] => { const x = chartInterpolate(minX, maxX, index / 32); return [x, predict(x)] }).filter(([x, y]) => Number.isFinite(x) && Number.isFinite(y))
  const originalMoments = exponential ? chartScaledMoments(pairs.map(([_x, y]) => y)) : yMoments
  const relativePredictions = pairs.map(([x]) => exponential ? Math.exp(predictScaled(x) * yMoments.scale - Math.log(originalMoments.scale)) : predictScaled(x))
  const r2 = trendlineRSquared(pairs.map(([_x, y]) => y / originalMoments.scale), relativePredictions)
  const formula = formatChartTrendlineEquation(options.type, { coefficients, xScale: xMoments.scale, xMean: xMoments.scaledMean, xSpread: spread, yScale: yMoments.scale, yMean: yMoments.scaledMean })
  return { points, formula, ...(r2 === undefined ? {} : { r2 }) }
}

/** Rust returns fit metadata; both numerical paths use this one formatter. */
export function formatChartTrendlineEquation(type: ChartTrendlineOptions['type'], fit: ChartTrendlineFit): string | undefined {
  if (type === 'movingAvg') return undefined
  for (const value of [fit.xScale, fit.xMean, fit.xSpread, fit.yScale, fit.yMean, ...fit.coefficients]) finiteChartValue(value, '趋势拟合参数')
  if (fit.coefficients.length < 2 || fit.coefficients.length > 7 || fit.xScale <= 0 || fit.xSpread <= 0 || fit.xSpread > 2 || fit.yScale <= 0 || Math.abs(fit.xMean) > 1 || Math.abs(fit.yMean) > 1) throw new RangeError('趋势拟合参数无效。')
  const { coefficients, xScale, xMean, xSpread, yScale, yMean } = fit
  const formatted = (value: number): string => { const rounded = Number(value.toPrecision(5)); return String(Number.isFinite(rounded) ? rounded : value) }
  const variable = type === 'log' || type === 'pow' ? 'ln(x)' : 'x', exponential = type === 'exp' || type === 'pow'
  const argument = `((${variable}/${formatted(xScale)}−${formatted(xMean)})/${formatted(xSpread)})`
  const terms = [formatted(yMean + coefficients[0]!), ...coefficients.slice(1).map((coefficient, index) => `${formatted(coefficient)}·${argument}^${index + 1}`)]
  const normalizedFormula = `${formatted(yScale)}·(${terms.join(' + ')})`
  let formula = `y = ${exponential ? `e^(${normalizedFormula})` : normalizedFormula}`
  if (type !== 'poly') {
    const span = xScale * xSpread, reciprocal = 1 / span
    const slope = Number.isFinite(span) && Number.isFinite(reciprocal) ? finiteTrendlineCoefficient([coefficients[1]!, yScale, reciprocal]) : undefined
    const intercept = finiteTrendlineCoefficient([yScale, yMean + coefficients[0]! - coefficients[1]! * xMean / xSpread])
    if (slope !== undefined && intercept !== undefined) {
      if (exponential) {
        const coefficient = Math.exp(intercept)
        formula = type === 'exp'
          ? Number.isFinite(coefficient) && coefficient > 0 ? `y = ${formatted(coefficient)}·e^(${formatted(slope)}x)` : `y = e^(${formatted(intercept)} + ${formatted(slope)}x)`
          : `y = ${Number.isFinite(coefficient) && coefficient > 0 ? formatted(coefficient) : `e^(${formatted(intercept)})`}·x^${formatted(slope)}`
      } else formula = `y = ${formatted(slope)}·${variable} + ${formatted(intercept)}`
    }
  }
  return formula
}

/** A conventional equation coefficient may be unrepresentable while its
 * normalized equation and the plotted predictions remain finite. */
function finiteTrendlineCoefficient(factors: number[]): number | undefined {
  try { return chartFiniteProduct(factors, '趋势公式系数') }
  catch (error) { if (error instanceof RangeError) return undefined; throw error }
}

function trendlineRSquared(actual: number[], predicted: number[]): number | undefined {
  if (predicted.some(value => !Number.isFinite(value))) return undefined
  const mean = chartScaledMoments(actual).mean
  const norm = (values: number[]): { scale: number; sum: number } => {
    const scale = values.reduce((maximum, value) => Math.max(maximum, Math.abs(value)), 0)
    return { scale, sum: scale ? chartFiniteSum(values.map(value => (value / scale) ** 2)) : 0 }
  }
  const total = norm(actual.map(value => value - mean)), residual = norm(actual.map((value, index) => value - predicted[index]!))
  if (!total.scale) return residual.scale <= 32 * Number.EPSILON ? 1 : 0
  if (!residual.scale) return 1
  const ratio = Math.exp(2 * (Math.log(residual.scale) - Math.log(total.scale)) + Math.log(residual.sum) - Math.log(total.sum))
  const r2 = 1 - ratio
  return Number.isFinite(r2) ? r2 : undefined
}

function polynomialFit(xs: number[], ys: number[], degree: number): number[] | undefined {
  const count = degree + 1
  const matrix = Array.from({ length: count }, (_unused, row) => Array.from({ length: count + 1 }, (_value, column) => column === count
    ? chartFiniteSum(xs.map((x, index) => x ** row * ys[index]!)) / xs.length
    : chartFiniteSum(xs.map(x => x ** (row + column))) / xs.length))
  for (let column = 0; column < count; column++) {
    let pivot = column
    for (let row = column + 1; row < count; row++) if (Math.abs(matrix[row]![column]!) > Math.abs(matrix[pivot]![column]!)) pivot = row
    if (Math.abs(matrix[pivot]![column]!) < 1e-12) return undefined
    ;[matrix[column], matrix[pivot]] = [matrix[pivot]!, matrix[column]!]
    const divisor = matrix[column]![column]!
    for (let entry = column; entry <= count; entry++) matrix[column]![entry]! /= divisor
    for (let row = 0; row < count; row++) {
      if (row === column) continue
      const multiplier = matrix[row]![column]!
      for (let entry = column; entry <= count; entry++) matrix[row]![entry]! -= multiplier * matrix[column]![entry]!
    }
  }
  const coefficients = matrix.map(row => row[count]!)
  return coefficients.every(Number.isFinite) ? coefficients : undefined
}

export function chartErrorIntervals(data: ChartData, series: ChartSeriesData, options: ChartErrorBarOptions): ChartErrorInterval[] {
  validateChartOptions({ errorBars: [options] })
  if (options.direction === 'x' && series.xValues === undefined) throw new RangeError('X 方向误差线需要数值 X 字段。')
  const usesPlus = options.type !== 'minus', usesMinus = options.type !== 'plus'
  if (options.valueType === 'cust') {
    if (!data.records || data.records.length !== data.categories.length || series.values.length !== data.categories.length) throw new RangeError('自定义误差线需要与当前数据点逐项对应的源记录；请调整分类聚合或误差字段。')
    const fieldCount = data.fields?.length ?? data.records[0]?.length ?? 0
    for (const field of [usesPlus ? options.plusIndex : undefined, usesMinus ? options.minusIndex : undefined]) if (field !== undefined && field >= fieldCount) throw new RangeError('自定义误差字段超出数据源范围。')
  }
  const samples = series.values.flatMap((y, index) => {
    const x = series.xValues === undefined ? index : series.xValues[index]
    return y !== null && x != null && Number.isFinite(x) && Number.isFinite(y) ? [options.direction === 'x' ? x : y] : []
  })
  if (!samples.length) return []
  const factor = options.value ?? 1
  let statisticalError = 0
  if ((options.valueType === 'stdDev' || options.valueType === 'stdErr') && factor && samples.length > 1) {
    const moments = chartScaledMoments(samples)
    const multiplier = options.valueType === 'stdDev' ? samples.length / (samples.length - 1) : 1 / (samples.length - 1)
    statisticalError = chartFiniteProduct([moments.scale, Math.sqrt(moments.scaledVariance * multiplier), factor], '统计误差值')
  }
  return series.values.flatMap((y, categoryIndex) => {
    const x = series.xValues === undefined ? categoryIndex : series.xValues[categoryIndex]
    if (y === null || x == null || !Number.isFinite(x) || !Number.isFinite(y)) return []
    const target = options.direction === 'x' ? x : y
    const error = options.valueType === 'fixedVal' ? factor : options.valueType === 'percentage' ? chartFiniteProduct([Math.abs(target), factor, 0.01], '百分比误差值') : statisticalError
    const record = data.records?.[categoryIndex]
    const plus = !usesPlus ? 0 : options.valueType === 'cust' ? chartNumber(record?.[options.plusIndex!]) : error
    const minus = !usesMinus ? 0 : options.valueType === 'cust' ? chartNumber(record?.[options.minusIndex!]) : error
    // Missing/invalid authored error cells omit the point's interval. Zero is
    // reserved for an actual zero or an unselected error direction.
    if (plus === null || minus === null) return []
    if (plus < 0 || minus < 0) throw new RangeError('自定义误差值不能为负数。')
    finiteChartValue(plus, '正误差值'); finiteChartValue(minus, '负误差值')
    if (usesPlus) finiteChartValue(target + plus, '误差上界')
    if (usesMinus) finiteChartValue(target - minus, '误差下界')
    return [{ categoryIndex, x, y, plus, minus }]
  })
}
