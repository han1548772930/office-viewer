import type { ChartErrorBarOptions, ChartModel, ChartSemanticOptions, ChartTrendlineOptions } from '../../domain/model'
import type { ChartData, ChartSeriesData } from './chart-engine'
import { chartNumber } from './chart-data'
import { chartFiniteProduct, chartFiniteSum, chartInterpolate, chartScaledMoments, finiteChartValue } from './chart-math'
import { validateChartOptions } from './chart-options'

export interface HistogramBin { kind: 'underflow' | 'regular' | 'overflow'; lowerBound?: number; upperBound?: number; count: number }
export interface HistogramData { validSampleCount: number; bins: HistogramBin[] }
export interface BoxplotData { min: number; q1: number; median: number; q3: number; max: number; outliers: number[] }
export interface ParetoData { categories: string[]; values: number[]; cumulativePercent: number[] }
export interface WaterfallPoint { start: number; end: number; value: number; role: 'change' | 'total'; categoryIndex: number }
export interface TrendlineData { points: Array<[number, number]>; formula?: string; r2?: number }
/** Coefficients fit centered normalized Y to centered/spread normalized X.
 * Log/exp/power transforms are applied before these scale factors. */
export interface ChartTrendlineFit { coefficients: number[]; xScale: number; xMean: number; xSpread: number; yScale: number; yMean: number }
export interface ChartErrorInterval { categoryIndex: number; x: number; y: number; minus: number; plus: number }

/** Small-data reference contract. Worker/Rust must use the same boundaries and statistics. */
export function chartHistogram(samples: readonly number[], options: ChartSemanticOptions['histogram'] = {}): HistogramData {
  validateChartOptions({ histogram: options })
  const valid = samples.filter(Number.isFinite)
  if (!valid.length) return { validSampleCount: 0, bins: [] }
  const regular = valid.filter(value => (options.underflowThreshold === undefined || value > options.underflowThreshold)
    && (options.overflowThreshold === undefined || value <= options.overflowThreshold))
  const bins: HistogramBin[] = []
  if (options.underflowThreshold !== undefined) bins.push({ kind: 'underflow', upperBound: options.underflowThreshold, count: valid.filter(value => value <= options.underflowThreshold!).length })
  if (regular.length) {
    let min = regular[0]!, max = min
    for (const value of regular) { min = Math.min(min, value); max = Math.max(max, value) }
    let lower = options.underflowThreshold ?? min, upper = options.overflowThreshold ?? max
    if (lower === upper) {
      const padding = options.binning?.mode === 'width' ? options.binning.width : Math.max(1, Math.abs(lower) * 0.1)
      const after = upper + padding, before = lower - padding
      if (options.overflowThreshold === undefined && Number.isFinite(after) && after > upper) upper = after
      else if (Number.isFinite(before) && before < lower) lower = before
      else throw new RangeError('直方图箱宽无法在当前数值精度内形成有效区间。')
    }
    const moments = chartScaledMoments(regular), scale = Math.max(Math.abs(lower), Math.abs(upper), moments.scale, options.binning?.mode === 'width' ? options.binning.width : 0)
    const range = upper / scale - lower / scale
    const scottWidth = 3.5 * Math.sqrt(moments.scaledVariance) * (moments.scale / scale) / Math.cbrt(regular.length)
    const defaultCount = range === 0 ? 1 : Math.max(1, Math.ceil(range / (scottWidth > 0 ? scottWidth : range / Math.sqrt(regular.length))))
    const count = options.binning?.mode === 'count' ? options.binning.count : options.binning?.mode === 'width' ? Math.max(1, Math.ceil(range / (options.binning.width / scale))) : min === max ? 1 : defaultCount
    if (!Number.isSafeInteger(count) || count < 1 || count > 10_000) throw new RangeError('直方图分箱数量须位于 1 到 10000。')
    const boundary = (index: number): number => {
      if (index === 0) return lower
      if (options.binning?.mode === 'width') {
        if (index === count && options.overflowThreshold !== undefined) return upper
        const normalized = lower / scale + index * (options.binning.width / scale)
        return finiteChartValue(normalized * scale, '直方图箱边界')
      }
      return chartInterpolate(lower, upper, index / count)
    }
    const entries = Array.from({ length: count }, (_unused, index): HistogramBin => ({ kind: 'regular', lowerBound: boundary(index), upperBound: boundary(index + 1), count: 0 }))
    if (entries.some(entry => entry.lowerBound! >= entry.upperBound!)) throw new RangeError('直方图箱数或箱宽无法在当前数值精度内形成有效区间。')
    for (const value of regular) {
      // The first interval includes its lower bound; later bins are (lower, upper].
      let first = 0, last = count - 1
      while (first < last) { const middle = Math.floor((first + last) / 2); if (value <= entries[middle]!.upperBound!) last = middle; else first = middle + 1 }
      entries[first]!.count++
    }
    bins.push(...entries)
  }
  if (options.overflowThreshold !== undefined) bins.push({ kind: 'overflow', lowerBound: options.overflowThreshold, count: valid.filter(value => value > options.overflowThreshold!).length })
  return { validSampleCount: valid.length, bins }
}

export function chartBoxplot(samples: readonly number[], options: ChartSemanticOptions['boxplot'] = {}): BoxplotData | undefined {
  validateChartOptions({ boxplot: options })
  const sorted = samples.filter(Number.isFinite).sort((a, b) => a - b)
  if (!sorted.length) return undefined
  const quartile = (p: number): number => {
    const position = options.quartileMethod === 'exclusive' ? p * (sorted.length + 1) - 1 : p * (sorted.length - 1)
    const lower = Math.max(0, Math.min(sorted.length - 1, Math.floor(position))), upper = Math.max(0, Math.min(sorted.length - 1, Math.ceil(position)))
    return chartInterpolate(sorted[lower]!, sorted[upper]!, Math.max(0, Math.min(1, position - Math.floor(position))))
  }
  const q1 = quartile(0.25), q3 = quartile(0.75), median = quartile(0.5)
  const scale = Math.max(Math.abs(q1), Math.abs(q3), ...[Math.abs(sorted[0]!), Math.abs(sorted.at(-1)!)]) || 1
  const distance = (q3 / scale - q1 / scale) * (options.whiskerMultiplier ?? 1.5)
  const isWithin = (value: number): boolean => value / scale >= q1 / scale - distance && value / scale <= q3 / scale + distance
  const within = sorted.filter(isWithin)
  return { min: within[0] ?? q1, q1, median, q3, max: within.at(-1) ?? q3, outliers: sorted.filter(value => !isWithin(value)) }
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
  const total = sorted.reduce((sum, [_name, group]) => sum + group.value / scale, 0)
  let cumulative = 0
  return { categories: sorted.map(([name]) => name), values: sorted.map(([_name, group]) => group.value), cumulativePercent: sorted.map(([_name, group], index) => { cumulative += group.value / scale; return total ? index === sorted.length - 1 ? 100 : Math.min(100, cumulative / total * 100) : 0 }) }
}

export function chartWaterfall(series: ChartSeriesData, chart: ChartModel): WaterfallPoint[] {
  validateChartOptions({ waterfall: chart.options?.waterfall })
  let running = 0
  return series.values.map((value, categoryIndex) => {
    const numeric = finiteChartValue(value ?? 0, '瀑布图数据')
    const role = chart.options?.waterfall?.pointRoles?.find(point => point.seriesIndex === series.field && point.dataPointIndex === categoryIndex)?.role ?? 'change'
    const start = role === 'total' ? 0 : running
    const end = role === 'total' ? numeric : finiteChartValue(running + numeric, '瀑布图累计值')
    running = end
    return { start, end, value: numeric, role, categoryIndex }
  })
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
