import type { ChartErrorBarOptions, ChartModel, ChartSemanticOptions, ChartTrendlineOptions } from '../../domain/model'
import type { ChartData, ChartSeriesData } from './chart-engine'
import { chartNumber } from './chart-data'
import { chartFiniteSum, chartInterpolate, chartScaledMoments, finiteChartValue } from './chart-math'
import { validateChartOptions } from './chart-options'

export interface HistogramBin { kind: 'underflow' | 'regular' | 'overflow'; lowerBound?: number; upperBound?: number; count: number }
export interface HistogramData { validSampleCount: number; bins: HistogramBin[] }
export interface BoxplotData { min: number; q1: number; median: number; q3: number; max: number; outliers: number[] }
export interface ParetoData { categories: string[]; values: number[]; cumulativePercent: number[] }
export interface WaterfallPoint { start: number; end: number; value: number; role: 'change' | 'total'; categoryIndex: number }
export interface TrendlineData { points: Array<[number, number]>; formula?: string; r2?: number }
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
    const lower = options.underflowThreshold ?? min
    const upper = options.overflowThreshold ?? max
    const moments = chartScaledMoments(regular), scale = Math.max(Math.abs(lower), Math.abs(upper), moments.scale)
    const range = upper / scale - lower / scale
    const scottWidth = 3.5 * Math.sqrt(moments.scaledVariance) * (moments.scale / scale) / Math.cbrt(regular.length)
    const defaultCount = range === 0 ? 1 : Math.max(1, Math.ceil(range / (scottWidth > 0 ? scottWidth : range / Math.sqrt(regular.length))))
    const count = range === 0 ? 1 : options.binning?.mode === 'count' ? options.binning.count : options.binning?.mode === 'width' ? Math.max(1, Math.ceil(range / (options.binning.width / scale))) : defaultCount
    if (!Number.isSafeInteger(count) || count < 1 || count > 10_000) throw new RangeError('直方图分箱数量须位于 1 到 10000。')
    const boundary = (index: number): number => {
      if (options.binning?.mode === 'width') return finiteChartValue((lower / scale + index * (options.binning.width / scale)) * scale, '直方图箱边界')
      if (range > 0) return chartInterpolate(lower, upper, index / count)
      const last = lower + Math.max(1, Math.abs(lower) * 0.1)
      return index === 0 || !Number.isFinite(last) ? lower : last
    }
    const entries = Array.from({ length: count }, (_unused, index): HistogramBin => ({ kind: 'regular', lowerBound: boundary(index), upperBound: boundary(index + 1), count: 0 }))
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
    if (value < 0) throw new RangeError('帕累托图权重不能为负数。')
    const name = categories[index] ?? String(index + 1), group = grouped.get(name) ?? { value: 0, index }
    group.value = finiteChartValue(group.value + value, '帕累托分类合计'); grouped.set(name, group)
  })
  const sorted = [...grouped].filter(([_name, group]) => includeZeroValues || group.value > 0).sort((a, b) => b[1].value - a[1].value || a[1].index - b[1].index)
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
      points.push([x, chartFiniteSum(values.map(value => value! / period), '移动平均')])
    }
    return { points }
  }
  if (pairs.length < 2) return { points: [] }
  const transformX = (x: number): number => options.type === 'log' || options.type === 'pow' ? Math.log(x) : x
  const transformY = (y: number): number => options.type === 'exp' || options.type === 'pow' ? Math.log(y) : y
  const degree = options.type === 'poly' ? options.order ?? 2 : 1
  if (pairs.length <= degree) return { points: [] }
  const xs = pairs.map(([x]) => transformX(x)), ys = pairs.map(([_x, y]) => transformY(y))
  const meanX = xs.reduce((sum, x) => sum + x / xs.length, 0)
  let scale = 0
  for (const x of xs) scale = Math.max(scale, Math.abs(x - meanX))
  if (!scale) return { points: [] }
  const coefficients = polynomialFit(xs.map(x => (x - meanX) / scale), ys, degree)
  if (!coefficients) return { points: [] }
  const predict = (x: number): number => {
    const argument = (transformX(x) - meanX) / scale
    const output = coefficients.reduceRight((sum, coefficient) => sum * argument + coefficient, 0)
    return options.type === 'exp' || options.type === 'pow' ? Math.exp(output) : output
  }
  let minX = pairs[0]![0], maxX = minX
  for (const [x] of pairs) { minX = Math.min(minX, x); maxX = Math.max(maxX, x) }
  const points = Array.from({ length: 33 }, (_unused, index): [number, number] => { const x = chartInterpolate(minX, maxX, index / 32); return [x, predict(x)] }).filter(([x, y]) => Number.isFinite(x) && Number.isFinite(y))
  const meanY = pairs.reduce((sum, [_x, y]) => sum + y / pairs.length, 0)
  const total = pairs.reduce((sum, [_x, y]) => sum + (y - meanY) ** 2, 0), residual = pairs.reduce((sum, [x, y]) => sum + (y - predict(x)) ** 2, 0)
  const slope = coefficients[1]! / scale, intercept = coefficients[0]! - slope * meanX
  const formatted = (value: number): string => String(Number(value.toPrecision(5)))
  const formula = options.type === 'poly' ? `y = ${coefficients.map((coefficient, power) => `${formatted(coefficient)}·((x−${formatted(meanX)})/${formatted(scale)})^${power}`).join(' + ')}`
    : options.type === 'exp' ? `y = ${formatted(Math.exp(intercept))}·e^(${formatted(slope)}x)`
      : options.type === 'log' ? `y = ${formatted(slope)}·ln(x) + ${formatted(intercept)}`
        : options.type === 'pow' ? `y = ${formatted(Math.exp(intercept))}·x^${formatted(slope)}`
          : `y = ${formatted(slope)}x + ${formatted(intercept)}`
  return { points, formula, r2: total === 0 ? residual < 1e-12 ? 1 : 0 : 1 - residual / total }
}

function polynomialFit(xs: number[], ys: number[], degree: number): number[] | undefined {
  const count = degree + 1
  const matrix = Array.from({ length: count }, (_unused, row) => Array.from({ length: count + 1 }, (_value, column) => column === count
    ? xs.reduce((sum, x, index) => sum + x ** row * ys[index]!, 0)
    : xs.reduce((sum, x) => sum + x ** (row + column), 0)))
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
  return matrix.map(row => row[count]!)
}

export function chartErrorIntervals(data: ChartData, series: ChartSeriesData, options: ChartErrorBarOptions): ChartErrorInterval[] {
  validateChartOptions({ errorBars: [options] })
  const samples = (options.direction === 'x' ? series.xValues ?? [] : series.values).filter((value): value is number => value !== null && Number.isFinite(value))
  if (!samples.length) return []
  const moments = chartScaledMoments(samples)
  const deviation = samples.length > 1 ? finiteChartValue(moments.deviation * Math.sqrt(samples.length / (samples.length - 1)), '样本标准差') : 0
  const factor = options.value ?? 1
  return series.values.flatMap((y, categoryIndex) => {
    const x = series.xValues === undefined ? categoryIndex : series.xValues[categoryIndex]
    if (y === null || x == null || !Number.isFinite(x) || !Number.isFinite(y)) return []
    const target = options.direction === 'x' ? x : y
    const error = options.valueType === 'fixedVal' ? factor : options.valueType === 'percentage' ? Math.abs(target) * (factor / 100)
      : options.valueType === 'stdDev' ? deviation * factor : options.valueType === 'stdErr' ? deviation / Math.sqrt(samples.length) * factor : 0
    const record = data.records?.[categoryIndex]
    const plus = options.valueType === 'cust' ? options.plusIndex === undefined ? 0 : chartNumber(record?.[options.plusIndex]) ?? 0 : error
    const minus = options.valueType === 'cust' ? options.minusIndex === undefined ? 0 : chartNumber(record?.[options.minusIndex]) ?? 0 : error
    if (plus < 0 || minus < 0) throw new RangeError('自定义误差值不能为负数。')
    finiteChartValue(plus, '正误差值'); finiteChartValue(minus, '负误差值')
    finiteChartValue(target + plus, '误差上界'); finiteChartValue(target - minus, '误差下界')
    return [{ categoryIndex, x, y, plus: options.type === 'minus' ? 0 : plus, minus: options.type === 'plus' ? 0 : minus }]
  })
}
