import type { ChartSemanticOptions } from '../../domain/model'
import type { HistogramBin, HistogramData } from './chart-statistics'
import { chartInterpolate, chartScaledMoments } from './chart-math'
import { formatChartNumber } from './chart-geometry'

export const MAX_CHART_HISTOGRAM_BINS = 1000
export const DEFAULT_CHART_HISTOGRAM_BIN_COUNT = 10
export const DEFAULT_CHART_HISTOGRAM_BIN_WIDTH = 1
export const DEFAULT_CHART_HISTOGRAM_BIN_GAP = 0.007
export const MAX_CHART_HISTOGRAM_BIN_GAP = 0.9
type HistogramOptions = NonNullable<ChartSemanticOptions['histogram']>
type RuntimeOptions = Pick<HistogramOptions, 'binning' | 'underflowThreshold' | 'overflowThreshold'>

/** Runtime fallbacks never rewrite the persisted author configuration. The
 * fixed release counts enabled edge bins in the total visible-bin budget. */
export function chartHistogramRuntimeOptions(options: HistogramOptions = {}): RuntimeOptions {
  const underflowThreshold = Number.isFinite(options.underflowThreshold) ? options.underflowThreshold : undefined
  const overflowThreshold = Number.isFinite(options.overflowThreshold) ? options.overflowThreshold : undefined
  const thresholds = underflowThreshold !== undefined && overflowThreshold !== undefined && underflowThreshold >= overflowThreshold
    ? {} : { underflowThreshold, overflowThreshold }
  const binning = options.binning
  return { ...thresholds, ...(binning?.mode === 'count' && Number.isInteger(binning.count) && binning.count >= 1 && binning.count <= MAX_CHART_HISTOGRAM_BINS
    ? { binning: { ...binning } } : binning?.mode === 'width' && Number.isFinite(binning.width) && binning.width > 0
      ? { binning: { ...binning } } : {}) }
}

export function chartHistogramBinGap(options?: HistogramOptions): number {
  const gap = options?.binGap
  return gap !== undefined && Number.isFinite(gap) && gap >= 0 && gap <= MAX_CHART_HISTOGRAM_BIN_GAP ? gap : DEFAULT_CHART_HISTOGRAM_BIN_GAP
}

/** Sample deviation uses n - 1; normalized moments retain finite extreme
 * observations without overflowing the intermediate sum or range. */
export function chartHistogramSampleStandardDeviation(samples: readonly number[]): number {
  if (samples.length < 2) return 0
  const moments = chartScaledMoments(samples)
  return Math.sqrt(moments.scaledVariance * (samples.length / (samples.length - 1))) * moments.scale
}

export function buildChartHistogram(samples: readonly (number | null)[], options: HistogramOptions = {}): HistogramData {
  const effective = chartHistogramRuntimeOptions(options), regular: number[] = []
  let validSampleCount = 0, underCount = 0, overCount = 0
  for (const value of samples) {
    if (typeof value !== 'number' || !Number.isFinite(value)) continue
    validSampleCount++
    if (effective.underflowThreshold !== undefined && value <= effective.underflowThreshold) underCount++
    else if (effective.overflowThreshold !== undefined && value > effective.overflowThreshold) overCount++
    else regular.push(value)
  }
  if (!validSampleCount) return { validSampleCount: 0, bins: [] }
  const bins: HistogramBin[] = []
  // Without regular observations the fixed release omits empty edge bins.
  if (effective.underflowThreshold !== undefined && (regular.length || underCount)) bins.push({ kind: 'underflow', upperBound: effective.underflowThreshold, count: underCount })
  if (regular.length) bins.push(...regularBins(regular, effective))
  if (effective.overflowThreshold !== undefined && (regular.length || overCount)) bins.push({ kind: 'overflow', lowerBound: effective.overflowThreshold, count: overCount })
  return { validSampleCount, bins }
}

function regularBins(samples: readonly number[], options: RuntimeOptions): HistogramBin[] {
  let min = samples[0]!, max = min
  for (const value of samples) { min = Math.min(min, value); max = Math.max(max, value) }
  const edgeCount = Number(options.underflowThreshold !== undefined) + Number(options.overflowThreshold !== undefined), capacity = MAX_CHART_HISTOGRAM_BINS - edgeCount
  const binning = options.binning
  let boundaries: number[] | undefined
  if (binning?.mode === 'width') {
    const lower = options.underflowThreshold ?? min, upper = options.overflowThreshold ?? (min === max ? min + binning.width : max)
    const scale = Math.max(Math.abs(lower), Math.abs(upper), binning.width)
    const count = Math.max(1, Math.ceil((upper / scale - lower / scale) / (binning.width / scale)))
    if (Number.isFinite(upper) && upper > lower && Number.isInteger(count) && count <= capacity) boundaries = histogramBoundaries(lower, upper, count, binning.width)
  }
  if (!boundaries) {
    if (min === max) return [{ kind: 'regular', lowerBound: min, upperBound: max, count: samples.length }]
    const count = binning?.mode === 'count' && binning.count > edgeCount ? binning.count - edgeCount : automaticCount(samples, min, max, capacity)
    boundaries = histogramBoundaries(min, max, count)
    if (!boundaries) throw new RangeError('直方图箱数无法在当前数值精度内形成有效区间。')
  }
  const bins = boundaries.slice(0, -1).map((lowerBound, index): HistogramBin => ({ kind: 'regular', lowerBound, upperBound: boundaries![index + 1]!, count: 0 }))
  for (const value of samples) {
    let first = 0, last = bins.length - 1
    while (first < last) { const middle = Math.floor((first + last) / 2); if (value <= bins[middle]!.upperBound!) last = middle; else first = middle + 1 }
    bins[first]!.count++
  }
  return bins
}

function automaticCount(samples: readonly number[], min: number, max: number, capacity: number): number {
  const moments = chartScaledMoments(samples), scale = Math.max(Math.abs(min), Math.abs(max), moments.scale)
  const range = max / scale - min / scale
  const width = 3.5 * Math.sqrt(moments.scaledVariance * (samples.length / (samples.length - 1))) * (moments.scale / scale) / Math.cbrt(samples.length)
  return Number.isFinite(width) && width > 0 ? Math.min(capacity, Math.max(1, Math.ceil(range / width))) : 1
}

function histogramBoundaries(lower: number, upper: number, count: number, width?: number): number[] | undefined {
  const boundaries = [lower], scale = Math.max(Math.abs(lower), Math.abs(upper), width ?? 0)
  for (let index = 1; index <= count; index++) {
    const boundary = index === count ? upper : width === undefined ? chartInterpolate(lower, upper, index / count)
      : Math.min(upper, (lower / scale + index * (width / scale)) * scale)
    if (!Number.isFinite(boundary) || boundary <= boundaries.at(-1)!) return undefined
    boundaries.push(boundary)
  }
  return boundaries
}

/** One materialized bin vocabulary is shared by drawing and point editing.
 * Labels do not inspect raw samples or guess effective thresholds from authors. */
export function chartHistogramCategories(histogram?: Pick<HistogramData, 'bins'>): string[] {
  return histogram?.bins.map((bin, index, bins) => {
    if (bin.kind === 'underflow') return `≤ ${formatChartNumber(bin.upperBound!)}`
    if (bin.kind === 'overflow') return `> ${formatChartNumber(bin.lowerBound!)}`
    const previous = bins[index - 1], includesLower = previous?.kind !== 'regular' && (previous?.kind !== 'underflow' || previous.upperBound !== bin.lowerBound)
    return `${includesLower ? '[' : '('}${formatChartNumber(bin.lowerBound!)}, ${formatChartNumber(bin.upperBound!)}]`
  }) ?? []
}
