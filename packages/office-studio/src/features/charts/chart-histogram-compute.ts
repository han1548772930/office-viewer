import type { ChartStatisticsJob, ChartStatisticsOutput, ChartStatisticsSegment } from '../../compute/chart-protocol'
import { ComputeError } from '../../compute/errors'
import { CHART_PREPARATION_CHUNK } from './chart-compute-lifecycle'
import { chartHistogramRuntimeOptions, MAX_CHART_HISTOGRAM_BINS } from './chart-histogram'
import type { ChartStatisticsPlan } from './chart-statistics-compute'
import type { HistogramBin, HistogramData } from './chart-statistics'

type HistogramJob = Extract<ChartStatisticsJob, { kind: 'histogram' }>
const reject = (message: string): never => { throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', message) }

/** Validate materialized Rust bins against retained source observations. This
 * does not recalculate Scott deviation or create a second set of boundaries. */
export function* decodeChartHistogramSteps(plan: ChartStatisticsPlan, output: ChartStatisticsOutput,
  segment: ChartStatisticsSegment, job: HistogramJob): Generator<void, HistogramData, void> {
  const source = plan.preparedData.series[segment.seriesPosition]
  if (!source || source.field !== segment.field) reject('直方图来源系列身份无效。')
  const effective = chartHistogramRuntimeOptions(job), bins: HistogramBin[] = []
  const valueAt = (offset: number): number => {
    const value = output.values[segment.valueOffset + offset]
    if (value === undefined || !Number.isFinite(value)) reject('直方图结果包含非有限边界。')
    return value
  }
  let count = 0, regularStart = -1, regularLength = 0, previousUpper: number | undefined, work = 0
  if (segment.valueLength / 2 > MAX_CHART_HISTOGRAM_BINS) reject('直方图分箱数量超出预算。')
  for (let offset = 0; offset < segment.valueLength; offset += 2) {
    const lower = valueAt(offset), upper = valueAt(offset + 1)
    const kind = output.indices[segment.indexOffset + offset]!, samples = output.indices[segment.indexOffset + offset + 1]!
    count += samples
    if (count > segment.validSampleCount) reject('直方图分箱计数超出样本数量。')
    if (kind === 0) {
      if (offset !== 0 || effective.underflowThreshold === undefined || lower !== 0 || upper !== effective.underflowThreshold)
        reject('直方图下溢边界无效。')
      bins.push({ kind: 'underflow', upperBound: upper, count: samples })
    } else if (kind === 1) {
      if (lower > upper || previousUpper !== undefined && (lower !== previousUpper || lower === upper)
        || regularLength === 1 && bins[regularStart]!.lowerBound === bins[regularStart]!.upperBound
        || effective.underflowThreshold !== undefined && lower < effective.underflowThreshold
        || effective.overflowThreshold !== undefined && upper > effective.overflowThreshold) reject('直方图普通分箱边界无效或不连续。')
      if (regularStart < 0) regularStart = bins.length
      regularLength++
      bins.push({ kind: 'regular', lowerBound: lower, upperBound: upper, count: samples }); previousUpper = upper
    } else if (kind === 2) {
      if (offset + 2 !== segment.valueLength || effective.overflowThreshold === undefined || lower !== effective.overflowThreshold || upper !== 0
        || previousUpper !== undefined && previousUpper > lower) reject('直方图上溢边界无效。')
      bins.push({ kind: 'overflow', lowerBound: lower, count: samples })
    } else reject('直方图分箱类型无效。')
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  if (count !== segment.validSampleCount) reject('直方图分箱计数没有覆盖有效样本。')
  const observed = new Uint32Array(bins.length)
  const underIndex = bins[0]?.kind === 'underflow' ? 0 : -1, overIndex = bins.at(-1)?.kind === 'overflow' ? bins.length - 1 : -1
  let sourceCount = 0, regularCount = 0, sourceMin = Infinity, sourceMax = -Infinity
  for (const value of source.values) {
    if (typeof value === 'number' && Number.isFinite(value)) {
      sourceCount++
      let index: number
      if (effective.underflowThreshold !== undefined && value <= effective.underflowThreshold) index = underIndex
      else if (effective.overflowThreshold !== undefined && value > effective.overflowThreshold) index = overIndex
      else {
        regularCount++; sourceMin = Math.min(sourceMin, value); sourceMax = Math.max(sourceMax, value)
        if (!regularLength || value < bins[regularStart]!.lowerBound! || value > bins[regularStart + regularLength - 1]!.upperBound!)
          reject('直方图普通分箱没有覆盖原始样本。')
        let first = regularStart, last = regularStart + regularLength - 1
        while (first < last) {
          const middle = Math.floor((first + last) / 2)
          if (value <= bins[middle]!.upperBound!) last = middle; else first = middle + 1
        }
        index = first
      }
      if (index < 0) reject('直方图缺少原始样本对应的边箱。')
      observed[index] = observed[index]! + 1
    }
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  if (sourceCount !== segment.validSampleCount || segment.available !== (sourceCount > 0)
    || (regularLength > 0) !== (regularCount > 0) || bins.some((bin, index) => bin.count !== observed[index]))
    reject('直方图分箱计数或可用状态与原始样本不一致。')
  if (regularCount) {
    if (effective.underflowThreshold !== undefined && underIndex < 0 || effective.overflowThreshold !== undefined && overIndex < 0)
      reject('直方图普通样本结果缺少启用的边箱。')
    const first = bins[regularStart]!, last = bins[regularStart + regularLength - 1]!, binning = effective.binning
    const sourceBounds = first.lowerBound === sourceMin && last.upperBound === sourceMax
    const widthBounds = binning?.mode === 'width' && first.lowerBound === (effective.underflowThreshold ?? sourceMin)
      && last.upperBound === (effective.overflowThreshold ?? (sourceMin === sourceMax ? sourceMin + binning.width : sourceMax))
    if (!sourceBounds && !widthBounds) reject('直方图普通分箱范围与原始样本或有效箱宽不一致。')
    if (sourceMin === sourceMax && !widthBounds && (regularLength !== 1 || first.lowerBound !== first.upperBound))
      reject('直方图恒定样本须保留一个恒定区间。')
    const edgeCount = Number(effective.underflowThreshold !== undefined) + Number(effective.overflowThreshold !== undefined)
    if (sourceMin !== sourceMax && binning?.mode === 'count' && binning.count > edgeCount && regularLength !== binning.count - edgeCount)
      reject('直方图箱数没有包含启用的边箱。')
  } else if (bins.some(bin => bin.count === 0)) reject('直方图没有普通样本时不应保留空边箱。')
  return { validSampleCount: sourceCount, bins }
}
