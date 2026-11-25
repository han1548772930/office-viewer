import { isRecord } from './protocol.ts'
import type { ChartStatisticsFit, ChartStatisticsOutput, ChartStatisticsParams, ChartStatisticsSegment } from './chart-protocol.ts'
import { estimateChartTaskMemory } from './chart-tasks.ts'
import { isUnsignedCount } from '../platform/wasm/task-protocol.ts'
import { CHART_COMPUTE_LIMITS } from './chart-protocol.ts'
import { isBoundedDenseArray } from './array-shape.ts'

function isFit(value: unknown): value is ChartStatisticsFit {
  if (!isRecord(value) || !isBoundedDenseArray(value.coefficients, 7,
    (coefficient): coefficient is number => typeof coefficient === 'number' && Number.isFinite(coefficient), 2)
    || !['xScale', 'xMean', 'xSpread', 'yScale', 'yMean'].every(key => typeof value[key] === 'number' && Number.isFinite(value[key]))) return false
  return Number(value.xScale) > 0 && Number(value.yScale) > 0 && Number(value.xSpread) > 0 && Number(value.xSpread) <= 2
    && Math.abs(Number(value.xMean)) <= 1 && Math.abs(Number(value.yMean)) <= 1
    && Object.keys(value).every(key => ['coefficients', 'xScale', 'xMean', 'xSpread', 'yScale', 'yMean'].includes(key))
}
function isSegment(value: unknown): value is ChartStatisticsSegment {
  return isRecord(value) && typeof value.kind === 'string' && typeof value.available === 'boolean'
    && ['jobIndex', 'field', 'seriesPosition', 'slot', 'valueOffset', 'valueLength', 'indexOffset', 'indexLength', 'validSampleCount'].every(key => isUnsignedCount(value[key]))
    && (value.fit === undefined || isFit(value.fit))
    && (value.r2 === undefined || typeof value.r2 === 'number' && Number.isFinite(value.r2))
    && Object.keys(value).every(key => ['kind', 'jobIndex', 'field', 'seriesPosition', 'slot', 'valueOffset', 'valueLength',
      'indexOffset', 'indexLength', 'validSampleCount', 'available', 'fit', 'r2'].includes(key))
}
function ownsBuffer(view: Float64Array | Uint32Array): boolean {
  if (!(view.buffer instanceof ArrayBuffer) || view.byteOffset !== 0 || view.byteLength !== view.buffer.byteLength) return false
  try { new Uint8Array(view.buffer); return true } catch { return false }
}
/** Validate small descriptors and bulk ownership without a second full scan on
 * the main thread. Decoding validates each value/category as it consumes it. */
export function readChartStatisticsOutput(value: unknown, params: ChartStatisticsParams): ChartStatisticsOutput | null {
  if (!isRecord(value) || value.kind !== 'chartStatistics' || !(value.values instanceof Float64Array) || !(value.indices instanceof Uint32Array)
    || !ownsBuffer(value.values) || !ownsBuffer(value.indices) || value.values.buffer === value.indices.buffer
    || !isBoundedDenseArray(value.segments, CHART_COMPUTE_LIMITS.maxJobs, isSegment) || value.segments.length !== params.jobs.length) return null
  const capacity = estimateChartTaskMemory(params)
  if (value.values.length > capacity.valueCapacity || value.indices.length > capacity.indexCapacity) return null
  const segments: ChartStatisticsSegment[] = []
  let valueOffset = 0, indexOffset = 0
  for (const [jobIndex, candidate] of value.segments.entries()) {
    const job = params.jobs[jobIndex]
    if (!job || !isSegment(candidate) || candidate.jobIndex !== jobIndex || candidate.kind !== job.kind
      || candidate.field !== params.series[job.seriesIndex]?.field || candidate.slot !== job.slot
      || candidate.seriesPosition !== params.series[job.seriesIndex]?.seriesPosition
      || candidate.valueOffset !== valueOffset || candidate.indexOffset !== indexOffset
      || candidate.validSampleCount > params.rowCount * (job.kind === 'waterfallLayout' ? job.seriesIndices.length : 1)) return null
    valueOffset += candidate.valueLength; indexOffset += candidate.indexLength
    if (valueOffset > value.values.length || indexOffset > value.indices.length) return null
    if (job.kind !== 'trendline' && (candidate.fit !== undefined || candidate.r2 !== undefined)) return null
    switch (job.kind) {
      case 'histogram':
        if (candidate.valueLength % 2 || candidate.indexLength !== candidate.valueLength || candidate.valueLength > 1000 * 2
          || (candidate.available ? candidate.validSampleCount === 0 || candidate.valueLength === 0 : candidate.valueLength !== 0 || candidate.validSampleCount !== 0)) return null
        break
      case 'boxplot':
        if (candidate.available
          ? candidate.valueLength < 5 || candidate.valueLength > params.rowCount + 5 || candidate.validSampleCount === 0
            || candidate.indexLength !== candidate.valueLength - 5 || candidate.indexLength > candidate.validSampleCount
          : candidate.valueLength !== 0 || candidate.indexLength !== 0 || candidate.validSampleCount !== 0) return null
        break
      case 'pareto':
        if (candidate.valueLength !== candidate.indexLength * 2 || candidate.indexLength > params.rowCount
          || candidate.available !== (candidate.indexLength > 0)) return null
        break
      case 'hierarchy': {
        if (candidate.indexLength < 2 || candidate.valueLength % 2) return null
        const nodes = value.indices[candidate.indexOffset], roots = value.indices[candidate.indexOffset + 1]
        if (nodes === undefined || roots === undefined || nodes > candidate.validSampleCount * job.hierarchyColumns.length
          || roots > Math.min(nodes, candidate.validSampleCount) || candidate.valueLength !== nodes * 2
          || candidate.indexLength !== 2 + nodes * 4 + roots || candidate.available !== (roots > 0)
          || (nodes === 0) !== (roots === 0)) return null
        break
      }
      case 'waterfall':
        if (candidate.valueLength !== params.rowCount * 3 || candidate.indexLength !== params.rowCount * 2
          || candidate.available !== (params.rowCount > 0)) return null
        break
      case 'chord': {
        if (candidate.indexLength < 2) return null
        const nodes = value.indices[candidate.indexOffset], links = value.indices[candidate.indexOffset + 1]
        if (nodes === undefined || links === undefined || nodes > candidate.validSampleCount * 2 || links > candidate.validSampleCount
          || candidate.valueLength !== nodes * 3 + links * 5 || candidate.indexLength !== 2 + nodes + links * 2
          || candidate.available !== (links > 0) || (nodes === 0) !== (links === 0) || links > 0 && nodes < 2
          || links === 0 && candidate.validSampleCount !== 0) return null
        break
      }
      case 'waterfallLayout': {
        const points = params.rowCount === 0 ? 0 : params.rowCount * job.seriesIndices.length
          + (job.useSubtotal ? job.stackType === 'stacked' ? 1 : job.seriesIndices.length : 0)
        const connectors = Math.max(0, (job.stackType === 'stacked' ? params.rowCount : params.rowCount * job.seriesIndices.length) - 1)
        if (candidate.valueLength !== 1 + points * 3 + connectors * 2 || candidate.indexLength !== 2 + points * 4 + connectors * 2
          || candidate.available !== (params.rowCount > 0) || value.indices[candidate.indexOffset] !== points
          || value.indices[candidate.indexOffset + 1] !== connectors || !Number.isFinite(value.values[candidate.valueOffset])
          || params.rowCount === 0 && (candidate.validSampleCount !== 0 || value.values[candidate.valueOffset] !== 0)) return null
        break
      }
      case 'trendline':
        if (candidate.indexLength !== 0 || candidate.valueLength % 2 || candidate.valueLength > (job.type === 'movingAvg' ? params.rowCount : 33) * 2) return null
        if (job.type === 'movingAvg') {
          if (candidate.fit !== undefined || candidate.r2 !== undefined) return null
        } else if (candidate.available
          ? !candidate.fit || candidate.fit.coefficients.length !== (job.type === 'poly' ? job.order ?? 2 : 1) + 1
          : candidate.fit !== undefined || candidate.r2 !== undefined || candidate.valueLength !== 0) return null
        break
      case 'errorBars':
        if (candidate.valueLength !== candidate.indexLength * 4 || candidate.indexLength > params.rowCount
          || !candidate.available && candidate.valueLength !== 0) return null
        break
    }
    segments.push(candidate)
  }
  if (valueOffset !== value.values.length || indexOffset !== value.indices.length) return null
  return { kind: 'chartStatistics', values: value.values, indices: value.indices, segments }
}
