import { ComputeError } from './errors.ts'
import { COMPUTE_LIMITS, isRecord } from './protocol.ts'
import { CHART_COMPUTE_LIMITS } from './chart-protocol.ts'
import type { ChartStatisticsJob, ChartStatisticsParams, ChartStatisticsSeries, ChartStatisticsTask } from './chart-protocol.ts'
import { isUnsignedCount } from '../platform/wasm/task-protocol.ts'
import { isBoundedDenseArray } from './array-shape.ts'

const keys = (value: Record<string, unknown>, allowed: readonly string[]): boolean => Object.keys(value).every(key => allowed.includes(key))
const index = (value: unknown, maximum = 0xffff_ffff): value is number => isUnsignedCount(value) && value <= maximum
const finite = (value: unknown): value is number => typeof value === 'number' && Number.isFinite(value)
const optionalIndex = (value: unknown, maximum = 0xffff_ffff): boolean => value === undefined || index(value, maximum)
const optionalFinite = (value: unknown, lower = -Number.MAX_VALUE, upper = Number.MAX_VALUE): boolean =>
  value === undefined || finite(value) && value >= lower && value <= upper
const oneOf = (value: unknown, values: readonly string[]): boolean => typeof value === 'string' && values.includes(value)
const optionalInteger = (value: unknown, lower: number, upper: number): boolean =>
  value === undefined || index(value, upper) && value >= lower

function isSeries(value: unknown): value is ChartStatisticsSeries {
  return isRecord(value) && keys(value, ['field', 'seriesPosition', 'column', 'xColumn'])
    && index(value.field, 1_048_575) && index(value.seriesPosition, CHART_COMPUTE_LIMITS.maxSeries - 1)
    && index(value.column) && optionalIndex(value.xColumn)
}
function isJob(value: unknown): value is ChartStatisticsJob {
  if (!isRecord(value) || !index(value.seriesIndex) || !index(value.slot, CHART_COMPUTE_LIMITS.maxJobs - 1)) return false
  const common = ['kind', 'seriesIndex', 'slot']
  switch (value.kind) {
    case 'histogram': {
      if (!keys(value, [...common, 'binning', 'underflowThreshold', 'overflowThreshold'])
        || !optionalFinite(value.underflowThreshold) || !optionalFinite(value.overflowThreshold)) return false
      if (value.binning === undefined) return true
      const bin = value.binning
      return isRecord(bin) && (bin.mode === 'count'
        ? keys(bin, ['mode', 'count']) && finite(bin.count)
        : bin.mode === 'width' && keys(bin, ['mode', 'width']) && finite(bin.width))
    }
    case 'boxplot':
      return keys(value, [...common, 'quartileMethod', 'whiskerMethod', 'whiskerMultiplier']) && oneOf(value.quartileMethod, ['medianOfHalves', 'inclusive', 'exclusive'])
        && (value.whiskerMethod === undefined || oneOf(value.whiskerMethod, ['range', 'iqr']))
        && finite(value.whiskerMultiplier) && value.whiskerMultiplier >= 0 && value.whiskerMultiplier <= 100
    case 'pareto':
      return keys(value, [...common, 'includeZeroValues']) && typeof value.includeZeroValues === 'boolean'
    case 'hierarchy':
      return keys(value, [...common, 'hierarchyColumns', 'labelColumns', 'keyCount', 'labelCount'])
        && isBoundedDenseArray(value.hierarchyColumns, CHART_COMPUTE_LIMITS.maxColumns,
          (column): column is number => index(column), 1)
        && isBoundedDenseArray(value.labelColumns, CHART_COMPUTE_LIMITS.maxColumns,
          (column): column is number => index(column), 1)
        && value.hierarchyColumns.length === value.labelColumns.length && index(value.keyCount, COMPUTE_LIMITS.maxElements)
        && index(value.labelCount, COMPUTE_LIMITS.maxElements)
    case 'waterfall':
      return keys(value, [...common, 'totalColumn']) && optionalIndex(value.totalColumn)
    case 'waterfallLayout':
      return keys(value, [...common, 'seriesIndices', 'stackType', 'useSubtotal', 'totalColumns'])
        && isBoundedDenseArray(value.seriesIndices, CHART_COMPUTE_LIMITS.maxSeries,
          (item): item is number => index(item, CHART_COMPUTE_LIMITS.maxSeries - 1), 1)
        && new Set(value.seriesIndices).size === value.seriesIndices.length && value.seriesIndices[0] === value.seriesIndex
        && oneOf(value.stackType, ['arrangement', 'stacked']) && typeof value.useSubtotal === 'boolean'
        && (value.totalColumns === undefined || isBoundedDenseArray(value.totalColumns, CHART_COMPUTE_LIMITS.maxSeries,
          (item): item is number | null => item === null || index(item)) && value.totalColumns.length === value.seriesIndices.length)
    case 'trendline':
      return keys(value, [...common, 'type', 'order', 'period']) && oneOf(value.type, ['exp', 'linear', 'log', 'movingAvg', 'poly', 'pow'])
        && optionalInteger(value.order, 2, 6) && optionalInteger(value.period, 2, 100_000)
    case 'errorBars':
      return keys(value, [...common, 'direction', 'type', 'valueType', 'value', 'plusColumn', 'minusColumn'])
        && oneOf(value.direction, ['x', 'y']) && oneOf(value.type, ['both', 'plus', 'minus'])
        && oneOf(value.valueType, ['fixedVal', 'percentage', 'stdDev', 'stdErr', 'cust'])
        && optionalFinite(value.value, 0) && optionalIndex(value.plusColumn) && optionalIndex(value.minusColumn)
    default: return false
  }
}
export function isChartStatisticsParams(value: unknown): value is ChartStatisticsParams {
  if (!isRecord(value) || !keys(value, ['rowCount', 'columnCount', 'categoryCount', 'series', 'jobs'])
    || !index(value.rowCount, COMPUTE_LIMITS.maxRows) || !index(value.columnCount, CHART_COMPUTE_LIMITS.maxColumns) || value.columnCount === 0
    || !index(value.categoryCount, value.rowCount) || (value.rowCount === 0 ? value.categoryCount !== 0 : value.categoryCount === 0)
    || !isBoundedDenseArray(value.series, CHART_COMPUTE_LIMITS.maxSeries, isSeries)
    || !isBoundedDenseArray(value.jobs, CHART_COMPUTE_LIMITS.maxJobs, isJob)
    || value.rowCount * value.columnCount > COMPUTE_LIMITS.maxElements) return false
  const positions = new Set<number>(), slots = new Set<string>()
  for (const series of value.series) {
    if (positions.has(series.seriesPosition) || series.column >= value.columnCount || series.xColumn !== undefined && series.xColumn >= value.columnCount) return false
    positions.add(series.seriesPosition)
  }
  for (const job of value.jobs) {
    const slot = job.kind + ':' + job.seriesIndex + ':' + job.slot
    if (slots.has(slot) || job.seriesIndex >= value.series.length) return false
    slots.add(slot)
    if (job.kind === 'waterfall' && job.totalColumn !== undefined && job.totalColumn >= value.columnCount) return false
    if (job.kind === 'waterfallLayout' && (job.seriesIndices.some(item => item >= value.series.length)
      || job.totalColumns?.some(column => column !== null && column >= value.columnCount))) return false
    if (job.kind === 'waterfallLayout') {
      const sourcePoints = value.rowCount * job.seriesIndices.length
      const points = value.rowCount === 0 ? 0 : sourcePoints + (job.useSubtotal ? job.stackType === 'stacked' ? 1 : job.seriesIndices.length : 0)
      if (points > 0xffff_ffff) return false
    }
    if (job.kind === 'hierarchy') {
      const columns = [...job.hierarchyColumns, ...job.labelColumns, value.series[job.seriesIndex]!.column]
      if (columns.some(column => column >= value.columnCount) || new Set(columns).size !== columns.length
        || job.keyCount > value.rowCount * job.hierarchyColumns.length || job.labelCount > value.rowCount * job.hierarchyColumns.length) return false
    }
    if (job.kind === 'errorBars') {
      if (job.direction === 'x' && value.series[job.seriesIndex]?.xColumn === undefined) return false
      if (job.valueType === 'cust') {
        if (job.type !== 'minus' && (job.plusColumn === undefined || job.plusColumn >= value.columnCount)) return false
        if (job.type !== 'plus' && (job.minusColumn === undefined || job.minusColumn >= value.columnCount)) return false
      }
    }
  }
  return true
}
export function validateChartTask(value: unknown): asserts value is ChartStatisticsTask {
  if (!isRecord(value) || value.kind !== 'chartStatistics' || !(value.values instanceof Float64Array)
    || !(value.states instanceof Uint8Array) || !(value.categoryIds instanceof Uint32Array) || !isChartStatisticsParams(value.params)
    || value.values.length !== value.params.rowCount * value.params.columnCount || value.states.length !== value.values.length
    || value.categoryIds.length !== value.params.rowCount) {
    throw new ComputeError('COMPUTE_INVALID_INPUT', '图表统计任务的列、分类或配置协议无效')
  }
}

export function chartOutputCapacity(job: ChartStatisticsJob, rows: number): { values: number; indices: number; scratch: number } {
  switch (job.kind) {
    case 'histogram': return { values: 1000 * 2, indices: 1000 * 2, scratch: rows * 8 + 1001 * 12 }
    case 'boxplot': return { values: rows + 5, indices: rows, scratch: rows * 16 }
    case 'pareto': return { values: rows * 2, indices: rows, scratch: rows * 16 }
    case 'hierarchy': {
      const nodes = rows * job.hierarchyColumns.length
      return { values: nodes * 2, indices: 2 + nodes * 4 + rows, scratch: nodes * 256 + (job.hierarchyColumns.length + 1) * 64 }
    }
    case 'waterfall': return { values: rows * 3, indices: rows * 2, scratch: 0 }
    case 'waterfallLayout': {
      const points = rows === 0 ? 0 : rows * job.seriesIndices.length + (job.useSubtotal ? job.stackType === 'stacked' ? 1 : job.seriesIndices.length : 0)
      const connectors = Math.max(0, (job.stackType === 'stacked' ? rows : rows * job.seriesIndices.length) - 1)
      return { values: 1 + points * 3 + connectors * 2, indices: 2 + points * 4 + connectors * 2, scratch: 0 }
    }
    case 'trendline': return job.type === 'movingAvg'
      ? { values: rows * 2, indices: 0, scratch: Math.min(rows, job.period ?? 2) * 112 }
      : { values: 33 * 2, indices: 0, scratch: rows * 16 + 4096 }
    case 'errorBars': return { values: rows * 4, indices: rows, scratch: 0 }
  }
}
export function estimateChartTaskMemory(params: ChartStatisticsParams): {
  inputBytes: number; outputBytes: number; estimatedPeakBytes: number; valueCapacity: number; indexCapacity: number
} {
  const inputBytes = params.rowCount * (params.columnCount * 9 + 4)
  let valueCapacity = 0, indexCapacity = 0, scratch = 0
  for (const job of params.jobs) {
    const capacity = chartOutputCapacity(job, params.rowCount)
    valueCapacity += capacity.values; indexCapacity += capacity.indices; scratch = Math.max(scratch, capacity.scratch)
  }
  const outputBytes = valueCapacity * 8 + indexCapacity * 4
  const controls = params.series.length * 32 + params.jobs.length * 512
    + params.jobs.reduce((sum, job) => sum + (job.kind === 'waterfallLayout' ? job.seriesIndices.length * 4 + (job.totalColumns?.length ?? 0) * 8
      : job.kind === 'hierarchy' ? job.hierarchyColumns.length * 8 : 0), 0)
  return { inputBytes, outputBytes, valueCapacity, indexCapacity,
    estimatedPeakBytes: inputBytes * 2 + outputBytes * 2 + scratch + CHART_COMPUTE_LIMITS.scratchAllowance
      + controls + CHART_COMPUTE_LIMITS.runtimeAllowance }
}
