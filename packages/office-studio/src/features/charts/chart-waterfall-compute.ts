import type { ChartStatisticsJob, ChartStatisticsOutput } from '../../compute/chart-protocol'
import { ComputeError } from '../../compute/errors'
import { CHART_PREPARATION_CHUNK } from './chart-compute-lifecycle'
import type { ChartStatisticsPlan } from './chart-statistics-compute'
import { CHART_WATERFALL_GLOBAL_SERIES, type WaterfallLayoutData, type WaterfallLayoutPoint } from './chart-statistics'

type WaterfallJob = Extract<ChartStatisticsJob, { kind: 'waterfallLayout' }>
type Segment = ChartStatisticsOutput['segments'][number]

const reject = (message: string): never => { throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', message) }
const same = (left: number, right: number): boolean => Number.isFinite(left) && Number.isFinite(right)
  && Math.abs(left - right) <= Number.EPSILON * 8 * Math.max(1, Math.abs(left), Math.abs(right))

/** Validate the Rust layout against the retained source and authoring metadata.
 * No values are generated here: only accepted Worker tuples reach rendering. */
export function* decodeChartWaterfallLayoutSteps(plan: ChartStatisticsPlan, output: ChartStatisticsOutput,
  segment: Segment, job: WaterfallJob): Generator<void, WaterfallLayoutData, void> {
  if (!plan.task) reject('没有可解码的瀑布图布局任务。')
  const params = plan.task.params, rows = params.rowCount, sources = job.seriesIndices.map(index => {
    const source = params.series[index]
    if (!source) reject('瀑布图任务引用了不存在的来源系列。')
    return source
  })
  if (sources.length !== plan.preparedData.series.length || sources.some((source, position) => source.seriesPosition !== position
    || source.field !== plan.preparedData.series[position]?.field)) reject('瀑布图任务没有完整覆盖当前来源系列。')
  const count = output.indices[segment.indexOffset]!, connectorCount = output.indices[segment.indexOffset + 1]!
  const realCount = rows * sources.length
  const expectedCount = realCount + (rows && job.useSubtotal ? job.stackType === 'arrangement' ? sources.length : 1 : 0)
  const expectedConnectors = Math.max(0, job.stackType === 'arrangement' ? realCount - 1 : rows - 1)
  if (count !== expectedCount || connectorCount !== expectedConnectors || segment.valueLength !== 1 + count * 3 + connectorCount * 2
    || segment.indexLength !== 2 + count * 4 + connectorCount * 2 || segment.available !== (rows > 0)) reject('瀑布图点、连接线或缓冲区长度无效。')
  const valueAt = (offset: number): number => {
    const value = output.values[segment.valueOffset + offset]
    if (value === undefined || !Number.isFinite(value)) reject('瀑布图结果包含非有限数值。')
    return value
  }
  const cumulative = valueAt(0), points: WaterfallLayoutPoint[] = [], realPoints: WaterfallLayoutPoint[] = [], categoryEnds: number[] = []
  let work = 0, validSamples = 0, previous = 0
  const sourceValue = (position: number, row: number): number => {
    const value = plan.preparedData.series[position]!.values[row]
    return typeof value === 'number' && Number.isFinite(value) ? value : 0
  }
  const pointAt = (offset: number, position: number, row: number, categoryIndex: number, synthetic: boolean): WaterfallLayoutPoint => {
    const index = segment.indexOffset + 2 + offset * 4
    const seriesPosition = output.indices[index]!, dataPointIndex = output.indices[index + 1]!, category = output.indices[index + 2]!, flags = output.indices[index + 3]!
    const total = synthetic || (job.stackType === 'arrangement' && plan.waterfallTotals?.get(position)?.has(row) === true)
    if (seriesPosition !== position || dataPointIndex !== row || category !== categoryIndex || flags !== (synthetic ? 3 : total ? 1 : 0)) {
      reject('瀑布图来源身份、显示列或项目含义与当前请求不一致。')
    }
    const start = valueAt(1 + offset * 3), end = valueAt(2 + offset * 3), value = valueAt(3 + offset * 3)
    if (synthetic) {
      if (start !== 0 || end !== cumulative || value !== cumulative) reject('瀑布图自动小计与最终累计值不一致。')
    } else {
      const source = plan.preparedData.series[position]!.values[row]
      if (typeof source === 'number' && Number.isFinite(source)) validSamples++
      if (value !== sourceValue(position, row)) reject('瀑布图值与原始项目不一致。')
    }
    return { seriesPosition, dataPointIndex, categoryIndex: category, synthetic, role: total ? 'total' : 'change', start, end, value }
  }
  if (job.stackType === 'arrangement') {
    let offset = 0
    for (const source of sources) {
      const position = source.seriesPosition
      for (let row = 0; row < rows; row++) {
        const point = pointAt(offset, position, row, offset, false)
        const expectedStart = point.role === 'total' ? 0 : previous, expectedEnd = point.role === 'total' ? point.value : point.start + point.value
        if (point.role === 'total' && point.start !== 0 || !same(point.start, expectedStart) || !same(point.end, expectedEnd)) reject('瀑布图差值或合计基线无效。')
        previous = point.end; points.push(point); realPoints.push(point); offset++
        if (++work % CHART_PREPARATION_CHUNK === 0) yield
      }
      if (rows && job.useSubtotal) {
        points.push(pointAt(offset, position, rows, offset, true)); offset++
        if (++work % CHART_PREPARATION_CHUNK === 0) yield
      }
    }
  } else {
    let offset = 0
    for (let row = 0; row < rows; row++) {
      let negativeBase = previous, nextCumulative = previous
      for (const source of sources) {
        const value = sourceValue(source.seriesPosition, row)
        if (value < 0) negativeBase += value
        nextCumulative += value
        if (++work % CHART_PREPARATION_CHUNK === 0) yield
      }
      if (!Number.isFinite(negativeBase) || !Number.isFinite(nextCumulative)) reject('瀑布图堆积累计超出有限数值范围。')
      let cursor = negativeBase
      for (const source of sources) {
        const point = pointAt(offset, source.seriesPosition, row, row, false)
        const low = Math.min(point.start, point.end), high = Math.max(point.start, point.end)
        if (!same(low, cursor) || !same(high, low + Math.abs(point.value))
          || point.value >= 0 && point.end < point.start || point.value < 0 && point.start < point.end) reject('瀑布图堆积片段有间隙、重叠或方向错误。')
        cursor = high; points.push(point); offset++
        if (++work % CHART_PREPARATION_CHUNK === 0) yield
      }
      previous = nextCumulative; categoryEnds.push(previous)
    }
    if (rows && job.useSubtotal) points.push(pointAt(offset, CHART_WATERFALL_GLOBAL_SERIES, rows, rows, true))
  }
  if (!same(previous, cumulative) || validSamples !== segment.validSampleCount) reject('瀑布图最终累计值或有效样本计数无效。')
  const connectors: WaterfallLayoutData['connectors'] = []
  for (let offset = 0; offset < connectorCount; offset++) {
    const index = segment.indexOffset + 2 + count * 4 + offset * 2
    const fromCategory = output.indices[index]!, toCategory = output.indices[index + 1]!
    const fromValue = valueAt(1 + count * 3 + offset * 2), toValue = valueAt(2 + count * 3 + offset * 2)
    if (job.stackType === 'arrangement') {
      const from = realPoints[offset]!, to = realPoints[offset + 1]!
      if (fromCategory !== from.categoryIndex || toCategory !== to.categoryIndex || !same(fromValue, from.end)
        || !same(toValue, to.role === 'total' ? to.end : to.start)) reject('瀑布图连接线没有对应真实项目端点。')
    } else if (fromCategory !== offset || toCategory !== offset + 1 || !same(fromValue, categoryEnds[offset]!)
      || !same(toValue, categoryEnds[offset]!)) reject('瀑布图堆积连接线与分类累计不一致。')
    connectors.push({ from: { categoryIndex: fromCategory, value: fromValue }, to: { categoryIndex: toCategory, value: toValue } })
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  return { stackType: job.stackType, useSubtotal: job.useSubtotal, sourceCategoryCount: rows,
    seriesFields: plan.preparedData.series.map(source => source.field), points, connectors, cumulative }
}
