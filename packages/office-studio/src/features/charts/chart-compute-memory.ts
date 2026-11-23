import type { ChartData } from './chart-engine'
import type { ChartStatisticsOutput } from '../../compute/chart-protocol.ts'
import { CHART_PREPARATION_CHUNK, consumeChartPreparation, type ChartComputePreparationContext } from './chart-compute-lifecycle'

/** Admission estimates deliberately include JS object/reference overhead.
 * They are not measured heap figures or a zero-copy claim. */
export function estimateChartDataMemory(data: ChartData, context: ChartComputePreparationContext = {}): Promise<number> {
  return consumeChartPreparation(chartDataMemorySteps(data), context)
}

function* chartDataMemorySteps(data: ChartData): Generator<void, number, void> {
  let bytes = 1024 + data.categories.length * 32, work = 0
  for (const category of data.categories) {
    bytes += category.length * 2
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  for (const field of data.fields ?? []) {
    bytes += 32 + field.length * 2
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  if (data.categoryValues) {
    bytes += 32 + data.categoryValues.length * 16
    for (const value of data.categoryValues) {
      if (typeof value === 'string') bytes += 32 + value.length * 2
      if (++work % CHART_PREPARATION_CHUNK === 0) yield
    }
  }
  if (data.categoryDateValues) bytes += 32 + data.categoryDateValues.length * 16
  for (const series of data.series) bytes += 128 + series.name.length * 2
    + (series.values.length + (series.xValues?.length ?? 0) + (series.sizeValues?.length ?? 0)) * 16
  for (const record of data.records ?? []) {
    bytes += 48 + record.length * 16
    for (const value of record) {
      if (typeof value === 'string') bytes += 32 + value.length * 2
      if (++work % CHART_PREPARATION_CHUNK === 0) yield
    }
  }
  for (const node of data.graph?.nodes ?? []) { bytes += 96 + node.name.length * 2; if (++work % CHART_PREPARATION_CHUNK === 0) yield }
  for (const edge of data.graph?.links ?? []) { bytes += 128 + (edge.source.length + edge.target.length) * 2; if (++work % CHART_PREPARATION_CHUNK === 0) yield }
  const stack = [...(data.hierarchy ?? [])]
  while (stack.length) {
    const node = stack.pop()!
    bytes += 96 + node.name.length * 2
    for (const child of node.children ?? []) { stack.push(child); if (++work % CHART_PREPARATION_CHUNK === 0) yield }
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  bytes += (data.candles?.length ?? 0) * 96
  return bytes
}

export function estimateChartDecodedMemory(output: ChartStatisticsOutput): number {
  let bytes = 1024 + output.segments.length * 512
  for (const segment of output.segments) {
    const rows = segment.kind === 'histogram' || segment.kind === 'pareto' || segment.kind === 'trendline' ? segment.valueLength / 2
      : segment.kind === 'waterfall' ? segment.valueLength / 3 : segment.kind === 'errorBars' ? segment.valueLength / 4 : Math.max(0, segment.valueLength - 5)
    bytes += rows * 128
  }
  return bytes
}
