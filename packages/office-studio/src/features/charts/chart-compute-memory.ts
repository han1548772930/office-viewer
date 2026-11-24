import type { ChartModel } from '../../domain/model'
import type { ChartData } from './chart-engine'
import type { ChartComputedStatistics } from './chart-derived'
import type { ChartStatisticsOutput } from '../../compute/chart-protocol.ts'
import { CHART_PREPARATION_CHUNK, consumeChartPreparation, type ChartComputePreparationContext } from './chart-compute-lifecycle'
import { CHART_DATE_AXIS_MAX_TICKS } from './chart-date-calendar'

/** Admission estimates deliberately include JS object/reference overhead.
 * They are not measured heap figures or a zero-copy claim. */
export function estimateChartDataMemory(data: ChartData, context: ChartComputePreparationContext = {}): Promise<number> {
  return consumeChartPreparation(chartDataMemorySteps(data), context)
}

function* chartDataMemorySteps(data: ChartData): Generator<void, number, void> {
  let bytes = 1024 + data.categories.length * 32
    + ((data.sourceCategoryFields?.length ?? 0) + (data.sourceValueFields?.length ?? 0)
      + (data.sourceNumericFields?.length ?? 0) + (data.visibleFields?.length ?? 0)) * 16, work = 0
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
  for (const series of data.series) {
    bytes += 128 + series.name.length * 2
      + (series.values.length + (series.xValues?.length ?? 0) + (series.sizeValues?.length ?? 0) + (series.formattedValues?.length ?? 0)) * 16
    for (const value of series.formattedValues ?? []) {
      if (value !== null) bytes += 32 + value.length * 2
      if (++work % CHART_PREPARATION_CHUNK === 0) yield
    }
  }
  for (const record of data.records ?? []) {
    bytes += 48 + record.length * 16
    for (const value of record) {
      if (typeof value === 'string') bytes += 32 + value.length * 2
      if (++work % CHART_PREPARATION_CHUNK === 0) yield
    }
  }
  // Conservative accounting includes both source rows and projected references.
  // Shared strings can be counted twice; these remain admission estimates.
  for (const record of data.formattedRecords ?? []) {
    bytes += 48 + record.length * 16
    for (const value of record) {
      if (value !== null) bytes += 32 + value.length * 2
      if (++work % CHART_PREPARATION_CHUNK === 0) yield
    }
  }
  for (const node of data.graph?.nodes ?? []) { bytes += 96 + node.name.length * 2; if (++work % CHART_PREPARATION_CHUNK === 0) yield }
  for (const edge of data.graph?.links ?? []) { bytes += 128 + (edge.source.length + edge.target.length) * 2; if (++work % CHART_PREPARATION_CHUNK === 0) yield }
  const stack = [...(data.hierarchy ?? [])]
  while (stack.length) {
    const node = stack.pop()!
    bytes += 256 + node.name.length * 2 + (node.id?.length ?? 0) * 2 + (node.path?.length ?? 0) * 16
      + ((node.formattedValue?.length ?? 0) + (node.formattedSelfValue?.length ?? 0)) * 2
    for (const child of node.children ?? []) { stack.push(child); if (++work % CHART_PREPARATION_CHUNK === 0) yield }
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  bytes += (data.hierarchyPaletteOrder?.length ?? 0) * 32 + (data.hierarchyValueFieldName?.length ?? 0) * 2
  bytes += (data.candles?.length ?? 0) * 96
  return bytes
}

export function estimateChartDecodedMemory(output: ChartStatisticsOutput): number {
  let bytes = 1024 + output.segments.length * 512
  for (const segment of output.segments) {
    if (segment.kind === 'waterfallLayout') {
      const points = output.indices[segment.indexOffset] ?? 0, connectors = output.indices[segment.indexOffset + 1] ?? 0
      bytes += points * 192 + connectors * 128
      continue
    }
    const rows = segment.kind === 'histogram' || segment.kind === 'pareto' || segment.kind === 'trendline' ? segment.valueLength / 2
      : segment.kind === 'waterfall' ? segment.valueLength / 3 : segment.kind === 'errorBars' ? segment.valueLength / 4 : Math.max(0, segment.valueLength - 5)
    bytes += rows * 128
  }
  return bytes
}

/** Additional peak for the final scene, above retained input and decoded
 * statistics. Counts use actual derived rows: a two-sample histogram can still
 * allocate one thousand bins, and trend/error extents outlive raw point arrays.
 * Object sizes remain conservative admission allowances, not heap measurements. */
export function estimateChartFinalLayoutMemory(source: ChartData, chart: ChartModel, statistics: ChartComputedStatistics): number {
  let categories = source.categories.length, seriesCount = source.series.length
  let points = source.series.reduce((count, series) => count + series.values.length, 0)
  let summaryValues = 0, extraSeriesPoints = 0
  if (chart.type === 'histogram') {
    categories = statistics.histogram?.bins.length ?? (points ? 1000 : 0)
    seriesCount = source.series.length ? 1 : 0; points = categories
  } else if (chart.type === 'boxplot') {
    categories = source.series.length; seriesCount = source.series.length ? 1 : 0; points = categories
    for (const box of statistics.boxplots ?? []) if (box.statistics) summaryValues += 5 + box.statistics.outliers.length
  } else if (chart.type === 'pareto') {
    categories = statistics.pareto?.categories.length ?? categories
    seriesCount = source.series.length ? 1 : 0; points = categories; extraSeriesPoints = categories
  } else if (statistics.waterfallLayout) {
    const layout = statistics.waterfallLayout
    categories = layout.sourceCategoryCount + Number(layout.useSubtotal && layout.sourceCategoryCount > 0)
    seriesCount = layout.seriesFields.length; points = layout.points.length
    summaryValues = layout.connectors.length * 2
  }
  let trendPoints = 0, errorPoints = 0
  for (const trend of statistics.trendlines ?? []) trendPoints += trend.statistics.points.length
  for (const error of statistics.errorBars ?? []) errorPoints += error.intervals.length
  const dateBytes = chart.options?.mapping?.useDateAxis
    ? categories * 192 + 2 * CHART_DATE_AXIS_MAX_TICKS * 160 : 0
  // Raw and projected point objects, stacked extents, numeric-X arrays and
  // source/derived category references can coexist during layout construction.
  return 4 * 1024 * 1024 + categories * 128 + seriesCount * 1024
    + points * 512 + extraSeriesPoints * 256 + summaryValues * 128
    + trendPoints * 96 + errorPoints * 192 + dateBytes
}
