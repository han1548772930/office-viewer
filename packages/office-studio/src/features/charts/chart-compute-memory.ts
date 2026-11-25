import type { ChartModel } from '../../domain/model'
import type { ChartData } from './chart-engine'
import type { ChartComputedStatistics } from './chart-derived'
import type { ChartStatisticsOutput } from '../../compute/chart-protocol.ts'
import { CHART_PREPARATION_CHUNK, consumeChartPreparation, type ChartComputePreparationContext } from './chart-compute-lifecycle'
import { CHART_DATE_AXIS_MAX_TICKS } from './chart-date-calendar'
import { chartSunburstTotal } from './chart-sunburst'
import { isChartLinearGradient } from './chart-color'
import { CHART_GRADIENT_COMMAND_BUDGET, CHART_GRADIENT_MAX_STRIPS } from './chart-gradient-geometry'
import { resolveChartChordOptions } from './chart-chord'

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
  if (data.chord) {
    bytes += 256 + data.chord.valueFieldName.length * 2
    for (const node of data.chord.nodes) {
      bytes += 256 + (node.id.length + node.name.length + node.formattedValue.length) * 2
      if (++work % CHART_PREPARATION_CHUNK === 0) yield
    }
    for (const link of data.chord.links) {
      bytes += 384 + (link.id.length + link.sourceId.length + link.targetId.length + link.source.length + link.target.length + link.formattedValue.length) * 2
      if (++work % CHART_PREPARATION_CHUNK === 0) yield
    }
  }
  // The graph compatibility projection references Chord's same arrays. Count
  // those owned facts once; unrelated graph arrays still receive their budget.
  if (data.graph?.nodes !== data.chord?.nodes) for (const node of data.graph?.nodes ?? []) {
    bytes += 96 + node.name.length * 2; if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  if (data.graph?.links !== data.chord?.links) for (const edge of data.graph?.links ?? []) {
    bytes += 128 + (edge.source.length + edge.target.length) * 2; if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
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
  if (data.wordCloud) for (const word of data.wordCloud.words) {
    bytes += 384 + (word.id.length + word.text.length + word.font.length) * 2
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  bytes += (data.relation?.nodes.length ?? 0) * 192 + (data.graph?.diagnostic?.length ?? 0) * 2
  for (const category of data.graph?.categories ?? []) bytes += 32 + category.length * 2
  return bytes
}

export function estimateChartDecodedMemory(output: ChartStatisticsOutput): number {
  let bytes = 1024 + output.segments.length * 512
  for (const segment of output.segments) {
    if (segment.kind === 'chord') {
      const nodes = output.indices[segment.indexOffset] ?? 0, links = output.indices[segment.indexOffset + 1] ?? 0
      bytes += nodes * 512 + links * 512
      continue
    }
    if (segment.kind === 'hierarchy') {
      bytes += (output.indices[segment.indexOffset] ?? 0) * 512
      continue
    }
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
  let hierarchyBytes = 0
  const sunburstTotal = chart.type === 'sunburst' ? chartSunburstTotal(source.hierarchy ?? []) : 0
  const labelBackground = chart.options?.sunburst?.label?.backgroundColor
  // A rotated label quad can become two triangles, with up to three clipped
  // triangles per gradient band. This estimates both staging and final paint.
  const sunburstLabelPaintBytes = isChartLinearGradient(labelBackground)
    ? 6 * (CHART_GRADIENT_MAX_STRIPS + labelBackground.stops.length + 2) * 384 : 0
  const sunburstSectorBytes = (value: number): number => {
    const slices = Math.max(2, Math.ceil(Math.min(1, sunburstTotal > 0 ? value / sunburstTotal : 0) * 96))
    // Each annular slice has two filled triangles plus inner/outer outline
    // segments; retained arc points and triangle staging can coexist.
    return slices * 1024 + 1024
  }
  const hierarchy = [...(source.hierarchy ?? [])]
  while (hierarchy.length) {
    const node = hierarchy.pop()!
    // Source items, render nodes, native option nodes and rectangle/label
    // staging can coexist. Parent-owned facts add one synthetic render leaf.
    hierarchyBytes += 2048 + (node.path?.length ?? 0) * 32
    if ((chart.type === 'treemap' || chart.type === 'sunburst') && node.children?.length && node.selfValue > 0) {
      hierarchyBytes += 1024
      if (chart.type === 'sunburst') hierarchyBytes += sunburstSectorBytes(node.selfValue)
    }
    if (chart.type === 'sunburst') {
      hierarchyBytes += sunburstSectorBytes(node.value)
      if (sunburstTotal > 0 && node.value / sunburstTotal >= 0.04) hierarchyBytes += sunburstLabelPaintBytes
    }
    for (const child of node.children ?? []) hierarchy.push(child)
  }
  // Raw and projected point objects, stacked extents, numeric-X arrays and
  // source/derived category references can coexist during layout construction.
  // At most 128 native 16px pattern tiles at DPR 4 remain in the adapter cache.
  const nativePatternBytes = chart.type === 'treemap' || chart.type === 'sunburst' ? 128 * 64 * 64 * 4 : 0
  let chordBytes = 0
  // Both the base and scaled/focused relation scene, actual line/marker hit
  // primitives, measured label quads and ECharts graphics may coexist.
  const relationBytes = chart.type === 'relation' ? (source.graph?.nodes.length ?? 0) * 24_576
    + (source.graph?.links.length ?? 0) * 24_576 : 0
  if (chart.type === 'chord' && source.chord) {
    const options = resolveChartChordOptions(chart.options?.chord, chart.options?.appearance?.font)
    let bandVertices = 0, triangles = 0, sectorSlices = 0
    for (const node of source.chord.nodes) sectorSlices += Math.max(2, Math.ceil(Math.abs(node.end - node.start) / (Math.PI / 48)))
    for (const link of source.chord.links) {
      // Two authoritative endpoint arcs plus two 32-segment center curves.
      const vertices = Math.max(2, Math.ceil(Math.abs(link.sourceEnd - link.sourceStart) / (Math.PI / 48)))
        + Math.max(2, Math.ceil(Math.abs(link.targetEnd - link.targetStart) / (Math.PI / 48))) + 65
      bandVertices += vertices; triangles += Math.max(0, vertices - 2)
    }
    const hasPattern = chart.options?.allSeriesStyle?.fillPattern !== undefined
      || Object.values(chart.options?.series ?? {}).some(style => style.fillPattern !== undefined || Object.values(style.dataPoints ?? {}).some(point => point.fillPattern !== undefined))
      || (chart.options?.namedSeries ?? []).some(entry => entry.style.fillPattern !== undefined || Object.values(entry.style.dataPoints ?? {}).some(point => point.fillPattern !== undefined))
    // A clipped gradient band can produce three triangles per input triangle;
    // the exact whole-scene command cap also applies to patterned paint.
    const fills = triangles * (options.ribbon.color === 'gradient' ? 3 * (CHART_GRADIENT_MAX_STRIPS + 2) : 1)
    const commands = hasPattern ? CHART_GRADIENT_COMMAND_BUDGET
      : Math.min(CHART_GRADIENT_COMMAND_BUDGET, fills + sectorSlices * 4 + source.chord.nodes.length * 4 + bandVertices)
    // Retained bands/bounds, triangulation/gradient staging, native graphics,
    // focus copies, measured label quads and sparse incident lists coexist.
    // These object allowances are estimates, not measured JS/native/GPU heap.
    chordBytes = bandVertices * 192 + commands * 768 + source.chord.nodes.length * 1024 + source.chord.links.length * 512
  }
  return 4 * 1024 * 1024 + nativePatternBytes + categories * 128 + seriesCount * 1024
    + points * 512 + extraSeriesPoints * 256 + summaryValues * 128
    + trendPoints * 96 + errorPoints * 192 + dateBytes + hierarchyBytes + chordBytes + relationBytes
}
