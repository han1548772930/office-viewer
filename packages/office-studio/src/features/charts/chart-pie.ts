import type { ChartLineOptions, ChartModel, ChartSeriesOptions } from '../../domain/model'
import type { ChartData } from './chart-engine'
import { chartCompoundPieData, chartPointStyle } from './chart-derived'
import { chartSeriesStyle, definedChartStyleFields } from './chart-options'

export function isChartCompoundPie(type: ChartModel['type']): boolean {
  return type === 'pieOfPie' || type === 'barOfPie'
}

/** Dedicated authoring wins over the historical synthetic-index appearance.
 * This keeps old snapshots readable without writing a generated index back
 * to the source-series dataPoints collection. */
export function chartPieAggregateStyle(chart: Pick<ChartModel, 'options'>, source: ChartData['series'][number],
  aggregateIndex: number, palette: readonly string[]): ChartSeriesOptions {
  const inherited = chartPointStyle(chartSeriesStyle(chart, source.field, aggregateIndex, palette, source.sourceName ?? source.name), aggregateIndex, 'pieOfPie')
  const authored = chart.options?.pie?.aggregateStyle
  return { ...inherited, ...definedChartStyleFields(authored),
    border: { ...definedChartStyleFields(inherited.border), ...definedChartStyleFields(authored?.border) } }
}
/** A materialized source slice has its own author layer. Legacy source
 * dataPoints remain readable below it and are never replaced by slice edits. */
export function chartPieSliceStyle(chart: Pick<ChartModel, 'options'>, source: ChartData['series'][number],
  sourcePointIndex: number, palette: readonly string[]): ChartSeriesOptions {
  const inherited = chartPointStyle(chartSeriesStyle(chart, source.field, sourcePointIndex, palette, source.sourceName ?? source.name), sourcePointIndex, 'pie')
  const authored = chart.options?.pie?.sliceStyles?.[sourcePointIndex]
  return { ...inherited, ...definedChartStyleFields(authored),
    border: { ...definedChartStyleFields(inherited.border), ...definedChartStyleFields(authored?.border) } }
}
export function chartPieSliceExplosion(chart: Pick<ChartModel, 'options'>, sourcePointIndex: number): number {
  return chart.options?.pie?.sliceStyles?.[sourcePointIndex]?.explosion ?? chart.options?.pie?.explosion ?? 0
}
export function chartPieSliceAllowsExplosion(chart: Pick<ChartModel, 'options' | 'type'>, sourcePointIndex: number,
  data?: ChartData, plot?: 'primary' | 'secondary'): boolean {
  if (!isChartCompoundPie(chart.type) || chart.type !== 'barOfPie' && chart.options?.pie?.secondaryPlot?.type !== 'bar') return true
  if (plot) return plot === 'primary'
  return !data || !chartCompoundPieData(data, chart).secondary.some(part => part.index === sourcePointIndex)
}
export function chartPieMaximumExplosion(data: ChartData, chart: Pick<ChartModel, 'options' | 'type'>): number {
  let maximum = Math.max(0, chart.options?.pie?.explosion ?? 0, isChartCompoundPie(chart.type) ? chartPieAggregateExplosion(chart) : 0)
  const inactiveSlices = new Set(isChartCompoundPie(chart.type)
    && (chart.type === 'barOfPie' || chart.options?.pie?.secondaryPlot?.type === 'bar')
    ? chartCompoundPieData(data, chart).secondary.map(part => part.index) : [])
  for (const [index, style] of Object.entries(chart.options?.pie?.sliceStyles ?? {})) {
    const point = Number(index)
    if (point < data.categories.length && (data.series[0]?.values[point] ?? 0) > 0 && !inactiveSlices.has(point)) maximum = Math.max(maximum, style?.explosion ?? 0)
  }
  return maximum
}

/** Omission retains the existing visible, neutral 1px connector. Effective
 * theme colors and defaults remain runtime-only. */
export function chartPieConnector(chart: Pick<ChartModel, 'options'>, themeColor: string): ChartLineOptions | undefined {
  const authored = chart.options?.pie?.secondaryPlot?.connector
  if (authored?.visible === false || authored?.width === 0 || authored?.opacity === 0) return undefined
  return { visible: true, color: themeColor, width: 1, opacity: 1, dashType: 'solid', ...definedChartStyleFields(authored) }
}

export function chartPieAggregateExplosion(chart: Pick<ChartModel, 'options'>): number {
  return chart.options?.pie?.aggregateStyle?.explosion ?? chart.options?.pie?.explosion ?? 0
}
