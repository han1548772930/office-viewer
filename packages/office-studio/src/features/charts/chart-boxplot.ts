import type { ChartAxisOptions, ChartModel, ChartSeriesOptions } from '../../domain/model'
import type { ChartLayout } from './chart-engine'
import type { ChartDerivedData } from './chart-derived'
import { chartPointStyle } from './chart-derived'
import { chartAxisRatio, chartColorOpacity } from './chart-geometry'
import { chartSeriesStyle } from './chart-options'

export type ChartBoxplotDistribution = NonNullable<ChartDerivedData['boxes']>[number]

/** Each raw numeric field owns one distribution, even when its header or output order changes. */
export function chartBoxplotStyle(chart: ChartModel, box: ChartBoxplotDistribution, categoryIndex: number, palette: readonly string[]): ChartSeriesOptions {
  return chartPointStyle(chartSeriesStyle(chart, box.field, box.seriesPosition, palette, box.sourceName), categoryIndex, chart.type)
}

export function chartBoxplotAxis(chart: ChartModel, style: ChartSeriesOptions): ChartAxisOptions {
  return style.rightAxis ? chart.options?.rightYAxis ?? { scaleType: 'linear' } : chart.yAxis
}

export function chartBoxplotValueDrawable(value: number, axis: ChartAxisOptions): boolean {
  return Number.isFinite(value) && (axis.scaleType !== 'log' || value > 0)
}

/** Invalid logarithmic summaries stay in the source statistics; they are never recomputed from positive-only samples. */
export function chartBoxplotSummaryDrawable(statistics: ChartBoxplotDistribution['statistics'], axis: ChartAxisOptions): boolean {
  return Boolean(statistics && [statistics.min, statistics.q1, statistics.median, statistics.q3, statistics.max].every(value => chartBoxplotValueDrawable(value, axis)))
}

export function chartBoxplotGeometry(chart: ChartModel, layout: ChartLayout, style: ChartSeriesOptions, categoryIndex: number, count: number) {
  const plot = layout.plot, step = layout.categoryBandwidth ?? plot.width / Math.max(1, count), width = step * 0.48
  const ratio = layout.categoryPositions ? layout.categoryPositions[categoryIndex] : (chart.xAxis.reverse ? count - categoryIndex - 0.5 : categoryIndex + 0.5) / Math.max(1, count)
  const x = ratio == null ? NaN : plot.x + ratio * plot.width
  const axis = chartBoxplotAxis(chart, style)
  const min = style.rightAxis ? layout.rightMin ?? layout.min : layout.min, max = style.rightAxis ? layout.rightMax ?? layout.max : layout.max
  const y = (value: number): number | undefined => {
    if (!Number.isFinite(x) || !chartBoxplotValueDrawable(value, axis)) return undefined
    const coordinate = plot.y + (1 - chartAxisRatio(value, min, max, axis)) * plot.height
    return Number.isFinite(coordinate) ? coordinate : undefined
  }
  const summary = (statistics: ChartBoxplotDistribution['statistics']) => {
    if (!statistics) return undefined
    const min = y(statistics.min), q1 = y(statistics.q1), median = y(statistics.median), q3 = y(statistics.q3), max = y(statistics.max)
    if (min === undefined || q1 === undefined || median === undefined || q3 === undefined || max === undefined) return undefined
    return { min, q1, median, q3, max }
  }
  return { x, width, y, summary }
}

/** Fill opacity is encoded only in the fill so the border retains its independent opacity. */
export function chartBoxplotEChartsStyle(style: ChartSeriesOptions) {
  const color = style.color ?? '#2563eb', border = style.border
  return { color: chartColorOpacity(color, style.fillOpacity ?? 0.45), opacity: 1,
    borderColor: chartColorOpacity(border?.color ?? color, border?.opacity ?? 1), borderWidth: border?.visible === false ? 0 : border?.width ?? 1.5,
    borderType: border?.dashType ?? 'solid' as const, borderJoin: 'round' as const }
}
