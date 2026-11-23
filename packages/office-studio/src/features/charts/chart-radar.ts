import type { ChartAxisOptions, ChartModel, ChartNullablePatch, ChartRadarAxisLabelOptions, ChartRadarOptions, ChartSemanticOptions } from '../../domain/model'
import type { ChartData } from './chart-engine'
import type { RenderRect } from '../../rendering/contracts'
import { chartAxisExtent, chartAxisRatio, chartAxisTicks, chartPolar } from './chart-geometry'
import { chartNullableValues } from './chart-cartesian-api'
import { cloneChartOptions, definedChartStyleFields } from './chart-options'

export const ChartRadarShapes = Object.freeze({ Polygon: 'polygon', Circle: 'circle' } as const)

export function patchChartRadarOptions(current: ChartRadarOptions | undefined, patch: ChartNullablePatch<ChartRadarOptions> | undefined): ChartRadarOptions | undefined {
  if (patch === undefined) return undefined
  const previous = cloneChartOptions({ radar: current })?.radar
  const values = chartNullableValues<ChartRadarOptions>(patch)
  const next: ChartRadarOptions = { ...previous, ...values }
  if (values.axisLabel) next.axisLabel = { ...previous?.axisLabel, ...values.axisLabel }
  return cloneChartOptions({ radar: next })!.radar
}

/** Dedicated scale-label fields inherit missing values from the shared value-axis appearance. */
export function chartRadarAxisLabel(chart: { options?: ChartSemanticOptions; yAxis?: Partial<ChartAxisOptions> }): ChartRadarAxisLabelOptions {
  const font = chart.options?.appearance?.font, legacy = chart.yAxis?.label, authored = chart.options?.radar?.axisLabel
  return { ...definedChartStyleFields(font), ...definedChartStyleFields(legacy), format: chart.yAxis?.format,
    visible: legacy?.visible ?? false, ...definedChartStyleFields(authored) }
}

export function chartRadarExtent(data: ChartData, axis: ChartAxisOptions = { scaleType: 'linear' }): { min: number; max: number } {
  if (axis.scaleType !== 'linear') throw new TypeError('雷达图须使用线性数值轴。')
  const values = function* () { for (const series of data.series) for (const value of series.values) if (value !== null) yield value }
  return chartAxisExtent(values(), axis)
}

export function chartRadarAxisLine(chart: Pick<ChartModel, 'xAxis' | 'yAxis'>) {
  const line = { ...chart.xAxis.line }
  for (const [key, value] of Object.entries(chart.yAxis.line ?? {})) if (value !== undefined) Object.assign(line, { [key]: value })
  return line
}

/** Both adapters use the same radius, direction, fixed range and four radial intervals. */
export function chartRadarGeometry(data: ChartData, chart: Pick<ChartModel, 'xAxis' | 'yAxis'>, plot: RenderRect) {
  const center = { x: plot.x + plot.width / 2, y: plot.y + plot.height / 2 }, radius = Math.min(plot.width, plot.height) * 0.36
  const extent = chartRadarExtent(data, chart.yAxis), ticks = chartAxisTicks(extent.min, extent.max, chart.yAxis)
  const angle = (categoryIndex: number): number => -Math.PI / 2 + (chart.xAxis.reverse ? -1 : 1) * categoryIndex * Math.PI * 2 / Math.max(1, data.categories.length)
  const vertex = (categoryIndex: number, ratio: number) => chartPolar(center, radius * ratio, angle(categoryIndex))
  const point = (categoryIndex: number, value: number) => vertex(categoryIndex, chartAxisRatio(value, extent.min, extent.max, chart.yAxis))
  return { center, radius, ...extent, ticks, angle, vertex, point }
}
