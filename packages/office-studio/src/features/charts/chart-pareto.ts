import type { ChartModel, ChartNullablePatch, ChartParetoBarStyleOptions, ChartParetoLineStyleOptions, ChartSeriesOptions, ChartType } from '../../domain/model'
import { chartNullableValues } from './chart-cartesian-api'
import { chartSeriesStyle, cloneChartOptions, definedChartStyleFields } from './chart-options'
import { chartPointStyle, type ChartComputedStatistics } from './chart-derived'
import { chartColorRepresentative } from './chart-color'
import { chartAxisExtent } from './chart-geometry'
import { chartMarkerShape } from './chart-markers'

export const CHART_PARETO_CUMULATIVE_NAME = '累计百分比'
export function assertChartParetoType(type: ChartType): void {
  if (type !== 'pareto') throw new TypeError('此配置仅适用于帕累托图。')
}

export function chartParetoBarStyleValues(value: ChartNullablePatch<ChartParetoBarStyleOptions> | undefined): ChartParetoBarStyleOptions | undefined {
  return value === undefined ? undefined : cloneChartOptions({ pareto: { barStyle: chartNullableValues<ChartParetoBarStyleOptions>(value) } })!.pareto!.barStyle
}

/** Fixed builder contract: a non-positive width is ignored; null is a local
 * reset extension and remains an explicit undefined for group merging. */
export function chartParetoLineStyleValues(value: ChartNullablePatch<ChartParetoLineStyleOptions> | undefined): ChartParetoLineStyleOptions | undefined {
  if (value === undefined) return undefined
  const patch = chartNullableValues<ChartParetoLineStyleOptions>(value)
  if (typeof patch.width === 'number' && Number.isFinite(patch.width) && patch.width <= 0) delete patch.width
  return cloneChartOptions({ pareto: { cumulativeLineStyle: patch } })!.pareto!.cumulativeLineStyle
}

/** Generic source-field/name styles describe the business bars. The dedicated
 * Pareto layer adds mark-specific overrides; point layers remain independent. */
export function chartParetoBarStyle(chart: Pick<ChartModel, 'options'>, field: number, position: number, palette: readonly string[], sourceName?: string): ChartSeriesOptions {
  const base = chartSeriesStyle(chart, field, position, palette, sourceName), bar = chart.options?.pareto?.barStyle
  if (!bar) return base
  const specific: ChartSeriesOptions = { color: bar.color, fillOpacity: bar.opacity, border: bar.border, label: bar.label, dataPoints: bar.dataPoints }
  return chartSeriesStyle({ options: { allSeriesStyle: base, series: { [field]: specific } } }, field, position, palette)
}

/** Fixed facade defaults: fourth theme color, 2px solid, opacity1 and 7px points.
 * The derived line owns its style independently from source business series. */
export function chartParetoCumulativeStyle(chart: Pick<ChartModel, 'options'>, palette: readonly string[]): ChartSeriesOptions {
  const line = chart.options?.pareto?.cumulativeLineStyle
  return { color: line?.color ?? palette[3 % palette.length] ?? '#d97706', fillOpacity: line?.opacity ?? 1,
    border: { width: line?.width !== undefined && line.width > 0 ? line.width : 2, opacity: line?.opacity ?? 1, dashType: line?.dashType ?? 'solid' },
    point: { size: 7, shape: 'circle', ...definedChartStyleFields(line?.point) }, label: { ...definedChartStyleFields(line?.label) },
    dataPoints: line?.dataPoints, rightAxis: true }
}

/** Linear percentages default to 0–100. Log axes need positive bounds, and a
 * single authored bound outside that default range needs a valid other edge. */
export function chartParetoCumulativeExtent(chart: Pick<ChartModel, 'options'>, values: readonly number[]): { min: number; max: number } {
  const axis = { scaleType: 'linear' as const, ...chart.options?.rightYAxis }
  const min = axis.min ?? 0, max = axis.max ?? 100
  return axis.scaleType === 'log' || min >= max ? chartAxisExtent(values, axis) : { min, max }
}

/** The local per-point contract addresses the sorted, aggregated Pareto marks.
 * Point colors can shade a segment; its symbols always use a solid fallback. */
export function chartParetoCumulativePointStyle(style: ChartSeriesOptions, index: number): ChartSeriesOptions {
  const point = chartPointStyle(style, index, 'pareto'), override = style.dataPoints?.[index]
  return { ...point, border: { ...point.border,
    opacity: override?.border?.opacity ?? override?.fillOpacity ?? style.border?.opacity ?? 1 } }
}

/** The curve's stroke is not a marker border. Only the materialized point's
 * authored border applies to its symbol, and its fallback uses the raw color
 * so fill opacity cannot be consumed a second time by the outline. */
export function chartParetoCumulativeMarkerStyle(style: ChartSeriesOptions, index?: number): ChartSeriesOptions {
  const current = index === undefined ? style : chartParetoCumulativePointStyle(style, index)
  const color = current.point?.color ?? chartColorRepresentative(current.color ?? '#d97706')
  const border = index === undefined ? undefined : style.dataPoints?.[index]?.border
  return { ...current, color, border: border ? { ...definedChartStyleFields(border), color: border.color ?? color,
    width: border.width ?? 1, opacity: border.opacity ?? current.fillOpacity ?? 1 }
    : chartMarkerShape(current.point?.shape ?? 'circle').hollow ? { color, width: 1, opacity: current.fillOpacity ?? 1 } : { visible: false, width: 0 } }
}

/** Label percentages project the materialized nonnegative business values.
 * They are independent of the fixed line's two-decimal truncation: equal
 * thirds remain equal thirds rather than 33.33, 33.33 and 33.34 percent. */
export function chartParetoValuePercentages(values: readonly number[]): number[] {
  let rawTotal = 0, scale = 0
  for (const value of values) {
    if (!Number.isFinite(value) || value < 0) throw new RangeError('帕累托业务占比须使用有限非负分类合计。')
    rawTotal += value; scale = Math.max(scale, value)
  }
  if (!scale) return values.map(() => 0)
  if (Number.isFinite(rawTotal)) return values.map(value => value / rawTotal * 100)
  const total = values.reduce((sum, value) => sum + value / scale, 0)
  return values.map(value => (value / scale) / total * 100)
}

/** Settings consume the current draft's Rust statistics result directly.
 * Its prepared source has already passed aggregation, Top N and normalization. */
export function chartParetoSettingsCategories(statistics?: Pick<ChartComputedStatistics, 'pareto'>): string[] {
  return statistics?.pareto?.categories ?? []
}
