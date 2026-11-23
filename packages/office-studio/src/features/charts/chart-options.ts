import type { ChartModel, ChartSemanticOptions, ChartSeriesOptions } from '../../domain/model'
import { validateChartOptions } from './chart-option-validation'
export { validateChartOptions, validateChartPalette, validateChartAxis, validateChartFont } from './chart-option-validation'

const DEFAULT_COLORS = ['#2563eb', '#dc2626', '#059669', '#d97706', '#7c3aed', '#0891b2']

export function cloneChartOptions(options: ChartSemanticOptions | undefined): ChartSemanticOptions | undefined {
  if (options === undefined) return undefined
  validateChartOptions(options)
  return cloneOptionValue(options) as ChartSemanticOptions
}

/** Groups omitted from a patch are retained. An explicitly supplied collection replaces that collection. */
export function mergeChartOptions(current: ChartSemanticOptions | undefined, patch: ChartSemanticOptions | undefined): ChartSemanticOptions | undefined {
  if (patch === undefined) return cloneChartOptions(current)
  const merged: ChartSemanticOptions = { ...cloneChartOptions(current), ...cloneChartOptions(patch) }
  for (const key of ['appearance', 'titleStyle', 'subtitleStyle', 'allSeriesStyle', 'rightYAxis', 'mapping', 'threeD', 'bar', 'area', 'pie', 'radar', 'bubble', 'funnel', 'relation', 'waterfall', 'pareto', 'histogram', 'boxplot', 'candlestick', 'heatmap', 'treemap', 'sunburst', 'gauge', 'chord', 'wordCloud', 'aggregation'] as const) {
    if (current?.[key] && patch[key]) Object.assign(merged, { [key]: mergeOptionGroup(current[key], patch[key]) })
  }
  return cloneChartOptions(merged)
}

export function cloneOptionValue(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(cloneOptionValue)
  if (value && typeof value === 'object') return Object.fromEntries(Object.entries(value).map(([key, entry]) => [key, cloneOptionValue(entry)]))
  return value
}

/** Explicit undefined clears a property; collections replace rather than merging by array position. */
function mergeOptionGroup(current: unknown, patch: unknown): unknown {
  if (!patch || typeof patch !== 'object' || Array.isArray(patch) || !current || typeof current !== 'object' || Array.isArray(current)) return cloneOptionValue(patch)
  const merged: Record<string, unknown> = { ...current, ...patch }
  for (const [key, value] of Object.entries(patch)) if (value && typeof value === 'object' && !Array.isArray(value)) merged[key] = mergeOptionGroup(Reflect.get(current, key), value)
  return merged
}

export function chartPalette(chart: Pick<ChartModel, 'palette'>, fallback: readonly string[] = DEFAULT_COLORS): string[] {
  return [...(chart.palette?.length ? chart.palette : fallback.length ? fallback : DEFAULT_COLORS)]
}

export function chartSeriesStyle(chart: ChartModel, field: number, index: number, palette: readonly string[] = DEFAULT_COLORS): ChartSeriesOptions {
  const shared = chart.options?.allSeriesStyle, specific = chart.options?.series?.[field]
  return {
    color: palette[index % palette.length],
    ...shared, ...specific,
    border: { ...shared?.border, ...specific?.border },
    point: { ...shared?.point, ...specific?.point },
    label: { ...shared?.label, ...specific?.label },
  }
}

