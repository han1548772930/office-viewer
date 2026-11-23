import type { ChartModel, ChartSemanticOptions, ChartSeriesOptions } from '../../domain/model'
import { validateChartOptions } from './chart-option-validation'
import { chartNamedSeriesStyle } from './chart-series-selector'
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
  for (const key of ['appearance', 'titleStyle', 'subtitleStyle', 'allSeriesStyle', 'axisPointer', 'rightYAxis', 'depthAxis', 'mapping', 'threeD', 'bar', 'area', 'pie', 'radar', 'bubble', 'funnel', 'relation', 'waterfall', 'pareto', 'histogram', 'boxplot', 'candlestick', 'heatmap', 'treemap', 'sunburst', 'gauge', 'chord', 'wordCloud', 'aggregation'] as const) {
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
  for (const [key, value] of Object.entries(patch)) if (key !== 'fillPattern' && value && typeof value === 'object' && !Array.isArray(value)) merged[key] = mergeOptionGroup(Reflect.get(current, key), value)
  return merged
}

export function chartPalette(chart: Pick<ChartModel, 'palette'>, fallback: readonly string[] = DEFAULT_COLORS): string[] {
  return [...(chart.palette?.length ? chart.palette : fallback.length ? fallback : DEFAULT_COLORS)]
}

export function chartSeriesStyle(chart: Pick<ChartModel, 'options'>, field: number, index: number, palette: readonly string[] = DEFAULT_COLORS, sourceName?: string): ChartSeriesOptions {
  const shared = chart.options?.allSeriesStyle, specific = chart.options?.series?.[field], named = chartNamedSeriesStyle(chart.options, sourceName)
  return {
    color: palette[index % palette.length],
    ...definedChartStyleFields(shared), ...definedChartStyleFields(specific), ...definedChartStyleFields(named),
    border: { ...definedChartStyleFields(shared?.border), ...definedChartStyleFields(specific?.border), ...definedChartStyleFields(named?.border) },
    point: { ...definedChartStyleFields(shared?.point), ...definedChartStyleFields(specific?.point), ...definedChartStyleFields(named?.point) },
    label: { ...definedChartStyleFields(shared?.label), ...definedChartStyleFields(specific?.label), ...definedChartStyleFields(named?.label) },
    dataPoints: mergeChartSeriesDataPoints(shared?.dataPoints, specific?.dataPoints, named?.dataPoints),
  }
}

export function chartSeriesColorIsAuthored(chart: Pick<ChartModel, 'options'>, field: number, sourceName?: string): boolean {
  return chartNamedSeriesStyle(chart.options, sourceName)?.color !== undefined || chart.options?.series?.[field]?.color !== undefined
    || chart.options?.allSeriesStyle?.color !== undefined
}

function mergeChartSeriesDataPoints(...groups: Array<ChartSeriesOptions['dataPoints']>): ChartSeriesOptions['dataPoints'] {
  const layers = groups.filter(group => group !== undefined)
  if (layers.length < 2) return layers[0]
  let result: ChartSeriesOptions['dataPoints']
  for (const group of layers) for (const [key, point] of Object.entries(group ?? {})) {
    if (point === undefined) continue
    const index = Number(key), previous = result?.[index]
    ;(result ??= {})[index] = { ...definedChartStyleFields(previous), ...definedChartStyleFields(point),
      border: { ...definedChartStyleFields(previous?.border), ...definedChartStyleFields(point.border) } }
  }
  return result
}

/** A cleared per-field override resumes inheritance instead of erasing the inherited value. */
export function definedChartStyleFields<T extends object>(style: T | undefined): Partial<T> {
  const defined: Partial<T> = {}
  for (const [key, value] of Object.entries(style ?? {})) if (value !== undefined) Object.assign(defined, { [key]: value })
  return defined
}

