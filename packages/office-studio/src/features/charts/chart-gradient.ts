import type { ChartColor, ChartLinearGradient, ChartSemanticOptions, ChartSeriesOptions, ChartType } from '../../domain/model'
import type { RenderRect } from '../../rendering/contracts'
import type { ChartSeriesLayout } from './chart-engine'
import { resolveChartType } from './chart-catalog'
import { chartColorChannels, chartSolidColorOpacity, isChartLinearGradient, validateChartColor } from './chart-color'
import { createChartAutomaticGradient, type ChartAutomaticGradient, type ChartAutomaticGradientMark } from './chart-auto-gradient'

const FILL_FAMILIES = ['column', 'bar', 'line', 'area', 'combination', 'pie', 'waterfall', 'histogram', 'pareto']

/** Fixed facade type map: line/area 2D data points are solid symbols; the
 * generic 3D/Combination data-point declarations accept fill-capable colors. */
export function supportsChartExplicitGradient(type: ChartType, target: 'series' | 'dataPoint' | 'symbol' = 'series'): boolean {
  const descriptor = resolveChartType(type)
  if (target === 'symbol' || !FILL_FAMILIES.includes(descriptor.family)) return false
  return target !== 'dataPoint' || descriptor.threeD || !['line', 'area'].includes(descriptor.family)
}

/** Validate a newly authored patch. Callers can separately retain inactive
 * colors from another chart type without claiming those marks support them. */
export function validateChartGradientSeriesPatch(type: ChartType, style: ChartSeriesOptions | undefined, name = '系列'): void {
  if (!style) return
  if (style.color !== undefined) validateChartColor(style.color, `${name}颜色`)
  if (isChartLinearGradient(style.color) && !supportsChartExplicitGradient(type)) throw new TypeError('当前图型的系列颜色只接受纯色。')
  for (const point of Object.values(style.dataPoints ?? {})) {
    if (point?.color !== undefined) validateChartColor(point.color, `${name}项目颜色`)
    if (isChartLinearGradient(point?.color) && !supportsChartExplicitGradient(type, 'dataPoint')) throw new TypeError('当前图型的数据点颜色只接受纯色。')
  }
  for (const [target, role] of Object.entries(style.waterfallStyles ?? {})) validateChartGradientSeriesPatch(type, role, `${name}.${target}`)
}

/** Checks authoring changes against the prior model. Equal inactive gradients
 * survive type switches and full collection replacement; new unsupported
 * colors reject before the candidate becomes content/history. */
export function validateChartGradientAuthoring(type: ChartType, options: ChartSemanticOptions | undefined, previous?: ChartSemanticOptions): void {
  const changed = (style: ChartSeriesOptions | undefined, before: ChartSeriesOptions | undefined): ChartSeriesOptions | undefined => {
    if (!style) return undefined
    const patch: ChartSeriesOptions = {}
    if (isChartLinearGradient(style.color) && !sameGradient(style.color, before?.color)) patch.color = style.color
    for (const [key, point] of Object.entries(style.dataPoints ?? {})) if (isChartLinearGradient(point?.color) && !sameGradient(point.color, before?.dataPoints?.[Number(key)]?.color)) {
      ;(patch.dataPoints ??= {})[Number(key)] = { color: point.color }
    }
    for (const target of ['positive', 'negative', 'subtotal'] as const) if (style.waterfallStyles?.[target]) {
      ;(patch.waterfallStyles ??= {})[target] = changed(style.waterfallStyles[target], before?.waterfallStyles?.[target])
    }
    return patch
  }
  validateChartGradientSeriesPatch(type, changed(options?.allSeriesStyle, previous?.allSeriesStyle), '所有系列')
  for (const [key, style] of Object.entries(options?.series ?? {})) validateChartGradientSeriesPatch(type, changed(style, previous?.series?.[Number(key)]), `系列 ${key}`)
  for (const entry of options?.namedSeries ?? []) validateChartGradientSeriesPatch(type, changed(entry.style, previous?.namedSeries?.find(before => before.name === entry.name)?.style), `名称系列 ${entry.name}`)
  for (const target of ['positive', 'negative', 'subtotal'] as const) validateChartGradientSeriesPatch(type,
    changed(options?.waterfall?.styles?.[target], previous?.waterfall?.styles?.[target]), `瀑布图.${target}`)
}

function sameGradient(color: ChartLinearGradient, previous: ChartColor | undefined): boolean {
  return isChartLinearGradient(previous) && color.start.x === previous.start.x && color.start.y === previous.start.y
    && color.end.x === previous.end.x && color.end.y === previous.end.y && color.stops.length === previous.stops.length
    && color.stops.every((stop, index) => { const before = previous.stops[index]!; return stop.offset === before.offset && stop.color === before.color && stop.opacity === before.opacity })
}

/** Explicit semantic colors take precedence over the automatic switch. Real
 * patterns retain the existing local precedence; symbols use solid fallbacks. */
export function createChartMarkGradient(
  chart: { type: ChartType; options?: ChartSemanticOptions },
  style: Pick<ChartSeriesOptions, 'color' | 'fillOpacity' | 'fillPattern'>,
  mark: ChartAutomaticGradientMark,
  direction: { negative?: boolean; reverse?: boolean; horizontal?: boolean; shade?: number; symbol?: boolean } = {},
  bounds?: RenderRect,
): ChartAutomaticGradient | undefined {
  if (style.fillPattern) return undefined
  if (isChartLinearGradient(style.color)) {
    if (!supportsChartExplicitGradient(chart.type, direction.symbol ? 'symbol' : 'series')) return undefined
    validateChartColor(style.color)
    const opacity = style.fillOpacity ?? 1, shade = direction.shade ?? 1
    const stops = style.color.stops.map(stop => ({ offset: stop.offset, color: shadeColor(chartSolidColorOpacity(stop.color, (stop.opacity ?? 1) * opacity), shade) }))
      .sort((first, second) => first.offset - second.offset)
    return { authored: true, start: { ...style.color.start }, end: { ...style.color.end }, stops, ...(bounds ? { bounds: { ...bounds } } : {}) }
  }
  const color = style.color === undefined ? undefined : shadeColor(style.color, direction.shade ?? 1)
  return createChartAutomaticGradient(chart, { ...style, color }, mark, direction, bounds)
}

function shadeColor(color: string, shade: number): string {
  if (shade === 1) return color
  const [red, green, blue, alpha] = chartColorChannels(color)
  return `rgba(${Math.round(Math.min(1, red * shade) * 255)}, ${Math.round(Math.min(1, green * shade) * 255)}, ${Math.round(Math.min(1, blue * shade) * 255)}, ${alpha})`
}

/** A line/area series uses one mark-space across its disconnected fragments. */
export function chartSeriesGradientBounds(series: Pick<ChartSeriesLayout, 'points'>, includeBase = false, additional?: Iterable<{ x: number; y: number }>): RenderRect | undefined {
  const points = function* (): Generator<{ x: number; y: number }> {
    for (const point of series.points) {
      yield point
      if (includeBase && point.baseY !== undefined) yield { x: point.baseX ?? point.x, y: point.baseY }
    }
    if (additional) yield* additional
  }
  return chartGradientPointBounds(points())
}

/** Projected faces use one detached mark rectangle, including depth extrusion. */
export function chartGradientPointBounds(points: Iterable<{ x: number; y: number }>): RenderRect | undefined {
  let left = Infinity, top = Infinity, right = -Infinity, bottom = -Infinity
  for (const { x, y } of points) {
    if (!Number.isFinite(x) || !Number.isFinite(y)) continue
    left = Math.min(left, x); top = Math.min(top, y); right = Math.max(right, x); bottom = Math.max(bottom, y)
  }
  return Number.isFinite(left) ? { x: left, y: top, width: right - left, height: bottom - top } : undefined
}
