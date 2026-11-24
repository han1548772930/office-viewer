import { color as echartsColor } from 'echarts/core'
import type { ChartColor, ChartLinearGradient } from '../../domain/model'

export const CHART_LINEAR_GRADIENT_TYPE = 'linear-gradient' as const
/** Local authoring budget; unsupported source metadata remains in the Office sidecar. */
export const CHART_GRADIENT_MAX_STOPS = 256
export type ChartLinearGradientInput = Omit<ChartLinearGradient, 'type'>

export function isChartLinearGradient(value: unknown): value is ChartLinearGradient {
  return value !== null && typeof value === 'object' && !Array.isArray(value) && Reflect.get(value, 'type') === CHART_LINEAR_GRADIENT_TYPE
}

/** A detached authoring object, never an ECharts/Canvas paint object. Stops keep
 * authored order; render adapters stable-sort a detached copy by offset. */
export function chartLinearGradient(input: ChartLinearGradientInput): ChartLinearGradient {
  plainColorObject(input, ['type', 'start', 'end', 'stops'], '图表渐变')
  if (Object.hasOwn(input, 'type') && Reflect.get(input, 'type') !== CHART_LINEAR_GRADIENT_TYPE) throw new TypeError('无效的图表渐变类型。')
  const gradient = { ...input, type: CHART_LINEAR_GRADIENT_TYPE }
  validateChartColor(gradient)
  return { type: CHART_LINEAR_GRADIENT_TYPE, start: { ...gradient.start }, end: { ...gradient.end },
    stops: gradient.stops.map(stop => ({ ...stop })) }
}

/** Finite coordinates may lie outside the mark. The local renderer accepts
 * 0–1 stops/alpha; exact upstream boundary validation remains unverified. */
export function validateChartColor(value: unknown, name = '图表颜色'): asserts value is ChartColor {
  if (typeof value === 'string') { chartColorChannels(value); return }
  plainColorObject(value, ['type', 'start', 'end', 'stops'], name)
  if (value.type !== CHART_LINEAR_GRADIENT_TYPE) throw new TypeError(`${name}需要 CSS 颜色或 linear-gradient 对象。`)
  for (const key of ['start', 'end'] as const) {
    const point = value[key]
    plainColorObject(point, ['x', 'y'], `${name}.${key}`)
    for (const coordinate of ['x', 'y'] as const) if (typeof point[coordinate] !== 'number' || !Number.isFinite(point[coordinate])) throw new RangeError(`${name}.${key}.${coordinate}须为有限数值。`)
  }
  const stops = value.stops
  if (!Array.isArray(stops) || stops.length < 2 || stops.length > CHART_GRADIENT_MAX_STOPS) throw new RangeError(`${name}需要 2 到 ${CHART_GRADIENT_MAX_STOPS} 个色标。`)
  if (Object.getPrototypeOf(stops) !== Array.prototype || Reflect.ownKeys(stops).some(key => typeof key !== 'string'
    || key !== 'length' && (!/^(0|[1-9]\d*)$/.test(key) || Number(key) >= stops.length)
    || Object.getOwnPropertyDescriptor(stops, key)?.get !== undefined || Object.getOwnPropertyDescriptor(stops, key)?.set !== undefined)) throw new TypeError(`${name}色标需要不含额外属性的普通数组。`)
  for (let index = 0; index < stops.length; index++) {
    if (!Object.hasOwn(stops, index)) throw new TypeError(`${name}色标不能含稀疏位置。`)
    const stop: unknown = stops[index]
    plainColorObject(stop, ['offset', 'color', 'opacity'], `${name}色标 ${index + 1}`)
    if (typeof stop.offset !== 'number' || !Number.isFinite(stop.offset) || stop.offset < 0 || stop.offset > 1) throw new RangeError(`${name}色标位置须为 0 到 1 的有限数值。`)
    if (typeof stop.color !== 'string') throw new TypeError(`${name}色标需要 CSS 颜色。`)
    chartColorChannels(stop.color)
    if (stop.opacity !== undefined && (typeof stop.opacity !== 'number' || !Number.isFinite(stop.opacity) || stop.opacity < 0 || stop.opacity > 1)) throw new RangeError(`${name}色标不透明度须为 0 到 1 的有限数值。`)
  }
}

function plainColorObject(value: unknown, keys: readonly string[], name: string): asserts value is Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value) || Object.getPrototypeOf(value) !== Object.prototype
    || Reflect.ownKeys(value).some(key => typeof key !== 'string' || !keys.includes(key)
      || Object.getOwnPropertyDescriptor(value, key)?.get !== undefined || Object.getOwnPropertyDescriptor(value, key)?.set !== undefined)) throw new TypeError(`${name}须为不含额外属性的普通对象。`)
}

/** Local legend/symbol fallback: first stop by offset, with its alpha consumed.
 * The complete authored gradient remains in the source model/description. */
export function chartColorRepresentative(color: ChartColor): string {
  if (typeof color === 'string') return color
  validateChartColor(color)
  let first = color.stops[0]!
  for (const stop of color.stops) if (stop.offset < first.offset) first = stop
  return chartSolidColorOpacity(first.color, first.opacity ?? 1)
}

export function chartSolidColorOpacity(color: string, opacity = 1): string {
  const [r, g, b, a] = chartColorChannels(color)
  return `rgba(${Math.round(r * 255)}, ${Math.round(g * 255)}, ${Math.round(b * 255)}, ${a * opacity})`
}

/** ECharts' pure CSS parser supports named colors and HSL without consulting DOM or the current theme. */
export function chartColorChannels(value: string): [number, number, number, number] {
  const parsed = echartsColor.parse(value)
  if (!parsed || parsed.length < 4 || parsed.some(channel => !Number.isFinite(channel))) throw new TypeError(`无效或尚不支持的图表颜色：${value}`)
  return [parsed[0]! / 255, parsed[1]! / 255, parsed[2]! / 255, parsed[3]!]
}

/** Canvas2D and the shared WebGPU geometry parser receive the same canonical color. */
export function normalizeChartColor(value: string): string {
  const [r, g, b, a] = chartColorChannels(value)
  return `rgba(${Math.round(r * 255)}, ${Math.round(g * 255)}, ${Math.round(b * 255)}, ${a})`
}
