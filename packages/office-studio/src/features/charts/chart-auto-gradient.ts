import type { ChartModel, ChartSeriesOptions, ChartType } from '../../domain/model'
import type { RenderRect } from '../../rendering/contracts'
import { resolveChartType } from './chart-catalog'
import { chartColorChannels } from './chart-color'

export interface ChartGradientPoint { x: number; y: number }
export interface ChartGradientStop { offset: number; color: string }
/** Derived paint only. The author saves autoGradientFill, never renderer objects. */
export interface ChartAutomaticGradient {
  /** Internal provenance, used by 3D geometry; never persisted in ChartColor. */
  authored?: boolean
  start: ChartGradientPoint
  end: ChartGradientPoint
  stops: ChartGradientStop[]
  /** Geometry fragments of one mark use one gradient coordinate space. */
  bounds?: RenderRect
}
export type ChartAutomaticGradientMark = 'column' | 'bar' | 'line' | 'area' | 'pie'

/** v1.6.0 ChartSetupPanel: area charts default on; other chart types default off. */
export function chartAutoGradientFill(chart: Pick<ChartModel, 'type' | 'options'>): boolean {
  return chart.options?.autoGradientFill ?? (resolveChartType(chart.type).family === 'area')
}

/** The setup switch follows chart-ui Qo/Zo, including inactive saved overrides. */
export function supportsChartAutoGradientControl(type: ChartType): boolean {
  return ['column', 'bar', 'area', 'combination', 'pie', 'waterfall'].includes(resolveChartType(type).family)
}

export function createChartAutomaticGradient(
  chart: Pick<ChartModel, 'type' | 'options'>,
  style: Pick<ChartSeriesOptions, 'color' | 'fillOpacity' | 'fillPattern'>,
  mark: ChartAutomaticGradientMark,
  direction: { negative?: boolean; reverse?: boolean; horizontal?: boolean } = {},
  bounds?: RenderRect,
): ChartAutomaticGradient | undefined {
  if (!['column', 'bar', 'line', 'area', 'combination', 'pie', 'waterfall'].includes(resolveChartType(chart.type).family)
    || !chartAutoGradientFill(chart) || style.fillPattern || typeof style.color !== 'string' || !style.color) return undefined
  const color = style.color, opacity = style.fillOpacity ?? 1
  let start = { x: 0, y: 1 }, end = { x: 0, y: 0 }
  let stops: ChartGradientStop[]
  if (mark === 'pie') {
    start = { x: 0, y: 0 }; end = { x: 1, y: 1 }
    stops = [{ offset: 0, color: gradientTint(color, 0.28, opacity) }, { offset: 1, color: gradientTint(color, -0.12, opacity) }]
  } else {
    if (mark === 'bar' || direction.horizontal) { start = { x: 0, y: 0 }; end = { x: 1, y: 0 } }
    if (Boolean(direction.negative) !== Boolean(direction.reverse)) [start, end] = [end, start]
    stops = mark === 'area'
      ? [{ offset: 0, color: gradientTint(color, 0, opacity * 0.06) }, { offset: 1, color: gradientTint(color, 0, opacity) }]
      : [{ offset: 0, color: gradientTint(color, 0.28, opacity) }, { offset: 1, color: gradientTint(color, -0.08, opacity) }]
  }
  return { start, end, stops, ...(bounds ? { bounds: { ...bounds } } : {}) }
}

function gradientTint(color: string, amount: number, opacity: number): string {
  const [red, green, blue, alpha] = chartColorChannels(color)
  const tint = (channel: number): number => Math.round((amount >= 0 ? channel + (1 - channel) * amount : channel * (1 + amount)) * 255)
  return 'rgba(' + tint(red) + ', ' + tint(green) + ', ' + tint(blue) + ', ' + alpha * opacity + ')'
}

export function chartGradientColorAt(gradient: ChartAutomaticGradient, position: number): string {
  const value = Math.min(1, Math.max(0, position)), stops = gradient.stops
  let first = stops[0]!, second = stops.at(-1)!
  if (value < first.offset) return first.color
  if (value >= second.offset) return second.color
  for (let index = 1; index < stops.length; index++) if (value < stops[index]!.offset) {
    first = stops[index - 1]!; second = stops[index]!; break
  }
  const ratio = second.offset === first.offset ? 0 : Math.min(1, Math.max(0, (value - first.offset) / (second.offset - first.offset)))
  const a = chartColorChannels(first.color), b = chartColorChannels(second.color)
  const channel = (index: number): number => a[index]! + (b[index]! - a[index]!) * ratio
  return 'rgba(' + Math.round(channel(0) * 255) + ', ' + Math.round(channel(1) * 255) + ', ' + Math.round(channel(2) * 255) + ', ' + channel(3) + ')'
}

/** Bounds keep fragmented geometry in chart coordinates; alpha is already in stops. */
export function chartEChartsAutomaticGradient(gradient: ChartAutomaticGradient | undefined) {
  if (!gradient) return undefined
  const bounds = gradient.bounds
  const color = { type: 'linear' as const,
    x: bounds ? bounds.x + gradient.start.x * bounds.width : gradient.start.x,
    y: bounds ? bounds.y + gradient.start.y * bounds.height : gradient.start.y,
    x2: bounds ? bounds.x + gradient.end.x * bounds.width : gradient.end.x,
    y2: bounds ? bounds.y + gradient.end.y * bounds.height : gradient.end.y,
    colorStops: gradient.stops.map(stop => ({ ...stop })), global: Boolean(bounds) }
  if (![color.x, color.y, color.x2, color.y2].every(Number.isFinite)) throw new RangeError('图表渐变坐标须为有限值。')
  return color
}
