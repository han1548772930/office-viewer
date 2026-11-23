import type {
  CellScalar, ChartFontOptions, ChartGaugeAnglesOptions, ChartGaugeAxisLabelOptions, ChartGaugeAxisLineOptions,
  ChartGaugeAxisTickOptions, ChartGaugeAnchorOptions, ChartGaugeDetailOptions, ChartGaugeMode, ChartGaugeOptions,
  ChartGaugePointerOptions, ChartGaugePointerShape, ChartGaugeProgressOptions, ChartGaugeSplitLineOptions,
  ChartGaugeTitleOptions, ChartLabelOptions,
} from '../../domain/model'
import { chartLabelText, type ChartVector } from './chart-geometry'
import { chartLinearRatio } from './chart-math'

export const ChartGaugeModes = Object.freeze({ Modern: 'modern', Classic: 'classic' } as const)
export const ChartGaugePointerShapes = Object.freeze({ Needle: 'needle', Slim: 'slim', Arrow: 'arrow' } as const)
export const CHART_GAUGE_PRESETS = Object.freeze([
  { value: 'modern', label: '现代仪表' }, { value: 'classic', label: '经典仪表' },
  { value: 'statusDial', label: '状态刻度' }, { value: 'progress', label: '进度仪表' }, { value: 'minimal', label: '指针仪表' },
] as const)
export type ChartGaugePreset = typeof CHART_GAUGE_PRESETS[number]['value']

type Resolved<T, K extends keyof T> = T & Required<Pick<T, K>>
export interface ChartResolvedGaugeOptions {
  mode: ChartGaugeMode
  min: number
  max: number
  startAngle: number
  endAngle: number
  splitNumber: number
  axisLine: Resolved<ChartGaugeAxisLineOptions, 'visible' | 'width'>
  axisTick: Resolved<ChartGaugeAxisTickOptions, 'visible' | 'splitNumber' | 'length' | 'width'>
  splitLine: Resolved<ChartGaugeSplitLineOptions, 'visible' | 'length' | 'width'>
  axisLabel: Resolved<ChartGaugeAxisLabelOptions, 'visible' | 'distance' | 'fontSize'>
  pointer: Resolved<ChartGaugePointerOptions, 'visible' | 'length' | 'width' | 'shape'>
  anchor: Resolved<ChartGaugeAnchorOptions, 'visible' | 'size' | 'borderWidth'>
  progress: Resolved<ChartGaugeProgressOptions, 'visible' | 'width' | 'rounded'>
  title: Resolved<ChartGaugeTitleOptions & ChartLabelOptions, 'visible' | 'offsetY' | 'fontSize'>
  detail: Resolved<ChartGaugeDetailOptions & ChartLabelOptions, 'visible' | 'offsetY' | 'fontSize'>
}

/** Fixed 1.6.0 public Gauge presets, adapted to the chart's theme and actual tile radius.
 * Authored dimensions stay in the model; annular widths and tick lengths are bounded for small tiles. */
export function resolveChartGaugeOptions(options: ChartGaugeOptions | undefined, radius: number, font: ChartFontOptions = {}, label: ChartLabelOptions = {}): ChartResolvedGaugeOptions {
  const mode = options?.mode ?? 'modern', legacy = mode === 'pointer' || mode === 'progress', classic = mode === 'classic'
  const width = legacy ? radius * 0.1 : classic ? 8 : 18
  const axisWidth = Math.min(radius, options?.axisLine?.width ?? width), innerRadius = Math.max(0, radius - axisWidth)
  const pointerWidth = options?.pointer?.width ?? (legacy ? 4 : 5)
  const labels: ChartLabelOptions = { ...gaugeFont(label, options?.label), visible: options?.label?.visible ?? label.visible,
    position: options?.label?.position ?? label.position, contentType: options?.label?.contentType ?? label.contentType,
    format: options?.label?.format ?? label.format }
  return {
    mode, min: options?.min ?? 0, max: options?.max ?? 100,
    startAngle: options?.startAngle ?? 225, endAngle: options?.endAngle ?? -45, splitNumber: options?.splitNumber ?? 5,
    axisLine: { ...options?.axisLine, visible: options?.axisLine?.visible ?? true, width: axisWidth },
    axisTick: { ...options?.axisTick, visible: options?.axisTick?.visible ?? classic,
      splitNumber: options?.axisTick?.splitNumber ?? (classic || legacy ? 5 : 2), length: Math.min(innerRadius, options?.axisTick?.length ?? (classic ? 8 : 5)), width: options?.axisTick?.width ?? 1 },
    splitLine: { ...options?.splitLine, visible: options?.splitLine?.visible ?? (classic || legacy),
      length: Math.min(innerRadius, options?.splitLine?.length ?? (legacy ? radius * 0.09 : classic ? 12 : 8)), width: options?.splitLine?.width ?? (legacy ? 1 : 2) },
    axisLabel: { ...options?.axisLabel, ...gaugeFont(font, options?.axisLabel), visible: options?.axisLabel?.visible ?? (classic || legacy),
      distance: options?.axisLabel?.distance ?? (legacy ? radius * 0.18 : 12), fontSize: options?.axisLabel?.fontSize ?? font.fontSize ?? (legacy ? 8 : 12) },
    pointer: { ...options?.pointer, visible: options?.pointer?.visible ?? mode !== 'progress', length: options?.pointer?.length ?? (legacy ? 0.7 : 0.62),
      width: pointerWidth, shape: options?.pointer?.shape ?? 'needle' },
    anchor: { ...options?.anchor, visible: options?.anchor?.visible ?? (classic || mode === 'pointer'),
      size: options?.anchor?.size ?? (legacy ? pointerWidth * 2.4 : 8), borderWidth: options?.anchor?.borderWidth ?? (legacy ? 0 : 2) },
    progress: { ...options?.progress, visible: options?.progress?.visible ?? mode !== 'pointer',
      width: Math.min(radius, options?.progress?.width ?? width), rounded: options?.progress?.rounded ?? mode === 'modern' },
    title: { ...labels, ...options?.title, ...gaugeFont(font, labels, options?.title), visible: options?.title?.visible ?? labels.visible ?? true,
      offsetY: options?.title?.offsetY ?? (legacy ? 0.56 : 0.95), fontSize: options?.title?.fontSize ?? labels.fontSize ?? font.fontSize ?? (legacy ? 10 : 12) },
    detail: { ...labels, ...options?.detail, ...gaugeFont({ bold: true }, font, labels, options?.detail), visible: options?.detail?.visible ?? labels.visible ?? true,
      offsetY: options?.detail?.offsetY ?? (legacy ? 0.3 : 0.65), fontSize: options?.detail?.fontSize ?? labels.fontSize ?? font.fontSize ?? (legacy ? 13 : 14),
      format: options?.detail?.format ?? labels.format },
  }
}

/** Null/undefined field resets restore inherited chart fonts rather than hiding them. */
function gaugeFont(...fonts: Array<ChartFontOptions | undefined>): ChartFontOptions {
  const result: ChartFontOptions = {}
  for (const font of fonts) for (const key of ['family', 'fontSize', 'color', 'bold', 'italic'] as const) {
    const value = font?.[key]
    if (value !== undefined) Object.assign(result, { [key]: value })
  }
  return result
}

/** Fixed public presets replace all known Gauge appearance, preserving source roles.
 * Explicit undefined fields clear older author values during the shared option merge. */
export function chartGaugePresetOptions(preset: ChartGaugePreset): ChartGaugeOptions {
  if (!CHART_GAUGE_PRESETS.some(item => item.value === preset)) throw new TypeError('仪表预设包含无效选项。')
  const base = resolveChartGaugeOptions({ mode: preset === 'classic' ? 'classic' : 'modern' }, 100)
  const font = { family: undefined, color: undefined, bold: false, italic: false }
  const options = {
    mode: base.mode, min: base.min, max: base.max, startAngle: base.startAngle, endAngle: base.endAngle, splitNumber: base.splitNumber,
    ranges: undefined, label: undefined,
    axisLine: { color: undefined, ...base.axisLine }, axisTick: { color: undefined, ...base.axisTick }, splitLine: { color: undefined, ...base.splitLine },
    axisLabel: { ...font, format: undefined, ...base.axisLabel }, pointer: { color: undefined, ...base.pointer },
    anchor: { color: undefined, borderColor: undefined, ...base.anchor }, progress: { color: undefined, ...base.progress },
    title: { ...font, ...base.title }, detail: { ...font, format: undefined, ...base.detail },
  } satisfies ChartGaugeOptions
  if (preset === 'statusDial') {
    options.splitNumber = 2; options.axisLine.width = 10
    options.splitLine.visible = true; options.axisLabel.visible = true; options.anchor.visible = true
  } else if (preset === 'progress') {
    options.pointer.visible = false; options.progress.width = 14
  } else if (preset === 'minimal') options.progress.visible = false
  return options
}

/** Canvas uses clockwise screen-space radians; authored angles use ECharts' degrees.
 * A forward sweep of at least one revolution is full. Reverse sweeps wrap clockwise;
 * exact reverse revolutions are full, matching zrender's normalizeArcAngles contract. */
export function chartGaugeSweep(options: ChartGaugeAnglesOptions): { start: number; end: number } {
  const startAngle = options.startAngle ?? 225, endAngle = options.endAngle ?? -45, difference = startAngle - endAngle
  const degrees = difference >= 360 ? 360 : difference < 0 ? 360 - (-difference % 360) : difference
  const start = -startAngle * Math.PI / 180
  return { start, end: start + degrees * Math.PI / 180 }
}

export function chartGaugeRatio(value: number, min: number, max: number): number { return Math.max(0, Math.min(1, chartLinearRatio(value, min, max))) }

/** Axis-label distance is measured from the end of the major tick in the native Gauge.
 * Labels whose authored distance would cross the center stop at the center in both adapters. */
export function chartGaugeAxisLabelRadius(options: ChartResolvedGaugeOptions, radius: number): number {
  return Math.max(0, radius - options.splitLine.length - options.axisLabel.distance)
}

/** Native GaugeView chooses unrotated label alignment from the radial direction. */
export function chartGaugeAxisLabelAlignment(angle: number): { align: 'left' | 'right' | 'center'; baseline: 'top' | 'bottom' | 'middle' } {
  const x = Math.cos(angle), y = Math.sin(angle)
  return { align: x < -0.4 ? 'left' : x > 0.4 ? 'right' : 'center', baseline: y < -0.8 ? 'top' : y > 0.8 ? 'bottom' : 'middle' }
}

/** Fill uncovered scale ranges with the authored axis color in both adapters. */
export function chartGaugeColorStops(options: ChartGaugeOptions | undefined, fallback: string): Array<[number, string]> {
  const min = options?.min ?? 0, max = options?.max ?? 100, stops: Array<[number, string]> = []
  let cursor = min
  for (const range of options?.ranges ?? []) {
    if (range.from > cursor) stops.push([chartGaugeRatio(range.from, min, max), fallback])
    stops.push([chartGaugeRatio(range.to, min, max), range.color]); cursor = range.to
  }
  if (cursor < max || !stops.length) stops.push([1, fallback])
  return stops
}

function pointerProfile(shape: ChartGaugePointerShape): ChartVector[] {
  if (shape === 'slim') return [{ x: 0, y: 0 }, { x: -0.5, y: -0.12 }, { x: 0, y: -1 }, { x: 0.5, y: -0.12 }]
  if (shape === 'arrow') return [{ x: -0.18, y: 0 }, { x: -0.18, y: -0.72 }, { x: -0.5, y: -0.72 },
    { x: 0, y: -1 }, { x: 0.5, y: -0.72 }, { x: 0.18, y: -0.72 }, { x: 0.18, y: 0 }]
  return [{ x: -0.5, y: 0 }, { x: 0, y: -1 }, { x: 0.5, y: 0 }]
}

export function chartGaugePointerPoints(center: ChartVector, length: number, width: number, angle: number, shape: ChartGaugePointerShape): ChartVector[] {
  const dx = Math.cos(angle), dy = Math.sin(angle)
  return pointerProfile(shape).map(point => ({ x: center.x - point.y * length * dx - point.x * width * dy,
    y: center.y - point.y * length * dy + point.x * width * dx }))
}

/** ECharts scales this same profile to its pointer width/length without keeping aspect ratio. */
export function chartGaugePointerIcon(shape: ChartGaugePointerShape): string {
  return `path://${pointerProfile(shape).map((point, index) => `${index ? 'L' : 'M'}${point.x * 100},${point.y * 100}`).join(' ')} Z`
}

/** Indicators use clamped ratios, while value text always receives the original observation. */
export function chartGaugeDetailText(value: number, category: string, series: string, label: ChartLabelOptions, sourceValue?: CellScalar): string {
  if (label.format === undefined && (label.contentType === undefined || label.contentType === 'value')
    && typeof sourceValue === 'string' && sourceValue.trim() && Number(sourceValue) === value) return sourceValue
  return chartLabelText(label, value, category, series)
}
