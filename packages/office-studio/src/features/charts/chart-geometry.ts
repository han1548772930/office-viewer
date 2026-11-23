import type { ChartAxisOptions, ChartFontOptions, ChartLabelOptions, ChartLineOptions } from '../../domain/model'
import type { DrawCommand, RenderRect } from '../../rendering/contracts'
import { chartColorChannels } from './chart-color'
import { chartInterpolate, chartLinearRatio } from './chart-math'

export interface ChartVector { x: number; y: number }
export interface ChartVector3 extends ChartVector { z: number }

export function chartFont(font: ChartFontOptions = {}, size = 11): string {
  return `${font.italic ? 'italic ' : ''}${font.bold ? '600 ' : ''}${font.fontSize ?? size}px ${font.family ?? 'Aptos, Arial, sans-serif'}`
}

export function chartColorOpacity(color: string, opacity = 1): string {
  if (opacity >= 1) return color
  const [r, g, b, a] = chartColorChannels(color)
  return `rgba(${Math.round(r * 255)}, ${Math.round(g * 255)}, ${Math.round(b * 255)}, ${a * opacity})`
}

export function chartShade(color: string, factor: number): string {
  const [r, g, b, a] = chartColorChannels(color)
  return `rgba(${Math.round(Math.min(1, r * factor) * 255)}, ${Math.round(Math.min(1, g * factor) * 255)}, ${Math.round(Math.min(1, b * factor) * 255)}, ${a})`
}

/** A small, deterministic numeric-format subset; the authored format is retained in the model. */
export function formatChartNumber(value: number, format?: string): string {
  if (!Number.isFinite(value)) return ''
  if (!format || format === 'General') return String(Number(value.toPrecision(5)))
  const section = format.split(';')[value < 0 ? 1 : 0] ?? format.split(';')[0]!
  const decimal = /[0#][.]([0#]+)/.exec(section)?.[1]?.length ?? 0
  const percent = section.includes('%')
  const numeric = (percent ? value * 100 : value).toLocaleString('en-US', {
    useGrouping: /[0#],[0#]/.test(section), minimumFractionDigits: Math.min(20, decimal), maximumFractionDigits: Math.min(20, decimal),
  })
  const currency = /[$€£¥￥]/.exec(section)?.[0] ?? ''
  const prefix = section.startsWith('(') && value < 0 ? '(' : ''
  const suffix = prefix ? ')' : ''
  return `${prefix}${currency}${prefix ? numeric.replace(/^-/, '') : numeric}${percent ? '%' : ''}${suffix}`
}

export function chartLabelText(label: ChartLabelOptions | undefined, value: number, category: string, series: string, percentage?: number): string {
  switch (label?.contentType) {
    case 'category': return category
    case 'series': return series
    case 'percentage': return `${formatChartNumber(percentage ?? 0, label.format)}%`
    default: return formatChartNumber(value, label?.format)
  }
}

export function chartAxisExtent(values: readonly number[], axis: ChartAxisOptions): { min: number; max: number } {
  let min = Infinity, max = -Infinity
  for (const value of values) {
    if (!Number.isFinite(value) || axis.scaleType === 'log' && value <= 0) continue
    min = Math.min(min, value); max = Math.max(max, value)
  }
  if (!Number.isFinite(min)) { min = axis.scaleType === 'log' ? 1 : 0; max = axis.scaleType === 'log' ? axis.logBase ?? 10 : 1 }
  if (axis.scaleType !== 'log') min = Math.min(0, min)
  min = axis.min ?? min; max = axis.max ?? max
  if (min === max) {
    if (axis.min !== undefined && axis.max === undefined) max = axis.scaleType === 'log' ? min * (axis.logBase ?? 10) : min + Math.max(1, Math.abs(min) * 0.1)
    else if (axis.max !== undefined && axis.min === undefined) min = axis.scaleType === 'log' ? max / (axis.logBase ?? 10) : max - Math.max(1, Math.abs(max) * 0.1)
    else max = axis.scaleType === 'log' ? min * (axis.logBase ?? 10) : min + Math.max(1, Math.abs(min) * 0.1)
  }
  return { min, max }
}

export function chartAxisRatio(value: number, min: number, max: number, axis: ChartAxisOptions): number {
  const ratio = axis.scaleType === 'log'
    ? (Math.log(value) - Math.log(min)) / (Math.log(max) - Math.log(min))
    : chartLinearRatio(value, min, max)
  return axis.reverse ? 1 - ratio : ratio
}

export function chartAxisTicks(min: number, max: number, axis: ChartAxisOptions): number[] {
  return Array.from({ length: 5 }, (_unused, index) => axis.scaleType === 'log'
    ? Math.exp(Math.log(min) + (Math.log(max) - Math.log(min)) * index / 4)
    : chartInterpolate(min, max, index / 4))
}

export function chartLine(commands: DrawCommand[], start: ChartVector, end: ChartVector, style: ChartLineOptions = {}, fallback = '#94a3b8', clip?: RenderRect): void {
  if (style.visible === false) return
  const color = chartColorOpacity(style.color ?? fallback, style.opacity ?? 1), width = style.width ?? 1
  if (width <= 0) return
  if (!style.dashType || style.dashType === 'solid') {
    commands.push({ kind: 'line', x1: start.x, y1: start.y, x2: end.x, y2: end.y, color, width, ...(clip ? { clip } : {}) }); return
  }
  const distance = Math.hypot(end.x - start.x, end.y - start.y)
  const length = style.dashType === 'dotted' ? Math.max(1, width) : Math.max(4, width * 4)
  const gap = Math.max(2, width * 2)
  for (let offset = 0; offset < distance; offset += length + gap) {
    const a = offset / distance, b = Math.min(distance, offset + length) / distance
    commands.push({ kind: 'line', x1: start.x + (end.x - start.x) * a, y1: start.y + (end.y - start.y) * a, x2: start.x + (end.x - start.x) * b, y2: start.y + (end.y - start.y) * b, color, width, ...(clip ? { clip } : {}) })
  }
}

export function chartPolygon(commands: DrawCommand[], points: ChartVector[], fill: string, border?: ChartLineOptions, clip?: RenderRect): void {
  if (points.length < 3) return
  // The shared WebGPU driver accepts triangle fans. Emit explicit triangles for
  // concave ribbons and annular sectors so both backends receive the same fill.
  for (const triangle of triangulateChartPolygon(points)) commands.push({ kind: 'polygon', points: triangle, fill, ...(clip ? { clip } : {}) })
  if (border && border.visible !== false) points.forEach((point, index) => chartLine(commands, point, points[(index + 1) % points.length]!, border, fill, clip))
}

export function triangulateChartPolygon(input: readonly ChartVector[]): ChartVector[][] {
  const points = input.filter((point, index) => index === 0 || Math.hypot(point.x - input[index - 1]!.x, point.y - input[index - 1]!.y) > 1e-8)
  if (points.length > 1 && Math.hypot(points[0]!.x - points.at(-1)!.x, points[0]!.y - points.at(-1)!.y) < 1e-8) points.pop()
  if (points.length < 3) return []
  const cross = (a: ChartVector, b: ChartVector, c: ChartVector): number => (b.x - a.x) * (c.y - a.y) - (b.y - a.y) * (c.x - a.x)
  const area = points.reduce((sum, point, index) => { const next = points[(index + 1) % points.length]!; return sum + point.x * next.y - next.x * point.y }, 0)
  const sign = area < 0 ? -1 : 1, indices = points.map((_point, index) => index), triangles: ChartVector[][] = []
  let attempts = 0
  while (indices.length > 3 && attempts < points.length * points.length) {
    let found = false
    for (let index = 0; index < indices.length; index++) {
      attempts++
      const aIndex = indices[(index + indices.length - 1) % indices.length]!, bIndex = indices[index]!, cIndex = indices[(index + 1) % indices.length]!
      const a = points[aIndex]!, b = points[bIndex]!, c = points[cIndex]!, turn = cross(a, b, c) * sign
      if (Math.abs(turn) < 1e-8) { indices.splice(index, 1); found = true; break }
      if (turn < 0) continue
      const occupied = indices.some(candidate => candidate !== aIndex && candidate !== bIndex && candidate !== cIndex
        && cross(a, b, points[candidate]!) * sign > 1e-8 && cross(b, c, points[candidate]!) * sign > 1e-8 && cross(c, a, points[candidate]!) * sign > 1e-8)
      if (occupied) continue
      triangles.push([a, b, c]); indices.splice(index, 1); found = true; break
    }
    if (!found) throw new RangeError('图表多边形无法三角化。')
  }
  if (indices.length === 3) triangles.push(indices.map(index => points[index]!))
  return triangles
}

export function chartCircle(commands: DrawCommand[], center: ChartVector, radius: number, fill: string, border?: ChartLineOptions, clip?: RenderRect): void {
  if (radius <= 0) return
  chartPolygon(commands, Array.from({ length: 36 }, (_unused, index) => chartPolar(center, radius, index * Math.PI / 18)), fill, border, clip)
}

export function chartPolar(center: ChartVector, radius: number, angle: number): ChartVector {
  return { x: center.x + radius * Math.cos(angle), y: center.y + radius * Math.sin(angle) }
}

export function chartArc(center: ChartVector, radius: number, start: number, end: number): ChartVector[] {
  const count = Math.max(2, Math.ceil(Math.abs(end - start) / (Math.PI / 48)))
  return Array.from({ length: count + 1 }, (_unused, index) => chartPolar(center, radius, start + (end - start) * index / count))
}

export function chartSector(center: ChartVector, inner: number, outer: number, start: number, end: number): ChartVector[] {
  return [...chartArc(center, outer, start, end), ...(inner > 0 ? chartArc(center, inner, end, start) : [center])]
}

export function chartBezier(start: ChartVector, first: ChartVector, second: ChartVector, end: ChartVector, count = 32): ChartVector[] {
  return Array.from({ length: count + 1 }, (_unused, index) => {
    const t = index / count, u = 1 - t
    return { x: u ** 3 * start.x + 3 * u * u * t * first.x + 3 * u * t * t * second.x + t ** 3 * end.x, y: u ** 3 * start.y + 3 * u * u * t * first.y + 3 * u * t * t * second.y + t ** 3 * end.y }
  })
}

export function chartSymbol(commands: DrawCommand[], point: ChartVector, size: number, color: string, shape = 'rect', border?: ChartLineOptions, clip?: RenderRect): void {
  if (shape === 'none' || size <= 0) return
  const radius = size / 2
  if (shape === 'circle') return chartCircle(commands, point, radius, color, border, clip)
  if (shape === 'diamond' || shape === 'triangle') {
    const points = shape === 'diamond' ? [{ x: point.x, y: point.y - radius }, { x: point.x + radius, y: point.y }, { x: point.x, y: point.y + radius }, { x: point.x - radius, y: point.y }]
      : [{ x: point.x, y: point.y - radius }, { x: point.x + radius, y: point.y + radius }, { x: point.x - radius, y: point.y + radius }]
    chartPolygon(commands, points, color, border, clip); return
  }
  commands.push({ kind: 'rect', x: point.x - radius, y: point.y - radius, width: size, height: size, fill: color, ...(border?.visible === false ? {} : { stroke: border?.color, strokeWidth: border?.width }), ...(clip ? { clip } : {}) })
}
