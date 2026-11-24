import type { DrawCommand, RenderRect } from '../../rendering/contracts'
import { chartGradientColorAt, type ChartAutomaticGradient, type ChartGradientPoint } from './chart-auto-gradient'

type Point = ChartGradientPoint
export const CHART_GRADIENT_MAX_STRIPS = 48
/** Gradient appends include the current target length in their budget. Borders,
 * other paint and separately generated scenes need their own whole-scene budget. */
export const CHART_GRADIENT_COMMAND_BUDGET = 262_144

function gradientBounds(points: Iterable<Point>): RenderRect | undefined {
  let left = Infinity, top = Infinity, right = -Infinity, bottom = -Infinity
  for (const point of points) {
    if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) continue
    left = Math.min(left, point.x); top = Math.min(top, point.y)
    right = Math.max(right, point.x); bottom = Math.max(bottom, point.y)
  }
  return Number.isFinite(left) ? { x: left, y: top, width: right - left, height: bottom - top } : undefined
}

function gradientSpace(gradient: ChartAutomaticGradient, fallback: RenderRect) {
  const bounds = gradient.bounds ?? fallback
  const start = { x: bounds.x + gradient.start.x * bounds.width, y: bounds.y + gradient.start.y * bounds.height }
  const vector = { x: (gradient.end.x - gradient.start.x) * bounds.width, y: (gradient.end.y - gradient.start.y) * bounds.height }
  const length = Math.hypot(vector.x, vector.y)
  if (![start.x, start.y, vector.x, vector.y, length].every(Number.isFinite)) throw new RangeError('图表渐变坐标须为有限值。')
  const count = length < 1e-6 ? 1 : Math.max(2, Math.min(CHART_GRADIENT_MAX_STRIPS, Math.ceil(length / 4)))
  const cuts = length < 1e-6 ? [0, 1] : [...new Set([...Array.from({ length: count + 1 }, (_unused, index) => index / count), ...gradient.stops.map(stop => stop.offset)])].sort((a, b) => a - b)
  return { cuts,
    coordinate: length < 1e-6 ? (_point: Point) => 0.5
      : (point: Point) => ((point.x - start.x) * (vector.x / length) + (point.y - start.y) * (vector.y / length)) / length }
}

function gradientBandIndex(value: number, cuts: readonly number[]): number {
  if (!Number.isFinite(value)) throw new RangeError('图表渐变采样坐标须为有限值。')
  if (value < cuts[0]!) return -1
  let first = 0, last = cuts.length - 1
  while (first < last) { const middle = Math.ceil((first + last) / 2); if (cuts[middle]! <= value) first = middle; else last = middle - 1 }
  return first
}

function appendWithinBudget(commands: DrawCommand[], generated: DrawCommand[], command: DrawCommand): void {
  if (commands.length + generated.length >= CHART_GRADIENT_COMMAND_BUDGET) throw new RangeError('图表渐变几何超出绘制指令预算。')
  generated.push(command)
}
function commitGradientCommands(commands: DrawCommand[], generated: DrawCommand[]): void {
  // Avoid spreading a large command array into a function argument list.
  for (const command of generated) commands.push(command)
}

/** Up to 48 base color bands, additional authored stop boundaries and two
 * clamped endpoint bands per triangle. Narrow/hard stops retain their boundary.
 * A failed budget leaves the caller's command array unchanged for this append. */
export function appendChartGradient(commands: DrawCommand[], triangles: readonly (readonly Point[])[], gradient: ChartAutomaticGradient, clip?: RenderRect): void {
  if (!triangles.length) return
  const valid = triangles.filter(triangle => triangle.length === 3 && triangle.every(point => Number.isFinite(point.x) && Number.isFinite(point.y)))
  const points = function* (): Generator<Point> { for (const triangle of valid) for (const point of triangle) yield point }
  const bounds = gradientBounds(points())
  if (!bounds) return
  const { coordinate, cuts } = gradientSpace(gradient, bounds), generated: DrawCommand[] = []
  for (const triangle of valid) {
    let minimum = Infinity, maximum = -Infinity
    for (const point of triangle) { const value = coordinate(point); minimum = Math.min(minimum, value); maximum = Math.max(maximum, value) }
    const first = gradientBandIndex(minimum, cuts), last = gradientBandIndex(maximum, cuts)
    for (let strip = first; strip <= last; strip++) {
      const lower = strip < 0 ? -Infinity : cuts[strip]!, upper = strip >= cuts.length - 1 ? Infinity : cuts[strip + 1]!
      const polygon = clipGradientHalfPlane(clipGradientHalfPlane(triangle, coordinate, lower, true), coordinate, upper, false)
      const fill = chartGradientColorAt(gradient, strip < 0 ? 0 : strip >= cuts.length - 1 ? 1 : (lower + upper) / 2)
      for (let index = 1; index + 1 < polygon.length; index++) {
        const points = [polygon[0]!, polygon[index]!, polygon[index + 1]!]
        if (Math.abs((points[1]!.x - points[0]!.x) * (points[2]!.y - points[0]!.y) - (points[1]!.y - points[0]!.y) * (points[2]!.x - points[0]!.x)) < 1e-10) continue
        appendWithinBudget(commands, generated, { kind: 'polygon', points, fill, ...(clip ? { clip } : {}) })
      }
    }
  }
  commitGradientCommands(commands, generated)
}

/** Solid/dashed line fragments sample one shared gradient space. */
export function appendChartGradientLine(commands: DrawCommand[], start: Point, end: Point, width: number, gradient: ChartAutomaticGradient, clip?: RenderRect): void {
  const bounds = gradientBounds([start, end])
  if (!bounds || width <= 0 || start.x === end.x && start.y === end.y) return
  const { coordinate, cuts: boundaries } = gradientSpace(gradient, bounds)
  const first = coordinate(start), delta = coordinate(end) - first, cuts = [0, 1], generated: DrawCommand[] = []
  if (!Number.isFinite(first) || !Number.isFinite(delta)) throw new RangeError('图表渐变采样坐标须为有限值。')
  if (Math.abs(delta) > 1e-12) for (const boundary of boundaries) {
    const ratio = (boundary - first) / delta
    if (ratio > 0 && ratio < 1) cuts.push(ratio)
  }
  cuts.sort((a, b) => a - b)
  for (let index = 1; index < cuts.length; index++) {
    const a = cuts[index - 1]!, b = cuts[index]!
    appendWithinBudget(commands, generated, { kind: 'line', x1: start.x + (end.x - start.x) * a, y1: start.y + (end.y - start.y) * a,
      x2: start.x + (end.x - start.x) * b, y2: start.y + (end.y - start.y) * b, width,
      color: chartGradientColorAt(gradient, first + delta * (a + b) / 2), ...(clip ? { clip } : {}) })
  }
  commitGradientCommands(commands, generated)
}

function clipGradientHalfPlane(points: readonly Point[], coordinate: (point: Point) => number, boundary: number, greater: boolean): Point[] {
  if (!Number.isFinite(boundary) || !points.length) return [...points]
  const result: Point[] = []
  let previous = points.at(-1)!, previousValue = coordinate(previous), previousInside = greater ? previousValue >= boundary : previousValue <= boundary
  for (const point of points) {
    const value = coordinate(point), inside = greater ? value >= boundary : value <= boundary
    if (inside !== previousInside) {
      const ratio = (boundary - previousValue) / (value - previousValue)
      result.push({ x: previous.x + (point.x - previous.x) * ratio, y: previous.y + (point.y - previous.y) * ratio })
    }
    if (inside) result.push({ ...point })
    previous = point; previousValue = value; previousInside = inside
  }
  return result
}
