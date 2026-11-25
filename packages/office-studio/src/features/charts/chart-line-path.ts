import type { ChartAreaLineStyle, ChartInvalidValueStrategy } from '../../domain/model'
import { chartAreaLineStyleValue, chartInvalidValueStrategyValue } from './chart-line-and-area'

export interface ChartLinePathPoint {
  x: number
  y: number
  baseY?: number
  categoryIndex: number
  sequenceIndex?: number
}
export interface ChartLinePathSample { x: number; y: number; baseY: number }
export const CHART_LINE_CURVE_STEPS = 12

export function chartLinePointsConnected(previous: ChartLinePathPoint, point: ChartLinePathPoint, strategy: ChartInvalidValueStrategy = 'break'): boolean {
  return strategy === 'link' || (point.sequenceIndex ?? point.categoryIndex) === (previous.sequenceIndex ?? previous.categoryIndex) + 1
}

/** Pure category/value geometry, reused before 3D projection. Both boundaries
 * share the step position and curve samples; no intermediate datum is authored. */
export function chartLineSegment(points: readonly ChartLinePathPoint[], index: number, style: ChartAreaLineStyle, strategy: ChartInvalidValueStrategy = 'break'): ChartLinePathSample[] {
  chartAreaLineStyleValue(style); chartInvalidValueStrategyValue(strategy)
  if (!Number.isSafeInteger(index) || index < 1 || index >= points.length) throw new RangeError('线段终点须属于当前图表系列。')
  const a = points[index - 1]!, b = points[index]!
  if (!chartLinePointsConnected(a, b, strategy)) return []
  const start = sample(a), end = sample(b)
  if (style === 'line') return [start, end]
  if (style === 'step') {
    const middle = a.x / 2 + b.x / 2
    return [start, { x: middle, y: start.y, baseY: start.baseY }, { x: middle, y: end.y, baseY: end.baseY }, end]
  }
  // Repeated dates/categories have no horizontal interval to smooth.
  const distance = b.x - a.x
  if (!Number.isFinite(distance) || distance === 0) return [start, end]
  const previous = index > 1 && chartLinePointsConnected(points[index - 2]!, a, strategy) ? points[index - 2] : undefined
  const next = index + 1 < points.length && chartLinePointsConnected(b, points[index + 1]!, strategy) ? points[index + 1] : undefined
  const top = tangents(previous, a, b, next, point => point.y)
  const base = tangents(previous, a, b, next, point => point.baseY ?? point.y)
  const result = [start]
  for (let step = 1; step < CHART_LINE_CURVE_STEPS; step++) {
    const t = step / CHART_LINE_CURVE_STEPS
    result.push({ x: a.x * (1 - t) + b.x * t,
      y: monotoneValue(start.y, end.y, top[0], top[1], distance, t),
      baseY: monotoneValue(start.baseY, end.baseY, base[0], base[1], distance, t) })
  }
  result.push(end)
  return result
}

function sample(point: ChartLinePathPoint): ChartLinePathSample { return { x: point.x, y: point.y, baseY: point.baseY ?? point.y } }

function tangents(previous: ChartLinePathPoint | undefined, a: ChartLinePathPoint, b: ChartLinePathPoint, next: ChartLinePathPoint | undefined,
  value: (point: ChartLinePathPoint) => number): [number, number] {
  const slope = (value(b) - value(a)) / (b.x - a.x)
  if (!Number.isFinite(slope)) return [0, 0]
  return [previous ? tangent(previous, a, b, value) : slope, next ? tangent(a, b, next, value) : slope]
}

/** Weighted monotone cubic tangents support irregular/reversed dates and do
 * not use a point across a Break/Zero gap as a smoothing neighbour. */
function tangent(a: ChartLinePathPoint, b: ChartLinePathPoint, c: ChartLinePathPoint, value: (point: ChartLinePathPoint) => number): number {
  const first = b.x - a.x, second = c.x - b.x
  if (first === 0 || second === 0 || Math.sign(first) !== Math.sign(second)) return 0
  const before = (value(b) - value(a)) / first, after = (value(c) - value(b)) / second
  if (!Number.isFinite(before) || !Number.isFinite(after) || before === 0 || after === 0 || Math.sign(before) !== Math.sign(after)) return 0
  const h0 = Math.abs(first), h1 = Math.abs(second), w0 = 2 * h1 + h0, w1 = h1 + 2 * h0
  const result = (w0 + w1) / (w0 / before + w1 / after)
  return Number.isFinite(result) ? result : 0
}

function monotoneValue(a: number, b: number, first: number, second: number, distance: number, t: number): number {
  const square = t * t, cube = square * t
  const value = (2 * cube - 3 * square + 1) * a + (cube - 2 * square + t) * distance * first
    + (-2 * cube + 3 * square) * b + (cube - square) * distance * second
  return Math.max(Math.min(a, b), Math.min(Math.max(a, b), value))
}
