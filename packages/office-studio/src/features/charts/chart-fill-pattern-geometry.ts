import type { ChartFillPatternPreset, ChartSeriesFillPattern } from '../../domain/model'
import type { DrawCommand, RenderRect } from '../../rendering/contracts'
import { isChartFillPatternPreset } from './chart-fill-pattern'

interface PatternPoint { x: number; y: number }
export interface ChartFillPatternTile { width: number; height: number; rectangles: readonly RenderRect[]; backgroundRectangles: readonly RenderRect[] }
const TILES = new Map<ChartFillPatternPreset, ChartFillPatternTile>()
const mod = (value: number, divisor: number): number => (value % divisor + divisor) % divisor
const CONFETTI = [[1, 1], [8, 2], [13, 5], [4, 7], [10, 10], [2, 13], [14, 14]] as const
const BAYER = [
  [0, 48, 12, 60, 3, 51, 15, 63], [32, 16, 44, 28, 35, 19, 47, 31], [8, 56, 4, 52, 11, 59, 7, 55], [40, 24, 36, 20, 43, 27, 39, 23],
  [2, 50, 14, 62, 1, 49, 13, 61], [34, 18, 46, 30, 33, 17, 45, 29], [10, 58, 6, 54, 9, 57, 5, 53], [42, 26, 38, 22, 41, 25, 37, 21],
] as const

/** Independent Office-style masks. Their precise pixels are not claimed to reproduce the closed renderer. */
function patternPixel(preset: ChartFillPatternPreset, x: number, y: number): boolean {
  if (preset.startsWith('pct')) return BAYER[y % 8]![x % 8]! < Math.round(Number(preset.slice(3)) * 64 / 100)
  const down = mod(x - y, 8), up = mod(x + y, 8), diamond = Math.abs(x % 8 - 4) + Math.abs(y % 8 - 4)
  switch (preset) {
    case 'horz': return y % 4 === 0
    case 'vert': return x % 4 === 0
    case 'ltHorz': return y % 8 === 0
    case 'ltVert': return x % 8 === 0
    case 'dkHorz': return y % 8 < 4
    case 'dkVert': return x % 8 < 4
    case 'narHorz': return y % 2 === 0
    case 'narVert': return x % 2 === 0
    case 'dashHorz': return y % 4 === 0 && x % 8 < 4
    case 'dashVert': return x % 4 === 0 && y % 8 < 4
    case 'cross': return x % 8 === 0 || y % 8 === 0
    case 'dnDiag': return down === 0
    case 'upDiag': return up === 0
    case 'ltDnDiag': return mod(x - y, 16) === 0
    case 'ltUpDiag': return mod(x + y, 16) === 0
    case 'dkDnDiag': return down < 4
    case 'dkUpDiag': return up < 4
    case 'wdDnDiag': return mod(x - y, 16) < 8
    case 'wdUpDiag': return mod(x + y, 16) < 8
    case 'dashDnDiag': return down === 0 && x % 8 < 4
    case 'dashUpDiag': return up === 0 && x % 8 < 4
    case 'diagCross': return down === 0 || up === 0
    case 'smCheck': return (Math.floor(x / 2) + Math.floor(y / 2)) % 2 === 0
    case 'lgCheck': return (Math.floor(x / 4) + Math.floor(y / 4)) % 2 === 0
    case 'smGrid': return x % 4 === 0 || y % 4 === 0
    case 'lgGrid': return x % 16 === 0 || y % 16 === 0
    case 'dotGrid': return x % 4 === 0 && y % 4 === 0
    case 'smConfetti': return CONFETTI.some(([a, b]) => x === a && y === b)
    case 'lgConfetti': return CONFETTI.some(([a, b]) => x >= a && x < a + 2 && y >= b && y < b + 2)
    case 'horzBrick': return y % 4 === 0 || mod(x + Math.floor(y / 4) * 4, 8) === 0
    case 'diagBrick': return down === 0 || mod(x + y + Math.floor((x - y + 16) / 8) * 4, 16) === 0
    case 'solidDmnd': return diamond <= 3
    case 'openDmnd': return diamond === 4
    case 'dotDmnd': return diamond === 4 && x % 2 === 0
    case 'plaid': return x % 8 < 2 || y % 8 < 2 || x % 8 === 4 && y % 2 === 0 || y % 8 === 4 && x % 2 === 0
    case 'sphere': {
      const a = x % 8 - 3.5, b = y % 8 - 3.5, radius = a * a + b * b
      return radius <= 12 && (radius >= 7 || a + b > -1)
    }
    case 'weave': return down < 2 && Math.floor(x / 4) % 2 === 0 || up < 2 && Math.floor(y / 4) % 2 === 1
    case 'divot': return y % 8 >= 2 && y % 8 <= 4 && Math.abs(x % 8 - 3) === 4 - y % 8
    case 'shingle': return y % 8 === 0 || y % 8 === mod(x + Math.floor(y / 8) * 4, 8)
    case 'wave': return y % 8 === Math.round(3 + 2 * Math.sin(x * Math.PI / 4))
    case 'trellis': return down < 2 || up < 2
    case 'zigZag': return y % 8 === 1 + Math.abs(x % 8 - 4)
    default: throw new TypeError('无效的 DrawingML 填充纹理。')
  }
}

/** Merge adjacent row runs into disjoint rectangles, so translucent foreground pixels never overlap. */
export function chartFillPatternTile(preset: ChartFillPatternPreset): ChartFillPatternTile {
  if (!isChartFillPatternPreset(preset)) throw new TypeError('无效的 DrawingML 填充纹理。')
  const cached = TILES.get(preset)
  if (cached) return cached
  const width = 16, height = 16
  const runs = (foreground: boolean): readonly RenderRect[] => {
    const rectangles: RenderRect[] = []
    let previous = new Map<string, RenderRect>()
    for (let y = 0; y < height; y++) {
      const current = new Map<string, RenderRect>()
      for (let x = 0; x < width;) {
        if (patternPixel(preset, x, y) !== foreground) { x++; continue }
        const start = x
        while (x < width && patternPixel(preset, x, y) === foreground) x++
        const key = `${start}:${x - start}`, existing = previous.get(key)
        const rectangle = existing ?? { x: start, y, width: x - start, height: 0 }
        rectangle.height++
        if (!existing) rectangles.push(rectangle)
        current.set(key, rectangle)
      }
      previous = current
    }
    return Object.freeze(rectangles.map(rect => Object.freeze(rect)))
  }
  const tile: ChartFillPatternTile = Object.freeze({ width, height, rectangles: runs(true), backgroundRectangles: runs(false) })
  TILES.set(preset, tile)
  return tile
}

function clipPatternRectangle(rectangle: RenderRect, triangle: readonly PatternPoint[]): PatternPoint[] {
  const cross = (a: PatternPoint, b: PatternPoint, c: PatternPoint): number => (b.x - a.x) * (c.y - a.y) - (b.y - a.y) * (c.x - a.x)
  const sign = cross(triangle[0]!, triangle[1]!, triangle[2]!) < 0 ? -1 : 1
  let polygon: PatternPoint[] = [{ x: rectangle.x, y: rectangle.y }, { x: rectangle.x + rectangle.width, y: rectangle.y },
    { x: rectangle.x + rectangle.width, y: rectangle.y + rectangle.height }, { x: rectangle.x, y: rectangle.y + rectangle.height }]
  for (let index = 0; index < 3 && polygon.length; index++) {
    const a = triangle[index]!, b = triangle[(index + 1) % 3]!, output: PatternPoint[] = []
    let last = polygon.at(-1)!, lastDistance = cross(a, b, last) * sign
    for (const point of polygon) {
      const distance = cross(a, b, point) * sign, inside = distance >= -1e-8, lastInside = lastDistance >= -1e-8
      if (inside !== lastInside) {
        const ratio = lastDistance / (lastDistance - distance)
        output.push({ x: last.x + (point.x - last.x) * ratio, y: last.y + (point.y - last.y) * ratio })
      }
      if (inside) output.push(point)
      last = point; lastDistance = distance
    }
    polygon = output
  }
  return polygon
}

/** Clip periodic foreground geometry to the already-triangulated mark. No new renderer protocol or DOM dependency. */
export function appendChartFillPattern(commands: DrawCommand[], triangles: readonly (readonly PatternPoint[])[], paint: ChartSeriesFillPattern, clip?: RenderRect): void {
  const tile = chartFillPatternTile(paint.preset), origin = { x: clip?.x ?? 0, y: clip?.y ?? 0 }
  for (const triangle of triangles) {
    if (triangle.length !== 3 || triangle.some(point => !Number.isFinite(point.x) || !Number.isFinite(point.y))) continue
    const left = Math.max(Math.min(...triangle.map(point => point.x)), clip?.x ?? -Infinity)
    const top = Math.max(Math.min(...triangle.map(point => point.y)), clip?.y ?? -Infinity)
    const right = Math.min(Math.max(...triangle.map(point => point.x)), clip ? clip.x + clip.width : Infinity)
    const bottom = Math.min(Math.max(...triangle.map(point => point.y)), clip ? clip.y + clip.height : Infinity)
    if (right <= left || bottom <= top) continue
    const firstX = Math.floor((left - origin.x) / tile.width), lastX = Math.ceil((right - origin.x) / tile.width)
    const firstY = Math.floor((top - origin.y) / tile.height), lastY = Math.ceil((bottom - origin.y) / tile.height)
    // Both colors occupy disjoint pixels. Painting a translucent foreground over a full background would apply opacity twice.
    for (const [rectangles, fill] of [[tile.rectangles, paint.foregroundColor], [tile.backgroundRectangles, paint.backgroundColor]] as const) {
      for (let y = firstY; y < lastY; y++) for (let x = firstX; x < lastX; x++) for (const rectangle of rectangles) {
        const rect = { x: origin.x + x * tile.width + rectangle.x, y: origin.y + y * tile.height + rectangle.y, width: rectangle.width, height: rectangle.height }
        if (rect.x >= right || rect.y >= bottom || rect.x + rect.width <= left || rect.y + rect.height <= top) continue
        const points = clipPatternRectangle(rect, triangle)
        const area = points.reduce((sum, point, index) => { const next = points[(index + 1) % points.length]!; return sum + point.x * next.y - next.x * point.y }, 0)
        if (points.length >= 3 && Math.abs(area) > 1e-8) commands.push({ kind: 'polygon', points, fill, ...(clip ? { clip } : {}) })
      }
    }
  }
}
