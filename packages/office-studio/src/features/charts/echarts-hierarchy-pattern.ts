import type { ChartSeriesFillPattern } from '../../domain/model'
import { chartFillPatternTile } from './chart-fill-pattern-geometry'
import { isChartFillPatternPreset } from './chart-fill-pattern'

export const CHART_HIERARCHY_PATTERN_KEY = '__officeHierarchyFillPattern'
interface NativePatternGraphic {
  setStyle?(style: Record<string, unknown>): void
  ensureState?(name: string): { style?: Record<string, unknown> }
  markRedraw?(): void
}
interface NativePatternPaint { image: HTMLCanvasElement | OffscreenCanvas; repeat: 'repeat'; scaleX: number; scaleY: number }
const tiles = new Map<string, NativePatternPaint>()
const MAX_NATIVE_PATTERN_TILES = 128

function isPaint(value: unknown): value is ChartSeriesFillPattern {
  return Boolean(value && typeof value === 'object' && 'preset' in value && isChartFillPatternPreset(value.preset)
    && 'foregroundColor' in value && typeof value.foregroundColor === 'string'
    && 'backgroundColor' in value && typeof value.backgroundColor === 'string')
}

/** Native shapes use the same disjoint Office-style mask as both shared Canvas
 * backends. Tiles are created synchronously in the browser adapter, so image
 * decoding cannot change a retained bitmap behind its sourceVersion. */
function patternTile(paint: ChartSeriesFillPattern, devicePixelRatio: number): NativePatternPaint {
  const scale = Math.max(1, Math.min(4, Number.isFinite(devicePixelRatio) ? devicePixelRatio : 1))
  const key = JSON.stringify([paint, scale]), retained = tiles.get(key)
  if (retained) { tiles.delete(key); tiles.set(key, retained); return retained }
  const tile = chartFillPatternTile(paint.preset)
  let image: HTMLCanvasElement | OffscreenCanvas, context: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D | null
  if (typeof OffscreenCanvas === 'function') {
    image = new OffscreenCanvas(Math.ceil(tile.width * scale), Math.ceil(tile.height * scale)); context = image.getContext('2d')
  } else if (typeof document !== 'undefined') {
    image = document.createElement('canvas'); image.width = Math.ceil(tile.width * scale); image.height = Math.ceil(tile.height * scale); context = image.getContext('2d')
  } else throw new Error('层级图案填充需要浏览器 Canvas 宿主。')
  if (!context) throw new Error('无法创建层级图案填充画布。')
  context.scale(image.width / tile.width, image.height / tile.height)
  for (const [rectangles, fill] of [[tile.rectangles, paint.foregroundColor], [tile.backgroundRectangles, paint.backgroundColor]] as const) {
    context.fillStyle = fill
    for (const rectangle of rectangles) context.fillRect(rectangle.x, rectangle.y, rectangle.width, rectangle.height)
  }
  const result: NativePatternPaint = { image, repeat: 'repeat', scaleX: tile.width / image.width, scaleY: tile.height / image.height }
  tiles.set(key, result)
  if (tiles.size > MAX_NATIVE_PATTERN_TILES) tiles.delete(tiles.keys().next().value!)
  return result
}

/** Applies only explicitly marked author patterns. Opacity is already encoded
 * into the disjoint tile colors; the native outline keeps its own alpha. */
export function applyChartHierarchyPattern(graphic: NativePatternGraphic, value: unknown, devicePixelRatio: number): void {
  if (value === undefined || value === null) return
  if (!isPaint(value)) throw new TypeError('层级图案填充数据无效。')
  const fill = patternTile(value, devicePixelRatio)
  graphic.setStyle?.({ fill, opacity: 1 })
  for (const name of ['emphasis', 'select', 'blur']) {
    const state = graphic.ensureState?.(name)
    if (state) state.style = { ...state.style, fill }
  }
  graphic.markRedraw?.()
}
