import type { ChartPointShape } from '../../domain/model'

/** Fixed v1.6.0 LinePointShape authoring vocabulary. Scalar sizes describe the
 * nominal box; dot/dash keep R6's inner rectangles within that box. */
export const CHART_POINT_SHAPE_CHOICES: ReadonlyArray<readonly [ChartPointShape, string]> = [
  ['dot', '点'], ['dash', '短划线'], ['circle', '圆形'], ['square', '正方形'],
  ['triangle', '三角形'], ['diamond', '菱形'], ['rect', '方形'],
  ['roundRect', '圆角矩形'], ['pin', '图钉'], ['arrow', '箭头'], ['none', '无标记'],
  ['emptyCircle', '空心圆形'], ['emptySquare', '空心正方形'], ['emptyTriangle', '空心三角形'],
  ['emptyDiamond', '空心菱形'], ['emptyRect', '空心方形'], ['emptyRoundRect', '空心圆角矩形'],
  ['emptyPin', '空心图钉'], ['emptyArrow', '空心箭头'],
]
export const CHART_POINT_SHAPES: readonly ChartPointShape[] = CHART_POINT_SHAPE_CHOICES.map(([shape]) => shape)
export interface ChartMarkerVector { x: number; y: number }
type ChartMarkerBase = 'dot' | 'circle' | 'rect' | 'triangle' | 'diamond' | 'dash' | 'roundRect' | 'pin' | 'arrow' | 'none'
export interface ChartMarkerShape { base: ChartMarkerBase; hollow: boolean }

export function isChartPointShape(shape: unknown): shape is ChartPointShape {
  return typeof shape === 'string' && CHART_POINT_SHAPES.includes(shape as ChartPointShape)
}

export function chartMarkerShape(shape: ChartPointShape): ChartMarkerShape {
  if (!isChartPointShape(shape)) throw new TypeError('图表标记包含无效形状。')
  const hollow = shape.startsWith('empty')
  const solid = hollow ? shape.slice(5, 6).toLowerCase() + shape.slice(6) : shape
  const base = solid === 'square' ? 'rect' : solid
  return { base: base as ChartMarkerBase, hollow }
}

/** New complex and hollow shapes use the same submitted scene in both chart
 * adapters. Native aliases are normalized before ECharts receives them. */
export function chartMarkerRequiresSharedGeometry(shape: ChartPointShape | undefined): boolean {
  if (shape === undefined) return false
  const { base, hollow } = chartMarkerShape(shape)
  return hollow || base === 'dot' || base === 'dash' || base === 'roundRect' || base === 'pin' || base === 'arrow'
}
export function chartNativeMarkerSymbol(shape: ChartPointShape | undefined, fallback: ChartPointShape = 'circle'): string {
  const current = shape ?? fallback, { base, hollow } = chartMarkerShape(current)
  // The zero-length anchors retain the fixed nominal 0..10 bounding box.
  if (base === 'dot') return 'path://M0 0L0 0M10 10L10 10M3 4H7V6H3Z'
  if (base === 'dash') return 'path://M0 0L0 0M10 10L10 10M1 4H9V6H1Z'
  return hollow ? 'empty' + base.slice(0, 1).toUpperCase() + base.slice(1) : base
}

/** Pure bounded geometry. It never paints a filled center for hollow marks;
 * the caller decides fill/outline color and alpha independently. */
export function chartMarkerPoints(center: ChartMarkerVector, size: number, shape: ChartPointShape): ChartMarkerVector[] {
  if (!Number.isFinite(size) || size < 0 || !Number.isFinite(center.x) || !Number.isFinite(center.y)) throw new RangeError('图表标记的位置须有限，大小须为有限非负数值。')
  const { base } = chartMarkerShape(shape)
  if (base === 'none' || size === 0) return []
  const radius = size / 2, points: ChartMarkerVector[] = []
  const add = (x: number, y: number): void => {
    const absolute = { x: center.x + radius * x, y: center.y + radius * y }
    if (!Number.isFinite(absolute.x) || !Number.isFinite(absolute.y)) throw new RangeError('图表标记的投影坐标溢出。')
    points.push(absolute)
  }
  if (base === 'circle') {
    for (let index = 0; index < 36; index++) add(Math.cos(index * Math.PI / 18), Math.sin(index * Math.PI / 18))
  } else if (base === 'rect' || base === 'dot' || base === 'dash') {
    const halfWidth = base === 'dot' ? 0.4 : base === 'dash' ? 0.8 : 1, halfHeight = base === 'rect' ? 1 : 0.2
    add(-halfWidth, -halfHeight); add(halfWidth, -halfHeight); add(halfWidth, halfHeight); add(-halfWidth, halfHeight)
  } else if (base === 'triangle') {
    add(0, -1); add(1, 1); add(-1, 1)
  } else if (base === 'diamond') {
    add(0, -1); add(1, 0); add(0, 1); add(-1, 0)
  } else if (base === 'roundRect') {
    const rounding = 0.4
    for (const [x, y, start] of [[1 - rounding, 1 - rounding, 0], [-1 + rounding, 1 - rounding, Math.PI / 2],
      [-1 + rounding, -1 + rounding, Math.PI], [1 - rounding, -1 + rounding, Math.PI * 1.5]]) {
      for (let step = 0; step <= 6; step++) {
        const angle = start! + step * Math.PI / 12
        add(x! + rounding * Math.cos(angle), y! + rounding * Math.sin(angle))
      }
    }
  } else if (base === 'pin') {
    add(0, 1)
    for (let index = 0; index <= 24; index++) {
      const angle = Math.PI / 6 - index * Math.PI / 18
      add(0.75 * Math.cos(angle), -0.25 + 0.75 * Math.sin(angle))
    }
  } else {
    for (const [x, y] of [[0, -1], [1, 0.1], [0.35, 0.1], [0.35, 1], [-0.35, 1], [-0.35, 0.1], [-1, 0.1]]) add(x!, y!)
  }
  return points
}
