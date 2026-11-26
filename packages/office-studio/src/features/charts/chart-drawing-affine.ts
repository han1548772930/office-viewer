import { drawingAffineBounds, drawingAffineFromFrame, drawingAffinePoint, invertDrawingAffine, normalizeDrawingAffineGeometry,
  type DrawingAffineGeometry, type DrawingAffineMatrix } from '../../domain/drawing-affine.ts'
import { planDrawingMove, sheetDrawingMoveDirectionForKey } from '../drawing-groups/drawing-group-move.ts'
import { SheetDrawingError } from '../drawing-groups/drawing-group-types.ts'
import type { DrawCommand, LineCommand, RenderRect } from '../../rendering/contracts.ts'
import { GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH } from '../../rendering/spreadsheet/layout.ts'
import type { ChartVector } from './chart-geometry.ts'

export interface ChartDrawingProjection {
  position: ChartVector
  size: { width: number; height: number }
  drawingGeometry?: DrawingAffineGeometry
}
export interface ChartDrawingViewport { zoom: number; scrollLeft: number; scrollTop: number; headerWidth?: number; headerHeight?: number }

/** Host footprints follow the full world geometry; renderer size remains the authored local size. */
export function chartDrawingWorldBounds(chart: ChartDrawingProjection): { left: number; top: number; width: number; height: number } {
  if (!chart.drawingGeometry) return { left: chart.position.x, top: chart.position.y, ...chart.size }
  const bounds = drawingAffineBounds(normalizeDrawingAffineGeometry(chart.drawingGeometry))
  return { ...bounds, left: GRID_HEADER_WIDTH + bounds.left, top: GRID_HEADER_HEIGHT + bounds.top }
}

/** Renderer coordinates remain local even when a drawing group scales or shears the chart. */
export function chartDrawingWorldPoint(chart: ChartDrawingProjection, point: ChartVector): ChartVector {
  const geometry = chart.drawingGeometry
  if (!geometry) return { x: chart.position.x + point.x, y: chart.position.y + point.y }
  const world = drawingAffinePoint(geometry.matrix, point.x * geometry.width / chart.size.width, point.y * geometry.height / chart.size.height)
  return { x: GRID_HEADER_WIDTH + world.x, y: GRID_HEADER_HEIGHT + world.y }
}

export function chartDrawingViewportPoint(chart: ChartDrawingProjection, point: ChartVector, view: ChartDrawingViewport): ChartVector {
  const world = chartDrawingWorldPoint(chart, point), scale = Math.max(0.1, view.zoom / 100)
  return { x: (view.headerWidth ?? GRID_HEADER_WIDTH) + (world.x - GRID_HEADER_WIDTH) * scale - view.scrollLeft,
    y: (view.headerHeight ?? GRID_HEADER_HEIGHT) + (world.y - GRID_HEADER_HEIGHT) * scale - view.scrollTop }
}

export function chartDrawingViewportBounds(chart: ChartDrawingProjection, rect: RenderRect, view: ChartDrawingViewport): RenderRect {
  const points = [[rect.x, rect.y], [rect.x + rect.width, rect.y], [rect.x + rect.width, rect.y + rect.height], [rect.x, rect.y + rect.height]]
    .map(([x, y]) => chartDrawingViewportPoint(chart, { x: x!, y: y! }, view)),
    left = Math.min(...points.map(point => point.x)), top = Math.min(...points.map(point => point.y))
  return { x: left, y: top, width: Math.max(...points.map(point => point.x)) - left, height: Math.max(...points.map(point => point.y)) - top }
}

export function chartDrawingLocalPoint(chart: ChartDrawingProjection, point: ChartVector, view: ChartDrawingViewport): ChartVector {
  const scale = Math.max(0.1, view.zoom / 100), x = (point.x + view.scrollLeft - (view.headerWidth ?? GRID_HEADER_WIDTH)) / scale,
    y = (point.y + view.scrollTop - (view.headerHeight ?? GRID_HEADER_HEIGHT)) / scale, geometry = chart.drawingGeometry
  if (!geometry) return { x: x + GRID_HEADER_WIDTH - chart.position.x, y: y + GRID_HEADER_HEIGHT - chart.position.y }
  const local = drawingAffinePoint(invertDrawingAffine(geometry.matrix), x, y)
  return { x: local.x * chart.size.width / geometry.width, y: local.y * chart.size.height / geometry.height }
}

export function chartDrawingLocalDelta(chart: ChartDrawingProjection, start: ChartVector, end: ChartVector, view: ChartDrawingViewport): ChartVector {
  const first = chartDrawingLocalPoint(chart, start, view), last = chartDrawingLocalPoint(chart, end, view)
  return { x: last.x - first.x, y: last.y - first.y }
}

/** Worksheet-pixel nudges use the complete four-corner bounds, independently of the renderer's local size. */
export function chartDrawingKeyboardTranslation(chart: ChartDrawingProjection, key: string,
  extent: { width: number; height: number }, largeStep = false): DrawingAffineMatrix | undefined {
  const direction = sheetDrawingMoveDirectionForKey(key)
  if (direction === undefined) return undefined
  const geometry = chart.drawingGeometry ?? drawingAffineFromFrame({ left: chart.position.x - GRID_HEADER_WIDTH,
    top: chart.position.y - GRID_HEADER_HEIGHT, ...chart.size }),
    plan = planDrawingMove({ sheetId: 'projection', objects: [{ ref: { kind: 'chart', id: 'projection' }, sheetId: 'projection', geometry }] },
      { selection: [{ kind: 'chart', id: 'projection' }], direction, step: largeStep ? 10 : 1, extent })
  if (plan.status === 'rejected') throw new SheetDrawingError(plan.reason, plan.message, plan.ref)
  const change = plan.status === 'ready' ? plan.transaction.geometryChanges[0] : undefined
  return change ? { a: 1, b: 0, c: 0, d: 1, e: change.after.matrix.e - change.before.matrix.e, f: change.after.matrix.f - change.before.matrix.f } : undefined
}

function clipLine(command: LineCommand, clip?: RenderRect): [ChartVector, ChartVector] | undefined {
  let start = 0, end = 1
  const dx = command.x2 - command.x1, dy = command.y2 - command.y1
  if (clip) {
    for (const [p, q] of [[-dx, command.x1 - clip.x], [dx, clip.x + clip.width - command.x1],
      [-dy, command.y1 - clip.y], [dy, clip.y + clip.height - command.y1]]) {
      if (p === 0) { if (q! < 0) return undefined; continue }
      const ratio = q! / p!
      if (p! < 0) start = Math.max(start, ratio)
      else end = Math.min(end, ratio)
      if (start > end) return undefined
    }
  }
  return [{ x: command.x1 + dx * start, y: command.y1 + dy * start }, { x: command.x1 + dx * end, y: command.y1 + dy * end }]
}

/** Clip in renderer space before projecting; an axis-aligned world clip would leak at rotated edges. */
export function chartDrawingOutlineCommands(chart: ChartDrawingProjection, chartId: string, commands: readonly DrawCommand[]): LineCommand[] {
  return commands.flatMap(command => {
    if (command.kind !== 'line') throw new TypeError('图表作者轮廓必须由线段组成。')
    const segment = clipLine(command, command.clip)
    if (!segment) return []
    const first = chartDrawingWorldPoint(chart, segment[0]), last = chartDrawingWorldPoint(chart, segment[1])
    return [{ kind: 'line' as const, x1: first.x, y1: first.y, x2: last.x, y2: last.y, color: command.color, width: command.width,
      drawingRef: { kind: 'chart' as const, id: chartId }, drawingOverlay: true }]
  })
}
