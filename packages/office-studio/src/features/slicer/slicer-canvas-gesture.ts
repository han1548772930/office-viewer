import { invertDrawingAffine, multiplyDrawingAffine, normalizeDrawingAffineGeometry, type DrawingAffineGeometry,
  type DrawingAffineMatrix } from '../../domain/drawing-affine.ts'
import type { GridViewport } from '../../rendering/spreadsheet/layout.ts'
import type { SlicerModel } from '../../domain/model.ts'
import type { ITimelineRange } from '../../domain/timeline-filters.ts'
import type { SlicerCanvasPeriod, SlicerObjectLayout } from './slicer-renderer.ts'
import { slicerObjectLocalPoint } from './slicer-renderer.ts'
import { normalizeTimelineDateRange } from './timeline-date-filter.ts'
import { timelinePeriodRange } from './timeline-periods.ts'

export interface SlicerGestureContext { unitId?: string; sheetId: string; revision?: number; userId: string; viewport: GridViewport }
export interface SlicerCanvasGesture extends SlicerGestureContext {
  signature: string
  layout: SlicerObjectLayout
  startPoint: { x: number; y: number }
  localStartPoint: { x: number; y: number }
}

const signature = (layout: SlicerObjectLayout) => JSON.stringify({ model: layout.model, geometry: layout.geometry })
export function captureSlicerCanvasGesture(layout: SlicerObjectLayout, startPoint: { x: number; y: number }, context: SlicerGestureContext): SlicerCanvasGesture {
  return { ...context, viewport: { ...context.viewport }, layout, startPoint: { ...startPoint },
    localStartPoint: slicerObjectLocalPoint(layout, startPoint), signature: signature(layout) }
}
export function canResumeSlicerCanvasGesture(gesture: SlicerCanvasGesture, layout: SlicerObjectLayout | undefined, context: SlicerGestureContext): boolean {
  return !!layout && gesture.unitId === context.unitId && gesture.sheetId === context.sheetId && gesture.revision === context.revision && gesture.userId === context.userId
    && gesture.signature === signature(layout) && (['zoom', 'scrollLeft', 'scrollTop', 'width', 'height'] as const)
      .every(key => gesture.viewport[key] === context.viewport[key])
}

/** World translation and local-axis resizing share the same authoritative geometry as the mixed drawing commands. */
export function slicerCanvasGestureGeometry(gesture: SlicerCanvasGesture, point: { x: number; y: number }, kind: 'move' | 'resize'): DrawingAffineGeometry {
  const before = gesture.layout.geometry, m = { ...before.matrix }
  if (kind === 'move') {
    m.e += point.x - gesture.startPoint.x
    m.f += point.y - gesture.startPoint.y
  } else {
    const local = slicerObjectLocalPoint(gesture.layout, point), dx = local.x - gesture.localStartPoint.x, dy = local.y - gesture.localStartPoint.y
    let sx = Math.max(80 / before.width, 1 + dx / before.width), sy = Math.max(48 / before.height, 1 + dy / before.height)
    if (gesture.layout.model.drawing?.keepRatio) {
      const scale = Math.abs(dx / before.width) >= Math.abs(dy / before.height) ? sx : sy
      sx = sy = Math.max(80 / before.width, 48 / before.height, scale)
    }
    m.a *= sx; m.b *= sx; m.c *= sy; m.d *= sy
  }
  return normalizeDrawingAffineGeometry({ width: before.width, height: before.height, matrix: m })
}

export function slicerCanvasGestureDelta(before: DrawingAffineGeometry, after: DrawingAffineGeometry): DrawingAffineMatrix {
  return multiplyDrawingAffine(after.matrix, invertDrawingAffine(before.matrix))
}

/** DOM Slider pointer input uses the same inverse affine as Canvas, including rotation and reflection. */
export function slicerCanvasScrollOffset(gesture: SlicerCanvasGesture, point: { x: number; y: number }, part: 'track' | 'thumb'): number {
  const layout = gesture.layout, track = layout.scrollbar, thumb = layout.scrollThumb
  if (!track || !thumb || layout.maxScrollOffset <= 0) return 0
  const local = slicerObjectLocalPoint(layout, point), horizontal = layout.model.type === 'timeline',
    coordinate = horizontal ? local.x : local.y, first = horizontal ? gesture.localStartPoint.x : gesture.localStartPoint.y,
    start = horizontal ? track.x : track.y, extent = horizontal ? track.width : track.height,
    travel = extent - (horizontal ? thumb.width : thumb.height),
    offset = part === 'thumb' ? layout.scrollOffset + (travel > 0 ? (coordinate - first) / travel * layout.maxScrollOffset : 0)
      : extent > 0 ? (coordinate - start) / extent * layout.maxScrollOffset : 0,
    step = horizontal ? 1 : Math.max(1, layout.model.columnCount ?? 1)
  return Math.min(layout.maxScrollOffset, Math.max(0, Math.round(offset / step) * step))
}

/** Selection grain expands the clicked display blocks; it never changes either grain or shrinks a displayed block. */
export function timelineCanvasSelectionRange(model: SlicerModel, first: SlicerCanvasPeriod, last: SlicerCanvasPeriod, date1904: boolean): ITimelineRange {
  const level = model.selectionLevel ?? model.level ?? 'month',
    start = timelinePeriodRange(first.startDate, level, date1904), end = timelinePeriodRange(last.startDate, level, date1904)
  return normalizeTimelineDateRange({ startDate: Math.min(first.startDate, start.startDate), endDate: Math.max(last.endDate, end.endDate),
    ...(model.timelineDateRange?.useWholeDay !== undefined ? { useWholeDay: model.timelineDateRange.useWholeDay } : {}) })
}

/** Moving one handle retains the other author's exact fractional bound and whole-day setting. */
export function timelineCanvasHandleRange(model: SlicerModel, period: SlicerCanvasPeriod, side: 'start' | 'end', date1904: boolean): ITimelineRange | undefined {
  const current = model.timelineDateRange
  if (current === undefined) return undefined
  const selected = timelineCanvasSelectionRange(model, period, period, date1904),
    range = { ...current, ...(side === 'start' ? { startDate: selected.startDate } : { endDate: selected.endDate }) }
  return range.startDate < range.endDate ? normalizeTimelineDateRange(range) : undefined
}
