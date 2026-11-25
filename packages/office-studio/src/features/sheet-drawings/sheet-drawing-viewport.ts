import type { CellRange } from '../../domain/model.ts'
import type { SheetDrawingAffineGeometry, SheetDrawingLeafRef, SheetDrawingRef } from '../../domain/sheet-drawings.ts'
import { drawingAffineBounds, drawingAffinePoint, invertDrawingAffine, multiplyDrawingAffine, normalizeDrawingAffineGeometry } from '../../domain/drawing-affine.ts'
import { drawingGeometryContains, drawingGraph, drawingHitOrder, drawingPickTarget, drawingRefKey, type SheetDrawingSceneInput } from '../drawing-groups/index.ts'
import type { RenderRect } from '../../rendering/contracts.ts'
import { intersectRect } from '../../rendering/geometry.ts'
import { GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH, type GridViewport, type SpreadsheetGeometry } from '../../rendering/spreadsheet/layout.ts'

export const SHEET_DRAWING_HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'] as const
export type SheetDrawingHandle = typeof SHEET_DRAWING_HANDLES[number] | 'rotate'
export interface SheetDrawingPaneLayout {
  frame: SheetDrawingAffineGeometry
  bounds: RenderRect
  clip: RenderRect
  frozenX: boolean
  frozenY: boolean
  handles: Record<SheetDrawingHandle, { x: number; y: number }>
}
export function sheetDrawingPaneLayouts(geometry: SpreadsheetGeometry, viewport: GridViewport, frame: SheetDrawingAffineGeometry): SheetDrawingPaneLayout[] {
  return geometry.panes(viewport).flatMap(pane => {
    if (!pane.clip.width || !pane.clip.height) return []
    const frozenX = geometry.frozenColumns > 0 && pane.clip.x === GRID_HEADER_WIDTH,
      frozenY = geometry.frozenRows > 0 && pane.clip.y === GRID_HEADER_HEIGHT,
      matrix = multiplyDrawingAffine({ a: viewport.zoom, b: 0, c: 0, d: viewport.zoom,
        e: GRID_HEADER_WIDTH - (frozenX ? 0 : viewport.scrollLeft), f: GRID_HEADER_HEIGHT - (frozenY ? 0 : viewport.scrollTop) }, frame.matrix),
      projected = { ...frame, matrix }, bounds = drawingAffineBounds(projected),
      rect = { x: bounds.left, y: bounds.top, width: bounds.width, height: bounds.height }, visible = intersectRect(rect, pane.clip)
    // Rotation handles can lie just outside the content's visible bounds.
    if (visible.width <= 0 || visible.height <= 0) return []
    const positions = { nw: [0, 0], n: [frame.width / 2, 0], ne: [frame.width, 0], e: [frame.width, frame.height / 2],
      se: [frame.width, frame.height], s: [frame.width / 2, frame.height], sw: [0, frame.height], w: [0, frame.height / 2] }
    const handles = Object.fromEntries(Object.entries(positions).map(([key, point]) => [key, drawingAffinePoint(matrix, point[0]!, point[1]!)])) as Record<SheetDrawingHandle, { x: number; y: number }>
    const center = drawingAffinePoint(matrix, frame.width / 2, frame.height / 2), top = handles.n,
      length = Math.hypot(top.x - center.x, top.y - center.y)
    handles.rotate = { x: top.x + (top.x - center.x) / length * 24, y: top.y + (top.y - center.y) / length * 24 }
    return [{ frame: projected, bounds: rect, clip: pane.clip, frozenX, frozenY, handles }]
  })
}
export function sheetDrawingWorldPoint(viewport: GridViewport, position: { x: number; y: number }, pane: Pick<SheetDrawingPaneLayout, 'frozenX' | 'frozenY'>): { x: number; y: number } {
  return { x: (position.x - GRID_HEADER_WIDTH + (pane.frozenX ? 0 : viewport.scrollLeft)) / viewport.zoom,
    y: (position.y - GRID_HEADER_HEIGHT + (pane.frozenY ? 0 : viewport.scrollTop)) / viewport.zoom }
}
export function sheetDrawingGeometryRange(geometry: SpreadsheetGeometry, frame: SheetDrawingAffineGeometry): CellRange {
  const bounds = drawingAffineBounds(frame)
  const locate = (axis: SpreadsheetGeometry['rows'], coordinate: number) => Math.max(0, Math.min(axis.count - 1, axis.locate(Math.max(0, coordinate))))
  const interior = (end: number, start: number) => end > start ? Math.max(start, end - Math.max(1e-7, Math.abs(end) * Number.EPSILON)) : start
  return { startRow: locate(geometry.rows, bounds.top), endRow: locate(geometry.rows, interior(bounds.top + bounds.height, bounds.top)),
    startCol: locate(geometry.columns, bounds.left), endCol: locate(geometry.columns, interior(bounds.left + bounds.width, bounds.left)) }
}
export function hitSheetDrawingFrame(layouts: readonly SheetDrawingPaneLayout[], position: { x: number; y: number }, controls = true): { layout: SheetDrawingPaneLayout; handle?: SheetDrawingHandle } | undefined {
  const inside = (clip: RenderRect) => position.x >= clip.x && position.y >= clip.y && position.x <= clip.x + clip.width && position.y <= clip.y + clip.height
  if (controls) for (const layout of [...layouts].reverse()) {
    if (!inside(layout.clip)) continue
    for (const handle of ['rotate', ...SHEET_DRAWING_HANDLES] as const) {
      const point = layout.handles[handle]
      if (Math.hypot(position.x - point.x, position.y - point.y) <= (handle === 'rotate' ? 7 : 6)) return { layout, handle }
    }
  }
  for (const layout of [...layouts].reverse()) if (inside(layout.clip) && drawingGeometryContains(layout.frame, position.x, position.y)) return { layout }
  return undefined
}
export function pickSheetDrawing(input: SheetDrawingSceneInput, geometry: SpreadsheetGeometry, viewport: GridViewport, position: { x: number; y: number },
  options: { editingGroupId?: string; visible?: (ref: SheetDrawingLeafRef, frame: SheetDrawingAffineGeometry) => boolean } = {}): { leaf: SheetDrawingLeafRef; target: SheetDrawingRef; layout: SheetDrawingPaneLayout } | undefined {
  const graph = drawingGraph(input)
  for (const leaf of drawingHitOrder(input)) {
    const object = graph.objects.get(drawingRefKey(leaf))
    if (!object || object.available === false || options.visible && !options.visible(leaf, object.geometry)) continue
    const hit = hitSheetDrawingFrame(sheetDrawingPaneLayouts(geometry, viewport, object.geometry), position, false)
    if (!hit) continue
    try { return { leaf, target: drawingPickTarget(input, leaf, options.editingGroupId), layout: hit.layout } }
    catch (error) {
      // A click outside the active group is handled by the caller when it leaves that editing context.
      if (options.editingGroupId) continue
      throw error
    }
  }
  return undefined
}
/** Resize in the retained local basis; rotation and nonuniform scaling never discard a prior shear. */
export function sheetDrawingGestureFrame(frame: SheetDrawingAffineGeometry, start: { x: number; y: number }, position: { x: number; y: number },
  handle?: SheetDrawingHandle, modifiers: { shiftKey?: boolean; altKey?: boolean } = {}): SheetDrawingAffineGeometry {
  if (!handle) {
    let dx = position.x - start.x, dy = position.y - start.y
    if (modifiers.shiftKey) { if (Math.abs(dx) >= Math.abs(dy)) dy = 0; else dx = 0 }
    return normalizeDrawingAffineGeometry({ ...frame, matrix: { ...frame.matrix, e: frame.matrix.e + dx, f: frame.matrix.f + dy } })
  }
  if (handle === 'rotate') {
    const center = drawingAffinePoint(frame.matrix, frame.width / 2, frame.height / 2)
    let angle = Math.atan2(position.y - center.y, position.x - center.x) - Math.atan2(start.y - center.y, start.x - center.x)
    if (modifiers.shiftKey) angle = Math.round(angle / (Math.PI / 12)) * Math.PI / 12
    const a = Math.cos(angle), b = Math.sin(angle), delta = { a, b, c: -b, d: a,
      e: center.x - a * center.x + b * center.y, f: center.y - b * center.x - a * center.y }
    return normalizeDrawingAffineGeometry({ ...frame, matrix: multiplyDrawingAffine(delta, frame.matrix) })
  }
  const inverse = invertDrawingAffine(frame.matrix), first = drawingAffinePoint(inverse, start.x, start.y), last = drawingAffinePoint(inverse, position.x, position.y),
    dx = last.x - first.x, dy = last.y - first.y, west = handle.includes('w'), east = handle.includes('e'), north = handle.includes('n'), south = handle.includes('s'),
    center = modifiers.altKey ? 2 : 1
  let width = Math.max(0.01, frame.width + (west ? -dx : east ? dx : 0) * center),
    height = Math.max(0.01, frame.height + (north ? -dy : south ? dy : 0) * center)
  if (modifiers.shiftKey) {
    const ratio = frame.width / frame.height
    if (!north && !south || (east || west) && Math.abs(width / frame.width - 1) >= Math.abs(height / frame.height - 1)) height = width / ratio
    else width = height * ratio
  }
  const x = modifiers.altKey ? (frame.width - width) / 2 : west ? frame.width - width : 0,
    y = modifiers.altKey ? (frame.height - height) / 2 : north ? frame.height - height : 0
  // Keep the same bounded local size for large selections; put the size ratio into the matrix.
  const localWidth = Math.min(100_000, width), localHeight = Math.min(100_000, height),
    local = { a: width / localWidth, b: 0, c: 0, d: height / localHeight, e: x, f: y }
  return normalizeDrawingAffineGeometry({ width: localWidth, height: localHeight, matrix: multiplyDrawingAffine(frame.matrix, local) })
}
