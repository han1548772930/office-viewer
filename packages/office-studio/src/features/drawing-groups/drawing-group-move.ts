import { drawingAffineBounds } from '../../domain/drawing-affine.ts'
import type { SheetDrawingMoveDirection, SheetDrawingRef } from '../../domain/sheet-drawings.ts'
import { assertDrawingEditable, assertDrawingSelectionDisjoint, drawingDescendants, drawingGeometryFor, drawingGraph,
  drawingRefKey, drawingSelection } from './drawing-group-model.ts'
import { planDrawingSelectionTransform } from './drawing-group-planner.ts'
import { drawingPlan } from './drawing-group-transaction.ts'
import { SheetDrawingError, type SheetDrawingPlanResult, type SheetDrawingSceneInput } from './drawing-group-types.ts'

export interface SheetDrawingMoveRequest {
  selection: readonly SheetDrawingRef[]
  direction: SheetDrawingMoveDirection
  /** Ordinary arrows use 1 CSS pixel; Shift arrows use the local 10-pixel extension. */
  step?: 1 | 10
  /** Worksheet content dimensions; headers, viewport scrolling and zoom are excluded. */
  extent: { width: number; height: number }
}

export function sheetDrawingMoveDirectionForKey(key: string): SheetDrawingMoveDirection | undefined {
  return key === 'ArrowUp' ? 0 : key === 'ArrowRight' ? 1 : key === 'ArrowDown' ? 2 : key === 'ArrowLeft' ? 3 : undefined
}

/** All selected objects receive one world translation constrained by their complete affine bounds. */
export function planDrawingMove(input: SheetDrawingSceneInput, request: SheetDrawingMoveRequest): SheetDrawingPlanResult {
  return drawingPlan(() => {
    if (![0, 1, 2, 3].includes(request.direction) || request.step !== undefined && request.step !== 1 && request.step !== 10) {
      throw new SheetDrawingError('invalid-input', '绘图移动方向或步长无效。')
    }
    if (!request.extent || ![request.extent.width, request.extent.height].every(value => Number.isFinite(value) && value >= 0)) {
      throw new SheetDrawingError('invalid-input', '绘图移动的工作表尺寸无效。')
    }
    const graph = drawingGraph(input), selection = drawingSelection(graph, request.selection)
    if (!selection.length) return { status: 'noop', reason: 'empty-selection' }
    assertDrawingSelectionDisjoint(graph, selection)
    assertDrawingEditable(graph, selection, false, true, true)
    let minX = -Infinity, minY = -Infinity, maxX = Infinity, maxY = Infinity
    for (const ref of selection) {
      // A group with an anchored child uses its whole frame, preserving its internal distances.
      const constrained = drawingDescendants(graph, [ref]).some(child => child.kind !== 'group'
        && graph.objects.get(drawingRefKey(child))?.constrainMovement !== false)
      if (!constrained) continue
      const bounds = drawingAffineBounds(drawingGeometryFor(graph, ref))
      minX = Math.max(minX, -bounds.left); minY = Math.max(minY, -bounds.top)
      maxX = Math.min(maxX, request.extent.width - bounds.left - bounds.width)
      maxY = Math.min(maxY, request.extent.height - bounds.top - bounds.height)
    }
    const step = request.step ?? 1, direction = request.direction,
      round = (value: number) => Math.round(value * 1_000_000) / 1_000_000,
      x = round(Math.max(minX, Math.min(direction === 3 ? -step : direction === 1 ? step : 0, maxX))),
      y = round(Math.max(minY, Math.min(direction === 0 ? -step : direction === 2 ? step : 0, maxY)))
    if (x === 0 && y === 0) return { status: 'noop', reason: 'at-boundary' }
    return planDrawingSelectionTransform(graph.scene, { selection, delta: { a: 1, b: 0, c: 0, d: 1, e: x, f: y } })
  })
}
