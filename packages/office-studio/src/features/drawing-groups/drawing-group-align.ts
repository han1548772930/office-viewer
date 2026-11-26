import { drawingAffineFrame } from '../../domain/drawing-affine.ts'
import { SHEET_DRAWING_ALIGN_TYPES, type SheetDrawingAlignType, type SheetDrawingRef } from '../../domain/sheet-drawings.ts'
import { assertDrawingEditable, assertDrawingSelectionDisjoint, cloneDrawingState, drawingGeometryFor, drawingGraph, drawingRefKey, drawingSelection } from './drawing-group-model.ts'
import { drawingGeometryEqual } from './drawing-group-geometry.ts'
import { planDrawingSelectionTransform } from './drawing-group-planner.ts'
import { applyDrawingTransaction, drawingPlan, drawingTransaction } from './drawing-group-transaction.ts'
import { SheetDrawingError, type SheetDrawingPlanResult, type SheetDrawingSceneInput } from './drawing-group-types.ts'

export interface SheetDrawingAlignRequest { selection: readonly SheetDrawingRef[]; alignType: SheetDrawingAlignType }

/** Fixed DrawingUpdateController semantics: transform frames and span/count distribution rounded to one decimal. */
export function planDrawingAlign(input: SheetDrawingSceneInput, request: SheetDrawingAlignRequest): SheetDrawingPlanResult {
  return drawingPlan(() => {
    if (!(Object.values(SHEET_DRAWING_ALIGN_TYPES) as readonly string[]).includes(request.alignType)) throw new SheetDrawingError('invalid-input', '绘图对齐方式无效。')
    const graph = drawingGraph(input), selection = drawingSelection(graph, request.selection), align = request.alignType
    if (align === '0' || selection.length < 2) return { status: 'noop', reason: selection.length ? 'too-few-objects' : 'empty-selection' }
    assertDrawingSelectionDisjoint(graph, selection); assertDrawingEditable(graph, selection, false, true, true)
    const items = selection.map((ref, index) => ({ ref, index, frame: drawingAffineFrame(drawingGeometryFor(graph, ref)) })),
      left = Math.min(...items.map(item => item.frame.left)), top = Math.min(...items.map(item => item.frame.top)),
      right = Math.max(...items.map(item => item.frame.left + item.frame.width)), bottom = Math.max(...items.map(item => item.frame.top + item.frame.height)),
      horizontal = Math.round((right - left) / items.length * 10) / 10, vertical = Math.round((bottom - top) / items.length * 10) / 10
    if (align === '7') items.sort((a, b) => a.frame.left + a.frame.width / 2 - b.frame.left - b.frame.width / 2 || a.index - b.index)
    if (align === '8') items.sort((a, b) => a.frame.top + a.frame.height / 2 - b.frame.top - b.frame.height / 2 || a.index - b.index)
    let scene = graph.scene
    const affected = new Map(selection.map(ref => [drawingRefKey(ref), ref]))
    for (const [index, item] of items.entries()) {
      const frame = item.frame, x = align === '1' ? left : align === '2' ? left + (right - left - frame.width) / 2
        : align === '3' ? right - frame.width : align === '7' ? left + horizontal * index : frame.left,
        y = align === '4' ? top : align === '5' ? top + (bottom - top - frame.height) / 2
          : align === '6' ? bottom - frame.height : align === '8' ? top + vertical * index : frame.top
      if (x === frame.left && y === frame.top) continue
      const plan = planDrawingSelectionTransform(scene, { selection: [item.ref], delta: { a: 1, b: 0, c: 0, d: 1, e: x - frame.left, f: y - frame.top } })
      if (plan.status === 'rejected') return plan
      if (plan.status !== 'ready') continue
      scene = applyDrawingTransaction(scene, plan.transaction).scene
      for (const ref of plan.transaction.affectedRefs) affected.set(drawingRefKey(ref), ref)
    }
    const finalObjects = new Map(scene.objects.map(object => [drawingRefKey(object.ref), object]))
    const changes = graph.scene.objects.flatMap(object => {
      const after = finalObjects.get(drawingRefKey(object.ref))!.geometry
      return drawingGeometryEqual(object.geometry, after) ? [] : [{ ref: object.ref, before: object.geometry, after }]
    })
    if (!changes.length) return { status: 'noop', reason: 'unchanged' }
    // Squash all derived moves into one guarded native-owner transaction and one content history entry.
    return drawingTransaction(graph, 'transform', cloneDrawingState(scene.state), selection, selection, changes, [...affected.values()])
  })
}
