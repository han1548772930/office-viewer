import type { SheetDrawingRef, SheetDrawingState } from '../../domain/sheet-drawings.ts'
import type { DrawingGraph } from './drawing-group-model.ts'
import { assertDrawingEditable, cloneDrawingRef, cloneDrawingState, drawingDescendants, drawingGraph, drawingRefKey, drawingStateEqual, indexDrawingState, normalizeDrawingState } from './drawing-group-model.ts'
import { cloneDrawingGeometry, drawingGeometryEqual, normalizeDrawingGeometry } from './drawing-group-geometry.ts'
import type { SheetDrawingApplyResult, SheetDrawingGeometryChange, SheetDrawingPlanResult, SheetDrawingSceneInput, SheetDrawingTransaction } from './drawing-group-types.ts'
import { SheetDrawingError } from './drawing-group-types.ts'

export function drawingPlan(action: () => SheetDrawingPlanResult): SheetDrawingPlanResult {
  try { return action() }
  catch (error) {
    const failure = error instanceof SheetDrawingError ? error : new SheetDrawingError('invalid-input', error instanceof Error ? error.message : 'Invalid drawing operation.')
    return { status: 'rejected', reason: failure.code, message: failure.message, ...(failure.ref ? { ref: cloneDrawingRef(failure.ref) } : {}) }
  }
}

export function drawingTransaction(graph: DrawingGraph, kind: SheetDrawingTransaction['kind'], next: SheetDrawingState,
  beforeSelection: readonly SheetDrawingRef[], afterSelection: readonly SheetDrawingRef[], changes: readonly SheetDrawingGeometryChange[], affected: readonly SheetDrawingRef[]): SheetDrawingPlanResult {
  next.revision = graph.scene.state.revision + 1
  const afterState = normalizeDrawingState({ ...graph.scene, state: next }), afterIndex = indexDrawingState(afterState)
  const refs = new Map<string, SheetDrawingRef>()
  for (const index of [graph, afterIndex]) {
    const existing = affected.filter(ref => index.refs.has(drawingRefKey(ref)))
    for (const ref of drawingDescendants(index, existing)) refs.set(drawingRefKey(ref), cloneDrawingRef(ref))
  }
  const expectedObjects = [...refs.values()].flatMap(ref => {
    if (ref.kind === 'group') return []
    const object = graph.objects.get(drawingRefKey(ref))
    if (!object) throw new SheetDrawingError('missing-object', 'An affected drawing owner is missing.', ref)
    return [{ ref: cloneDrawingRef(ref), geometry: cloneDrawingGeometry(object.geometry), ...(object.revision === undefined ? {} : { revision: object.revision }) }]
  })
  return { status: 'ready', transaction: { kind, sheetId: graph.scene.sheetId, beforeState: cloneDrawingState(graph.scene.state), afterState,
    geometryChanges: changes.map(change => ({ ref: cloneDrawingRef(change.ref), before: cloneDrawingGeometry(change.before), after: cloneDrawingGeometry(change.after) })),
    expectedObjects, selectionBefore: beforeSelection.map(cloneDrawingRef), selectionAfter: afterSelection.map(cloneDrawingRef), affectedRefs: [...refs.values()],
  } }
}

/** Pure publication recipe: the host applies returned leaf geometry and state together in its document mutation. */
export function applyDrawingTransaction(input: SheetDrawingSceneInput, transaction: SheetDrawingTransaction,
  options: { direction?: 'forward' | 'backward' } = {}): SheetDrawingApplyResult {
  const graph = drawingGraph(input), backward = options.direction === 'backward'
  if (options.direction !== undefined && options.direction !== 'forward' && options.direction !== 'backward') {
    throw new SheetDrawingError('invalid-input', 'Unknown drawing transaction direction.')
  }
  if (transaction.sheetId !== graph.scene.sheetId) throw new SheetDrawingError('stale-state', 'A drawing transaction cannot cross worksheets.')
  const expectedState = backward ? transaction.afterState : transaction.beforeState, nextState = backward ? transaction.beforeState : transaction.afterState
  if (!drawingStateEqual(graph.scene.state, expectedState)) throw new SheetDrawingError('stale-state', 'Drawing hierarchy or order changed after planning.')
  const patches = new Map<string, SheetDrawingGeometryChange>(), guards = new Map(transaction.expectedObjects.map(guard => [drawingRefKey(guard.ref), guard]))
  if (guards.size !== transaction.expectedObjects.length) throw new SheetDrawingError('invalid-input', 'Duplicate drawing transaction guards.')
  for (const change of transaction.geometryChanges) {
    const key = drawingRefKey(change.ref), guard = guards.get(key)
    if (patches.has(key) || !guard || !drawingGeometryEqual(guard.geometry, change.before)) throw new SheetDrawingError('invalid-input', 'Invalid drawing geometry transaction.', change.ref)
    normalizeDrawingGeometry(change.before); normalizeDrawingGeometry(change.after); patches.set(key, change)
  }
  for (const guard of guards.values()) {
    const object = graph.objects.get(drawingRefKey(guard.ref)), change = patches.get(drawingRefKey(guard.ref))
    const expectedGeometry = backward && change ? change.after : guard.geometry
    if (!object || !drawingGeometryEqual(object.geometry, expectedGeometry)
      || (!backward && guard.revision !== undefined && object.revision !== guard.revision)) {
      throw new SheetDrawingError('stale-object', 'An affected drawing changed after planning.', guard.ref)
    }
  }
  assertDrawingEditable(graph, [...guards.values()].map(guard => guard.ref), false, transaction.kind === 'transform', transaction.kind === 'transform')
  const objects = graph.scene.objects.map(object => {
    const change = patches.get(drawingRefKey(object.ref))
    return { ...object, ref: cloneDrawingRef(object.ref), geometry: cloneDrawingGeometry(change ? (backward ? change.before : change.after) : object.geometry) }
  })
  const state = normalizeDrawingState({ sheetId: graph.scene.sheetId, objects, state: nextState })
  if (!drawingStateEqual(state, nextState)) throw new SheetDrawingError('invalid-state', 'The drawing transaction omits known leaf owners.')
  return { scene: { sheetId: graph.scene.sheetId, state, objects }, selection: (backward ? transaction.selectionBefore : transaction.selectionAfter).map(cloneDrawingRef) }
}
