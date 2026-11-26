import type { SheetDrawingAffineGeometry, SheetDrawingGroup, SheetDrawingLeafRef, SheetDrawingRef, SheetDrawingState } from '../../domain/sheet-drawings.ts'
import { SHEET_DRAWING_LEAF_KINDS } from '../../domain/sheet-drawings.ts'
import { cloneModel } from '../../domain/serialization.ts'
import type { SheetDrawingObject, SheetDrawingScene, SheetDrawingSceneInput } from './drawing-group-types.ts'
import { SheetDrawingError } from './drawing-group-types.ts'
import { cloneDrawingGeometry, normalizeDrawingGeometry } from './drawing-group-geometry.ts'

export const drawingRefKey = (ref: SheetDrawingRef): string => ref.kind + ':' + ref.id
export const cloneDrawingRef = <T extends SheetDrawingRef>(ref: T): T => ({ ...ref })
export function cloneDrawingState(state: SheetDrawingState): SheetDrawingState {
  return { version: 1, revision: state.revision, rootOrder: state.rootOrder.map(cloneDrawingRef), groups: state.groups.map(group => ({
    id: group.id, members: group.members.map(cloneDrawingRef), frame: cloneDrawingGeometry(group.frame),
    ...(group.metadata === undefined ? {} : { metadata: cloneModel(group.metadata) }),
  })) }
}

export function drawingStateEqual(left: SheetDrawingState, right: SheetDrawingState): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}

export function drawingId(value: unknown): string {
  if (typeof value !== 'string' || !value.trim() || value.length > 1024 || /[\u0000-\u001f]/.test(value)) {
    throw new SheetDrawingError('invalid-input', 'A drawing identifier must be a nonempty string without control characters.')
  }
  return value
}

function record(value: unknown, label: string): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new SheetDrawingError('invalid-state', label + ' must be an object.')
  return value as Record<string, unknown>
}

export function normalizeDrawingRef(value: unknown): SheetDrawingRef {
  const raw = record(value, 'Drawing reference'), id = drawingId(raw.id)
  if (raw.kind !== 'group' && !(SHEET_DRAWING_LEAF_KINDS as readonly unknown[]).includes(raw.kind)) {
    throw new SheetDrawingError('unknown-kind', 'Unknown drawing kinds cannot be discarded during normalization.')
  }
  return { kind: raw.kind as SheetDrawingRef['kind'], id } as SheetDrawingRef
}

export function normalizeDrawingObjects(sheetId: string, values: readonly SheetDrawingObject[]): SheetDrawingObject[] {
  drawingId(sheetId)
  if (!Array.isArray(values)) throw new SheetDrawingError('invalid-input', 'Drawing objects must be an array.')
  const seen = new Set<string>()
  return values.map(value => {
    const raw = record(value, 'Drawing object'), ref = normalizeDrawingRef(raw.ref)
    if (ref.kind === 'group') throw new SheetDrawingError('invalid-input', 'Group geometry belongs to drawing state, not the leaf object list.', ref)
    if (raw.sheetId !== sheetId) throw new SheetDrawingError('invalid-input', 'Drawing operations must stay within one worksheet.', ref)
    const key = drawingRefKey(ref)
    if (seen.has(key)) throw new SheetDrawingError('id-conflict', 'Duplicate drawing object reference.', ref)
    seen.add(key)
    for (const name of ['editable', 'groupable', 'transformable', 'available', 'constrainMovement'] as const) {
      if (raw[name] !== undefined && typeof raw[name] !== 'boolean') throw new SheetDrawingError('invalid-input', 'Invalid drawing capability.', ref)
    }
    if (raw.revision !== undefined && (typeof raw.revision !== 'number' || !Number.isSafeInteger(raw.revision) || raw.revision < 0)) {
      throw new SheetDrawingError('invalid-input', 'Invalid drawing object revision.', ref)
    }
    return { ref, sheetId, geometry: normalizeDrawingGeometry(raw.geometry),
      ...(raw.revision === undefined ? {} : { revision: raw.revision as number }),
      ...(raw.editable === undefined ? {} : { editable: raw.editable as boolean }),
      ...(raw.groupable === undefined ? {} : { groupable: raw.groupable as boolean }),
      ...(raw.transformable === undefined ? {} : { transformable: raw.transformable as boolean }),
      ...(raw.available === undefined ? {} : { available: raw.available as boolean }),
      ...(raw.constrainMovement === undefined ? {} : { constrainMovement: raw.constrainMovement as boolean }),
    }
  })
}

/** Structural parsing is deliberately separate from checking that current leaf owners still exist. */
export function readDrawingState(value: unknown): SheetDrawingState {
  const raw = record(value, 'Drawing state')
  if (raw.version !== 1 || typeof raw.revision !== 'number' || !Number.isSafeInteger(raw.revision) || raw.revision < 0
    || !Array.isArray(raw.rootOrder) || !Array.isArray(raw.groups)) throw new SheetDrawingError('invalid-state', 'Invalid drawing state version, revision, or containers.')
  return { version: 1, revision: raw.revision, rootOrder: raw.rootOrder.map(normalizeDrawingRef), groups: raw.groups.map(value => {
    const group = record(value, 'Drawing group')
    if (!Array.isArray(group.members) || !group.members.length) throw new SheetDrawingError('invalid-state', 'A persisted group must contain at least one member.')
    const metadata = group.metadata === undefined ? undefined : record(group.metadata, 'Drawing group metadata')
    return { id: drawingId(group.id), members: group.members.map(normalizeDrawingRef), frame: normalizeDrawingGeometry(group.frame),
      ...(metadata === undefined ? {} : { metadata: cloneModel(metadata) }),
    }
  }) }
}

export interface DrawingStateIndex {
  groups: Map<string, SheetDrawingGroup>
  parents: Map<string, string | null>
  refs: Map<string, SheetDrawingRef>
}

export function indexDrawingState(state: SheetDrawingState): DrawingStateIndex {
  const groups = new Map<string, SheetDrawingGroup>(), parents = new Map<string, string | null>(), refs = new Map<string, SheetDrawingRef>()
  for (const group of state.groups) {
    if (groups.has(group.id)) throw new SheetDrawingError('id-conflict', 'Duplicate drawing group identifier.', { kind: 'group', id: group.id })
    groups.set(group.id, group)
  }
  const claim = (ref: SheetDrawingRef, parent: string | null) => {
    const key = drawingRefKey(ref)
    if (parents.has(key)) throw new SheetDrawingError('invalid-state', 'A drawing may have exactly one parent container.', ref)
    if (ref.kind === 'group' && !groups.has(ref.id)) throw new SheetDrawingError('missing-object', 'A referenced drawing group is missing.', ref)
    parents.set(key, parent); refs.set(key, ref)
  }
  state.rootOrder.forEach(ref => claim(ref, null))
  for (const group of state.groups) {
    if (!group.members.length) throw new SheetDrawingError('invalid-state', 'An empty group must be removed by an explicit reconciliation.')
    group.members.forEach(ref => claim(ref, group.id))
  }
  for (const group of state.groups) if (!parents.has(drawingRefKey({ kind: 'group', id: group.id }))) {
    throw new SheetDrawingError('invalid-state', 'A drawing group is detached from its worksheet hierarchy.', { kind: 'group', id: group.id })
  }
  // A parent walk detects disconnected cycles as well as cycles reachable from roots, without recursion.
  const complete = new Set<string>()
  for (const group of state.groups) {
    let id: string | null = group.id
    const path = new Set<string>()
    while (id !== null && !complete.has(id)) {
      if (path.has(id)) throw new SheetDrawingError('invalid-state', 'Drawing groups cannot contain a cycle.', { kind: 'group', id })
      path.add(id); id = parents.get(drawingRefKey({ kind: 'group', id })) ?? null
    }
    path.forEach(id => complete.add(id))
  }
  return { groups, parents, refs }
}

export function normalizeDrawingState(input: SheetDrawingSceneInput): SheetDrawingState {
  const objects = normalizeDrawingObjects(input.sheetId, input.objects)
  if (input.state === undefined) return { version: 1, revision: 0, rootOrder: objects.map(object => cloneDrawingRef(object.ref)), groups: [] }
  const state = readDrawingState(input.state), index = indexDrawingState(state), objectMap = new Map(objects.map(object => [drawingRefKey(object.ref), object]))
  for (const ref of index.refs.values()) if (ref.kind !== 'group' && !objectMap.has(drawingRefKey(ref))) {
    throw new SheetDrawingError('missing-object', 'A drawing owner is missing; deletion must be explicit rather than inferred from missing media.', ref)
  }
  // Legacy snapshots and newly inserted known leaves get a deterministic topmost registration.
  for (const object of objects) if (!index.refs.has(drawingRefKey(object.ref))) state.rootOrder.push(cloneDrawingRef(object.ref))
  return state
}

export interface DrawingGraph extends DrawingStateIndex {
  scene: SheetDrawingScene
  objects: Map<string, SheetDrawingObject>
}
export function drawingGraph(input: SheetDrawingSceneInput): DrawingGraph {
  const objects = normalizeDrawingObjects(input.sheetId, input.objects), state = normalizeDrawingState({ ...input, objects })
  return { ...indexDrawingState(state), scene: { sheetId: input.sheetId, objects, state }, objects: new Map(objects.map(object => [drawingRefKey(object.ref), object])) }
}

export function drawingGeometryFor(graph: DrawingGraph, ref: SheetDrawingRef): SheetDrawingAffineGeometry {
  const geometry = ref.kind === 'group' ? graph.groups.get(ref.id)?.frame : graph.objects.get(drawingRefKey(ref))?.geometry
  if (!geometry) throw new SheetDrawingError('not-found', 'Drawing geometry is missing.', ref)
  return geometry
}

export function drawingContainer(state: SheetDrawingState, parent: string | null): SheetDrawingRef[] {
  if (parent === null) return state.rootOrder
  const members = state.groups.find(group => group.id === parent)?.members
  if (!members) throw new SheetDrawingError('not-found', 'Drawing parent group is missing.', { kind: 'group', id: parent })
  return members
}

export function drawingSelection(graph: DrawingStateIndex, selection: readonly SheetDrawingRef[]): SheetDrawingRef[] {
  if (!Array.isArray(selection)) throw new SheetDrawingError('invalid-input', 'Drawing selection must be an array.')
  const seen = new Set<string>()
  return selection.map(raw => {
    const ref = normalizeDrawingRef(raw), key = drawingRefKey(ref)
    if (!graph.refs.has(key)) throw new SheetDrawingError('not-found', 'A selected drawing is missing.', ref)
    if (seen.has(key)) throw new SheetDrawingError('invalid-input', 'A drawing selection must not contain duplicate references.', ref)
    seen.add(key); return ref
  })
}

export function drawingDescendants(index: DrawingStateIndex, selection: readonly SheetDrawingRef[]): SheetDrawingRef[] {
  const stack = [...selection].reverse(), result: SheetDrawingRef[] = [], seen = new Set<string>()
  while (stack.length) {
    const ref = stack.pop()!, key = drawingRefKey(ref)
    if (seen.has(key)) continue
    seen.add(key); result.push(cloneDrawingRef(ref))
    if (ref.kind === 'group') stack.push(...[...(index.groups.get(ref.id)?.members ?? [])].reverse())
  }
  return result
}

export function assertDrawingSelectionDisjoint(index: DrawingStateIndex, selection: readonly SheetDrawingRef[]): void {
  const selected = new Set(selection.map(drawingRefKey))
  for (const ref of selection) {
    let parent = index.parents.get(drawingRefKey(ref)) ?? null
    while (parent !== null) {
      if (selected.has(drawingRefKey({ kind: 'group', id: parent }))) throw new SheetDrawingError('overlapping-selection', 'A selection cannot contain both a group and its descendant.', ref)
      parent = index.parents.get(drawingRefKey({ kind: 'group', id: parent })) ?? null
    }
  }
}

export function assertDrawingEditable(graph: DrawingGraph, refs: readonly SheetDrawingRef[], groupable = false, requireAvailable = false, requireTransformable = false): void {
  for (const ref of drawingDescendants(graph, refs)) if (ref.kind !== 'group') {
    const object = graph.objects.get(drawingRefKey(ref))
    if (!object) throw new SheetDrawingError('missing-object', 'An affected drawing owner is missing.', ref)
    if (object.editable === false) throw new SheetDrawingError('read-only', 'An affected drawing is not editable.', ref)
    if (groupable && object.groupable === false) throw new SheetDrawingError('not-groupable', 'An affected drawing cannot participate in grouping.', ref)
    if (requireTransformable && object.transformable === false) throw new SheetDrawingError('not-transformable', 'The drawing owner cannot retain a complete affine transform yet.', ref)
    if (requireAvailable && object.available === false) throw new SheetDrawingError('unavailable-object', 'An affected drawing resource is unavailable and must remain preserved.', ref)
  }
}

export function drawingPaintOrder(input: SheetDrawingSceneInput): SheetDrawingLeafRef[] {
  const graph = drawingGraph(input)
  return drawingDescendants(graph, graph.scene.state.rootOrder).filter((ref): ref is SheetDrawingLeafRef => ref.kind !== 'group')
}
export function drawingHitOrder(input: SheetDrawingSceneInput): SheetDrawingLeafRef[] { return drawingPaintOrder(input).reverse() }

/** Resolve an actual painted leaf hit to the root group, or to its direct child in a group-editing context. */
export function drawingPickTarget(input: SheetDrawingSceneInput, leaf: SheetDrawingLeafRef, editingGroupId?: string): SheetDrawingRef {
  const graph = drawingGraph(input), ref = drawingSelection(graph, [leaf])[0]!
  let target = ref, parent = graph.parents.get(drawingRefKey(ref)) ?? null
  while (parent !== null) {
    if (parent === editingGroupId) return cloneDrawingRef(target)
    target = { kind: 'group', id: parent }; parent = graph.parents.get(drawingRefKey(target)) ?? null
  }
  if (editingGroupId !== undefined) throw new SheetDrawingError('different-parent', 'The hit is outside the active drawing group.', leaf)
  return cloneDrawingRef(target)
}
