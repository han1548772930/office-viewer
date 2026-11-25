import type { SheetDrawingState } from '../../domain/sheet-drawings.ts'
import type { SheetDrawingCopyMapping, SheetDrawingReconcileChanges, SheetDrawingSceneInput } from './drawing-group-types.ts'
import { SheetDrawingError } from './drawing-group-types.ts'
import { cloneDrawingRef, cloneDrawingState, drawingDescendants, drawingRefKey, drawingSelection, drawingStateEqual, indexDrawingState, normalizeDrawingObjects, normalizeDrawingRef, normalizeDrawingState, readDrawingState } from './drawing-group-model.ts'
import { cloneDrawingGeometry, drawingGroupFrame, normalizeDrawingGeometry } from './drawing-group-geometry.ts'
import { cloneModel } from '../../domain/serialization.ts'

/** Reconcile after the host has applied leaf insert/delete/copy changes, without inferring deletion from missing media. */
export function reconcileDrawingState(input: SheetDrawingSceneInput, changes: SheetDrawingReconcileChanges = {}): SheetDrawingState {
  const objects = normalizeDrawingObjects(input.sheetId, input.objects), objectMap = new Map(objects.map(object => [drawingRefKey(object.ref), object]))
  if (input.state === undefined) {
    if (changes.removedRefs?.length || changes.copies?.length) throw new SheetDrawingError('invalid-state', 'Explicit deletion or copy mappings require the previous drawing state.')
    return normalizeDrawingState({ ...input, objects })
  }
  const before = readDrawingState(input.state), source = indexDrawingState(before), next = cloneDrawingState(before)
  const removed = drawingSelection(source, changes.removedRefs ?? []), removedRefs = new Set(drawingDescendants(source, removed).map(drawingRefKey))
  for (const key of removedRefs) {
    const ref = source.refs.get(key)!
    if (ref.kind !== 'group' && objectMap.has(key)) throw new SheetDrawingError('removed-object-present', 'A declared deleted leaf is still present in the host object list.', ref)
  }
  const mappings = new Map<string, SheetDrawingCopyMapping>(), targets = new Set<string>(), explicitFrames = new Set<string>(), copiedGroups = new Set<string>()
  if (changes.copies !== undefined && !Array.isArray(changes.copies)) throw new SheetDrawingError('invalid-input', 'Drawing copy mappings must be an array.')
  for (const raw of changes.copies ?? []) {
    const original = normalizeDrawingRef(raw.source), target = normalizeDrawingRef(raw.target), sourceKey = drawingRefKey(original), targetKey = drawingRefKey(target)
    if (!source.refs.has(sourceKey)) throw new SheetDrawingError('not-found', 'The drawing copy source is missing.', original)
    if (original.kind !== target.kind) throw new SheetDrawingError('invalid-input', 'A copied drawing must retain its kind.', target)
    if (source.refs.has(targetKey) || targets.has(targetKey) || mappings.has(sourceKey)) throw new SheetDrawingError('id-conflict', 'Copy mappings require distinct new drawing identities.', target)
    if (target.kind === 'group' && objects.some(object => object.ref.id === target.id)) throw new SheetDrawingError('id-conflict', 'The copied group identifier conflicts with a leaf owner.', target)
    if (target.kind !== 'group' && !objectMap.has(targetKey)) throw new SheetDrawingError('missing-object', 'The copied leaf owner must be published before reconciling its state.', target)
    if (target.kind !== 'group' && raw.frame !== undefined) throw new SheetDrawingError('invalid-input', 'Only a copied group accepts an explicit world frame.', target)
    mappings.set(sourceKey, { source: original, target, ...(raw.frame === undefined ? {} : { frame: normalizeDrawingGeometry(raw.frame) }) }); targets.add(targetKey)
  }
  for (const mapping of mappings.values()) if (mapping.source.kind === 'group') {
    for (const child of drawingDescendants(source, [mapping.source])) if (!mappings.has(drawingRefKey(child))) {
      throw new SheetDrawingError('incomplete-copy', 'Copying a group requires identity mappings for its complete subtree.', child)
    }
  }
  const changedGroups = new Set<string>()
  for (const ref of drawingDescendants(source, removed)) {
    let parent = source.parents.get(drawingRefKey(ref)) ?? null
    while (parent !== null) { changedGroups.add(parent); parent = source.parents.get(drawingRefKey({ kind: 'group', id: parent })) ?? null }
  }
  next.rootOrder = next.rootOrder.filter(ref => !removedRefs.has(drawingRefKey(ref)))
  next.groups = next.groups.filter(group => !removedRefs.has(drawingRefKey({ kind: 'group', id: group.id })))
    .map(group => ({ ...group, members: group.members.filter(ref => !removedRefs.has(drawingRefKey(ref))) }))
  for (const mapping of mappings.values()) if (mapping.source.kind === 'group' && mapping.target.kind === 'group') {
    const original = source.groups.get(mapping.source.id)!, id = mapping.target.id
    next.groups.push({ id, members: original.members.map(ref => cloneDrawingRef(mappings.get(drawingRefKey(ref))!.target)),
      frame: mapping.frame ? cloneDrawingGeometry(mapping.frame) : cloneDrawingGeometry(original.frame),
      ...(original.metadata === undefined ? {} : { metadata: cloneModel(original.metadata) }),
    })
    copiedGroups.add(id); if (mapping.frame !== undefined) explicitFrames.add(id)
  }
  // New copied group roots are topmost; mapped descendants remain in their copied containers.
  const ownedGroups = new Set(next.groups.flatMap(group => group.members.filter(ref => ref.kind === 'group').map(ref => ref.id)))
  for (const id of copiedGroups) if (!ownedGroups.has(id)) next.rootOrder.push({ kind: 'group', id })
  removeExplicitlyEmptiedGroups(next, changedGroups)
  const intermediate = indexDrawingState(next)
  for (const ref of intermediate.refs.values()) if (ref.kind !== 'group' && !objectMap.has(drawingRefKey(ref))) {
    throw new SheetDrawingError('missing-object', 'A drawing owner disappeared without an explicit deletion; its reference cannot be discarded.', ref)
  }
  for (const object of objects) if (!intermediate.refs.has(drawingRefKey(object.ref))) next.rootOrder.push(cloneDrawingRef(object.ref))
  const index = indexDrawingState(next), groupOrder = drawingDescendants(index, next.rootOrder).filter(ref => ref.kind === 'group').reverse()
  for (const ref of groupOrder) {
    if (!changedGroups.has(ref.id) && (!copiedGroups.has(ref.id) || explicitFrames.has(ref.id))) continue
    const group = index.groups.get(ref.id)!
    group.frame = drawingGroupFrame(group.members.map(member => member.kind === 'group'
      ? index.groups.get(member.id)!.frame : objectMap.get(drawingRefKey(member))!.geometry), group.frame)
  }
  next.revision = before.revision
  if (!drawingStateEqual(next, before)) next.revision++
  return normalizeDrawingState({ ...input, objects, state: next })
}

function removeExplicitlyEmptiedGroups(state: SheetDrawingState, changed: Set<string>): void {
  const groups = new Map(state.groups.map(group => [group.id, group])), parents = new Map<string, string | null>()
  for (const ref of state.rootOrder) if (ref.kind === 'group') parents.set(ref.id, null)
  for (const group of state.groups) for (const ref of group.members) if (ref.kind === 'group') parents.set(ref.id, group.id)
  const queue = state.groups.filter(group => !group.members.length).map(group => group.id)
  for (let cursor = 0; cursor < queue.length; cursor++) {
    const id = queue[cursor]!, group = groups.get(id)
    if (!group || group.members.length) continue
    groups.delete(id)
    const parent = parents.get(id) ?? null, container = parent === null ? state.rootOrder : groups.get(parent)?.members
    if (container) {
      const position = container.findIndex(ref => ref.kind === 'group' && ref.id === id)
      if (position >= 0) container.splice(position, 1)
      if (parent !== null) { changed.add(parent); if (!container.length) queue.push(parent) }
    }
  }
  state.groups = state.groups.filter(group => groups.has(group.id))
}
