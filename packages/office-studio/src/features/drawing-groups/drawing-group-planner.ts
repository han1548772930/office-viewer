import type { SheetDrawingAffineGeometry, SheetDrawingLeafRef, SheetDrawingRef } from '../../domain/sheet-drawings.ts'
import { assertDrawingEditable, assertDrawingSelectionDisjoint, cloneDrawingRef, cloneDrawingState, drawingContainer, drawingDescendants, drawingGeometryFor, drawingGraph, drawingId, drawingRefKey, drawingSelection } from './drawing-group-model.ts'
import { cloneDrawingGeometry, drawingGeometryEqual, drawingGroupFrame, drawingGroupFrameDelta, normalizeDrawingGeometry, normalizeDrawingMatrix, transformDrawingGeometry } from './drawing-group-geometry.ts'
import { drawingPlan, drawingTransaction } from './drawing-group-transaction.ts'
import type { SheetDrawingArrangeRequest, SheetDrawingGeometryChange, SheetDrawingGroupRequest, SheetDrawingGroupTransformRequest, SheetDrawingPlanResult, SheetDrawingSceneInput, SheetDrawingSelectionTransformRequest, SheetDrawingUngroupRequest } from './drawing-group-types.ts'
import { SheetDrawingError } from './drawing-group-types.ts'

export function planDrawingGroup(input: SheetDrawingSceneInput, request: SheetDrawingGroupRequest): SheetDrawingPlanResult {
  return drawingPlan(() => {
    const graph = drawingGraph(input), selection = drawingSelection(graph, request.selection)
    if (selection.length < 2) return { status: 'noop', reason: selection.length ? 'too-few-objects' : 'empty-selection' }
    assertDrawingSelectionDisjoint(graph, selection)
    assertDrawingEditable(graph, selection, true, true)
    const groupId = drawingId(request.groupId), ref = { kind: 'group' as const, id: groupId }
    if (graph.groups.has(groupId) || [...graph.objects.values()].some(object => object.ref.id === groupId)) throw new SheetDrawingError('id-conflict', 'The new group identifier already exists.', ref)
    const parent = graph.parents.get(drawingRefKey(selection[0]!))!
    if (selection.some(ref => graph.parents.get(drawingRefKey(ref)) !== parent)) throw new SheetDrawingError('different-parent', 'Only sibling drawings can be grouped.')
    const next = cloneDrawingState(graph.scene.state), container = drawingContainer(next, parent), selected = new Set(selection.map(drawingRefKey))
    const members = container.filter(ref => selected.has(drawingRefKey(ref)))
    let lastIndex = -1
    for (let index = 0; index < container.length; index++) if (selected.has(drawingRefKey(container[index]!))) lastIndex = index
    const insertion = container.slice(0, lastIndex + 1).filter(ref => !selected.has(drawingRefKey(ref))).length
    const remaining = container.filter(ref => !selected.has(drawingRefKey(ref)))
    remaining.splice(insertion, 0, ref); container.splice(0, container.length, ...remaining)
    next.groups.push({ id: groupId, members: members.map(cloneDrawingRef), frame: drawingGroupFrame(members.map(ref => drawingGeometryFor(graph, ref))) })
    return drawingTransaction(graph, 'group', next, selection, [ref], [], [...selection, ref])
  })
}

export function planDrawingUngroup(input: SheetDrawingSceneInput, request: SheetDrawingUngroupRequest): SheetDrawingPlanResult {
  return drawingPlan(() => {
    const graph = drawingGraph(input), selection = drawingSelection(graph, request.selection), groups = selection.filter(ref => ref.kind === 'group')
    if (!groups.length) return { status: 'noop', reason: selection.length ? 'unchanged' : 'empty-selection' }
    assertDrawingSelectionDisjoint(graph, selection); assertDrawingEditable(graph, groups)
    const next = cloneDrawingState(graph.scene.state), removed = new Set(groups.map(ref => ref.id)), selectionAfter: SheetDrawingRef[] = []
    for (const ref of selection) {
      if (ref.kind !== 'group') { selectionAfter.push(cloneDrawingRef(ref)); continue }
      const group = graph.groups.get(ref.id)!, parent = graph.parents.get(drawingRefKey(ref))!, container = drawingContainer(next, parent)
      const position = container.findIndex(item => drawingRefKey(item) === drawingRefKey(ref))
      container.splice(position, 1, ...group.members.map(cloneDrawingRef)); selectionAfter.push(...group.members.map(cloneDrawingRef))
    }
    next.groups = next.groups.filter(group => !removed.has(group.id))
    return drawingTransaction(graph, 'ungroup', next, selection, selectionAfter, [], groups)
  })
}

export function planDrawingArrange(input: SheetDrawingSceneInput, request: SheetDrawingArrangeRequest): SheetDrawingPlanResult {
  return drawingPlan(() => {
    const graph = drawingGraph(input), selection = drawingSelection(graph, request.selection)
    if (!selection.length) return { status: 'noop', reason: 'empty-selection' }
    assertDrawingSelectionDisjoint(graph, selection)
    const parent = graph.parents.get(drawingRefKey(selection[0]!))!
    if (selection.some(ref => graph.parents.get(drawingRefKey(ref)) !== parent)) throw new SheetDrawingError('different-parent', 'Layer order is scoped to one drawing container.')
    if ((request.arrangeType === undefined) === (request.zOrder === undefined)) throw new SheetDrawingError('invalid-input', 'Specify one arrange action or one absolute order.')
    const next = cloneDrawingState(graph.scene.state), container = drawingContainer(next, parent), before = container.map(cloneDrawingRef), selected = new Set(selection.map(drawingRefKey))
    const isSelected = (ref: SheetDrawingRef) => selected.has(drawingRefKey(ref))
    if (request.zOrder !== undefined) {
      if (selection.length !== 1 || typeof request.zOrder !== 'number') throw new SheetDrawingError('invalid-input', 'Absolute drawing order requires one drawing and a numeric position.')
      const position = container.findIndex(isSelected), target = Math.max(0, Math.min(container.length - 1, Math.floor(request.zOrder))) || 0
      const [ref] = container.splice(position, 1); container.splice(target, 0, ref!)
    } else if (request.arrangeType === 'front' || request.arrangeType === 'back') {
      const chosen = container.filter(isSelected), others = container.filter(ref => !isSelected(ref))
      container.splice(0, container.length, ...(request.arrangeType === 'front' ? [...others, ...chosen] : [...chosen, ...others]))
    } else if (request.arrangeType === 'forward') {
      for (let index = container.length - 2; index >= 0; index--) if (isSelected(container[index]!) && !isSelected(container[index + 1]!)) {
        [container[index], container[index + 1]] = [container[index + 1]!, container[index]!]
      }
    } else if (request.arrangeType === 'backward') {
      for (let index = 1; index < container.length; index++) if (isSelected(container[index]!) && !isSelected(container[index - 1]!)) {
        [container[index], container[index - 1]] = [container[index - 1]!, container[index]!]
      }
    } else throw new SheetDrawingError('invalid-input', 'Unknown drawing arrange action.')
    const changed = before.filter((ref, index) => drawingRefKey(ref) !== drawingRefKey(container[index]!))
    if (!changed.length) return { status: 'noop', reason: 'at-boundary' }
    assertDrawingEditable(graph, [...selection, ...changed])
    return drawingTransaction(graph, 'arrange', next, selection, selection, [], [...selection, ...changed])
  })
}

export function planDrawingGroupTransform(input: SheetDrawingSceneInput, request: SheetDrawingGroupTransformRequest): SheetDrawingPlanResult {
  return drawingPlan(() => {
    const ref = { kind: 'group' as const, id: drawingId(request.groupId) }
    return planDrawingSelectionTransform(input, { ...request, selection: [ref] })
  })
}

export function planDrawingSelectionTransform(input: SheetDrawingSceneInput, request: SheetDrawingSelectionTransformRequest): SheetDrawingPlanResult {
  return drawingPlan(() => {
    const graph = drawingGraph(input), selection = drawingSelection(graph, request.selection)
    if (!selection.length) return { status: 'noop', reason: 'empty-selection' }
    assertDrawingSelectionDisjoint(graph, selection)
    if ((request.frame === undefined) === (request.delta === undefined)) throw new SheetDrawingError('invalid-input', 'Specify one target frame or one world delta.')
    assertDrawingEditable(graph, selection, false, true, true)
    const beforeFrame = selection.length === 1 ? drawingGeometryFor(graph, selection[0]!) : drawingGroupFrame(selection.map(ref => drawingGeometryFor(graph, ref)))
    const frame = request.frame === undefined ? undefined : normalizeDrawingGeometry(request.frame)
    const delta = frame ? drawingGroupFrameDelta(beforeFrame, frame) : normalizeDrawingMatrix(request.delta)
    const descendants = drawingDescendants(graph, selection), changes: SheetDrawingGeometryChange[] = [], next = cloneDrawingState(graph.scene.state)
    const groupMap = new Map(next.groups.map(group => [group.id, group])), geometries = new Map<string, SheetDrawingAffineGeometry>()
    for (const child of descendants) {
      const before = drawingGeometryFor(graph, child), after = frame && selection.length === 1 && child.kind === 'group' && child.id === selection[0]!.id
        ? frame : transformDrawingGeometry(before, delta)
      geometries.set(drawingRefKey(child), after)
      if (child.kind === 'group') groupMap.get(child.id)!.frame = cloneDrawingGeometry(after)
      else if (!drawingGeometryEqual(before, after)) changes.push({ ref: child as SheetDrawingLeafRef, before, after })
    }
    const affected = [...descendants], ancestorIds = new Set<string>(), transformed = new Set(descendants.map(drawingRefKey))
    for (const ref of selection) {
      let parent = graph.parents.get(drawingRefKey(ref)) ?? null
      while (parent !== null) { ancestorIds.add(parent); parent = graph.parents.get(drawingRefKey({ kind: 'group', id: parent })) ?? null }
    }
    const ancestorOrder = drawingDescendants(graph, graph.scene.state.rootOrder).filter(ref => ref.kind === 'group' && ancestorIds.has(ref.id)).reverse()
    for (const ancestorRef of ancestorOrder) {
      if (transformed.has(drawingRefKey(ancestorRef))) continue
      const ancestor = groupMap.get(ancestorRef.id)!
      ancestor.frame = drawingGroupFrame(ancestor.members.map(member => geometries.get(drawingRefKey(member))
        ?? (member.kind === 'group' ? groupMap.get(member.id)!.frame : drawingGeometryFor(graph, member))), ancestor.frame)
      geometries.set(drawingRefKey(ancestorRef), ancestor.frame); affected.push(ancestorRef)
    }
    if (!changes.length && next.groups.every(group => drawingGeometryEqual(group.frame, graph.groups.get(group.id)!.frame))) return { status: 'noop', reason: 'unchanged' }
    assertDrawingEditable(graph, affected)
    return drawingTransaction(graph, 'transform', next, selection, selection, changes, affected)
  })
}
