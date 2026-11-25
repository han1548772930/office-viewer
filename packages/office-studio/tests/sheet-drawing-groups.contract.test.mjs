import assert from 'node:assert/strict'
import { test } from 'vitest'
import { drawingAffineFromFrame } from '../src/domain/drawing-affine.ts'
import { applyDrawingTransaction, drawingGeometryBounds, drawingGeometryContains, drawingGraph, drawingHitOrder,
  drawingPaintOrder, drawingPickTarget, normalizeDrawingState, planDrawingArrange, planDrawingGroup,
  planDrawingGroupTransform, planDrawingSelectionTransform, planDrawingUngroup } from '../src/features/drawing-groups/index.ts'

// Deferred model contracts. No tests, type checks, servers, builds, browser or WASM were run in this batch.
const ref = (kind, id) => ({ kind, id })
const frame = (left, top, width, height, rotation = 0) => drawingAffineFromFrame({ left, top, width, height }, rotation)
const leaf = (kind, id, left = 0, top = 0, width = 20, height = 10, flags = {}) =>
  ({ ref: ref(kind, id), sheetId: 'data', geometry: frame(left, top, width, height), ...flags })
const scene = objects => ({ sheetId: 'data', objects, state: normalizeDrawingState({ sheetId: 'data', objects }) })
const ready = result => { assert.equal(result.status, 'ready', JSON.stringify(result)); return result.transaction }
const publish = (input, result) => applyDrawingTransaction(input, ready(result)).scene
const grouped = (input, selection, groupId) => publish(input, planDrawingGroup(input, { selection, groupId }))
const keys = refs => refs.map(item => item.kind + ':' + item.id)
const order = input => keys(input.state.rootOrder)
const object = (input, kind, id) => input.objects.find(item => item.ref.kind === kind && item.ref.id === id)
const code = expected => error => error.code === expected
const four = () => scene([leaf('image', 'a', 10, 20), leaf('shape', 'b', 30, 20), leaf('chart', 'c', 50, 40, 30, 20), leaf('image', 'd', 90, 20)])

test('noncontiguous mixed members keep their order and insert the group at the highest selected layer; inverse restores state and geometry', () => {
  const before = four(), transaction = ready(planDrawingGroup(before, { selection: [ref('chart', 'c'), ref('image', 'a')], groupId: 'g' }))
  assert.equal(transaction.geometryChanges.length, 0)
  const applied = applyDrawingTransaction(before, transaction)
  assert.deepEqual(order(applied.scene), ['shape:b', 'group:g', 'image:d'])
  assert.deepEqual(keys(applied.scene.state.groups[0].members), ['image:a', 'chart:c'])
  assert.deepEqual(applied.selection, [ref('group', 'g')]); assert.deepEqual(applied.scene.objects, before.objects)
  const inverse = applyDrawingTransaction(applied.scene, transaction, { direction: 'backward' })
  assert.deepEqual(inverse.scene.state, before.state); assert.deepEqual(inverse.scene.objects, before.objects)
  assert.deepEqual(inverse.selection, transaction.selectionBefore)
})

test('typed image/chart identities remain distinct even when IDs match; a newly created group cannot reuse a leaf ID', () => {
  const before = scene([leaf('image', 'same'), leaf('chart', 'same', 40)])
  const after = grouped(before, [ref('image', 'same'), ref('chart', 'same')], 'mixed')
  assert.deepEqual(keys(after.state.groups[0].members), ['image:same', 'chart:same'])
  assert.deepEqual(keys(drawingPaintOrder(after)), ['image:same', 'chart:same'])
  assert.equal(planDrawingGroup(before, { selection: before.state.rootOrder, groupId: 'same' }).reason, 'id-conflict')
})

test('unknown kinds, missing owners, duplicate parents and disconnected group cycles reject without deleting data', () => {
  const before = four(), original = structuredClone(before)
  assert.throws(() => normalizeDrawingState({ ...before, state: { ...before.state, rootOrder: [ref('unknown', 'a')] } }), code('unknown-kind'))
  assert.throws(() => normalizeDrawingState({ ...before, objects: before.objects.slice(1) }), code('missing-object'))
  const group = { id: 'g', members: [ref('image', 'a')], frame: frame(0, 0, 20, 10) }
  assert.throws(() => drawingGraph({ ...before, state: { ...before.state, rootOrder: [...before.state.rootOrder, ref('group', 'g')], groups: [group] } }), code('invalid-state'))
  const cycle = { version: 1, revision: 0, rootOrder: [], groups: [
    { id: 'g1', members: [ref('group', 'g2'), ...before.state.rootOrder], frame: frame(0, 0, 120, 80) },
    { id: 'g2', members: [ref('group', 'g1')], frame: frame(0, 0, 120, 80) },
  ] }
  assert.throws(() => drawingGraph({ ...before, state: cycle }), code('invalid-state')); assert.deepEqual(before, original)
})

test('nested groups flatten for paint/hit order and resolve hits to the root or direct child of the active editing group', () => {
  const before = four(), inner = grouped(before, [ref('image', 'a'), ref('chart', 'c')], 'inner')
  const outer = grouped(inner, [ref('shape', 'b'), ref('group', 'inner')], 'outer')
  assert.deepEqual(keys(drawingPaintOrder(outer)), ['shape:b', 'image:a', 'chart:c', 'image:d'])
  assert.deepEqual(keys(drawingHitOrder(outer)), ['image:d', 'chart:c', 'image:a', 'shape:b'])
  assert.deepEqual(drawingPickTarget(outer, ref('chart', 'c')), ref('group', 'outer'))
  assert.deepEqual(drawingPickTarget(outer, ref('chart', 'c'), 'outer'), ref('group', 'inner'))
  assert.deepEqual(drawingPickTarget(outer, ref('chart', 'c'), 'inner'), ref('chart', 'c'))
  assert.throws(() => drawingPickTarget(outer, ref('image', 'd'), 'inner'), code('different-parent'))
  const ungrouped = publish(outer, planDrawingUngroup(outer, { selection: [ref('group', 'outer')] }))
  assert.deepEqual(ungrouped.state.rootOrder, inner.state.rootOrder); assert.deepEqual(ungrouped.state.groups, inner.state.groups)
  assert.equal(ungrouped.state.revision, outer.state.revision + 1); assert.deepEqual(ungrouped.objects, before.objects)
})

test('ungrouping several sibling groups in one transaction preserves mixed non-group selection and each container position', () => {
  const before = scene([leaf('image', 'a'), leaf('chart', 'b', 25), leaf('image', 'c', 50), leaf('shape', 'd', 75), leaf('shape', 'e', 100)])
  const first = grouped(before, [ref('image', 'a'), ref('chart', 'b')], 'first')
  const both = grouped(first, [ref('image', 'c'), ref('shape', 'd')], 'second')
  const transaction = ready(planDrawingUngroup(both, { selection: [ref('group', 'first'), ref('group', 'second'), ref('shape', 'e')] }))
  const applied = applyDrawingTransaction(both, transaction)
  assert.deepEqual(order(applied.scene), order(before)); assert.equal(applied.scene.state.groups.length, 0)
  assert.deepEqual(keys(applied.selection), ['image:a', 'chart:b', 'image:c', 'shape:d', 'shape:e'])
  assert.equal(transaction.geometryChanges.length, 0); assert.deepEqual(applied.scene.objects, before.objects)
})

test('ancestor/descendant selections and grouping or arranging across different parents reject', () => {
  const before = grouped(four(), [ref('image', 'a'), ref('chart', 'c')], 'g')
  const overlap = [ref('group', 'g'), ref('image', 'a')]
  assert.equal(planDrawingGroup(before, { selection: overlap, groupId: 'bad' }).reason, 'overlapping-selection')
  assert.equal(planDrawingUngroup(before, { selection: overlap }).reason, 'overlapping-selection')
  assert.equal(planDrawingSelectionTransform(before, { selection: overlap, delta: { a: 1, b: 0, c: 0, d: 1, e: 1, f: 0 } }).reason, 'overlapping-selection')
  assert.equal(planDrawingGroup(before, { selection: [ref('image', 'a'), ref('image', 'd')], groupId: 'cross' }).reason, 'different-parent')
  assert.equal(planDrawingArrange(before, { selection: [ref('image', 'a'), ref('image', 'd')], arrangeType: 'front' }).reason, 'different-parent')
})

test('front/back and one-layer actions preserve the selected objects relative order, including consecutive blocks', () => {
  const before = four(), selection = [ref('chart', 'c'), ref('image', 'a')]
  assert.deepEqual(order(publish(before, planDrawingArrange(before, { selection, arrangeType: 'front' }))), ['shape:b', 'image:d', 'image:a', 'chart:c'])
  assert.deepEqual(order(publish(before, planDrawingArrange(before, { selection, arrangeType: 'back' }))), ['image:a', 'chart:c', 'shape:b', 'image:d'])
  assert.deepEqual(order(publish(before, planDrawingArrange(before, { selection: [ref('image', 'a'), ref('shape', 'b')], arrangeType: 'forward' }))), ['chart:c', 'image:a', 'shape:b', 'image:d'])
  assert.deepEqual(order(publish(before, planDrawingArrange(before, { selection: [ref('shape', 'b'), ref('chart', 'c')], arrangeType: 'backward' }))), ['shape:b', 'chart:c', 'image:a', 'image:d'])
})

test('arrangement boundaries are no-ops and absolute order floors and clamps the fixed public NaN/Infinity behavior', () => {
  const before = four(), chosen = [ref('shape', 'b')]
  assert.equal(planDrawingArrange(before, { selection: [ref('image', 'd')], arrangeType: 'front' }).status, 'noop')
  assert.equal(planDrawingArrange(before, { selection: [], arrangeType: 'back' }).reason, 'empty-selection')
  for (const zOrder of [NaN, -Infinity, -40]) assert.deepEqual(order(publish(before, planDrawingArrange(before, { selection: chosen, zOrder }))), ['shape:b', 'image:a', 'chart:c', 'image:d'])
  assert.deepEqual(order(publish(before, planDrawingArrange(before, { selection: chosen, zOrder: Infinity }))), ['image:a', 'chart:c', 'image:d', 'shape:b'])
  assert.deepEqual(order(publish(before, planDrawingArrange(before, { selection: chosen, zOrder: 2.8 }))), ['image:a', 'chart:c', 'shape:b', 'image:d'])
})

test('displaced read-only siblings block arrangement; unavailable media remains sortable but cannot group or transform', () => {
  const before = four(); object(before, 'shape', 'b').editable = false
  assert.equal(planDrawingArrange(before, { selection: [ref('image', 'a')], arrangeType: 'front' }).reason, 'read-only')
  const unavailable = four(); object(unavailable, 'image', 'a').available = false
  assert.equal(planDrawingGroup(unavailable, { selection: [ref('image', 'a'), ref('chart', 'c')], groupId: 'g' }).reason, 'unavailable-object')
  assert.equal(planDrawingSelectionTransform(unavailable, { selection: [ref('image', 'a')], delta: { a: 1, b: 0, c: 0, d: 1, e: 10, f: 0 } }).reason, 'unavailable-object')
  assert.deepEqual(order(publish(unavailable, planDrawingArrange(unavailable, { selection: [ref('image', 'a')], arrangeType: 'front' }))), ['shape:b', 'chart:c', 'image:d', 'image:a'])
})

test('slicer/timeline owners participate in mixed order while their unconnected affine capabilities reject grouping and transforms', () => {
  const before = scene([leaf('image', 'a'), leaf('slicer', 's', 40, 0, 20, 20, { groupable: false, transformable: false }), leaf('timeline', 't', 70, 0, 20, 20, { groupable: false, transformable: false })])
  assert.equal(planDrawingGroup(before, { selection: [ref('image', 'a'), ref('slicer', 's')], groupId: 'bad' }).reason, 'not-groupable')
  assert.equal(planDrawingSelectionTransform(before, { selection: [ref('timeline', 't')], frame: frame(100, 0, 50, 20) }).reason, 'not-transformable')
  assert.deepEqual(order(publish(before, planDrawingArrange(before, { selection: [ref('timeline', 't')], arrangeType: 'back' }))), ['timeline:t', 'image:a', 'slicer:s'])
})

test('resizing a mixed group changes each world matrix once, retains leaf local dimensions and survives ungroup without a second transform', () => {
  const before = scene([leaf('image', 'a', 10, 20, 20, 10), leaf('chart', 'c', 50, 40, 30, 20)])
  const group = grouped(before, before.state.rootOrder, 'g'), target = frame(100, 200, 140, 120)
  const transaction = ready(planDrawingGroupTransform(group, { groupId: 'g', frame: target })), after = applyDrawingTransaction(group, transaction).scene
  assert.equal(transaction.geometryChanges.length, 2); assert.deepEqual(after.state.groups[0].frame, target)
  assert.deepEqual(drawingGeometryBounds(object(after, 'image', 'a').geometry), { left: 100, top: 200, width: 40, height: 30 })
  assert.deepEqual(drawingGeometryBounds(object(after, 'chart', 'c').geometry), { left: 180, top: 260, width: 60, height: 60 })
  assert.equal(object(after, 'image', 'a').geometry.width, 20); assert.equal(object(after, 'chart', 'c').geometry.height, 20)
  const expanded = publish(after, planDrawingUngroup(after, { selection: [ref('group', 'g')] }))
  assert.deepEqual(expanded.objects, after.objects); assert.equal(expanded.state.groups.length, 0)
})

test('shear/reflection and nonuniform scaling of a rotated leaf preserve the full affine matrix rather than decomposing it', () => {
  const before = scene([leaf('image', 'a', 10, 20), leaf('chart', 'c', 50, 40)]), delta = { a: 1, b: 0.5, c: 0.25, d: 1, e: 3, f: -2 }
  const after = publish(before, planDrawingSelectionTransform(before, { selection: before.state.rootOrder, delta }))
  assert.deepEqual(object(after, 'image', 'a').geometry.matrix, { a: 1, b: 0.5, c: 0.25, d: 1, e: 18, f: 23 })
  const rotated = scene([leaf('shape', 'r')]); rotated.objects[0].geometry = frame(10, 20, 20, 10, 45)
  const scaled = publish(rotated, planDrawingSelectionTransform(rotated, { selection: [ref('shape', 'r')], delta: { a: 2, b: 0, c: 0, d: 3, e: 0, f: 0 } }))
  const m = scaled.objects[0].geometry.matrix
  assert.ok(Math.abs(m.a * m.c + m.b * m.d - 2.5) < 1e-9)
  const reflected = publish(after, planDrawingSelectionTransform(after, { selection: [ref('image', 'a')], delta: { a: -1, b: 0, c: 0, d: 1, e: 200, f: 0 } }))
  assert.deepEqual(object(reflected, 'image', 'a').geometry.matrix, { a: -1, b: 0.5, c: -0.25, d: 1, e: 182, f: 23 })
})

test('nested group transforms patch every leaf exactly once; independently moved children refit each ancestor from child to parent', () => {
  const before = scene([leaf('image', 'a', 0, 0), leaf('shape', 'b', 30, 0), leaf('chart', 'c', 60, 0)])
  const inner = grouped(before, [ref('image', 'a'), ref('shape', 'b')], 'inner'), outer = grouped(inner, [ref('group', 'inner'), ref('chart', 'c')], 'outer')
  const transaction = ready(planDrawingGroupTransform(outer, { groupId: 'outer', delta: { a: 1, b: 0.5, c: 0.25, d: 1, e: 3, f: -2 } }))
  assert.deepEqual(keys(transaction.geometryChanges.map(change => change.ref)), ['image:a', 'shape:b', 'chart:c'])
  const childMove = publish(outer, planDrawingSelectionTransform(outer, { selection: [ref('image', 'a')], delta: { a: 1, b: 0, c: 0, d: 1, e: 100, f: 0 } }))
  const graph = drawingGraph(childMove)
  assert.deepEqual(drawingGeometryBounds(graph.groups.get('inner').frame), { left: 30, top: 0, width: 90, height: 10 })
  assert.deepEqual(drawingGeometryBounds(graph.groups.get('outer').frame), { left: 30, top: 0, width: 90, height: 10 })
  assert.deepEqual(object(childMove, 'shape', 'b').geometry, object(before, 'shape', 'b').geometry)
})

test('a mixed multiselection can transform directly without creating a group or changing its selection/order', () => {
  const before = four(), selection = [ref('image', 'a'), ref('shape', 'b')], delta = { a: 1, b: 0, c: 0, d: 1, e: 7, f: -3 }
  const transaction = ready(planDrawingSelectionTransform(before, { selection, delta })), after = applyDrawingTransaction(before, transaction)
  assert.equal(after.scene.state.groups.length, 0); assert.deepEqual(order(after.scene), order(before)); assert.deepEqual(after.selection, selection)
  assert.equal(transaction.geometryChanges.length, 2)
  assert.deepEqual(object(after.scene, 'image', 'a').geometry.matrix, { a: 1, b: 0, c: 0, d: 1, e: 17, f: 17 })
  assert.deepEqual(object(after.scene, 'chart', 'c').geometry, object(before, 'chart', 'c').geometry)
})

test('transactions reject stale hierarchy, geometry, owner revisions and malformed patches before publication', () => {
  const before = four(); object(before, 'chart', 'c').revision = 4
  const transaction = ready(planDrawingSelectionTransform(before, { selection: [ref('chart', 'c')], delta: { a: 1, b: 0, c: 0, d: 1, e: 10, f: 0 } }))
  const staleState = structuredClone(before); staleState.state.revision++
  assert.throws(() => applyDrawingTransaction(staleState, transaction), code('stale-state'))
  const staleGeometry = structuredClone(before); object(staleGeometry, 'chart', 'c').geometry.matrix.e++
  assert.throws(() => applyDrawingTransaction(staleGeometry, transaction), code('stale-object'))
  const staleRevision = structuredClone(before); object(staleRevision, 'chart', 'c').revision++
  assert.throws(() => applyDrawingTransaction(staleRevision, transaction), code('stale-object'))
  const noGuards = { ...transaction, expectedObjects: [] }
  assert.throws(() => applyDrawingTransaction(before, noGuards), code('invalid-input'))
  assert.throws(() => applyDrawingTransaction(before, transaction, { direction: 'sideways' }), code('invalid-input'))
  const after = applyDrawingTransaction(before, transaction).scene
  assert.deepEqual(applyDrawingTransaction(after, transaction, { direction: 'backward' }).scene.state, before.state)
})

test('singular/nonfinite transforms reject and hit testing inverts shear/reflection without using an AABB as the hit shape', () => {
  const before = four()
  assert.equal(planDrawingSelectionTransform(before, { selection: [ref('image', 'a')], delta: { a: 1, b: 1, c: 1, d: 1, e: 0, f: 0 } }).reason, 'invalid-geometry')
  assert.equal(planDrawingSelectionTransform(before, { selection: [ref('image', 'a')], delta: { a: 1, b: 0, c: 0, d: 1, e: Infinity, f: 0 } }).reason, 'invalid-geometry')
  const geometry = { width: 20, height: 10, matrix: { a: -1, b: 0.5, c: 0.25, d: 1, e: 100, f: 20 } }
  assert.equal(drawingGeometryContains(geometry, 91.25, 30), true)
  assert.equal(drawingGeometryContains(geometry, 80, 20), false)
  assert.equal(drawingGeometryContains(geometry, NaN, 20), false)
})
