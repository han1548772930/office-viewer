import assert from 'node:assert/strict'
import { test } from 'vitest'
import { drawingAffineFrame, drawingAffineFromFrame } from '../src/domain/drawing-affine.ts'
import { applyDrawingTransaction, normalizeDrawingState, planDrawingAlign, planDrawingGroup, planDrawingSelectionTransform } from '../src/features/drawing-groups/index.ts'

// Deferred model contracts, not executed in this implementation batch.
const ref = (kind, id) => ({ kind, id })
const leaf = (kind, id, left, top, width, height, extra = {}) => ({ ref: ref(kind, id), sheetId: 'data',
  geometry: drawingAffineFromFrame({ left, top, width, height }), ...extra })
const scene = objects => ({ sheetId: 'data', objects, state: normalizeDrawingState({ sheetId: 'data', objects }) })
const seed = () => scene([leaf('image', 'a', 10, 20, 20, 10), leaf('chart', 'c', 50, 40, 30, 20), leaf('shape', 'b', 90, 5, 10, 30)])
const ready = plan => { assert.equal(plan.status, 'ready', JSON.stringify(plan)); return plan.transaction }
const apply = (input, plan) => applyDrawingTransaction(input, ready(plan)).scene
const frames = input => input.objects.map(object => drawingAffineFrame(object.geometry))
const align = (input, alignType, selection = input.state.rootOrder) => planDrawingAlign(input, { selection, alignType })

for (const [type, coordinate, expected] of [
  ['1', 'left', [10, 10, 10]], ['2', 'left', [45, 40, 50]], ['3', 'left', [80, 70, 90]],
  ['4', 'top', [5, 5, 5]], ['5', 'top', [27.5, 22.5, 17.5]], ['6', 'top', [50, 40, 30]],
]) test(`align ${type} matches fixed transform-frame ${coordinate} semantics and preserves the other coordinate`, () => {
  const before = seed(), original = structuredClone(before), transaction = ready(align(before, type)), after = applyDrawingTransaction(before, transaction).scene
  assert.deepEqual(frames(after).map(frame => frame[coordinate]), expected)
  const other = coordinate === 'left' ? 'top' : 'left'
  assert.deepEqual(frames(after).map(frame => frame[other]), frames(before).map(frame => frame[other]))
  assert.equal(transaction.afterState.revision, transaction.beforeState.revision + 1)
  assert.deepEqual(transaction.selectionAfter, before.state.rootOrder); assert.deepEqual(before, original)
  assert.deepEqual(applyDrawingTransaction(after, transaction, { direction: 'backward' }).scene, before)
})

test('horizontal distribution sorts centers and uses the fixed rounded span/count step even for uneven widths', () => {
  const before = seed(), after = apply(before, align(before, '7', [...before.state.rootOrder].reverse()))
  assert.deepEqual(frames(after).map(frame => frame.left), [10, 40, 70])
  assert.equal(frames(after)[2].left + frames(after)[2].width, 80)
})

test('vertical distribution sorts centers with a stable selection-order tie and rounds the interval to one decimal', () => {
  const before = scene([leaf('image', 'a', 0, 0, 20, 20), leaf('chart', 'c', 40, 5, 20, 10), leaf('shape', 'b', 80, 60, 20, 41)])
  const after = apply(before, align(before, '8', [ref('chart', 'c'), ref('image', 'a'), ref('shape', 'b')]))
  assert.deepEqual(frames(after).map(frame => frame.top), [33.7, 0, 67.4])
  assert.deepEqual(frames(after).map(frame => frame.left), [0, 40, 80])
})

test('reflected and sheared owners align transform frames without rebaking their matrix bases or local author extents', () => {
  let before = seed()
  before = apply(before, planDrawingSelectionTransform(before, { selection: [ref('image', 'a'), ref('chart', 'c')],
    delta: { a: -1.2, b: 0.25, c: 0.4, d: 1.1, e: 300, f: 30 } }))
  const original = structuredClone(before), after = apply(before, align(before, '1'))
  const target = Math.min(...frames(before).map(frame => frame.left))
  for (const [index, object] of after.objects.entries()) {
    assert.ok(Math.abs(drawingAffineFrame(object.geometry).left - target) < 1e-8)
    assert.equal(object.geometry.width, before.objects[index].geometry.width); assert.equal(object.geometry.height, before.objects[index].geometry.height)
    for (const component of ['a', 'b', 'c', 'd']) assert.equal(object.geometry.matrix[component], before.objects[index].geometry.matrix[component])
  }
  assert.deepEqual(before, original)
})

test('a group and separate leaf align once, retaining its child order and member offsets', () => {
  const before = seed(), grouped = apply(before, planDrawingGroup(before, { selection: [ref('image', 'a'), ref('chart', 'c')], groupId: 'g' }))
  const transaction = ready(align(grouped, '3', [ref('group', 'g'), ref('shape', 'b')])), after = applyDrawingTransaction(grouped, transaction).scene
  assert.equal(transaction.afterState.revision, grouped.state.revision + 1)
  assert.equal(new Set(transaction.geometryChanges.map(change => `${change.ref.kind}:${change.ref.id}`)).size, transaction.geometryChanges.length)
  assert.deepEqual(after.state.groups[0].members, grouped.state.groups[0].members)
  assert.equal(frames(after)[1].left - frames(after)[0].left, 40)
  assert.deepEqual(applyDrawingTransaction(after, transaction, { direction: 'backward' }).scene, grouped)
})

test('an ancestor and its selected descendant reject before publishing any partial alignment', () => {
  const before = seed(), grouped = apply(before, planDrawingGroup(before, { selection: [ref('image', 'a'), ref('chart', 'c')], groupId: 'g' })), original = structuredClone(grouped)
  assert.equal(align(grouped, '1', [ref('group', 'g'), ref('image', 'a')]).reason, 'overlapping-selection')
  assert.deepEqual(grouped, original)
})

for (const [flag, reason] of [['editable', 'read-only'], ['available', 'unavailable-object'], ['transformable', 'not-transformable']])
  test(`a ${flag}=false member rejects the complete batch, including one already at the target`, () => {
    const before = seed(); before.objects[0][flag] = false
    const original = structuredClone(before)
    assert.equal(align(before, '1').reason, reason); assert.deepEqual(before, original)
  })

test('none, empty, one-object and already aligned requests return no-op recipes', () => {
  const before = seed()
  assert.equal(align(before, '0').status, 'noop'); assert.equal(align(before, '1', []).reason, 'empty-selection')
  assert.equal(align(before, '1', [ref('image', 'a')]).reason, 'too-few-objects')
  const after = apply(before, align(before, '1')); assert.equal(align(after, '1').reason, 'unchanged')
})

test('invalid types, duplicate typed references and missing owners reject while reused cross-kind IDs remain legal', () => {
  const before = seed()
  assert.equal(align(before, 'left').reason, 'invalid-input')
  assert.equal(align(before, '1', [ref('image', 'a'), ref('image', 'a')]).reason, 'invalid-input')
  assert.equal(align(before, '1', [ref('image', 'missing'), ref('chart', 'c')]).reason, 'not-found')
  const reused = scene([leaf('image', 'same', 10, 10, 20, 10), leaf('chart', 'same', 50, 30, 20, 10)])
  assert.equal(align(reused, '1').status, 'ready')
})

test('one final transaction guards the original hierarchy, owner geometry and permissions', () => {
  const before = seed(), transaction = ready(align(before, '1'))
  const changed = structuredClone(before); changed.objects[1].geometry.matrix.e += 1
  assert.throws(() => applyDrawingTransaction(changed, transaction), error => error.code === 'stale-object')
  const locked = structuredClone(before); locked.objects[1].editable = false
  assert.throws(() => applyDrawingTransaction(locked, transaction), error => error.code === 'read-only')
  const reordered = structuredClone(before); reordered.state.rootOrder.reverse()
  assert.throws(() => applyDrawingTransaction(reordered, transaction), error => error.code === 'stale-state')
})
