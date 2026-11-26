import assert from 'node:assert/strict'
import { test } from 'vitest'
import { drawingAffineBounds, drawingAffineFromFrame } from '../src/domain/drawing-affine.ts'
import { applyDrawingTransaction, normalizeDrawingState, planDrawingGroup, planDrawingMove,
  sheetDrawingMoveDirectionForKey } from '../src/features/drawing-groups/index.ts'

// Deferred pure-model drafts. No test runner has been invoked.
const ref = (kind, id) => ({ kind, id })
const leaf = (kind, id, left, top, width = 20, height = 10, extra = {}) => ({ ref: ref(kind, id), sheetId: 'data',
  geometry: drawingAffineFromFrame({ left, top, width, height }), ...extra })
const scene = objects => ({ sheetId: 'data', objects, state: normalizeDrawingState({ sheetId: 'data', objects }) })
const ready = plan => { assert.equal(plan.status, 'ready', JSON.stringify(plan)); return plan.transaction }
const moved = (input, request) => applyDrawingTransaction(input, ready(planDrawingMove(input, {
  selection: input.state.rootOrder, extent: { width: 200, height: 100 }, ...request }))).scene
const near = (value, expected) => assert.ok(Math.abs(value - expected) < 1e-8, `${value} != ${expected}`)

for (const [direction, dx, dy] of [[0, 0, -1], [1, 1, 0], [2, 0, 1], [3, -1, 0]])
  test(`fixed Direction ${direction} moves every selected native geometry by one worksheet pixel`, () => {
    const before = scene([leaf('image', 'a', 30, 40), leaf('chart', 'c', 80, 50)]), original = structuredClone(before)
    const transaction = ready(planDrawingMove(before, { selection: before.state.rootOrder, direction, extent: { width: 200, height: 100 } }))
    const after = applyDrawingTransaction(before, transaction).scene
    for (const [index, object] of after.objects.entries()) {
      near(object.geometry.matrix.e, before.objects[index].geometry.matrix.e + dx)
      near(object.geometry.matrix.f, before.objects[index].geometry.matrix.f + dy)
    }
    assert.deepEqual(before, original)
    assert.deepEqual(applyDrawingTransaction(after, transaction, { direction: 'backward' }).scene, before)
  })

test('keyboard direction mapping is explicit and unrelated keys remain outside the drawing command', () => {
  assert.deepEqual(['ArrowUp', 'ArrowRight', 'ArrowDown', 'ArrowLeft'].map(sheetDrawingMoveDirectionForKey), [0, 1, 2, 3])
  for (const key of ['Enter', 'Tab', 'Left', '']) assert.equal(sheetDrawingMoveDirectionForKey(key), undefined)
})

test('reflected and sheared four-corner bounds constrain movement without changing author extent or matrix bases', () => {
  const before = scene([leaf('chart', 'c', 0, 0)])
  before.objects[0].geometry = { width: 30, height: 20, matrix: { a: -1, b: 0.2, c: 0.3, d: 1, e: 30.25, f: 20 } }
  const after = moved(before, { direction: 3 }), old = before.objects[0].geometry, next = after.objects[0].geometry
  near(drawingAffineBounds(next).left, 0); near(next.matrix.e - old.matrix.e, -0.25)
  for (const name of ['a', 'b', 'c', 'd']) assert.equal(next.matrix[name], old.matrix[name])
  assert.equal(next.width, old.width); assert.equal(next.height, old.height)
})

test('one selected owner at the boundary constrains the whole mixed selection and retains distances', () => {
  const before = scene([leaf('image', 'a', 30, 20), leaf('chart', 'c', 179.5, 40)])
  const after = moved(before, { direction: 1, step: 10 })
  near(after.objects[0].geometry.matrix.e, 30.5); near(after.objects[1].geometry.matrix.e, 180)
  near(after.objects[1].geometry.matrix.e - after.objects[0].geometry.matrix.e, 149.5)
  assert.equal(planDrawingMove(after, { selection: after.state.rootOrder, direction: 1, extent: { width: 200, height: 100 } }).reason, 'at-boundary')
})

test('nested groups use their outer frame while each descendant receives exactly one common translation', () => {
  let before = scene([leaf('image', 'a', 0.5, 30), leaf('chart', 'c', 30, 30), leaf('shape', 's', 70, 30)])
  before = applyDrawingTransaction(before, ready(planDrawingGroup(before, { selection: [ref('image', 'a'), ref('chart', 'c')], groupId: 'inner' }))).scene
  before = applyDrawingTransaction(before, ready(planDrawingGroup(before, { selection: [ref('group', 'inner'), ref('shape', 's')], groupId: 'outer' }))).scene
  const transaction = ready(planDrawingMove(before, { selection: [ref('group', 'outer')], direction: 3, step: 10, extent: { width: 200, height: 100 } }))
  const after = applyDrawingTransaction(before, transaction).scene
  assert.equal(transaction.geometryChanges.length, 3)
  for (const [index, object] of after.objects.entries()) near(object.geometry.matrix.e - before.objects[index].geometry.matrix.e, -0.5)
  assert.deepEqual(after.state.groups.map(group => group.members), before.state.groups.map(group => group.members))
  assert.deepEqual(applyDrawingTransaction(after, transaction, { direction: 'backward' }).scene, before)
})

test('None-anchored owners can leave the sheet while an anchored peer still constrains a mixed move', () => {
  const absolute = scene([leaf('image', 'a', -20, -30, 20, 10, { constrainMovement: false })])
  const after = moved(absolute, { direction: 3, step: 10 })
  near(after.objects[0].geometry.matrix.e, -30); near(after.objects[0].geometry.matrix.f, -30)
  const mixed = scene([...absolute.objects, leaf('chart', 'c', 0, 20)])
  assert.equal(planDrawingMove(mixed, { selection: mixed.state.rootOrder, direction: 3, extent: { width: 200, height: 100 } }).reason, 'at-boundary')
})

test('fractional boundary deltas follow fixed six-decimal precision and empty selections have no transaction', () => {
  const before = scene([leaf('chart', 'c', 179.87654321, 20)])
  const after = moved(before, { direction: 1 }), delta = after.objects[0].geometry.matrix.e - before.objects[0].geometry.matrix.e
  near(delta, 0.123457)
  assert.equal(planDrawingMove(before, { selection: [], direction: 1, extent: { width: 200, height: 100 } }).reason, 'empty-selection')
})

for (const [flag, reason] of [['editable', 'read-only'], ['available', 'unavailable-object'], ['transformable', 'not-transformable']])
  test(`a ${flag}=false peer rejects the entire move before any owner changes`, () => {
    const before = scene([leaf('image', 'a', 20, 30), leaf('chart', 'c', 60, 30, 20, 10, { [flag]: false })]), original = structuredClone(before)
    const plan = planDrawingMove(before, { selection: before.state.rootOrder, direction: 1, extent: { width: 200, height: 100 } })
    assert.equal(plan.status, 'rejected'); assert.equal(plan.reason, reason); assert.deepEqual(before, original)
  })

test('ancestor overlap, duplicate typed refs and missing owners cannot be double-moved', () => {
  const before = scene([leaf('image', 'a', 20, 30), leaf('chart', 'c', 60, 30)])
  const grouped = applyDrawingTransaction(before, ready(planDrawingGroup(before, { selection: before.state.rootOrder, groupId: 'g' }))).scene
  const move = selection => planDrawingMove(grouped, { selection, direction: 1, extent: { width: 200, height: 100 } })
  assert.equal(move([ref('group', 'g'), ref('image', 'a')]).reason, 'overlapping-selection')
  assert.equal(move([ref('image', 'a'), ref('image', 'a')]).reason, 'invalid-input')
  assert.equal(move([ref('chart', 'missing')]).reason, 'not-found')
})

test('invalid direction, step, extent and movement capabilities reject instead of silently becoming another command', () => {
  const before = scene([leaf('chart', 'c', 30, 20)])
  for (const request of [{ direction: 'right' }, { direction: 4 }, { direction: 1, step: 2 },
    { direction: 1, extent: { width: NaN, height: 100 } }, { direction: 1, extent: { width: 100, height: -1 } }]) {
    assert.equal(planDrawingMove(before, { selection: before.state.rootOrder, extent: { width: 200, height: 100 }, ...request }).reason, 'invalid-input')
  }
  before.objects[0].constrainMovement = 'false'
  assert.equal(planDrawingMove(before, { selection: before.state.rootOrder, direction: 1, extent: { width: 200, height: 100 } }).reason, 'invalid-input')
})
