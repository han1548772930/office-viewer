import assert from 'node:assert/strict'
import { test } from 'vitest'
import { drawingAffineFromFrame } from '../src/domain/drawing-affine.ts'
import { applyDrawingTransaction, drawingGeometryBounds, drawingGraph, normalizeDrawingState,
  planDrawingGroup, reconcileDrawingState } from '../src/features/drawing-groups/index.ts'

// Deferred structural contracts; no runner or platform verification was executed in this batch.
const ref = (kind, id) => ({ kind, id })
const leaf = (kind, id, left = 0) => ({ ref: ref(kind, id), sheetId: 'data', geometry: drawingAffineFromFrame({ left, top: 0, width: 20, height: 10 }) })
const scene = objects => ({ sheetId: 'data', objects, state: normalizeDrawingState({ sheetId: 'data', objects }) })
const group = (input, selection, groupId) => {
  const result = planDrawingGroup(input, { selection, groupId }); assert.equal(result.status, 'ready')
  return applyDrawingTransaction(input, result.transaction).scene
}
const code = expected => error => error.code === expected
const keys = refs => refs.map(item => item.kind + ':' + item.id)
function nested() {
  const before = scene([leaf('image', 'a'), leaf('shape', 'b', 30), leaf('chart', 'c', 60)])
  const inner = group(before, [ref('image', 'a'), ref('shape', 'b')], 'inner')
  const outer = group(inner, [ref('group', 'inner'), ref('chart', 'c')], 'outer')
  outer.state.groups.find(item => item.id === 'inner').metadata = { label: 'Inner', keep: { custom: 7 } }
  outer.state.groups.find(item => item.id === 'outer').metadata = { label: 'Outer' }
  return outer
}
const targetLeaves = () => [leaf('image', 'a2', 100), leaf('shape', 'b2', 130), leaf('chart', 'c2', 160)]
const complete = () => [
  { source: ref('group', 'outer'), target: ref('group', 'outer2') },
  { source: ref('group', 'inner'), target: ref('group', 'inner2') },
  { source: ref('image', 'a'), target: ref('image', 'a2') },
  { source: ref('shape', 'b'), target: ref('shape', 'b2') },
  { source: ref('chart', 'c'), target: ref('chart', 'c2') },
]

test('an explicit deletion keeps a singleton group and refits its frame; deleting its last member removes empty ancestors', () => {
  const before = scene([leaf('image', 'a'), leaf('chart', 'c', 30), leaf('shape', 'outside', 60)])
  const grouped = group(before, [ref('image', 'a'), ref('chart', 'c')], 'g')
  const current = { ...grouped, objects: grouped.objects.filter(item => item.ref.id !== 'a') }
  const one = reconcileDrawingState(current, { removedRefs: [ref('image', 'a')] })
  assert.deepEqual(keys(one.rootOrder), ['group:g', 'shape:outside']); assert.deepEqual(one.groups[0].members, [ref('chart', 'c')])
  assert.deepEqual(drawingGeometryBounds(one.groups[0].frame), { left: 30, top: 0, width: 20, height: 10 })
  const empty = reconcileDrawingState({ ...current, state: one, objects: current.objects.filter(item => item.ref.id !== 'c') }, { removedRefs: [ref('chart', 'c')] })
  assert.deepEqual(empty.groups, []); assert.deepEqual(keys(empty.rootOrder), ['shape:outside'])
  assert.equal(empty.revision, grouped.state.revision + 2)
})

test('explicit subtree removal deletes only that subtree and recursively removes newly empty parent groups', () => {
  const before = nested(), current = { ...before, objects: before.objects.filter(item => item.ref.id === 'c') }
  const state = reconcileDrawingState(current, { removedRefs: [ref('group', 'inner')] })
  assert.deepEqual(keys(state.rootOrder), ['group:outer']); assert.deepEqual(state.groups.map(item => item.id), ['outer'])
  assert.deepEqual(state.groups[0].members, [ref('chart', 'c')]); assert.deepEqual(state.groups[0].metadata, { label: 'Outer' })
  const empty = reconcileDrawingState({ ...current, state, objects: [] }, { removedRefs: [ref('chart', 'c')] })
  assert.deepEqual(empty.groups, []); assert.deepEqual(empty.rootOrder, [])
})

test('missing media remains represented, while an owner disappearing without explicit deletion is a contract error', () => {
  const before = nested(), unavailable = { ...before, objects: before.objects.map(item => item.ref.id === 'a' ? { ...item, available: false } : item) }
  assert.deepEqual(reconcileDrawingState(unavailable), before.state)
  assert.throws(() => reconcileDrawingState({ ...before, objects: before.objects.filter(item => item.ref.id !== 'a') }), code('missing-object'))
  assert.deepEqual(keys(before.state.groups.find(item => item.id === 'inner').members), ['image:a', 'shape:b'])
})

test('declared deletions must already be absent from native owners and require a prior structural state', () => {
  const before = nested()
  assert.throws(() => reconcileDrawingState(before, { removedRefs: [ref('image', 'a')] }), code('removed-object-present'))
  assert.throws(() => reconcileDrawingState({ sheetId: before.sheetId, objects: [] }, { removedRefs: [ref('image', 'a')] }), code('invalid-state'))
  assert.throws(() => reconcileDrawingState(before, { removedRefs: [ref('image', 'missing')] }), code('not-found'))
})

test('a complete nested copy preserves metadata and member order, fits the copied native leaf geometry and appends a single new root', () => {
  const before = nested(), current = { ...before, objects: [...before.objects, ...targetLeaves()] }
  const state = reconcileDrawingState(current, { copies: complete() }), graph = drawingGraph({ ...current, state })
  assert.deepEqual(keys(state.rootOrder), ['group:outer', 'group:outer2'])
  assert.deepEqual(keys(graph.groups.get('outer2').members), ['group:inner2', 'chart:c2'])
  assert.deepEqual(keys(graph.groups.get('inner2').members), ['image:a2', 'shape:b2'])
  assert.deepEqual(graph.groups.get('inner2').metadata, { label: 'Inner', keep: { custom: 7 } })
  assert.deepEqual(drawingGeometryBounds(graph.groups.get('inner2').frame), { left: 100, top: 0, width: 50, height: 10 })
  assert.deepEqual(drawingGeometryBounds(graph.groups.get('outer2').frame), { left: 100, top: 0, width: 80, height: 10 })
  assert.deepEqual(graph.groups.get('outer'), before.state.groups.find(item => item.id === 'outer'))
  assert.equal(state.revision, before.state.revision + 1)
})

test('incomplete subtree mappings, reused group/leaf IDs and cross-kind copies reject atomically', () => {
  const before = nested(), current = { ...before, objects: [...before.objects, ...targetLeaves()] }, snapshot = structuredClone(current)
  assert.throws(() => reconcileDrawingState(current, { copies: complete().filter(item => item.source.id !== 'b') }), code('incomplete-copy'))
  const collision = complete(); collision[0].target = ref('group', 'a2')
  assert.throws(() => reconcileDrawingState(current, { copies: collision }), code('id-conflict'))
  assert.throws(() => reconcileDrawingState(current, { copies: [{ source: ref('image', 'a'), target: ref('shape', 'b2') }] }), code('invalid-input'))
  assert.throws(() => reconcileDrawingState(current, { copies: [...complete(), complete()[2]] }), code('id-conflict'))
  assert.deepEqual(current, snapshot)
})

test('an independently copied member becomes a root while its source remains in the original group', () => {
  const before = nested(), current = { ...before, objects: [...before.objects, leaf('image', 'a2', 100)] }
  const state = reconcileDrawingState(current, { copies: [{ source: ref('image', 'a'), target: ref('image', 'a2') }] })
  assert.deepEqual(keys(state.rootOrder), ['group:outer', 'image:a2'])
  assert.deepEqual(state.groups, before.state.groups)
  assert.deepEqual(state.groups.find(item => item.id === 'inner').members, [ref('image', 'a'), ref('shape', 'b')])
})

test('an explicit copied group frame preserves its authored shear/rotation instead of refitting or duplicating the child transforms', () => {
  const before = nested(), current = { ...before, objects: [...before.objects, ...targetLeaves()] }, copies = complete()
  const explicit = { width: 80, height: 10, matrix: { a: 1, b: 0.5, c: 0.25, d: 1, e: 100, f: 40 } }
  copies[0].frame = explicit
  const state = reconcileDrawingState(current, { copies })
  assert.deepEqual(state.groups.find(item => item.id === 'outer2').frame, explicit)
  assert.deepEqual(current.objects.slice(-3), targetLeaves())
  assert.deepEqual(state.groups.find(item => item.id === 'outer').frame, before.state.groups.find(item => item.id === 'outer').frame)
})

test('the model can reconcile an explicitly removed source subtree and its complete fresh copy in one publication', () => {
  const before = nested(), current = { ...before, objects: targetLeaves() }
  const state = reconcileDrawingState(current, { removedRefs: [ref('group', 'outer')], copies: complete() })
  assert.deepEqual(keys(state.rootOrder), ['group:outer2']); assert.deepEqual(new Set(state.groups.map(item => item.id)), new Set(['outer2', 'inner2']))
  assert.deepEqual(new Set(keys([...drawingGraph({ ...current, state }).refs.values()])),
    new Set(['group:outer2', 'group:inner2', 'image:a2', 'shape:b2', 'chart:c2']))
})
