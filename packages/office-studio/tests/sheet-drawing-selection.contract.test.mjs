import assert from 'node:assert/strict'
import { test } from 'vitest'
import { effectScope, ref, shallowRef } from 'vue'
import { drawingAffineFromFrame } from '../src/domain/drawing-affine.ts'
import { applyDrawingTransaction, normalizeDrawingState, planDrawingGroup } from '../src/features/drawing-groups/index.ts'
import { bindSheetDrawingSelection, clearSheetDrawingSelections, getSheetDrawingSelection, hasSheetDrawingSelection,
  setSheetDrawingSelection } from '../src/api/sheet-drawing-selection.ts'
import { normalizeSheetDrawingUiSelection, useSheetDrawingSelection } from '../src/components/spreadsheet/use-sheet-drawing-selection.ts'

// Deferred effect-scope/broker contracts. No DOM, server or test runner has been started.
const drawingRef = (kind, id) => ({ kind, id })
const a = drawingRef('image', 'a'), c = drawingRef('chart', 'c'), b = drawingRef('shape', 'b')
const json = value => JSON.parse(JSON.stringify(value))
function scene(sheetId = 'data') {
  const objects = [a, c, b].map((ref, index) => ({ ref, sheetId,
    geometry: drawingAffineFromFrame({ left: index * 40, top: 10, width: 20, height: 10 }) }))
  return { sheetId, objects, state: normalizeDrawingState({ sheetId, objects }) }
}
const owner = () => ({ isDisposed: false, activeWorkbookId: 'book', editor: { state: { activeKind: 'workbook' } } })
function fixture(initial = []) {
  const scope = effectScope(), slot = shallowRef(owner()), unitId = ref('book'), sheetId = ref('data'), active = ref(true), native = shallowRef(json(initial))
  const scenes = shallowRef({ data: scene(), other: scene('other') }), writes = []
  const ui = scope.run(() => useSheetDrawingSelection({ studio: () => slot.value, unitId: () => unitId.value, sheetId: () => sheetId.value,
    active: () => active.value, scene: () => scenes.value[sheetId.value], nativeSelection: () => native.value,
    applyNative: refs => { writes.push(json(refs)); native.value = json(refs) },
  }))
  return { scope, slot, unitId, sheetId, active, native, scenes, ui, writes,
    dispose() { scope.stop(); clearSheetDrawingSelections(slot.value) } }
}

test('initial native selection is normalized once and broker feedback cannot recursively reapply it', () => {
  const f = fixture([a, a, c])
  try {
    assert.deepEqual(f.ui.refs.value, [a, c]); assert.deepEqual(f.native.value, [a, c])
    const count = f.writes.length; f.ui.select(b, true)
    assert.deepEqual(f.ui.refs.value, [a, c, b]); assert.deepEqual(getSheetDrawingSelection(f.slot.value, 'book', 'data'), [a, c, b])
    assert.ok(f.writes.length - count <= 3)
    f.native.value = [c]
    assert.deepEqual(f.ui.refs.value, [c]); assert.deepEqual(getSheetDrawingSelection(f.slot.value, 'book', 'data'), [c])
  } finally { f.dispose() }
})

test('deduplication distinguishes object kinds, removes missing refs and removes descendants of a selected group', () => {
  const before = scene(), plan = planDrawingGroup(before, { selection: [a, c], groupId: 'g' })
  assert.equal(plan.status, 'ready'); const grouped = applyDrawingTransaction(before, plan.transaction).scene, g = drawingRef('group', 'g')
  assert.deepEqual(normalizeSheetDrawingUiSelection(grouped, [a, b, g, c, b, drawingRef('image', 'missing')]), [b, g])
  const reused = scene(); reused.objects[1].ref = drawingRef('chart', 'a'); reused.state = normalizeDrawingState({ sheetId: reused.sheetId, objects: reused.objects })
  assert.deepEqual(normalizeSheetDrawingUiSelection(reused, [a, drawingRef('chart', 'a')]), [a, drawingRef('chart', 'a')])
})

test('additive parent/child selection replaces its related target, toggles it off and preserves unrelated refs', () => {
  const f = fixture(), before = scene(), plan = planDrawingGroup(before, { selection: [a, c], groupId: 'g' }), g = drawingRef('group', 'g')
  try {
    f.scenes.value = { ...f.scenes.value, data: applyDrawingTransaction(before, plan.transaction).scene }
    f.ui.set([g, b]); f.ui.select(a, true); assert.deepEqual(f.ui.refs.value, [b, a])
    f.ui.select(g, true); assert.deepEqual(f.ui.refs.value, [b, g])
    f.ui.select(g, true); assert.deepEqual(f.ui.refs.value, [b])
    f.ui.select(a, true); f.ui.select(b, false, true); assert.deepEqual(f.ui.refs.value, [a, b])
  } finally { f.dispose() }
})

test('sheets restore isolated selections while a new sheet with reused IDs starts empty', () => {
  const f = fixture([a])
  try {
    f.ui.select(c, true); f.sheetId.value = 'other'; assert.deepEqual(f.ui.refs.value, [])
    f.ui.set([b]); f.sheetId.value = 'data'; assert.deepEqual(f.ui.refs.value, [a, c])
    f.sheetId.value = 'other'; assert.deepEqual(f.ui.refs.value, [b])
    f.unitId.value = 'different-book'; assert.deepEqual(f.ui.refs.value, [])
    f.unitId.value = 'book'; assert.deepEqual(f.ui.refs.value, [b])
  } finally { f.dispose() }
})

test('an explicitly empty API selection remains empty when its host is reactivated', () => {
  const f = fixture([a])
  try {
    f.active.value = false
    setSheetDrawingSelection(f.slot.value, 'book', 'data', [])
    assert.equal(hasSheetDrawingSelection(f.slot.value, 'book', 'data'), true)
    f.active.value = true; assert.deepEqual(f.ui.refs.value, []); assert.deepEqual(f.native.value, [])
  } finally { f.dispose() }
})

test('an API publication updates only the matching visible sheet and inactive publications are restored later', () => {
  const f = fixture([a])
  try {
    setSheetDrawingSelection(f.slot.value, 'book', 'other', [b]); assert.deepEqual(f.ui.refs.value, [a])
    f.sheetId.value = 'other'; assert.deepEqual(f.ui.refs.value, [b])
    setSheetDrawingSelection(f.slot.value, 'book', 'other', [c]); assert.deepEqual(f.ui.refs.value, [c])
    const captured = getSheetDrawingSelection(f.slot.value, 'book', 'other'); captured.push(b); captured[0].id = 'detached-copy'
    assert.notDeepEqual(captured, getSheetDrawingSelection(f.slot.value, 'book', 'other'))
  } finally { f.dispose() }
})

test('scene changes prune deleted objects without retaining an invalid native selection', () => {
  const f = fixture([a, c])
  try {
    const next = scene(); next.objects = next.objects.filter(object => object.ref.kind !== 'image')
    next.state = normalizeDrawingState({ sheetId: next.sheetId, objects: next.objects })
    f.scenes.value = { ...f.scenes.value, data: next }
    assert.deepEqual(f.ui.refs.value, [c]); assert.deepEqual(f.native.value, [c])
  } finally { f.dispose() }
})

test('replacing the Studio owner with the same IDs cannot revive the previous selection session', () => {
  const f = fixture([a]), old = f.slot.value
  try {
    f.ui.set([a, c]); f.slot.value = owner()
    assert.deepEqual(f.ui.refs.value, []); assert.deepEqual(getSheetDrawingSelection(f.slot.value, 'book', 'data'), [])
    assert.deepEqual(getSheetDrawingSelection(old, 'book', 'data'), [a, c])
    clearSheetDrawingSelections(old)
  } finally { f.dispose() }
})

test('stopping an old scope does not dispose a replacement broker owner and stops native synchronization', () => {
  const f = fixture([a]); let current = { sheetId: 'data', refs: [b] }
  const replacement = bindSheetDrawingSelection(f.slot.value, 'book', () => current, value => { current = json(value) })
  try {
    f.scope.stop(); f.native.value = [c]; assert.deepEqual(f.ui.refs.value, [a])
    setSheetDrawingSelection(f.slot.value, 'book', 'data', [c]); assert.deepEqual(current.refs, [c])
    assert.deepEqual(f.ui.refs.value, [a])
  } finally { replacement.dispose(); f.dispose() }
})
