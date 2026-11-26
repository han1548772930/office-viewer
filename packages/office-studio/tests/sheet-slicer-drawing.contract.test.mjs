import assert from 'node:assert/strict'
import { test } from 'vitest'
import { drawingAffineBounds, multiplyDrawingAffine } from '../src/domain/drawing-affine.ts'
import { drawingPaintOrder, planDrawingArrange, planDrawingGroup, planDrawingGroupTransform,
  planDrawingSelectionTransform, planDrawingUngroup } from '../src/features/drawing-groups/index.ts'
import { applySheetDrawingTransaction, createSheetDrawingScene, removeSheetDrawings } from '../src/features/sheet-drawings/sheet-drawing-host.ts'
import { updateSlicerPresentation } from '../src/features/slicer/slicer-engine.ts'
import { slicerDrawingAuthor, slicerDrawingRefs, slicerDrawingWorkbook } from './fixtures/slicer-drawing-model-fixture.mjs'

// Model/transaction drafts only; not run, built, browser-checked or accepted as Office parity.
const ref = (kind, id) => ({ kind, id })
const scene = workbook => createSheetDrawingScene(workbook, 'data')
const ready = plan => { assert.equal(plan.status, 'ready', JSON.stringify(plan)); return plan.transaction }
const close = (actual, expected) => { for (const key of Object.keys(expected)) assert.ok(Math.abs(actual[key] - expected[key]) < 1e-8, `${key}: ${actual[key]} / ${expected[key]}`) }
function grouped(workbook) {
  applySheetDrawingTransaction(workbook, 'data', ready(planDrawingGroup(scene(workbook), {
    selection: [ref('image', 'a'), ref('shape', 'b')], groupId: 'inner',
  })))
  applySheetDrawingTransaction(workbook, 'data', ready(planDrawingGroup(scene(workbook), {
    selection: [ref('group', 'inner'), ...slicerDrawingRefs.slice(2)], groupId: 'outer',
  })))
  return scene(workbook)
}

test('Table slicer and date Pivot timeline are groupable affine owners with retained positive local control sizes', () => {
  const workbook = slicerDrawingWorkbook(), before = structuredClone(workbook), input = scene(workbook)
  for (const kind of ['slicer', 'timeline']) {
    const owner = input.objects.find(object => object.ref.kind === kind), model = workbook.slicers.find(control => control.id === owner.ref.id)
    assert.equal(owner.groupable, true); assert.equal(owner.transformable, true)
    assert.deepEqual(owner.geometry, model.drawingGeometry); assert.ok(owner.geometry.width > 0 && owner.geometry.height > 0)
  }
  assert.deepEqual(workbook, before)
})

test('nested mixed transform changes all five leaf kinds once and preserves real filter sources, selection and authors', () => {
  const workbook = slicerDrawingWorkbook({ filtered: true }), input = grouped(workbook), authors = workbook.slicers.map(slicerDrawingAuthor)
  const source = structuredClone({ tables: workbook.tables, pivotTables: workbook.pivotTables, cells: workbook.sheets.map(sheet => sheet.cells) })
  const delta = { a: 1, b: 0.5, c: 0.25, d: 1, e: 30, f: 15 }
  const transaction = ready(planDrawingGroupTransform(input, { groupId: 'outer', delta }))
  assert.equal(transaction.geometryChanges.length, 5)
  assert.equal(new Set(transaction.geometryChanges.map(change => change.ref.kind)).size, 5)
  applySheetDrawingTransaction(workbook, 'data', transaction)
  for (const change of transaction.geometryChanges) {
    const actual = scene(workbook).objects.find(object => object.ref.kind === change.ref.kind && object.ref.id === change.ref.id)
    close(actual.geometry.matrix, multiplyDrawingAffine(delta, change.before.matrix))
    assert.equal(actual.geometry.width, change.before.width); assert.equal(actual.geometry.height, change.before.height)
  }
  assert.deepEqual(workbook.slicers.map(slicerDrawingAuthor), authors)
  assert.deepEqual({ tables: workbook.tables, pivotTables: workbook.pivotTables, cells: workbook.sheets.map(sheet => sheet.cells) }, source)
})

test('forward/backward/forward guards slicer revisions and restores the exact scene geometry and group order', () => {
  const workbook = slicerDrawingWorkbook(), input = grouped(workbook), revisions = workbook.slicers.map(control => control.revision)
  const transaction = ready(planDrawingGroupTransform(input, { groupId: 'outer', delta: { a: 1, b: 0, c: 0, d: 1, e: 35, f: 10 } }))
  applySheetDrawingTransaction(workbook, 'data', transaction)
  const after = scene(workbook)
  assert.deepEqual(workbook.slicers.map(control => control.revision), revisions.map(value => value + 1))
  applySheetDrawingTransaction(workbook, 'data', transaction, 'backward')
  assert.deepEqual(workbook.slicers.map(control => control.revision), revisions)
  for (const original of input.objects) close(scene(workbook).objects.find(object => object.ref.kind === original.ref.kind && object.ref.id === original.ref.id).geometry.matrix, original.geometry.matrix)
  assert.deepEqual(scene(workbook).state, input.state)
  applySheetDrawingTransaction(workbook, 'data', transaction); assert.deepEqual(scene(workbook), after)
})

for (const id of ['s', 't']) test(`newer ${id} content blocks the inverse without partially restoring any mixed leaf`, () => {
  const workbook = slicerDrawingWorkbook(), input = grouped(workbook)
  const transaction = ready(planDrawingGroupTransform(input, { groupId: 'outer', delta: { a: 1, b: 0, c: 0, d: 1, e: 30, f: 15 } }))
  applySheetDrawingTransaction(workbook, 'data', transaction)
  const model = workbook.slicers.find(control => control.id === id); model.revision++; model.caption = 'Newer content'
  const current = structuredClone(workbook)
  assert.throws(() => applySheetDrawingTransaction(workbook, 'data', transaction, 'backward'), error => error.code === 'stale-object')
  assert.deepEqual(workbook, current)
})

test('locked position prevents an entire mixed transform but allows grouping and shared stacking operations', () => {
  const workbook = slicerDrawingWorkbook(); workbook.slicers[0].lockedPosition = true
  const input = grouped(workbook), before = structuredClone(workbook)
  assert.equal(planDrawingGroupTransform(input, { groupId: 'outer', delta: { a: 1, b: 0, c: 0, d: 1, e: 5, f: 0 } }).reason, 'not-transformable')
  assert.deepEqual(workbook, before)
  applySheetDrawingTransaction(workbook, 'data', ready(planDrawingArrange(input, { selection: [ref('group', 'outer')], arrangeType: 'back' })))
  assert.deepEqual(workbook.slicers, before.slicers)
})

test('native slicer arrangement and mixed drawing arrangement use one root order and no legacy zIndex rewrite', () => {
  const workbook = slicerDrawingWorkbook(), slicer = workbook.slicers.find(control => control.id === 's'), zIndex = slicer.drawing.zIndex
  updateSlicerPresentation(slicer, { arrange: 'front' }, workbook)
  assert.deepEqual(drawingPaintOrder(scene(workbook)).at(-1), ref('slicer', 's')); assert.equal(slicer.drawing.zIndex, zIndex)
  applySheetDrawingTransaction(workbook, 'data', ready(planDrawingArrange(scene(workbook), { selection: [ref('slicer', 's')], arrangeType: 'back' })))
  assert.deepEqual(drawingPaintOrder(scene(workbook))[0], ref('slicer', 's')); assert.equal(slicer.drawing.zIndex, zIndex)
  const revision = slicer.revision, state = structuredClone(workbook.sheets[0].drawingState)
  updateSlicerPresentation(slicer, { arrange: 'back' }, workbook)
  assert.equal(slicer.revision, revision); assert.deepEqual(workbook.sheets[0].drawingState, state)
})

test('ungrouping only changes membership and keeps every already published affine native geometry', () => {
  const workbook = slicerDrawingWorkbook(), input = grouped(workbook)
  applySheetDrawingTransaction(workbook, 'data', ready(planDrawingGroupTransform(input, { groupId: 'outer', delta: { a: -1, b: 0.5, c: 0.25, d: 1, e: 850, f: 10 } })))
  const before = structuredClone({ slicers: workbook.slicers, charts: workbook.charts, images: workbook.sheets[0].images, shapes: workbook.sheets[0].shapes })
  applySheetDrawingTransaction(workbook, 'data', ready(planDrawingUngroup(scene(workbook), { selection: [ref('group', 'outer')] })))
  assert.deepEqual({ slicers: workbook.slicers, charts: workbook.charts, images: workbook.sheets[0].images, shapes: workbook.sheets[0].shapes }, before)
  assert.equal(workbook.sheets[0].drawingState.groups.length, 1)
})

test('removing the nested mixed group removes control views while retaining Table, Pivot, source filters and unrelated media', () => {
  const workbook = slicerDrawingWorkbook({ filtered: true }); grouped(workbook)
  const before = structuredClone(workbook), result = removeSheetDrawings(workbook, 'data', [ref('group', 'outer')])
  assert.equal(workbook.slicers.length, 0); assert.equal(workbook.charts.length, 0)
  assert.deepEqual(workbook.tables, before.tables); assert.deepEqual(workbook.pivotTables, before.pivotTables)
  assert.deepEqual(workbook.sheets.map(sheet => sheet.cells), before.sheets.map(sheet => sheet.cells))
  assert.deepEqual(workbook.sheets[0].images.map(image => image.id), ['d']); assert.equal(result.removedRefs.length, 7)
})

test('native layout/position/size updates retain reflection and shear while keeping geometry and layout AABB consistent', () => {
  const workbook = slicerDrawingWorkbook(), slicer = workbook.slicers.find(control => control.id === 's')
  updateSlicerPresentation(slicer, { drawingGeometry: { width: 180, height: 160, matrix: { a: -1, b: 0.25, c: 0.3, d: 1, e: 400, f: 60 } } }, workbook)
  const author = slicerDrawingAuthor(slicer), local = { width: slicer.drawingGeometry.width, height: slicer.drawingGeometry.height }
  updateSlicerPresentation(slicer, { layout: { left: 244, top: 226, width: 400, height: 300 } }, workbook)
  close(drawingAffineBounds(slicer.drawingGeometry), { left: 200, top: 200, width: 400, height: 300 })
  assert.ok(slicer.drawingGeometry.matrix.a < 0); assert.notEqual(slicer.drawingGeometry.matrix.b, 0); assert.notEqual(slicer.drawingGeometry.matrix.c, 0)
  assert.deepEqual({ width: slicer.drawingGeometry.width, height: slicer.drawingGeometry.height }, local)
  updateSlicerPresentation(slicer, { position: { row: 5, column: 3 }, size: { rows: 6, columns: 4 } }, workbook)
  close(drawingAffineBounds(slicer.drawingGeometry), { left: 300, top: 100, width: 400, height: 120 })
  assert.deepEqual(slicerDrawingAuthor(slicer), { ...author, size: { rows: 6, columns: 4 } })
})

test('invalid or mixed geometry patches reject atomically, and repeated presentation values do not increment revisions', () => {
  const workbook = slicerDrawingWorkbook(), slicer = workbook.slicers[0]
  for (const patch of [
    { drawingGeometry: { width: 100, height: 100, matrix: { a: 1, b: 1, c: 1, d: 1, e: 0, f: 0 } }, caption: 'Rejected' },
    { drawingGeometry: slicer.drawingGeometry, layout: { left: 200, top: 200, width: 100, height: 100 } },
    { layout: { left: 200, top: 200, width: 0, height: 100 }, caption: 'Rejected' },
  ]) {
    const before = structuredClone(workbook); assert.throws(() => updateSlicerPresentation(slicer, patch, workbook)); assert.deepEqual(workbook, before)
  }
  const revision = slicer.revision
  updateSlicerPresentation(slicer, { caption: slicer.caption }, workbook); assert.equal(slicer.revision, revision)
  slicer.lockedPosition = true
  const before = structuredClone(workbook)
  assert.throws(() => updateSlicerPresentation(slicer, { layout: { left: 200, top: 200, width: 100, height: 100 }, lockedPosition: false }, workbook), /锁定/)
  assert.deepEqual(workbook, before)
})
