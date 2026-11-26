import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createCell } from '../src/domain/model.ts'
import { drawingAffineFrame } from '../src/domain/drawing-affine.ts'
import { createChartDefinition } from '../src/features/charts/chart-engine.ts'
import { drawingGroupFrame, planDrawingGroup } from '../src/features/drawing-groups/index.ts'
import { createSheetDrawingScene, applySheetDrawingTransaction } from '../src/features/sheet-drawings/sheet-drawing-host.ts'
import { initializeWorkbookDrawingStates, reconcileWorkbookDrawingStates } from '../src/features/sheet-drawings/sheet-drawing-model-lifecycle.ts'
import { reconcileWorkbookImageGeometryChanges } from '../src/features/sheet-drawings/sheet-drawing-image-geometry.ts'
import { sheetClipboardObjectAxes } from '../src/features/sheet-workbench/sheet-clipboard-object-geometry.ts'
import { sheetImageFrame } from '../src/features/sheet-images/sheet-image-placement.ts'
import { transformSheetImagesForAxisEdit } from '../src/features/sheet-images/sheet-image-structure.ts'
import { sheetImageAffine, sheetImageFixture } from './fixtures/sheet-image-fixture.mjs'

// Deferred contracts for the final computed model boundary, not executed this batch.
const ref = (kind, id) => ({ kind, id })
const close = (actual, expected) => { for (const key of Object.keys(expected)) assert.ok(Math.abs(actual[key] - expected[key]) < 1e-8, `${key}: ${actual[key]} !== ${expected[key]}`) }
function seed() {
  const f = sheetImageFixture(); f.sheet.defaultRowHeight = 20; f.sheet.defaultColumnWidth = 100; f.sheet.defaultRowAutoHeight = false
  const axes = sheetClipboardObjectAxes(f.workbook, f.sheet)
  f.sheet.images = ['0', '1', '2'].map((anchorType, index) => sheetImageAffine('image-' + anchorType, axes, { anchorType,
    drawingGeometry: { width: 60, height: 40, matrix: { a: 1, b: 0.25, c: 0.3, d: 1, e: 150 + index * 120, f: 60 } } }))
  initializeWorkbookDrawingStates(f.workbook)
  return f.workbook
}
const image = (workbook, type) => workbook.sheets[0].images.find(image => image.id === 'image-' + type)
const authors = value => ({ size: value.size, crop: value.crop, resource: value.resource, rotation: value.rotation,
  flipX: value.flipX, flipY: value.flipY, metadata: value.metadata })

test('final row/column dimensions translate Position once and keep None geometry and every author payload', () => {
  const before = seed(), original = structuredClone(before), after = structuredClone(before)
  after.sheets[0].columnWidths = { 0: 150 }; after.sheets[0].rowHeights = { 0: 32 }
  reconcileWorkbookImageGeometryChanges(before, after)
  close(image(after, '0').drawingGeometry.matrix, { e: image(before, '0').drawingGeometry.matrix.e + 50, f: image(before, '0').drawingGeometry.matrix.f + 12 })
  assert.deepEqual(image(after, '2'), image(before, '2'))
  for (const type of ['0', '1', '2']) assert.deepEqual(authors(image(after, type)), authors(image(before, type)))
  assert.deepEqual(before, original)
})

test('Both fits its projected attachment span while retaining the full matrix basis and local author size/crop', () => {
  const before = seed(), after = structuredClone(before), sheet = after.sheets[0]
  sheet.columnWidths = { 2: 150, 3: 150 }; sheet.rowHeights = { 3: 36, 4: 36 }
  const source = image(before, '1'), plain = { ...source }; delete plain.drawingGeometry
  const expected = sheetImageFrame(plain, sheetClipboardObjectAxes(after, sheet))
  reconcileWorkbookImageGeometryChanges(before, after)
  close(drawingAffineFrame(image(after, '1').drawingGeometry), expected)
  for (const key of ['a', 'b', 'c', 'd']) assert.equal(image(after, '1').drawingGeometry.matrix[key], source.drawingGeometry.matrix[key])
  assert.deepEqual(authors(image(after, '1')), authors(source))
})

for (const [name, mutate] of [
  ['computed auto-height', sheet => { sheet.rowAutoHeightSpans = [{ start: 0, end: 0, value: true }]; sheet.rowAutoHeights = { 0: 45 } }],
  ['sheet filtering', sheet => { sheet.autoFilter = { range: { startRow: 0, startCol: 0, endRow: 5, endCol: 1 }, columns: {}, filteredRows: [1] } }],
  ['collapsed outline', sheet => { sheet.outline = { summaryAbove: false, summaryLeft: false,
    groups: [{ id: 'grouped-rows', axis: 'row', start: 1, end: 1, collapsed: true }] } }],
]) test(`${name} participates in the final image geometry boundary`, () => {
  const before = seed(), after = structuredClone(before); mutate(after.sheets[0])
  const oldAxes = sheetClipboardObjectAxes(before, before.sheets[0]), newAxes = sheetClipboardObjectAxes(after, after.sheets[0]), source = image(before, '0')
  const shift = newAxes.rows.offset(source.anchor.row) - oldAxes.rows.offset(source.anchor.row)
  assert.notEqual(shift, 0)
  reconcileWorkbookImageGeometryChanges(before, after)
  close(image(after, '0').drawingGeometry.matrix, { e: source.drawingGeometry.matrix.e, f: source.drawingGeometry.matrix.f + shift })
})

test('table filtering uses the same final sparse axes even when the sheet filter itself is unchanged', () => {
  const before = seed(); before.tables = [{ id: 'table', sheetId: 'source', name: 'Table', range: { startRow: 0, startCol: 0, endRow: 5, endCol: 1 }, columns: [] }]
  const after = structuredClone(before); after.tables[0].autoFilter = { columns: {}, filteredRows: [1] }
  reconcileWorkbookImageGeometryChanges(before, after)
  assert.equal(image(after, '0').drawingGeometry.matrix.f, image(before, '0').drawingGeometry.matrix.f - 20)
})

test('explicit or structural placement already changed in the candidate is not projected a second time', () => {
  const before = seed(), after = structuredClone(before), sheet = after.sheets[0], source = image(before, '0')
  sheet.rowHeights = { 0: 40 }
  sheet.images = transformSheetImagesForAxisEdit(sheet.images, { axis: 'row', kind: 'insert', start: 0, count: 1 }, sheet,
    { before: sheetClipboardObjectAxes(before, before.sheets[0]), after: sheetClipboardObjectAxes(after, sheet) })
  const structural = structuredClone(sheet.images)
  reconcileWorkbookImageGeometryChanges(before, after); assert.deepEqual(sheet.images, structural)
  const explicit = structuredClone(before); explicit.sheets[0].rowHeights = { 0: 40 }; image(explicit, '0').drawingGeometry.matrix.e += 75
  const authored = structuredClone(image(explicit, '0')); reconcileWorkbookImageGeometryChanges(before, explicit)
  assert.deepEqual(image(explicit, '0'), authored); assert.notEqual(authored.drawingGeometry.matrix.e, source.drawingGeometry.matrix.e)
})

test('collapsed Both preserves its reversible author geometry, has no available hit owner and restores when rows are shown', () => {
  const before = seed(), hidden = structuredClone(before); hidden.sheets[0].rowHiddenSpans = [{ start: 0, end: 10, value: true }]
  reconcileWorkbookDrawingStates(before, hidden)
  const hiddenObject = createSheetDrawingScene(hidden, 'source').objects.find(object => object.ref.id === 'image-1')
  assert.equal(hiddenObject.available, false); assert.deepEqual(image(hidden, '1').drawingGeometry, image(before, '1').drawingGeometry)
  const shown = structuredClone(hidden); shown.sheets[0].rowHiddenSpans = []
  reconcileWorkbookDrawingStates(hidden, shown)
  const restored = createSheetDrawingScene(shown, 'source').objects.find(object => object.ref.id === 'image-1')
  assert.notEqual(restored.available, false); close(drawingAffineFrame(restored.geometry), drawingAffineFrame(image(before, '1').drawingGeometry))
  assert.deepEqual(authors(image(shown, '1')), authors(image(before, '1')))
})

test('parent groups fit the final leaf geometry after automatic dimensions, preserving child order', () => {
  const before = seed(), sheet = before.sheets[0]
  sheet.cells['0:0'] = createCell(0, 0, 'Label'); sheet.cells['0:1'] = createCell(0, 1, 'Value')
  sheet.cells['1:0'] = createCell(1, 0, 'A'); sheet.cells['1:1'] = createCell(1, 1, 1)
  before.charts = [createChartDefinition(before, 'source', { type: 'line', title: 'Chart',
    source: { sheetId: 'source', orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 } },
    position: { x: 64, y: 48 }, size: { width: 200, height: 100 }, valueFields: [1] }, 'chart')]
  initializeWorkbookDrawingStates(before)
  const plan = planDrawingGroup(createSheetDrawingScene(before, 'source'), { selection: [ref('image', 'image-0'), ref('chart', 'chart')], groupId: 'g' })
  assert.equal(plan.status, 'ready'); applySheetDrawingTransaction(before, 'source', plan.transaction)
  const after = structuredClone(before); after.sheets[0].columnWidths = { 0: 200 }; after.sheets[0].rowHeights = { 0: 50 }
  reconcileWorkbookDrawingStates(before, after)
  const finalScene = createSheetDrawingScene(after, 'source'), group = finalScene.state.groups.find(group => group.id === 'g'), prior = before.sheets[0].drawingState.groups.find(group => group.id === 'g')
  const geometries = group.members.map(ref => finalScene.objects.find(object => object.ref.kind === ref.kind && object.ref.id === ref.id).geometry)
  assert.deepEqual(group.frame, drawingGroupFrame(geometries, prior.frame)); assert.deepEqual(group.members, prior.members)
  assert.notDeepEqual(group.frame, prior.frame)
})

test('unrelated edits do not clone or alter image placements and per-sheet axes stay isolated', () => {
  const before = seed(), after = structuredClone(before), retained = after.sheets[0].images
  after.sheets[0].metadata = { note: 'unrelated' }; after.sheets[1].rowHeights = { 0: 100 }
  reconcileWorkbookImageGeometryChanges(before, after)
  assert.equal(after.sheets[0].images, retained); assert.deepEqual(retained, before.sheets[0].images)
})
