import assert from 'node:assert/strict'
import { test } from 'vitest'
import { drawingAffineBounds } from '../src/domain/drawing-affine.ts'
import { drawingGroupFrame, planDrawingGroup } from '../src/features/drawing-groups/index.ts'
import { applySheetDrawingTransaction, createSheetDrawingScene } from '../src/features/sheet-drawings/sheet-drawing-host.ts'
import { reconcileWorkbookDrawingStates } from '../src/features/sheet-drawings/sheet-drawing-model-lifecycle.ts'
import { reconcileWorkbookSlicerGeometryChanges } from '../src/features/sheet-drawings/sheet-drawing-slicer-geometry.ts'
import { sheetClipboardObjectAxes } from '../src/features/sheet-workbench/sheet-clipboard-object-geometry.ts'
import { transformWorkbookCellShiftResources } from '../src/features/sheet-workbench/sheet-cell-shift-resources.ts'
import { resolveSlicerAnchorBounds, resolveSlicerBounds } from '../src/features/slicer/slicer-drawing.ts'
import { updateSlicerPresentation } from '../src/features/slicer/slicer-engine.ts'
import { GRID_HEADER_WIDTH, GRID_HEADER_HEIGHT } from '../src/rendering/spreadsheet/layout.ts'
import { slicerDrawingAuthor, slicerDrawingWorkbook } from './fixtures/slicer-drawing-model-fixture.mjs'

// Deferred final-axis/structure contracts; AABB anchoring is a local contract, not verified Office major-axis behavior.
const close = (actual, expected) => { for (const key of Object.keys(expected)) assert.ok(Math.abs(actual[key] - expected[key]) < 1e-8, `${key}: ${actual[key]} / ${expected[key]}`) }
const control = (workbook, id = 's') => workbook.slicers.find(control => control.id === id)
const bounds = (workbook, id = 's') => resolveSlicerBounds(control(workbook, id), sheetClipboardObjectAxes(workbook, workbook.sheets[0]))
function seed(anchorType = 'position', affine = true) {
  const workbook = slicerDrawingWorkbook({ affine })
  for (const slicer of workbook.slicers) updateSlicerPresentation(slicer, { anchorType }, workbook)
  return workbook
}

test('Position translates both control matrices once after final row and column dimensions while retaining every author field', () => {
  const before = seed(), untouched = structuredClone(before), after = structuredClone(before)
  after.sheets[0].rowHeights = { 0: 32 }; after.sheets[0].columnWidths = { 0: 150 }
  reconcileWorkbookSlicerGeometryChanges(before, after)
  for (const id of ['s', 't']) {
    const original = control(before, id), next = control(after, id)
    close(next.drawingGeometry.matrix, { ...original.drawingGeometry.matrix, e: original.drawingGeometry.matrix.e + 50, f: original.drawingGeometry.matrix.f + 12 })
    assert.equal(next.revision, original.revision + 1); assert.deepEqual(slicerDrawingAuthor(next), slicerDrawingAuthor(original))
  }
  const projected = structuredClone(after)
  reconcileWorkbookSlicerGeometryChanges(before, after); assert.deepEqual(after, projected); assert.deepEqual(before, untouched)
})

test('Both world AABB projection retains reflection/shear and the local control extent', () => {
  const before = seed('both'), original = control(before)
  updateSlicerPresentation(original, { drawingGeometry: { ...original.drawingGeometry,
    matrix: { a: -1, b: .25, c: .2, d: 1, e: 520, f: 80 } } }, before)
  const after = structuredClone(before)
  after.sheets[0].rowHeights = { 4: 40 }; after.sheets[0].columnWidths = { 2: 180 }
  const target = resolveSlicerAnchorBounds(control(after), sheetClipboardObjectAxes(after, after.sheets[0]))
  reconcileWorkbookSlicerGeometryChanges(before, after)
  const next = control(after), m = original.drawingGeometry.matrix, expected = drawingAffineBounds(original.drawingGeometry)
  close(drawingAffineBounds(next.drawingGeometry), { ...target, left: target.left - GRID_HEADER_WIDTH, top: target.top - GRID_HEADER_HEIGHT })
  close(next.drawingGeometry.matrix, { a: m.a * target.width / expected.width, c: m.c * target.width / expected.width,
    b: m.b * target.height / expected.height, d: m.d * target.height / expected.height })
  assert.equal(next.drawingGeometry.width, original.drawingGeometry.width); assert.equal(next.drawingGeometry.height, original.drawingGeometry.height)
  assert.deepEqual(slicerDrawingAuthor(next), slicerDrawingAuthor(original))
})

test('None control matrices and placement models ignore dimension and filter projection', () => {
  const before = seed('none'), after = structuredClone(before)
  after.sheets[0].rowHeights = { 0: 70 }; after.sheets[0].columnWidths = { 0: 180 }
  after.tables[0].autoFilter = { columns: {}, filteredRows: [1, 2] }
  reconcileWorkbookSlicerGeometryChanges(before, after); assert.deepEqual(after.slicers, before.slicers)
})

for (const [name, mutate] of [
  ['computed auto-height', workbook => { const sheet = workbook.sheets[0]; sheet.rowAutoHeightSpans = [{ start: 0, end: 0, value: true }]; sheet.rowAutoHeights = { 0: 45 } }],
  ['sheet filter', workbook => { workbook.sheets[0].autoFilter = { range: { startRow: 0, startCol: 0, endRow: 3, endCol: 1 }, columns: {}, filteredRows: [1] } }],
  ['Table filter', workbook => { workbook.tables[0].autoFilter = { columns: {}, filteredRows: [1] } }],
  ['outline collapse', workbook => { workbook.sheets[0].outline = { summaryAbove: false, summaryLeft: false, groups: [{ id: 'rows', axis: 'row', start: 1, end: 1, collapsed: true }] } }],
]) test(`${name} enters the final slicer/timeline projection after visible axes are resolved`, () => {
  const before = seed(), after = structuredClone(before); mutate(after)
  const priorAxes = sheetClipboardObjectAxes(before, before.sheets[0]), finalAxes = sheetClipboardObjectAxes(after, after.sheets[0])
  reconcileWorkbookSlicerGeometryChanges(before, after)
  for (const id of ['s', 't']) {
    const original = control(before, id), row = original.drawing.from.row
    const delta = finalAxes.rows.offset(row) - priorAxes.rows.offset(row)
    assert.notEqual(delta, 0); close(control(after, id).drawingGeometry.matrix, { f: original.drawingGeometry.matrix.f + delta })
  }
})

test('collapsed Both retains positive authored models, removes hit owners and recovers the original basis on show', () => {
  const before = seed('both'), hidden = structuredClone(before)
  hidden.sheets[0].rowHiddenSpans = [{ start: 0, end: 30, value: true }]
  reconcileWorkbookDrawingStates(before, hidden)
  for (const id of ['s', 't']) {
    const model = control(hidden, id), owner = createSheetDrawingScene(hidden, 'data').objects.find(object => object.ref.id === id)
    assert.equal(owner.available, false); assert.deepEqual(model.drawingGeometry, control(before, id).drawingGeometry)
    assert.ok(model.layout.width > 0 && model.layout.height > 0); assert.equal(bounds(hidden, id).height, 0)
    updateSlicerPresentation(model, { caption: 'Still editable settings' }, hidden)
  }
  const shown = structuredClone(hidden); shown.sheets[0].rowHiddenSpans = []
  reconcileWorkbookDrawingStates(hidden, shown)
  for (const id of ['s', 't']) {
    assert.notEqual(createSheetDrawingScene(shown, 'data').objects.find(object => object.ref.id === id).available, false)
    close(control(shown, id).drawingGeometry.matrix, control(before, id).drawingGeometry.matrix)
  }
})

test('an explicit new control matrix in the candidate is authoritative and is not translated a second time', () => {
  const before = seed(), after = structuredClone(before)
  after.sheets[0].rowHeights = { 0: 50 }; control(after).drawingGeometry.matrix.f += 75
  const authored = structuredClone(control(after)); reconcileWorkbookSlicerGeometryChanges(before, after)
  assert.deepEqual(control(after), authored)
})

test('cell markers moved by a structural edit still project when the author matrix itself has not changed', () => {
  const before = seed(), after = structuredClone(before), model = control(after), original = control(before)
  model.drawing.from.row++; model.drawing.to.row++
  reconcileWorkbookSlicerGeometryChanges(before, after)
  close(model.drawingGeometry.matrix, { f: original.drawingGeometry.matrix.f + 20 }); assert.equal(model.layout.top, original.layout.top + 20)
})

test('legacy Both collapse during a local delete keeps a positive settings layout instead of publishing zero dimensions', () => {
  const before = seed('both', false), slicer = control(before)
  slicer.drawing.from = { row: 5, column: 2, offsetX: 0, offsetY: 0 }
  slicer.drawing.to = { row: 7, column: 2, offsetX: 80, offsetY: 0 }
  const after = structuredClone(before), authored = structuredClone(slicer.layout)
  transformWorkbookCellShiftResources(after, before, 'data', { kind: 'delete', axis: 'row', start: 5, count: 3, bandStart: 2, bandEnd: 2 })
  assert.equal(bounds(after).height, 0); assert.deepEqual(control(after).layout, authored)
  updateSlicerPresentation(control(after), { caption: 'Valid collapsed settings' }, after)
  assert.equal(control(after).caption, 'Valid collapsed settings')
})

test('parent fitting runs after final slicer geometry and preserves the complete member order', () => {
  const before = seed(), plan = planDrawingGroup(createSheetDrawingScene(before, 'data'), {
    selection: [{ kind: 'image', id: 'a' }, { kind: 'slicer', id: 's' }, { kind: 'timeline', id: 't' }], groupId: 'g',
  })
  assert.equal(plan.status, 'ready'); applySheetDrawingTransaction(before, 'data', plan.transaction)
  const after = structuredClone(before); after.sheets[0].rowHeights = { 0: 60 }
  reconcileWorkbookDrawingStates(before, after)
  const input = createSheetDrawingScene(after, 'data'), actual = input.state.groups.find(group => group.id === 'g'), prior = before.sheets[0].drawingState.groups.find(group => group.id === 'g')
  const geometries = actual.members.map(ref => input.objects.find(object => object.ref.kind === ref.kind && object.ref.id === ref.id).geometry)
  assert.deepEqual(actual.frame, drawingGroupFrame(geometries, prior.frame)); assert.deepEqual(actual.members, prior.members)
})

test('other-sheet dimensions and unrelated metadata do not rewrite control placements or revisions', () => {
  const before = seed(), after = structuredClone(before)
  after.sheets[1].rowHeights = { 0: 100 }; after.sheets[0].metadata = { unrelated: 'metadata' }
  const slicers = after.slicers; reconcileWorkbookSlicerGeometryChanges(before, after)
  assert.equal(after.slicers, slicers); assert.deepEqual(after.slicers, before.slicers)
})
