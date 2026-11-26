import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankSheet } from '../src/domain/model.ts'
import { drawingAffineBounds } from '../src/domain/drawing-affine.ts'
import { captureSheetClipboard, createSheetClipboardPastePlan, applySheetClipboardPastePlan, assertSheetClipboardSnapshot } from '../src/features/sheet-workbench/sheet-clipboard.ts'
import { assertSheetClipboardWorkbookResources } from '../src/features/sheet-workbench/sheet-clipboard-resource-validation.ts'
import { sheetClipboardObjectAxes, sheetClipboardRangeBounds } from '../src/features/sheet-workbench/sheet-clipboard-object-geometry.ts'
import { chartDrawingWorldBounds } from '../src/features/charts/chart-drawing-affine.ts'
import { resolveSlicerBounds } from '../src/features/slicer/slicer-drawing.ts'
import { updateSlicerPresentation } from '../src/features/slicer/slicer-engine.ts'
import { slicerDrawingWorkbook } from './fixtures/slicer-drawing-model-fixture.mjs'

// Deferred range-resource contracts. These do not exercise the OS clipboard or Office binary IO.
const sourceRange = { startRow: 0, endRow: 25, startCol: 0, endCol: 8 }
const targetCell = { startRow: 30, endRow: 30, startCol: 10, endCol: 10 }
const context = { userId: 'local-user', region: 'en-US' }
const close = (actual, expected) => { for (const key of Object.keys(expected)) assert.ok(Math.abs(actual[key] - expected[key]) < 1e-8, `${key}: ${actual[key]} / ${expected[key]}`) }
function setup() {
  const workbook = slicerDrawingWorkbook(), destination = createBlankSheet('Paste')
  destination.id = 'paste'; destination.rowCount = 100; destination.columnCount = 30
  destination.defaultRowHeight = 20; destination.defaultColumnWidth = 100; destination.defaultRowAutoHeight = false
  workbook.sheets.push(destination)
  workbook.charts[0].drawingGeometry = { width: 30, height: 20, matrix: { a: -1, b: .3, c: .2, d: 1, e: 160, f: 60 } }
  return workbook
}
const target = workbook => ({ unitId: workbook.id, sheetId: 'paste', range: targetCell })
function capture(workbook, copyType = 'copy', range = sourceRange) {
  return captureSheetClipboard(workbook, { unitId: workbook.id, sheetId: 'data', range }, { copyType, userId: context.userId })
}
function deltaFor(workbook, snapshot, result) {
  const destination = workbook.sheets.find(sheet => sheet.id === 'paste'), bounds = sheetClipboardRangeBounds(workbook, destination, result.ranges[0])
  return { x: bounds.left - snapshot.sourceOrigin.left, y: bounds.top - snapshot.sourceOrigin.top }
}

test('range capture retains complete chart/slicer affine authors and resolved control AABBs', () => {
  const workbook = setup(), before = structuredClone(workbook), snapshot = capture(workbook)
  assertSheetClipboardSnapshot(snapshot)
  assert.deepEqual(snapshot.workbookResources.charts.map(chart => chart.id), ['c'])
  assert.deepEqual(snapshot.workbookResources.slicers.map(slicer => slicer.id), ['s', 't'])
  assert.deepEqual(snapshot.workbookResources.charts[0].drawingGeometry, workbook.charts[0].drawingGeometry)
  for (const slicer of snapshot.workbookResources.slicers) {
    const original = workbook.slicers.find(model => model.id === slicer.id)
    assert.deepEqual(slicer.drawingGeometry, original.drawingGeometry)
    close(slicer.layout, resolveSlicerBounds(original, sheetClipboardObjectAxes(workbook, workbook.sheets[0])))
  }
  snapshot.workbookResources.slicers[0].drawingGeometry.matrix.a = 99
  assert.deepEqual(workbook, before)
})

for (const copyType of ['copy', 'cut']) test(`${copyType} range placement translates full chart and control matrices once and preserves native bindings`, () => {
  const workbook = setup(), before = structuredClone(workbook), snapshot = capture(workbook, copyType)
  const snapshotBefore = structuredClone(snapshot)
  const plan = createSheetClipboardPastePlan(workbook, target(workbook), snapshot, 'all', context)
  assert.equal(plan.result.accepted, true); assert.equal(plan.result.changed, true); assert.deepEqual(workbook, before)
  const movement = deltaFor(workbook, snapshot, plan.result), after = structuredClone(workbook)
  applySheetClipboardPastePlan(after, plan)
  const charts = after.charts.filter(chart => chart.sheetId === 'paste'), slicers = after.slicers.filter(slicer => slicer.hostSheetId === 'paste')
  assert.equal(charts.length, 1); assert.equal(slicers.length, 2)
  const chart = charts[0], originalChart = before.charts[0]
  close(chart.drawingGeometry.matrix, { ...originalChart.drawingGeometry.matrix,
    e: originalChart.drawingGeometry.matrix.e + movement.x, f: originalChart.drawingGeometry.matrix.f + movement.y })
  assert.deepEqual(chart.options, originalChart.options); assert.deepEqual(chart.size, originalChart.size)
  close(chartDrawingWorldBounds(chart), { ...chartDrawingWorldBounds(originalChart),
    left: chartDrawingWorldBounds(originalChart).left + movement.x, top: chartDrawingWorldBounds(originalChart).top + movement.y })
  assert.equal(chart.source.sheetId, 'paste')
  for (const next of slicers) {
    const original = before.slicers.find(model => model.type === next.type), m = original.drawingGeometry.matrix
    close(next.drawingGeometry.matrix, { ...m, e: m.e + movement.x, f: m.f + movement.y })
    assert.equal(next.drawingGeometry.width, original.drawingGeometry.width); assert.equal(next.drawingGeometry.height, original.drawingGeometry.height)
    assert.equal(next.caption, original.caption); assert.equal(next.style, original.style); assert.deepEqual(next.timelineRange, original.timelineRange)
    const projected = drawingAffineBounds(next.drawingGeometry)
    close(next.layout, { left: projected.left + 44, top: projected.top + 26, width: projected.width, height: projected.height })
    if (next.type === 'slicer') {
      const table = after.tables.find(table => table.id === next.tableId)
      assert.equal(table.sheetId, 'paste'); assert.ok(table.columns.some(column => column.id === next.sourceFieldId))
    } else assert.equal(next.pivotTableId, original.pivotTableId)
    assert.equal(next.unitId, after.id)
    assert.equal(next.id === original.id, copyType === 'cut')
  }
  assert.equal(after.tables.find(table => table.sheetId === 'paste').id === before.tables[0].id, copyType === 'cut')
  assert.deepEqual(after.pivotTables, before.pivotTables)
  assert.deepEqual(snapshot, snapshotBefore)
})

test('copying a region that contains only the old chart rectangle excludes its larger transformed footprint', () => {
  const workbook = setup(), chart = workbook.charts[0]
  chart.drawingGeometry = { width: 30, height: 20, matrix: { a: 1, b: .3, c: .8, d: 1, e: 180, f: 40 } }
  const snapshot = capture(workbook, 'copy', { startRow: 0, endRow: 5, startCol: 0, endCol: 1 })
  assert.deepEqual(snapshot.workbookResources.charts, [])
  assert.deepEqual(workbook.charts[0], chart)
})

test('collapsed Both controls remain in the source model and do not enter captured visible drawing resources', () => {
  const workbook = setup()
  for (const slicer of workbook.slicers) updateSlicerPresentation(slicer, { anchorType: 'both' }, workbook)
  workbook.sheets[0].rowHiddenSpans = [{ start: 3, end: 30, value: true }]
  const before = structuredClone(workbook), snapshot = capture(workbook)
  assert.deepEqual(snapshot.workbookResources.slicers, [])
  assert.equal(workbook.slicers.length, 2); assert.deepEqual(workbook, before)
  assert.ok(workbook.slicers.every(slicer => slicer.layout.width > 0 && slicer.layout.height > 0))
})

for (const [name, corrupt] of [
  ['singular matrix', snapshot => { snapshot.workbookResources.slicers[0].drawingGeometry.matrix = { a: 1, b: 1, c: 1, d: 1, e: 100, f: 100 } }],
  ['non-finite matrix', snapshot => { snapshot.workbookResources.charts[0].drawingGeometry.matrix.a = Infinity }],
  ['zero local extent', snapshot => { snapshot.workbookResources.slicers[1].drawingGeometry.width = 0 }],
  ['falsely small layout', snapshot => { snapshot.workbookResources.slicers[0].layout.width = 1 }],
  ['forged layout origin', snapshot => { snapshot.workbookResources.slicers[0].layout.left += 100 }],
  ['partially selected affine chart', snapshot => { snapshot.workbookResources.charts[0].drawingGeometry.matrix.e = snapshot.sourceBounds.left + snapshot.sourceBounds.width }],
]) test(`untrusted clipboard rejects ${name} before planning changes`, () => {
  const workbook = setup(), before = structuredClone(workbook), snapshot = capture(workbook); corrupt(snapshot)
  assert.throws(() => createSheetClipboardPastePlan(workbook, target(workbook), snapshot, 'all', context))
  assert.deepEqual(workbook, before)
})

test('affine source validation rejects a control layout that was forged to hide its true containment', () => {
  const workbook = setup(), snapshot = capture(workbook), slicer = snapshot.workbookResources.slicers[0]
  slicer.drawingGeometry.matrix.e = 1000
  slicer.layout = { ...slicer.layout, left: snapshot.sourceBounds.left, top: snapshot.sourceBounds.top, width: 1, height: 1 }
  assert.throws(() => assertSheetClipboardWorkbookResources(snapshot.workbookResources, {
    source: snapshot.source, rows: snapshot.rows, columns: snapshot.columns, sourceOrigin: snapshot.sourceOrigin, sourceBounds: snapshot.sourceBounds,
  }), /布局|几何|affine|仿射/)
})

for (const mode of ['values', 'formulas', 'formats', 'columnWidths']) test(`${mode} paste leaves native drawing matrices and resource bindings untouched`, () => {
  const workbook = setup(), snapshot = capture(workbook), plan = createSheetClipboardPastePlan(workbook, target(workbook), snapshot, mode, context)
  const after = structuredClone(workbook); applySheetClipboardPastePlan(after, plan)
  assert.deepEqual(after.charts, workbook.charts); assert.deepEqual(after.slicers, workbook.slicers)
  assert.deepEqual(after.tables, workbook.tables); assert.deepEqual(after.pivotTables, workbook.pivotTables)
})

test('an unresolved cross-workbook timeline dependency rejects the detached paste instead of retaining a foreign pivot ID', () => {
  const original = setup(), snapshot = capture(original), targetBook = setup(); targetBook.id = 'foreign'
  for (const control of targetBook.slicers) control.unitId = 'foreign'
  for (const pivot of targetBook.pivotTables) { pivot.source.unitId = 'foreign'; pivot.anchor.unitId = 'foreign' }
  const before = structuredClone(targetBook)
  assert.throws(() => createSheetClipboardPastePlan(targetBook, target(targetBook), snapshot, 'all', context), /透视数据源/)
  assert.deepEqual(targetBook, before)
})
