import assert from 'node:assert/strict'
import { test } from 'vitest'
import { drawingGeometryBounds, planDrawingArrange, planDrawingGroup, planDrawingGroupTransform,
  planDrawingSelectionTransform } from '../src/features/drawing-groups/index.ts'
import { applySheetDrawingTransaction, createSheetDrawingScene, removeSheetDrawings,
  sheetDrawingTransactionAffectedRanges } from '../src/features/sheet-drawings/sheet-drawing-host.ts'
import { sheetImageNormalizedCrop } from '../src/features/sheet-images/sheet-image-placement.ts'
import { CHART_SHEET_METADATA_KEY } from '../src/features/charts/chart-object-commands.ts'
import { drawingFrame, drawingRef as ref, sheetDrawingWorkbook } from './fixtures/sheet-drawing-model-fixture.mjs'

// Native host drafts only. No unit/type/build/browser/Office/WASM/performance verification was run.
const ready = result => { assert.equal(result.status, 'ready', JSON.stringify(result)); return result.transaction }
const snapshot = value => JSON.parse(JSON.stringify(value))
const code = expected => error => error.code === expected
const rangeKey = range => [range.startRow, range.startCol, range.endRow, range.endCol].join(':')
function group(workbook, selection = [ref('image', 'a'), ref('chart', 'c')], groupId = 'g') {
  const transaction = ready(planDrawingGroup(createSheetDrawingScene(workbook, 'data'), { selection, groupId }))
  applySheetDrawingTransaction(workbook, 'data', transaction)
  return createSheetDrawingScene(workbook, 'data')
}

test('the host derives header-free image/chart/shape/slicer geometry from native models and actual axes without mutating them', () => {
  const workbook = sheetDrawingWorkbook({ slicer: true }), before = snapshot(workbook), scene = createSheetDrawingScene(workbook, 'data')
  const bounds = (kind, id) => drawingGeometryBounds(scene.objects.find(item => item.ref.kind === kind && item.ref.id === id).geometry)
  assert.deepEqual(bounds('image', 'a'), { left: 10, top: 20, width: 20, height: 10 })
  assert.deepEqual(bounds('chart', 'c'), { left: 50, top: 40, width: 30, height: 20 })
  assert.deepEqual(bounds('shape', 'b'), { left: 30, top: 20, width: 20, height: 10 })
  assert.deepEqual(bounds('slicer', 's'), { left: 120, top: 80, width: 80, height: 40 })
  assert.equal(scene.objects.find(item => item.ref.id === 's').transformable, true)
  assert.equal(scene.objects.find(item => item.ref.id === 's').groupable, true); assert.deepEqual(snapshot(workbook), before)
})

test('mixed group shear publishes native geometry while retaining image resource/crop/flip, chart source/options/size and shape authors', () => {
  const workbook = sheetDrawingWorkbook(), before = snapshot(workbook)
  const scene = group(workbook, [ref('image', 'a'), ref('shape', 'b'), ref('chart', 'c')])
  const transaction = ready(planDrawingGroupTransform(scene, { groupId: 'g', delta: { a: 1, b: 0.5, c: 0.25, d: 1, e: 3, f: -2 } }))
  applySheetDrawingTransaction(workbook, 'data', transaction)
  const image = workbook.sheets[0].images[0], chart = workbook.charts[0], shape = workbook.sheets[0].shapes[0]
  assert.deepEqual(image.drawingGeometry.matrix, { a: 1, b: 0.5, c: 0.25, d: 1, e: 18, f: 23 })
  assert.deepEqual(image.resource, before.sheets[0].images[0].resource); assert.equal(image.flipX, true); assert.equal(image.flipY, false)
  assert.equal(image.opacity, 0.8); assert.equal(image.altText, 'Keep original'); assert.equal(image.title, 'Image')
  const actualCrop = sheetImageNormalizedCrop(image.crop, image.size), expectedCrop = sheetImageNormalizedCrop(before.sheets[0].images[0].crop, before.sheets[0].images[0].size)
  for (const key of ['x', 'y', 'width', 'height']) assert.ok(Math.abs(actualCrop[key] - expectedCrop[key]) < 1e-9)
  assert.deepEqual(chart.source, before.charts[0].source); assert.deepEqual(chart.options, before.charts[0].options)
  assert.deepEqual(chart.size, before.charts[0].size); assert.equal(chart.revision, before.charts[0].revision + 1)
  assert.deepEqual(shape.source, before.sheets[0].shapes[0].source); assert.deepEqual(shape.style, before.sheets[0].shapes[0].style)
  assert.deepEqual(shape.text, before.sheets[0].shapes[0].text); assert.deepEqual(workbook.sheets[0].images[1], before.sheets[0].images[1])
})

test('native forward/backward/forward restores the original guarded chart revision and reapplies the same transform', () => {
  const workbook = sheetDrawingWorkbook(), before = group(workbook), originalRevision = workbook.charts[0].revision
  const transaction = ready(planDrawingGroupTransform(before, { groupId: 'g', frame: drawingFrame(100, 200, 140, 120) }))
  applySheetDrawingTransaction(workbook, 'data', transaction)
  const once = createSheetDrawingScene(workbook, 'data')
  assert.equal(workbook.charts[0].revision, originalRevision + 1)
  applySheetDrawingTransaction(workbook, 'data', transaction, 'backward')
  assert.equal(workbook.charts[0].revision, originalRevision)
  assert.deepEqual(createSheetDrawingScene(workbook, 'data').state, before.state)
  applySheetDrawingTransaction(workbook, 'data', transaction)
  assert.deepEqual(createSheetDrawingScene(workbook, 'data'), once); assert.equal(workbook.charts[0].revision, originalRevision + 1)
})

test('a newer native chart revision blocks inverse publication without partially writing image, shape or group state', () => {
  const workbook = sheetDrawingWorkbook(), before = group(workbook, [ref('image', 'a'), ref('shape', 'b'), ref('chart', 'c')])
  const transaction = ready(planDrawingGroupTransform(before, { groupId: 'g', delta: { a: 1, b: 0, c: 0, d: 1, e: 25, f: 0 } }))
  applySheetDrawingTransaction(workbook, 'data', transaction)
  workbook.charts[0].revision++; workbook.charts[0].title = 'Newer author'
  const current = snapshot(workbook)
  assert.throws(() => applySheetDrawingTransaction(workbook, 'data', transaction, 'backward'), code('stale-object'))
  assert.deepEqual(snapshot(workbook), current)
})

test('hidden rows/columns may collapse a legacy Both image; its unavailable state placeholder preserves original anchors and resource', () => {
  const workbook = sheetDrawingWorkbook(), sheet = workbook.sheets[0], image = sheet.images[0]
  image.anchorType = '1'; image.anchor = { row: 0, column: 0, offsetX: 0, offsetY: 0 }; image.endAnchor = { row: 1, column: 1, offsetX: 0, offsetY: 0 }
  delete image.absolutePosition; sheet.hiddenRows = [0]; sheet.hiddenColumns = [0]
  const before = snapshot(image), scene = createSheetDrawingScene(workbook, 'data', { available: () => true }), owner = scene.objects.find(item => item.ref.id === 'a')
  assert.equal(owner.available, false); assert.deepEqual(drawingGeometryBounds(owner.geometry), { left: 0, top: 0, width: 1, height: 1 })
  assert.deepEqual(snapshot(image), before)
  assert.equal(planDrawingGroup(scene, { selection: [ref('image', 'a'), ref('chart', 'c')], groupId: 'bad' }).reason, 'unavailable-object')
  assert.equal(planDrawingSelectionTransform(scene, { selection: [ref('image', 'a')], delta: { a: 1, b: 0, c: 0, d: 1, e: 10, f: 0 } }).reason, 'unavailable-object')
  applySheetDrawingTransaction(workbook, 'data', ready(planDrawingArrange(scene, { selection: [ref('image', 'a')], arrangeType: 'front' })))
  assert.deepEqual(snapshot(image), before)
  sheet.hiddenRows = []; sheet.hiddenColumns = []
  const visible = createSheetDrawingScene(workbook, 'data').objects.find(item => item.ref.id === 'a')
  assert.equal(visible.available, true); assert.deepEqual(drawingGeometryBounds(visible.geometry), { left: 0, top: 0, width: 100, height: 20 })
})

test('permissions use both old and new world footprints with one header conversion and the real row/column indexes', () => {
  const workbook = sheetDrawingWorkbook(), scene = createSheetDrawingScene(workbook, 'data')
  const transaction = ready(planDrawingSelectionTransform(scene, { selection: [ref('image', 'a')], delta: { a: 1, b: 0, c: 0, d: 1, e: 205, f: 45 } }))
  assert.deepEqual(new Set(sheetDrawingTransactionAffectedRanges(workbook, 'data', transaction).map(rangeKey)), new Set(['1:0:1:0', '3:2:3:2']))
  const before = snapshot(workbook)
  assert.throws(() => applySheetDrawingTransaction(workbook, 'data', transaction, 'forward', { canEdit: () => false }), code('read-only'))
  assert.deepEqual(snapshot(workbook), before)
})

test('Delete expands a mixed group and selected slicer, retains the source table/cells and keeps another owner of the same image resource', () => {
  const workbook = sheetDrawingWorkbook({ slicer: true }), before = snapshot(workbook)
  group(workbook, [ref('image', 'a'), ref('shape', 'b'), ref('chart', 'c')])
  const result = removeSheetDrawings(workbook, 'data', [ref('group', 'g'), ref('slicer', 's')])
  assert.equal(workbook.charts.length, 0); assert.equal(workbook.slicers.length, 0); assert.equal(workbook.sheets[0].shapes.length, 0)
  assert.deepEqual(workbook.tables, before.tables); assert.deepEqual(workbook.sheets[0].cells, before.sheets[0].cells)
  assert.deepEqual(workbook.sheets[0].images, [before.sheets[0].images[1]]); assert.deepEqual(result.selection, [])
  assert.deepEqual(result.scene.state.groups, []); assert.deepEqual(result.scene.state.rootOrder, [ref('image', 'd')])
  assert.deepEqual(new Set(result.removedRefs.map(item => item.kind + ':' + item.id)), new Set(['group:g', 'image:a', 'shape:b', 'chart:c', 'slicer:s']))
})

test('a primary chart blocks deletion of its containing group atomically; empty Delete does not initialize or write drawingState', () => {
  const empty = sheetDrawingWorkbook(), untouched = snapshot(empty)
  assert.deepEqual(removeSheetDrawings(empty, 'data', []).removedRefs, []); assert.deepEqual(snapshot(empty), untouched)
  assert.equal(empty.sheets[0].drawingState, undefined)
  const workbook = sheetDrawingWorkbook(); group(workbook)
  workbook.sheets[0].metadata = { [CHART_SHEET_METADATA_KEY]: { version: 1, primaryChartId: 'c' } }
  const before = snapshot(workbook)
  assert.throws(() => removeSheetDrawings(workbook, 'data', [ref('group', 'g')]), code('read-only'))
  assert.deepEqual(snapshot(workbook), before)
})

test('group deletion shares native chart resource ownership: another chart retains its mask, and removing the final owner prunes only the mask registry', () => {
  const workbook = sheetDrawingWorkbook(), resource = { id: 'mask', sourceType: 'url', source: 'https://example.invalid/mask.png', mimeType: 'image/png', sizeBytes: 8 }
  workbook.charts[0].type = 'wordCloud'; workbook.charts[0].options.wordCloud = { maskImageId: 'mask' }
  workbook.charts.push({ ...snapshot(workbook.charts[0]), id: 'other-chart' })
  workbook.metadata = { keep: 'other metadata', chartImageResources: { mask: resource } }
  group(workbook)
  removeSheetDrawings(workbook, 'data', [ref('group', 'g')])
  assert.deepEqual(workbook.metadata.chartImageResources.mask, resource); assert.equal(workbook.charts[0].id, 'other-chart')
  removeSheetDrawings(workbook, 'data', [ref('chart', 'other-chart')])
  assert.equal(workbook.metadata.chartImageResources, undefined); assert.equal(workbook.metadata.keep, 'other metadata')
})
