import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'
import { createBlankSheet } from '../src/domain/model.ts'
import { sheetDrawingWorkbook, drawingRef as ref } from './fixtures/sheet-drawing-model-fixture.mjs'

// Deferred real registry/Facade/lifecycle integration. The hooks below have not been invoked.
let server, api, host, broker, images, geometry, axes
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[api, host, broker, images, geometry, axes] = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/index.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/sheet-drawings/sheet-drawing-host.ts'),
    server.ssrLoadModule('/packages/office-studio/src/api/sheet-drawing-selection.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/sheet-images/sheet-image-placement.ts'),
    server.ssrLoadModule('/packages/office-studio/src/domain/drawing-affine.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/sheet-workbench/sheet-clipboard-object-geometry.ts'),
  ])
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
const count = (studio, unitId) => studio.editor.history.units[unitId]?.undo.length ?? 0
const selection = [ref('image', 'a'), ref('chart', 'c')]
const delta = { a: 1, b: 0, c: 0, d: 1, e: 30, f: 15 }
function setup(model = sheetDrawingWorkbook()) {
  const studio = new api.OfficeStudio({ workbook: model }), facade = api.FOffice.newAPI(studio), book = facade.getActiveWorkbook(), sheet = book.getActiveSheet()
  const unitId = book.getId(), events = [], event = studio.events.addEvent('CommandExecuted', value => events.push(value.id))
  return { studio, facade, book, sheet, unitId, events, dispose() { event.dispose(); studio.dispose() } }
}
const scene = f => host.createSheetDrawingScene(f.studio.getWorkbookModel(f.unitId), 'data')

test('generic Facade alignment calls the canonical registry once and restores content through undo/redo and reopening', () => {
  const f = setup(); let reopened
  try {
    const before = json(f.book.save()), history = count(f.studio, f.unitId)
    assert.equal(f.sheet.alignDrawings(selection, f.facade.Enum.DrawingAlignType.left), f.sheet)
    assert.equal(f.events.filter(id => id === api.SHEET_DRAWING_GROUP_COMMAND_IDS.align).length, 1)
    assert.equal(count(f.studio, f.unitId), history + 1)
    const afterScene = scene(f), a = afterScene.objects.find(object => object.ref.kind === 'image' && object.ref.id === 'a'), c = afterScene.objects.find(object => object.ref.kind === 'chart')
    assert.equal(geometry.drawingAffineFrame(a.geometry).left, geometry.drawingAffineFrame(c.geometry).left)
    assert.deepEqual(f.book.save().charts[0].source, before.charts[0].source); assert.deepEqual(f.book.save().sheets[0].images[0].crop, before.sheets[0].images[0].crop)
    const saved = json(f.book.save()); f.book.undo(); assert.deepEqual(json(f.book.save()), before)
    f.book.redo(); assert.deepEqual(json(f.book.save()), saved)
    reopened = new api.OfficeStudio({ workbook: saved })
    assert.deepEqual(json(host.createSheetDrawingScene(reopened.getWorkbookModel(f.unitId), 'data')), json(afterScene))
  } finally { reopened?.dispose(); f.dispose() }
})

test('group, transform, arrange, ungroup and remove wrappers share native owners and one history entry per actual change', () => {
  const f = setup()
  try {
    const initial = count(f.studio, f.unitId), author = json(f.book.save())
    assert.equal(f.sheet.groupDrawings(selection, 'g'), f.sheet)
    assert.equal(f.sheet.transformDrawings([ref('group', 'g')], { delta }), f.sheet)
    assert.equal(f.sheet.arrangeDrawings([ref('group', 'g')], 'back'), f.sheet)
    const transformed = json(f.book.save())
    assert.equal(f.sheet.ungroupDrawings([ref('group', 'g')]), f.sheet)
    assert.deepEqual(f.book.save().charts, transformed.charts); assert.deepEqual(f.book.save().sheets[0].images, transformed.sheets[0].images)
    assert.equal(f.sheet.removeDrawings(selection), f.sheet)
    assert.equal(count(f.studio, f.unitId), initial + 5)
    assert.equal(f.book.save().charts.length, 0); assert.deepEqual(f.book.save().sheets[0].images.map(image => image.id), ['d'])
    for (let i = 0; i < 5; i++) f.book.undo()
    assert.deepEqual(json(f.book.save()), author)
  } finally { f.dispose() }
})

test('no-op alignment and invalid alignment/transform parameters add no content history', () => {
  const f = setup()
  try {
    const before = json(f.book.save()), history = count(f.studio, f.unitId)
    f.sheet.alignDrawings([], '1'); f.sheet.alignDrawings([selection[0]], '1'); f.sheet.alignDrawings(selection, '0')
    assert.throws(() => f.sheet.alignDrawings(selection, 'left'), TypeError)
    assert.throws(() => f.sheet.transformDrawings(selection, {}), TypeError)
    assert.throws(() => f.sheet.transformDrawings(selection, { frame: scene(f).objects[0].geometry, delta }), TypeError)
    assert.deepEqual(json(f.book.save()), before); assert.equal(count(f.studio, f.unitId), history)
  } finally { f.dispose() }
})

test('a bound worksheet cannot be redirected by extra transform fields or a different active workbook', () => {
  const f = setup()
  try {
    const other = f.facade.createWorkbook(sheetDrawingWorkbook({ id: 'other' })), saved = json(other.save())
    f.sheet.transformDrawings([ref('image', 'a')], { delta, unitId: other.getId(), sheetId: 'wrong', selection: [ref('image', 'd')] })
    assert.deepEqual(json(other.save()), saved)
    const own = scene(f), moved = own.objects.find(object => object.ref.kind === 'image' && object.ref.id === 'a')
    assert.equal(geometry.drawingAffineFrame(moved.geometry).left, 40)
    assert.equal(geometry.drawingAffineFrame(own.objects.find(object => object.ref.kind === 'image' && object.ref.id === 'd').geometry).left, 90)
  } finally { f.dispose() }
})

test('alignment rejects stale expected selection and aborted commands before content or history changes', async () => {
  const f = setup()
  try {
    broker.setSheetDrawingSelection(f.studio, f.unitId, 'data', [ref('image', 'd')])
    const before = json(f.book.save()), history = count(f.studio, f.unitId), input = { unitId: f.unitId, sheetId: 'data', selection, alignType: '1' }
    await assert.rejects(f.studio.executeCommand(api.SHEET_DRAWING_GROUP_COMMAND_IDS.align, { ...input, expectedSelection: selection }), /选择/)
    const controller = new AbortController(); controller.abort()
    await assert.rejects(f.studio.executeCommand(api.SHEET_DRAWING_GROUP_COMMAND_IDS.align, { ...input, signal: controller.signal }), /取消/)
    assert.deepEqual(json(f.book.save()), before); assert.equal(count(f.studio, f.unitId), history)
  } finally { f.dispose() }
})

test('a real axis-size commit publishes image projection and parent fitting in the same saved history snapshot', () => {
  const model = sheetDrawingWorkbook(), sheet = model.sheets[0]
  const ownAxes = axes.sheetClipboardObjectAxes(model, sheet)
  sheet.images[0] = images.sheetImageWithFrame(sheet.images[0], { left: 150, top: 60, width: 80, height: 40 }, ownAxes, '0')
  sheet.images[0] = images.sheetImageWithDrawingGeometry(sheet.images[0], geometry.drawingAffineFromFrame({ left: 150, top: 60, width: 80, height: 40 }), ownAxes)
  const f = setup(model)
  try {
    f.sheet.groupDrawings(selection, 'g')
    const before = json(f.book.save()), history = count(f.studio, f.unitId), oldMatrix = before.sheets[0].images[0].drawingGeometry.matrix
    f.sheet.setColumnWidth(0, 150)
    assert.equal(count(f.studio, f.unitId), history + 1)
    const saved = json(f.book.save()), projected = saved.sheets[0].images[0]
    assert.equal(projected.drawingGeometry.matrix.e, oldMatrix.e + 50)
    assert.deepEqual(projected.size, before.sheets[0].images[0].size); assert.deepEqual(projected.crop, before.sheets[0].images[0].crop)
    assert.notDeepEqual(saved.sheets[0].drawingState.groups[0].frame, before.sheets[0].drawingState.groups[0].frame)
    f.book.undo(); assert.deepEqual(json(f.book.save()), before)
    f.book.redo(); assert.deepEqual(json(f.book.save()), saved)
  } finally { f.dispose() }
})

test('filter refresh projects an untouched affine image after its final hidden-row set is computed', () => {
  const model = sheetDrawingWorkbook(), sheet = model.sheets[0], ownAxes = axes.sheetClipboardObjectAxes(model, sheet)
  sheet.images[0] = images.sheetImageWithFrame(sheet.images[0], { left: 150, top: 60, width: 80, height: 40 }, ownAxes, '0')
  sheet.images[0] = images.sheetImageWithDrawingGeometry(sheet.images[0], geometry.drawingAffineFromFrame({ left: 150, top: 60, width: 80, height: 40 }), ownAxes)
  const f = setup(model)
  try {
    f.sheet.setAutoFilter('A1:B3')
    const before = json(f.book.save()), history = count(f.studio, f.unitId)
    f.sheet.setFilterCriteria(0, { kind: 'values', selectedItems: ['North'] })
    const saved = json(f.book.save())
    assert.deepEqual(saved.sheets[0].autoFilter.filteredRows, [2])
    assert.equal(saved.sheets[0].images[0].drawingGeometry.matrix.f, before.sheets[0].images[0].drawingGeometry.matrix.f - 20)
    assert.equal(count(f.studio, f.unitId), history + 1)
    f.book.undo(); assert.deepEqual(json(f.book.save()), before); f.book.redo(); assert.deepEqual(json(f.book.save()), saved)
  } finally { f.dispose() }
})

test('worksheet copy remaps grouped native owners and reopening retains complete affine author data', () => {
  const f = setup(); let reopened
  try {
    f.sheet.groupDrawings(selection, 'g'); f.sheet.transformDrawings([ref('group', 'g')], { delta: { a: 1, b: 0.2, c: 0.3, d: 1, e: 50, f: 30 } })
    const original = json(f.book.save()), copied = f.book.copySheet(f.sheet), model = json(f.book.save()), destination = model.sheets.find(sheet => sheet.id === copied.getSheetId())
    assert.ok(destination); assert.notEqual(destination.drawingState.groups[0].id, 'g')
    assert.deepEqual(destination.images.map(image => image.drawingGeometry), original.sheets[0].images.map(image => image.drawingGeometry))
    assert.ok(destination.images.every(image => !original.sheets[0].images.some(source => source.id === image.id)))
    const nativeChart = model.charts.find(chart => chart.sheetId === destination.id)
    assert.deepEqual(nativeChart.drawingGeometry, original.charts[0].drawingGeometry)
    reopened = new api.OfficeStudio({ workbook: model })
    assert.deepEqual(json(host.createSheetDrawingScene(reopened.getWorkbookModel(f.unitId), destination.id)), json(host.createSheetDrawingScene(model, destination.id)))
  } finally { reopened?.dispose(); f.dispose() }
})

test('commands addressed to a protected inactive worksheet cannot fall back to an editable active sheet', () => {
  const model = sheetDrawingWorkbook(), locked = createBlankSheet('Locked')
  locked.id = 'locked'; locked.shapes = [{ id: 'locked-shape', sheetId: 'locked', geometry: model.sheets[0].shapes[0].geometry, source: { kind: 'preset', preset: 'rect' } }]
  locked.worksheetProtection = { id: 'lock', permissionId: 'lock', unitType: 'worksheet', ownerUserId: 'peer', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView' }
  model.sheets.push(locked)
  const f = setup(model)
  try {
    const other = f.book.getSheetBySheetId('locked'), before = json(f.book.save()), history = count(f.studio, f.unitId)
    assert.throws(() => other.transformDrawings([ref('shape', 'locked-shape')], { delta }))
    assert.deepEqual(json(f.book.save()), before); assert.equal(count(f.studio, f.unitId), history)
  } finally { f.dispose() }
})
