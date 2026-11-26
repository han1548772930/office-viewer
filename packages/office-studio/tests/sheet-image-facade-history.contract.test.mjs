import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { sheetImageAffine, sheetImageCopy as copy, sheetImageFixture } from './fixtures/sheet-image-fixture.mjs'

// Deferred real Studio/Facade/command/history contracts. This file has not started Vite or executed tests.
let server, OfficeStudio, FOffice, images, objects, geometry
const studios = new Set()
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/sheet-image-facade-history', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio } = await server.ssrLoadModule('/packages/office-studio/src/api/office-studio.ts'))
  ;({ FOffice } = await server.ssrLoadModule('/packages/office-studio/src/api/facade.ts'))
  ;[images, objects, geometry] = await Promise.all(['features/sheet-images/sheet-image-commands',
    'features/sheet-images/sheet-image-object-clipboard', 'features/sheet-workbench/sheet-clipboard-object-geometry']
    .map(path => server.ssrLoadModule('/packages/office-studio/src/' + path + '.ts')))
})
afterEach(() => { for (const studio of studios) studio.dispose(); studios.clear() })
afterAll(async () => server?.close())
function open(source, publicSnapshot = false) {
  if (!source) {
    const fixture = sheetImageFixture(); fixture.sheet.images = [sheetImageAffine('a', fixture.axes), sheetImageAffine('b', fixture.axes)]
    source = fixture.workbook
  }
  source.calculationMode = 'manual'
  const studio = new OfficeStudio({ ...(publicSnapshot ? {} : { workbook: source }), currentUserId: 'editor' }); studios.add(studio)
  const api = FOffice.newAPI(studio), workbook = publicSnapshot ? api.createWorkbook(source) : api.getActiveWorkbook(),
    sheet = workbook.getSheetBySheetId('source'), destination = workbook.getSheetBySheetId('destination')
  return { studio, workbook, sheet, destination }
}
const history = (studio, id) => ({ undo: studio.editor.history.units[id]?.undo.length ?? 0, redo: studio.editor.history.units[id]?.redo.length ?? 0 })
const sourceModel = state => state.studio.getWorkbookModel(state.workbook.id).sheets.find(sheet => sheet.id === 'source')
const destinationModel = state => state.studio.getWorkbookModel(state.workbook.id).sheets.find(sheet => sheet.id === 'destination')
const target = state => ({ unitId: state.workbook.id, sheetId: 'destination', range: { startRow: 10, endRow: 10, startCol: 6, endCol: 6 } })
function capture(state, cut = false, patch = {}) {
  return state.studio.executeCommandSync(cut ? objects.SHEET_IMAGE_OBJECT_COMMANDS.cut : objects.SHEET_IMAGE_OBJECT_COMMANDS.copy,
    { unitId: state.workbook.id, sheetId: 'source', imageIds: ['a'], range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 }, ...patch })
}

test('Facade setters share one history source and preserve affine media through undo and redo', () => {
  const state = open(), image = state.sheet.getImageById('a'), before = copy(sourceModel(state).images[0])
  assert.equal(image.setFlipX(false), true); assert.equal(image.setOpacity(0.25), true); assert.equal(image.setTitle('Changed'), true)
  assert.equal(history(state.studio, state.workbook.id).undo, 3)
  assert.equal(state.studio.undo('workbook', state.workbook.id), true); assert.equal(image.getSnapshot().title, before.title)
  assert.equal(state.studio.redo('workbook', state.workbook.id), true); assert.equal(image.getSnapshot().title, 'Changed')
  assert.deepEqual(sourceModel(state).images[0].drawingGeometry, before.drawingGeometry)
  assert.deepEqual(sourceModel(state).images[0].crop, before.crop); assert.deepEqual(sourceModel(state).images[0].resource, before.resource)
})

test('a later invalid image in a batch cannot publish an earlier valid edit or add history', () => {
  const state = open(), before = copy(state.studio.getWorkbookModel(state.workbook.id)), edits = copy(sourceModel(state).images)
  edits[0].title = 'Unpublished'; edits[1].size.width = 0
  assert.throws(() => state.sheet.updateImages(edits), /宽度|尺寸/)
  assert.deepEqual(state.studio.getWorkbookModel(state.workbook.id), before); assert.deepEqual(history(state.studio, state.workbook.id), { undo: 0, redo: 0 })
})

test('stale explicit copy revisions reject without replacing an existing clipboard or changing content', () => {
  const state = open(), first = capture(state), before = copy(state.studio.getWorkbookModel(state.workbook.id)),
    refused = capture(state, false, { expectedRevision: state.studio.getWorkbookRevision(state.workbook.id) + 1 })
  assert.equal(first.accepted, true); assert.equal(refused.accepted, false)
  assert.equal(state.studio.sheetClipboard.getImageClipboard().id, first.imageClipboard.id)
  assert.deepEqual(state.studio.getWorkbookModel(state.workbook.id), before); assert.equal(history(state.studio, state.workbook.id).undo, 0)
})

test('immediate object cut and later paste use separate single transactions and a fresh destination image identity', () => {
  const state = open(), original = copy(sourceModel(state).images[0]), cut = capture(state, true)
  assert.equal(cut.accepted, true); assert.equal(cut.changed, true); assert.equal(sourceModel(state).images.some(image => image.id === 'a'), false)
  assert.equal(history(state.studio, state.workbook.id).undo, 1)
  assert.equal(state.studio.undo('workbook', state.workbook.id), true); assert.ok(state.sheet.getImageById('a'))
  assert.equal(state.studio.redo('workbook', state.workbook.id), true); assert.equal(state.sheet.getImageById('a'), null)
  const pasted = state.studio.executeCommandSync(objects.SHEET_IMAGE_OBJECT_COMMANDS.paste, target(state))
  assert.equal(pasted.accepted, true); assert.equal(pasted.changed, true); assert.equal(pasted.imageIds.length, 1)
  assert.notEqual(pasted.imageIds[0], original.id); assert.equal(history(state.studio, state.workbook.id).undo, 2)
  const image = destinationModel(state).images[0]
  for (const key of ['a', 'b', 'c', 'd']) assert.equal(image.drawingGeometry.matrix[key], original.drawingGeometry.matrix[key])
  assert.deepEqual(image.crop, original.crop); assert.deepEqual(image.resource, original.resource)
})

test('internal HTML reconstructs an image author in a newly opened Studio after the source broker is disposed', () => {
  const first = open(), copied = capture(first), saved = first.workbook.save(), original = copy(sourceModel(first).images[0])
  first.studio.dispose(); studios.delete(first.studio)
  const state = open(copy(saved)), result = state.studio.executeCommandSync('sheet.command.paste', { ...target(state), html: copied.html, text: copied.plain })
  assert.equal(result.accepted, true); assert.equal(result.changed, true); assert.equal(destinationModel(state).images.length, 1)
  const pasted = destinationModel(state).images[0]
  assert.notEqual(pasted.id, original.id); assert.deepEqual(pasted.crop, original.crop); assert.equal(pasted.flipX, original.flipX)
  for (const key of ['a', 'b', 'c', 'd']) assert.equal(pasted.drawingGeometry.matrix[key], original.drawingGeometry.matrix[key])
  assert.equal(history(state.studio, state.workbook.id).undo, 1)
})

test('source copy permissions cover the transformed footprint beyond the anchor cell', () => {
  const state = open(), model = state.studio.getWorkbookModel(state.workbook.id), source = sourceModel(state),
    clipboard = objects.captureSheetImageObjects(model, source, ['a'], 'editor'), footprint = geometry.sheetClipboardObjectRange(model, source.id, clipboard.entries[0].bounds)
  source.rangeProtectionRules = [{ id: 'hidden-corner', permissionId: 'hidden-corner', ownerUserId: 'other', allowedUserIds: [],
    editState: 'onlyMe', viewState: 'onlyMe', description: '', ranges: [{ startRow: footprint.endRow, endRow: footprint.endRow,
      startCol: footprint.endCol, endCol: footprint.endCol }] }]
  const before = copy(model)
  assert.throws(() => capture(state), /权限/); assert.deepEqual(state.studio.getWorkbookModel(state.workbook.id), before)
  assert.equal(history(state.studio, state.workbook.id).undo, 0)
})

test('destination paste checks the complete new footprint instead of just the selected target cell', () => {
  const state = open(), copied = capture(state), model = state.studio.getWorkbookModel(state.workbook.id), destination = destinationModel(state),
    plan = objects.planPasteSheetImageObjects(model, destination, target(state).range, copied.imageClipboard), footprint = plan.ranges[0]
  destination.rangeProtectionRules = [{ id: 'locked-corner', permissionId: 'locked-corner', ownerUserId: 'other', allowedUserIds: [],
    editState: 'onlyMe', viewState: 'othersCanView', description: '', ranges: [{ startRow: footprint.endRow, endRow: footprint.endRow,
      startCol: footprint.endCol, endCol: footprint.endCol }] }]
  const before = copy(model)
  assert.throws(() => state.studio.executeCommandSync(objects.SHEET_IMAGE_OBJECT_COMMANDS.paste, target(state)), /权限/)
  assert.deepEqual(state.studio.getWorkbookModel(state.workbook.id), before); assert.equal(history(state.studio, state.workbook.id).undo, 0)
})

test('a cleared history generation rejects an earlier captured paste target without changing the author model', () => {
  const state = open(); capture(state)
  const generation = state.studio.editor.history.captureSavepoint(state.workbook.id).generation
  state.studio.editor.history.clear(state.workbook.id)
  const before = copy(state.studio.getWorkbookModel(state.workbook.id)), result = state.studio.executeCommandSync(objects.SHEET_IMAGE_OBJECT_COMMANDS.paste,
    { ...target(state), expectedHistoryGeneration: generation })
  assert.equal(result.accepted, false); assert.equal(result.reason, 'stale-target')
  assert.deepEqual(state.studio.getWorkbookModel(state.workbook.id), before); assert.equal(history(state.studio, state.workbook.id).undo, 0)
})

test('history reentrancy before the nested insert publication cannot publish an image from an earlier target generation', () => {
  const state = open(); capture(state)
  const before = copy(state.studio.getWorkbookModel(state.workbook.id))
  state.studio.events.addEvent('BeforeCommandExecute', event => {
    if (event.id === images.SHEET_IMAGE_COMMANDS.insertPrepared) state.studio.editor.history.clear(state.workbook.id)
  })
  assert.throws(() => state.studio.executeCommandSync(objects.SHEET_IMAGE_OBJECT_COMMANDS.paste, target(state)), /历史|更改/)
  assert.deepEqual(state.studio.getWorkbookModel(state.workbook.id), before); assert.equal(history(state.studio, state.workbook.id).undo, 0)
})

test('malformed author media cannot partially insert images or consume clipboard content', () => {
  const state = open(), copied = capture(state), invalid = copy(copied.imageClipboard), before = copy(state.studio.getWorkbookModel(state.workbook.id))
  invalid.id = 'invalid-author'; invalid.entries[0].image.resource.source = 'data:image/png;base64,AAAA'
  const html = `<div data-office-image-clipboard-id="${invalid.id}" data-office-image-clipboard-model="${encodeURIComponent(JSON.stringify(invalid))}"></div>`
  assert.throws(() => state.studio.executeCommandSync(objects.SHEET_IMAGE_OBJECT_COMMANDS.paste, { ...target(state), html, text: '' }))
  assert.deepEqual(state.studio.getWorkbookModel(state.workbook.id), before); assert.equal(history(state.studio, state.workbook.id).undo, 0)
  assert.equal(state.studio.sheetClipboard.getImageClipboard().id, copied.imageClipboard.id)
})

test('internal and public snapshots reopen the editable affine author with crop, flips and unknown metadata', () => {
  const state = open(), original = copy(sourceModel(state).images[0]), internal = open(copy(state.workbook.save())), publicState = open(copy(state.workbook.savePublicSnapshot()), true)
  for (const reopened of [internal, publicState]) {
    const image = sourceModel(reopened).images.find(image => image.id === original.id)
    assert.ok(image); assert.deepEqual(image.drawingGeometry, original.drawingGeometry); assert.deepEqual(image.crop, original.crop)
    assert.equal(image.flipX, true); assert.equal(image.flipY, true); assert.deepEqual(image.metadata.opaqueAuthor, original.metadata.opaqueAuthor)
    assert.equal(reopened.sheet.getImageById(original.id).setTitle('After reopen'), true)
    assert.equal(reopened.studio.undo('workbook', reopened.workbook.id), true); assert.equal(reopened.sheet.getImageById(original.id).getSnapshot().title, original.title)
  }
})
