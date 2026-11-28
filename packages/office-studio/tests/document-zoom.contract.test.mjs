import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Deferred real API/view-history/public-snapshot contracts. Not executed.
let server, OfficeStudio, FDocument, zoom
const studios = new Set()
beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio } = await server.ssrLoadModule('/packages/office-studio/src/api/office-studio.ts'))
  ;({ FDocument } = await server.ssrLoadModule('/packages/office-studio/src/api/facade.ts'))
  zoom = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/document-zoom.ts')
})
afterEach(() => { for (const studio of studios) studio.dispose(); studios.clear() })
afterAll(async () => { await server?.close() })
const copy = value => JSON.parse(JSON.stringify(value))
const history = (studio, id) => ({ undo: studio.editor.history.units[id]?.undo.length ?? 0,
  redo: studio.editor.history.units[id]?.redo.length ?? 0, dirty: studio.editor.history.isDirty(id) })
const snapshot = (id, zoomRatio, flavor = 1) => ({ id, title: id, documentStyle: { documentFlavor: flavor },
  body: { dataStream: '甲乙\r\n' }, ...(zoomRatio === undefined ? {} : { settings: { zoomRatio, customSetting: { kept: true } } }) })
function open(input = snapshot('zoom-doc', 1)) {
  const studio = new OfficeStudio({ currentUserId: 'reader' })
  studios.add(studio)
  const model = studio.createDocument(copy(input))
  studio.activateDocument(model.id)
  return { studio, document: new FDocument(studio, model.id) }
}

test.each([1, 2])('public settings and detached snapshots keep per-document zoom in flavor %s', flavor => {
  const { studio, document } = open(snapshot('zoom-flavor-' + flavor, 0.8, flavor))
  assert.equal(document.getZoomRatio(), 0.8)
  assert.equal(document.getModel().metadata.documentZoomRatio, 0.8)
  assert.equal(document.setZoomRatio(1.3), true)
  const saved = document.save(), internal = document.saveModel()
  assert.equal(saved.settings.zoomRatio, 1.3)
  assert.deepEqual(saved.settings.customSetting, { kept: true })
  saved.settings.customSetting.kept = false
  assert.deepEqual(document.getSnapshot().settings.customSetting, { kept: true })
  assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
  assert.equal(open(document.getSnapshot()).document.getZoomRatio(), 1.3)
  const nativeStudio = new OfficeStudio({ document: internal, currentUserId: 'reader' })
  studios.add(nativeStudio)
  assert.equal(new FDocument(nativeStudio, internal.id).getZoomRatio(), 1.3)
})

test('missing and null settings use one default without consuming content history', () => {
  const original = open(snapshot('zoom-missing'))
  assert.equal(original.document.getZoomRatio(), 1)
  assert.equal(original.document.getSnapshot().settings.zoomRatio, 1)
  const nullable = open({ ...snapshot('zoom-null'), settings: { zoomRatio: null } })
  assert.equal(nullable.document.getZoomRatio(), 1)
  assert.equal(original.studio.executeCommandSync('doc.command.set-zoom-ratio', { documentId: original.document.id }), true)
  assert.deepEqual(history(original.studio, original.document.id), { undo: 0, redo: 0, dirty: false })
  assert.deepEqual(zoom.DOCUMENT_ZOOM_SHORTCUTS, [50, 80, 100, 130, 150, 170, 200, 400])
})

test.each([0, -1, 0.099, 4.001, '1', Number.NaN, Number.POSITIVE_INFINITY])('invalid ratio %s is rejected before view or snapshot publication', invalid => {
  const { studio, document } = open()
  const before = document.getSnapshot()
  assert.throws(() => document.setZoomRatio(invalid))
  assert.deepEqual(document.getSnapshot(), before)
  assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
  assert.throws(() => studio.createDocument({ ...snapshot('zoom-invalid'), settings: { zoomRatio: invalid } }))
})

test('command, operation and Facade isolate units and publish the same ratio event', () => {
  const { studio, document } = open(snapshot('zoom-a', 0.5))
  const otherModel = studio.createDocument(snapshot('zoom-b', 2))
  const other = new FDocument(studio, otherModel.id)
  studio.activateDocument(other.id)
  const observed = []
  const subscription = document.onZoomRatioChanged(event => observed.push(copy(event)))
  assert.equal(studio.executeCommandSync('doc.operation.set-zoom-ratio', { unitId: document.id, zoomRatio: 1.5 }), true)
  assert.equal(document.getZoomRatio(), 1.5); assert.equal(other.getZoomRatio(), 2)
  assert.equal(studio.activeDocumentId, other.id)
  assert.equal(studio.executeCommandSync('doc.command.set-zoom-ratio', { documentId: document.id, zoomRatio: 1.7 }), true)
  assert.equal(document.setZoomRatio(0.8), true)
  assert.deepEqual(observed.map(event => [event.documentId, event.previousZoomRatio, event.zoomRatio]),
    [['zoom-a', 0.5, 1.5], ['zoom-a', 1.5, 1.7], ['zoom-a', 1.7, 0.8]])
  subscription.dispose()
  document.setZoomRatio(1)
  assert.equal(observed.length, 3)
  assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
  assert.deepEqual(history(studio, other.id), { undo: 0, redo: 0, dirty: false })
})

test('read-only documents can zoom and view-denied documents cannot', () => {
  const deniedEdit = snapshot('zoom-reader', 1)
  deniedEdit.officeStudio = { version: 1, permissionRules: [{ id: 'deny-edit', scope: 'document', action: 'edit', effect: 'deny', userIds: ['reader'] }] }
  const readOnly = open(deniedEdit)
  assert.equal(readOnly.document.setZoomRatio(1.3), true)
  assert.equal(readOnly.document.getZoomRatio(), 1.3)
  assert.deepEqual(history(readOnly.studio, readOnly.document.id), { undo: 0, redo: 0, dirty: false })
  const deniedView = snapshot('zoom-hidden', 1)
  deniedView.officeStudio = { version: 1, permissionRules: [{ id: 'deny-view', scope: 'document', action: 'view', effect: 'deny', userIds: ['reader'] }] }
  const inaccessible = open(deniedView)
  assert.throws(() => inaccessible.document.setZoomRatio(1.3))
  assert.deepEqual(history(inaccessible.studio, inaccessible.document.id), { undo: 0, redo: 0, dirty: false })
})

test('view changes preserve the text input session and survive content undo and redo', () => {
  const { studio, document } = open()
  const id = document.getModel().blocks[0].id
  const selection = { anchor: { blockId: id, offset: 1 }, focus: { blockId: id, offset: 1 } }
  document.setInputSelection(selection)
  document.executeCommand({ id: 'doc.command.insert-text', text: '新增' }, selection)
  const input = copy(document.getInputState()), revision = document.getModelRevision(), contentHistory = history(studio, document.id)
  document.setZoomRatio(1.7)
  assert.deepEqual(document.getInputState(), input)
  assert.equal(document.getModelRevision(), revision)
  assert.deepEqual(history(studio, document.id), contentHistory)
  assert.equal(document.undo(), true)
  assert.equal(document.getZoomRatio(), 1.7)
  assert.equal(document.getText(), '甲乙')
  assert.equal(document.redo(), true)
  assert.equal(document.getZoomRatio(), 1.7)
  assert.match(document.getText(), /新增/)
})

test('cancel and same-ID replacement during BeforeCommandExecute cannot write into a reopened unit', () => {
  const { studio, document } = open()
  const cancel = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === 'doc.operation.set-zoom-ratio') event.cancel = true })
  assert.throws(() => studio.executeCommandSync('doc.operation.set-zoom-ratio', { unitId: document.id, zoomRatio: 2 }))
  cancel.dispose()
  assert.equal(document.getZoomRatio(), 1)
  const replacement = studio.events.addEvent('BeforeCommandExecute', event => {
    if (event.id === 'doc.operation.set-zoom-ratio') { replacement.dispose(); studio.createDocument(snapshot(document.id, 0.5)) }
  })
  assert.throws(() => studio.executeCommandSync('doc.operation.set-zoom-ratio', { unitId: document.id, zoomRatio: 2 }))
  assert.equal(new FDocument(studio, document.id).getZoomRatio(), 0.5)
  assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
})
