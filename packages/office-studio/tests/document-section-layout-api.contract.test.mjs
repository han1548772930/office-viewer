import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Deferred public-command/history/layout drafts, 2026-11-22. Not executed.
let server, OfficeStudio, FDocument, domain, docModel, layoutModule
const studios = new Set()
beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio } = await server.ssrLoadModule('/packages/office-studio/src/api/office-studio.ts'))
  ;({ FDocument } = await server.ssrLoadModule('/packages/office-studio/src/api/facade.ts'))
  domain = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
  docModel = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
  layoutModule = await server.ssrLoadModule('/packages/office-studio/src/features/document/layout/document-layout.ts')
})
afterEach(() => { for (const studio of studios) studio.dispose(); studios.clear() })
afterAll(async () => { await server?.close() })
const copy = value => JSON.parse(JSON.stringify(value))
const measurer = { measure: text => Array.from(text).length * 10 }
const range = blockId => ({ anchor: { blockId, offset: 0 }, focus: { blockId, offset: 0 } })
const history = (studio, id) => ({ undo: studio.editor.history.units[id]?.undo.length ?? 0, redo: studio.editor.history.units[id]?.redo.length ?? 0, dirty: studio.editor.history.isDirty(id) })
function sourceModel(id = 'section-layout-api') {
  const model = domain.createBlankDocument(), first = docModel.createParagraph('A'.repeat(300)), last = docModel.createParagraph('Second section')
  model.id = id
  const section = { ...docModel.createParagraph(), id: id + '-section', type: 'sectionBreak', runs: [], sectionBreak: { type: 'nextPage' } }
  for (const block of [first, last]) Object.assign(block.style, { spaceBefore: 0, spaceAfter: 0, lineHeight: 20, spacingRule: 'exact', wordWrap: true })
  model.page = { width: 320, height: 360, marginLeft: 30, marginRight: 30, marginTop: 30, marginBottom: 30 }
  model.blocks = [first, section, last]
  model.wordCount = first.runs[0].text.length + last.runs[0].text.replace(/\s/g, '').length
  return model
}
function open(model = sourceModel()) {
  const studio = new OfficeStudio({ document: model, currentUserId: 'editor' })
  studios.add(studio); studio.activateDocument(model.id)
  return { studio, document: new FDocument(studio, model.id), body: range(model.blocks[0].id), sectionId: model.blocks[1].id }
}
const layout = document => layoutModule.layoutDocument(document.getModel(), measurer)
const geometry = document => layout(document).pages.map(page => ({ width: page.width, height: page.height, pageNumber: page.pageNumber, sectionId: page.sectionId, columns: page.sections.map(section => section.columns.length), blank: page.isParityBlank }))

test('section commands change real geometry, share one history and persist through internal save/reopen', () => {
  const { studio, document, body, sectionId } = open(), original = copy(document.getModel()), originalGeometry = geometry(document)
  const first = document.executeCommand({ id: 'doc.command.set-section-setup', sectionId: 'first', patch: { columnCount: 2, columnGap: 20, columnSeparator: 'line' } }, body)
  assert.equal(first.changed, true); assert.deepEqual(first.selection, body)
  assert.deepEqual(layout(document).pages[0].sections[0].columns.map(column => column.width), [120, 120])
  assert.deepEqual(history(studio, document.id), { undo: 1, redo: 0, dirty: true })
  const firstGeometry = geometry(document)
  document.executeCommand({ id: 'doc.command.set-section-setup', sectionId, patch: { width: 620, height: 480, orientation: 'landscape', marginLeft: 50, pageNumberStart: 7 } }, body)
  const edited = copy(document.getModel()), editedGeometry = geometry(document)
  assert.equal(layout(document).pages.at(-1).width, 620)
  assert.equal(layout(document).pages.at(-1).pageNumber, 7)
  assert.deepEqual(history(studio, document.id), { undo: 2, redo: 0, dirty: true })
  assert.equal(document.undo(), true); assert.deepEqual(geometry(document), firstGeometry)
  assert.equal(document.undo(), true); assert.deepEqual(document.getModel(), original); assert.deepEqual(geometry(document), originalGeometry)
  assert.equal(document.redo(), true); assert.equal(document.redo(), true)
  assert.deepEqual(document.getModel(), edited)
  const saved = copy(document.save()), reopened = open(saved)
  assert.deepEqual(geometry(reopened.document), editedGeometry)
  assert.deepEqual(reopened.document.getModel().sectionSetup, saved.sectionSetup)
  assert.deepEqual(reopened.document.getModel().blocks[1].sectionBreak, saved.blocks[1].sectionBreak)
  assert.deepEqual(history(studio, document.id), { undo: 2, redo: 0, dirty: false })
  assert.deepEqual(history(reopened.studio, saved.id), { undo: 0, redo: 0, dirty: false })
})

test('invalid effective geometry, permission denial and canceled commands publish neither model nor history', () => {
  const { studio, document, body, sectionId } = open(), before = copy(document.getModel()), beforeGeometry = geometry(document)
  for (const patch of [{ width: 40 }, { marginTop: 180, marginBottom: 180 }, { columnCount: 12, columnGap: 1000 }]) {
    assert.throws(() => document.executeCommand({ id: 'doc.command.set-section-setup', sectionId, patch }, body))
    assert.deepEqual(document.getModel(), before)
    assert.deepEqual(geometry(document), beforeGeometry)
    assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
  }
  const cancel = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === 'doc.command.set-section-setup') event.cancel = true })
  assert.throws(() => document.executeCommand({ id: 'doc.command.set-section-setup', sectionId, patch: { columnCount: 2 } }, body))
  cancel.dispose()
  assert.deepEqual(document.getModel(), before)
  assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
  const blockedModel = sourceModel('section-layout-denied')
  blockedModel.permissionRules = [{ id: 'deny-section', scope: 'section', targetId: blockedModel.blocks[1].id, action: 'edit', effect: 'deny', userIds: [] }]
  const blocked = open(blockedModel), blockedBefore = copy(blocked.document.getModel())
  assert.throws(() => blocked.document.executeCommand({ id: 'doc.command.set-section-setup', sectionId: blocked.sectionId, patch: { pageNumberStart: 8 } }, blocked.body))
  assert.deepEqual(blocked.document.getModel(), blockedBefore)
  assert.deepEqual(history(blocked.studio, blocked.document.id), { undo: 0, redo: 0, dirty: false })
})

test('layout edits and history remain scoped to the target document when another document is active', () => {
  const first = open(), secondModel = sourceModel('section-layout-other')
  first.studio.createDocument(secondModel); first.studio.activateDocument(secondModel.id)
  const second = new FDocument(first.studio, secondModel.id), secondGeometry = geometry(second)
  first.document.executeCommand({ id: 'doc.command.set-section-setup', sectionId: 'first', patch: { columnCount: 3, columnGap: 10 } }, first.body)
  assert.equal(first.studio.activeDocumentId, second.id)
  assert.deepEqual(geometry(second), secondGeometry)
  assert.deepEqual(history(first.studio, second.id), { undo: 0, redo: 0, dirty: false })
  assert.equal(first.document.undo(), true)
  assert.deepEqual(geometry(second), secondGeometry)
  assert.deepEqual(history(first.studio, first.document.id), { undo: 0, redo: 1, dirty: false })
})
