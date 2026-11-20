import assert from 'node:assert/strict'
import { after, before, test } from 'node:test'
import { createServer } from 'vite'

let server
let api

before(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})
after(async () => { await server?.close() })

function runtime(t, options) {
  const studio = new api.OfficeStudio(options)
  t.after(() => studio.dispose())
  return { studio, univer: api.FUniver.newAPI(studio), editor: studio.editor }
}

test('editor new-workbook and new-document actions synchronously register new API units and retain old facades', (t) => {
  const { studio, univer, editor } = runtime(t)
  const oldBook = univer.getActiveWorkbook()
  const oldDocument = univer.getActiveDocument()
  oldBook.getActiveSheet().getRange('A1').setValue('old workbook content')
  oldDocument.insertText(0, 'old document content')
  const oldBookId = oldBook.getId()
  const oldDocumentId = oldDocument.getId()
  const created = []
  univer.addEvent(univer.Event.WorkbookCreated, (event) => created.push(event.workbookId))
  univer.addEvent(univer.Event.DocCreated, (event) => created.push(event.documentId))
  editor.newWorkbook()
  const newBookId = editor.state.workbook.id
  assert.notEqual(newBookId, oldBookId)
  assert.equal(studio.activeWorkbookId, newBookId)
  assert.equal(univer.getActiveWorkbook().getId(), newBookId)
  assert.equal(oldBook.getActiveSheet().getRange('A1').getValue(), 'old workbook content')
  assert.equal(univer.getActiveWorkbook().getActiveSheet().getRange('A1').getValue(), null)
  editor.newDocument()
  const newDocumentId = editor.state.document.id
  assert.notEqual(newDocumentId, oldDocumentId)
  assert.equal(studio.activeDocumentId, newDocumentId)
  assert.equal(univer.getDocument(newDocumentId).getText(), '')
  assert.equal(oldDocument.getText(), 'old document content')
  assert.deepEqual(created, [newBookId, newDocumentId])
  studio.activateWorkbook(oldBookId)
  assert.equal(editor.state.workbook.id, oldBookId)
  assert.equal(editor.state.workbook.modified, true)
  oldBook.undo()
  assert.equal(oldBook.getActiveSheet().getRange('A1').getValue(), null)
  assert.equal(editor.state.workbook.modified, false)
  studio.activateDocument(oldDocumentId)
  assert.equal(oldDocument.getText(), 'old document content')
  assert.equal(oldDocument.undo(), true)
  assert.equal(oldDocument.getText(), '')
})

test('an imported normalized model becomes the matching active facade without replacing the former unit', async (t) => {
  const { studio, univer, editor } = runtime(t)
  const oldDocument = univer.getActiveDocument()
  oldDocument.insertText(0, 'keep old content')
  const imported = api.createBlankDocument()
  imported.id = 'synthetic-normalized-import'
  imported.name = 'Imported.docx'
  imported.blocks[0].runs[0].text = 'imported content'
  // This isolates the normalized parser boundary; actual WASM/file parsing is tested separately.
  editor.parser.parse = async () => ({ kind: 'document', model: imported, format: 'docx', fileName: imported.name, size: 128, parserVersion: 'fixture', diagnostics: [], capabilities: [], stats: {} })
  assert.equal(await editor.importFile({ name: imported.name }), true)
  assert.equal(studio.activeDocumentId, imported.id)
  assert.equal(univer.getActiveDocument().getId(), imported.id)
  assert.equal(univer.getDocument(imported.id).getText(), 'imported content')
  assert.equal(oldDocument.getText(), 'keep old content')
  assert.equal(editor.canUndo, false)
  studio.activateDocument(oldDocument.getId())
  assert.equal(editor.canUndo, true)
  assert.equal(oldDocument.undo(), true)
  assert.equal(oldDocument.getText(), '')
  assert.equal(univer.getDocument(imported.id).getText(), 'imported content')
})

test('sheet navigation has no content-history entry and undo/redo preserves the later active sheet', (t) => {
  const model = api.createBlankWorkbook()
  model.id = 'navigation-book'
  model.sheets.push(api.createBlankSheet('Second'))
  const { univer, editor } = runtime(t, { workbook: model })
  const book = univer.getActiveWorkbook()
  const [first, second] = book.getSheets()
  const originalRevision = editor.captureSavepoint('workbook').revision
  book.setActiveSheet(second)
  assert.equal(editor.state.workbook.activeSheetId, second.getSheetId())
  assert.equal(editor.captureSavepoint('workbook').revision, originalRevision)
  assert.equal(editor.canUndo, false)
  assert.equal(editor.state.workbook.modified, false)
  book.setActiveSheet(first)
  first.getRange('A1').setValue('edited')
  book.setActiveSheet(second)
  assert.equal(editor.state.workbook.modified, true)
  book.undo()
  assert.equal(first.getRange('A1').getValue(), null)
  assert.equal(book.getActiveSheet().getSheetId(), second.getSheetId())
  assert.equal(editor.state.workbook.modified, false)
  assert.equal(editor.canUndo, false)
  book.redo()
  assert.equal(first.getRange('A1').getValue(), 'edited')
  assert.equal(book.getActiveSheet().getSheetId(), second.getSheetId())
  assert.equal(editor.state.workbook.modified, true)
})

test('a savepoint completed for an inactive unit is reflected on activation without saving its neighbor', (t) => {
  const { studio, univer, editor } = runtime(t)
  const first = univer.getActiveWorkbook()
  first.getActiveSheet().getRange('A1').setValue('exported')
  const exported = editor.captureSavepoint('workbook')
  const second = univer.createWorkbook({ name: 'Second' })
  second.getActiveSheet().getRange('A1').setValue('unsaved')
  assert.equal(editor.markSaved(exported), true)
  assert.equal(studio.activeWorkbookId, second.getId())
  assert.equal(editor.state.workbook.modified, true)
  studio.activateWorkbook(first.getId())
  assert.equal(editor.state.workbook.modified, false)
  studio.activateWorkbook(second.getId())
  assert.equal(editor.state.workbook.modified, true)
  first.save()
  assert.equal(studio.activeWorkbookId, second.getId())
  assert.equal(editor.state.workbook.modified, true)
})

test('disposing the active workbook or document cannot resurrect it while creating the replacement unit', (t) => {
  const { studio, univer } = runtime(t)
  const book = univer.getActiveWorkbook()
  const document = univer.getActiveDocument()
  const oldBookId = book.getId()
  const oldDocumentId = document.getId()
  assert.equal(univer.disposeUnit(oldBookId), true)
  assert.notEqual(studio.activeWorkbookId, oldBookId)
  assert.equal(univer.getWorkbook(oldBookId), null)
  assert.throws(() => book.getName(), /not found/)
  assert.equal(univer.disposeUnit(oldDocumentId), true)
  assert.notEqual(studio.activeDocumentId, oldDocumentId)
  assert.equal(univer.getDocument(oldDocumentId), null)
  assert.throws(() => document.getName(), /not found/)
})

test('runtime disposal stops model synchronization, clears event listeners and rejects further commands', (t) => {
  const { studio, univer, editor } = runtime(t)
  const created = []
  const lifecycle = []
  univer.addEvent(univer.Event.WorkbookCreated, (event) => created.push(event.workbookId))
  univer.addEvent(univer.Event.LifeCycleChanged, (event) => lifecycle.push(event.stage))
  const activeId = studio.activeWorkbookId
  studio.dispose()
  studio.dispose()
  assert.deepEqual(lifecycle, ['disposed'])
  assert.equal(univer.getActiveWorkbook(), null)
  assert.equal(univer.getActiveDocument(), null)
  assert.equal(studio.getWorkbookModel(activeId), undefined)
  assert.throws(() => univer.createWorkbook(), /disposed/i)
  assert.throws(() => univer.executeCommandSync('sheet.command.set-workbook-name', { unitId: activeId, name: 'blocked' }), /disposed/i)
  editor.newWorkbook()
  assert.equal(studio.activeWorkbookId, activeId, 'the disposed model watcher must not re-register new units')
  assert.equal(univer.getWorkbook(editor.state.workbook.id), null)
  studio.events.emit('WorkbookCreated', { workbookId: 'after-dispose' })
  assert.deepEqual(created, [])
})

test('disposing a command registration removes the same command path used by API calls', (t) => {
  const { univer } = runtime(t)
  let executions = 0
  const registration = univer.registerCommand('test.operation.count', ({ amount }) => { executions += amount; return executions })
  assert.equal(univer.executeCommandSync('test.operation.count', { amount: 2 }), 2)
  registration.dispose()
  registration.dispose()
  assert.throws(() => univer.executeCommandSync('test.operation.count', { amount: 3 }), /Unknown command/)
  assert.equal(executions, 2)
})

test('calculation revisions follow content edits rather than selection, active-sheet navigation or save state', (t) => {
  const model = api.createBlankWorkbook()
  model.id = 'compute-revision-book'
  model.sheets.push(api.createBlankSheet('Second'))
  const { studio, univer } = runtime(t, { workbook: model })
  const book = univer.getActiveWorkbook()
  const [first, second] = book.getSheets()
  const initial = studio.getWorkbookRevision(book.getId())
  book.setActiveSheet(second)
  second.getRange('B3:C4').activate()
  book.save()
  studio.activateWorkbook(book.getId())
  assert.equal(studio.getWorkbookRevision(book.getId()), initial)
  first.getRange('A1').setValue(42)
  const edited = studio.getWorkbookRevision(book.getId())
  assert.ok(edited > initial)
  book.save()
  second.getRange('D5').activate()
  assert.equal(studio.getWorkbookRevision(book.getId()), edited)
  book.undo()
  const undone = studio.getWorkbookRevision(book.getId())
  assert.ok(undone > edited)
  book.redo()
  assert.ok(studio.getWorkbookRevision(book.getId()) > undone)
})
