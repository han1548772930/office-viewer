import assert from 'node:assert/strict'
import { after, before, test } from 'node:test'
import { resolve } from 'node:path'
import { createServer } from 'vite'
import { createWorkspaceFileOperations } from '../../../examples/workbench/src/workspace-file-operations.ts'

let server
let createPinia
let useEditorStore
let cloneModel
let toExcelWorkbookView

before(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(), server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ createPinia } = await server.ssrLoadModule('pinia'))
  ;({ useEditorStore } = await server.ssrLoadModule('/packages/office-studio/src/stores/editorStore.ts'))
  ;({ cloneModel } = await server.ssrLoadModule('/packages/office-studio/src/domain/serialization.ts'))
  ;({ toExcelWorkbookView } = await server.ssrLoadModule('/packages/office-studio/src/services/excelViewAdapter.ts'))
})
after(async () => { await server?.close() })

function deferred() {
  let resolve
  let reject
  const promise = new Promise((res, rej) => { resolve = res; reject = rej })
  return { promise, resolve, reject }
}

function editCell(editor, text) {
  const view = toExcelWorkbookView(editor.state.workbook)
  view.sheets[0].cells.A1 = { value: text }
  editor.replaceWorkbookFromView(view)
}

function fileSession(editor, writer) {
  const downloaded = []
  const notices = []
  const files = createWorkspaceFileOperations({
    importFile: async () => false,
    cancelImport: () => {},
    capture: (kind = editor.state.activeKind) => ({ kind, model: cloneModel(kind === 'workbook' ? editor.state.workbook : editor.state.document), fileName: 'Saved.xlsx', savepoint: editor.captureSavepoint(kind) }),
    write: () => writer.promise,
    download: (_bytes, snapshot) => { downloaded.push(snapshot) },
    markSaved: (savepoint) => editor.markSaved(savepoint),
    onExportTask: () => {},
    notify: (type, message) => { notices.push({ type, message }) },
  })
  return { files, downloaded, notices }
}

test('file save preserves edits made during writing and leaves the other active document dirty', async () => {
  const editor = useEditorStore(createPinia())
  editCell(editor, 'saved content')
  const savedRevision = editor.captureSavepoint('workbook').revision
  const writer = deferred()
  const session = fileSession(editor, writer)
  const save = session.files.exportFile('save')
  editCell(editor, 'newer unsaved edit')
  editor.setActiveKind('document')
  editor.replaceDocumentFromView({ ...cloneModel(editor.state.document), name: 'Also unsaved.docx' })
  writer.resolve({ bytes: new Uint8Array([0x50, 0x4b]), warnings: [] })
  assert.equal(await save, true)
  assert.equal(editor.state.activeKind, 'document')
  assert.equal(editor.state.document.modified, true)
  assert.equal(editor.state.workbook.modified, true)
  assert.equal(toExcelWorkbookView(session.downloaded[0].model).sheets[0].cells.A1.value, 'saved content')
  assert.equal(toExcelWorkbookView(editor.state.workbook).sheets[0].cells.A1.value, 'newer unsaved edit')
  assert.equal(editor.history.units[editor.state.workbook.id].savedRevision, savedRevision)
  editor.setActiveKind('workbook')
  editor.undo()
  assert.equal(editor.state.workbook.modified, false)
  assert.equal(toExcelWorkbookView(editor.state.workbook).sheets[0].cells.A1.value, 'saved content')
})

test('a writer failure does not clear actual editor history dirty state', async () => {
  const editor = useEditorStore(createPinia())
  editCell(editor, 'not yet saved')
  const originalSavedRevision = editor.history.units[editor.state.workbook.id].savedRevision
  const writer = deferred()
  const session = fileSession(editor, writer)
  const save = session.files.exportFile('save')
  writer.reject(new Error('WASM writer failed'))
  assert.equal(await save, false)
  assert.equal(session.downloaded.length, 0)
  assert.equal(editor.state.workbook.modified, true)
  assert.equal(editor.history.units[editor.state.workbook.id].savedRevision, originalSavedRevision)
  assert.deepEqual(session.notices, [{ type: 'error', message: 'WASM writer failed' }])
})
