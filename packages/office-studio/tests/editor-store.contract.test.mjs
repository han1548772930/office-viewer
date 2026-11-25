import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createPinia } from 'pinia'
import { reactive } from 'vue'
import { createServer } from 'vite'

let server
let useEditorStore
let toExcelWorkbookView
let createBlankSheet
let createCell
let moveSheetCellRange
let structureEdit
let transformWorkbookReferencesForSheetEdit

beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ useEditorStore } = await server.ssrLoadModule('/packages/office-studio/src/stores/editorStore.ts'))
  ;({ toExcelWorkbookView } = await server.ssrLoadModule('/packages/office-studio/src/services/excelViewAdapter.ts'))
  ;({ createBlankSheet, createCell } = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts'))
  ;({ moveSheetCellRange, structureEdit, transformWorkbookReferencesForSheetEdit } = await server.ssrLoadModule('/packages/office-studio/src/features/sheet-workbench/sheet-cell-range-move.ts'))
})

afterAll(async () => { await server?.close() })
const freshEditor = () => useEditorStore(createPinia())
const snapshot = (value) => JSON.parse(JSON.stringify(value))

test('unchanged workbook and document view callbacks do not create edits', () => {
  const editor = freshEditor()
  const original = snapshot(editor.state.workbook)
  assert.equal(editor.replaceWorkbookFromView(toExcelWorkbookView(editor.state.workbook)), false)
  assert.deepEqual(snapshot(editor.state.workbook), original, 'viewport normalization must not change untouched source geometry')
  assert.equal(editor.state.workbook.modified, false)
  assert.equal(editor.canUndo, false)
  assert.equal(editor.replaceDocumentFromView(snapshot(editor.state.document)), false)
  editor.setActiveKind('document')
  assert.equal(editor.state.document.modified, false)
  assert.equal(editor.canUndo, false)
})

test('filename edits use the same per-unit history, including immediate kind switches', () => {
  const editor = freshEditor()
  const bookName = editor.state.fileName
  editor.state.fileName = '已重命名.xlsx'
  assert.equal(editor.state.workbook.modified, true)
  assert.equal(editor.canUndo, true)
  editor.setActiveKind('document')
  const docName = editor.state.fileName
  editor.state.fileName = '已重命名.docx'
  editor.setActiveKind('workbook')
  editor.undo()
  assert.equal(editor.state.workbook.name, bookName)
  assert.equal(editor.state.document.name, '已重命名.docx')
  assert.equal(editor.state.workbook.modified, false)
  assert.equal(editor.state.document.modified, true)
  editor.setActiveKind('document')
  editor.undo()
  assert.equal(editor.state.document.name, docName)
  assert.equal(editor.state.document.modified, false)
})

test('saving an exported snapshot after switching kinds never saves the other document', () => {
  const editor = freshEditor()
  editor.state.fileName = 'first.xlsx'
  const ticket = editor.captureSavepoint('workbook')
  editor.state.fileName = 'second.xlsx'
  editor.setActiveKind('document')
  editor.state.fileName = 'unsaved.docx'
  assert.equal(editor.markSaved(ticket), true)
  assert.equal(editor.state.document.modified, true)
  assert.equal(editor.state.workbook.modified, true)
  editor.setActiveKind('workbook')
  editor.undo()
  assert.equal(editor.state.workbook.name, 'first.xlsx')
  assert.equal(editor.state.workbook.modified, false)
  editor.redo()
  assert.equal(editor.state.workbook.modified, true)
})

test('a workbook view that only switches sheets is an operation', () => {
  const editor = freshEditor()
  const model = snapshot(editor.state.workbook)
  const second = createBlankSheet('Second')
  model.sheets.push(second)
  editor.workbook.replace(model)
  const view = toExcelWorkbookView(model)
  view.activeSheetId = second.id
  assert.equal(editor.replaceWorkbookFromView(view), false)
  assert.equal(editor.state.workbook.activeSheetId, second.id)
  assert.equal(editor.canUndo, false)
  assert.equal(editor.state.workbook.modified, false)
})

test('canceled imports return false instead of reporting successful replacement', async () => {
  const editor = freshEditor()
  const originalId = editor.state.workbook.id
  editor.parser.parse = async () => undefined
  assert.equal(await editor.importFile({ name: 'canceled.xlsx' }), false)
  assert.equal(editor.state.workbook.id, originalId)
})

test('content undo and redo preserve a later sheet-navigation operation', () => {
  const editor = freshEditor()
  const model = snapshot(editor.state.workbook)
  const second = createBlankSheet('Second')
  model.sheets.push(second)
  editor.workbook.replace(model)
  const edited = toExcelWorkbookView(model)
  edited.sheets[0].cells.A1 = { value: 'content edit' }
  assert.equal(editor.replaceWorkbookFromView(edited), true)
  const navigated = toExcelWorkbookView(editor.state.workbook)
  navigated.activeSheetId = second.id
  assert.equal(editor.replaceWorkbookFromView(navigated), false)
  editor.undo()
  assert.equal(editor.state.workbook.activeSheetId, second.id)
  assert.deepEqual(editor.state.workbook.sheets[0].cells, {})
  assert.equal(editor.state.workbook.modified, false)
  editor.redo()
  assert.equal(editor.state.workbook.activeSheetId, second.id)
  assert.equal(editor.state.workbook.sheets[0].cells['0:0'].value, 'content edit')
})

test('range-cell shifts enter workbook history as one atomic undoable edit', () => {
  const editor = freshEditor()
  const model = snapshot(editor.state.workbook)
  const sheet = model.sheets[0]
  model.calculationMode = 'manual'
  model.defaultStyle = { font: { bold: true }, fill: '#eeddcc' }
  sheet.cells = { '0:0': createCell(0, 0, 'before'), '0:1': createCell(0, 1, 'move') }
  sheet.cells['0:1'].formula = '=A1+B1'
  model.metadata = { ...model.metadata, definedNames: [{ name: 'Moved', refersTo: `${sheet.name}!B1` }] }
  editor.workbook.replace(model)
  const original = toExcelWorkbookView(model)
  const moved = moveSheetCellRange(original.sheets[0], { startRow: 0, startColumn: 1, endRow: 0, endColumn: 1 }, 'right')
  assert.equal(moved.status, 'applied')
  const edit = structureEdit({ startRow: 0, startColumn: 1, endRow: 0, endColumn: 1 }, 'right')
  const after = transformWorkbookReferencesForSheetEdit({ ...original, sheets: [moved.sheet] }, sheet.id, edit)
  assert.equal(after.sheets[0].cells.C1.formula, '=A1+C1')
  assert.equal(after.metadata.definedNames[0].refersTo, `${sheet.name}!C1`)
  assert.equal(editor.replaceWorkbookFromView(after), true)
  assert.equal(editor.canUndo, true)
  editor.undo()
  assert.equal(editor.state.workbook.sheets[0].cells['0:1'].value, 'move')
  assert.equal(editor.state.workbook.sheets[0].cells['0:2'], undefined)
  editor.redo()
  assert.equal(editor.state.workbook.sheets[0].cells['0:2'].value, 'move')
  assert.equal(editor.state.workbook.sheets[0].cells['0:2'].formula, '=A1+C1')
  assert.deepEqual(editor.state.workbook.sheets[0].cells['0:2'].authoredStyle, {})
  assert.equal(editor.state.workbook.sheets[0].cells['0:2'].style.fill, '#eeddcc')
  assert.equal(editor.state.workbook.sheets[0].cells['0:1'], undefined)
})

test('adding a sheet applies its initial activation and a canceled save token cannot mark a new unit', () => {
  const editor = freshEditor()
  const token = editor.captureSavepoint()
  const view = toExcelWorkbookView(editor.state.workbook)
  const second = createBlankSheet('Inserted')
  const nextSheet = toExcelWorkbookView({ ...editor.state.workbook, sheets: [second] }).sheets[0]
  view.sheets.push(nextSheet)
  view.activeSheetId = second.id
  editor.replaceWorkbookFromView(view)
  assert.equal(editor.state.workbook.activeSheetId, second.id)
  editor.newWorkbook()
  editor.state.fileName = 'new unit.xlsx'
  editor.markSaved(token)
  assert.equal(editor.state.workbook.modified, true)
  assert.equal(editor.canUndo, true)
})

test('malformed view input fails before changing the model, revision or redo stack', () => {
  const editor = freshEditor()
  editor.state.fileName = 'edit.xlsx'
  editor.undo()
  const original = snapshot(editor.state.workbook)
  const revision = editor.captureSavepoint().revision
  const view = toExcelWorkbookView(editor.state.workbook)
  view.sheets[0].cells.A1 = null
  assert.throws(() => editor.replaceWorkbookFromView(view), TypeError)
  assert.deepEqual(snapshot(editor.state.workbook), original)
  assert.equal(editor.captureSavepoint().revision, revision)
  assert.equal(editor.canUndo, false)
  assert.equal(editor.canRedo, true)
})

test('reactive view payloads and later caller mutations cannot corrupt recorded snapshots', () => {
  const editor = freshEditor()
  const bookView = reactive(toExcelWorkbookView(editor.state.workbook))
  bookView.sheets[0].cells.A1 = { value: 'recorded cell' }
  editor.replaceWorkbookFromView(bookView)
  bookView.sheets[0].cells.A1.value = 'uncommitted external edit'
  assert.equal(editor.state.workbook.sheets[0].cells['0:0'].value, 'recorded cell')
  editor.undo()
  editor.redo()
  assert.equal(editor.state.workbook.sheets[0].cells['0:0'].value, 'recorded cell')
  editor.setActiveKind('document')
  const documentView = reactive(snapshot(editor.state.document))
  documentView.blocks[0].runs[0].text = 'recorded text'
  editor.replaceDocumentFromView(documentView)
  documentView.blocks[0].runs[0].text = 'uncommitted external edit'
  assert.equal(editor.state.document.blocks[0].runs[0].text, 'recorded text')
  editor.undo()
  assert.equal(editor.state.document.blocks[0].runs[0].text, '')
  editor.redo()
  assert.equal(editor.state.document.blocks[0].runs[0].text, 'recorded text')
})

test('only changing a document dirty flag is not a content edit', () => {
  const editor = freshEditor()
  editor.setActiveKind('document')
  const view = snapshot(editor.state.document)
  view.modified = true
  assert.equal(editor.replaceDocumentFromView(view), false)
  assert.equal(editor.state.document.modified, false)
  assert.equal(editor.canUndo, false)
})

test('regional view edits retain one history source through callbacks, undo and redo', () => {
  const editor = freshEditor()
  const model = snapshot(editor.state.workbook)
  const cell = createCell(0, 0, 1234.5)
  cell.style.numberFormat = '#,##0.00'; cell.display = '1.234,50'
  model.sheets[0].cells['0:0'] = cell
  editor.workbook.replace(model)
  const view = toExcelWorkbookView(model)
  view.sheets[0].cells.A1.bold = true
  assert.equal(editor.replaceWorkbookFromView(view, 'de-DE'), true)
  assert.equal(editor.state.workbook.sheets[0].cells['0:0'].display, '1.234,50')
  assert.equal(editor.replaceWorkbookFromView(toExcelWorkbookView(editor.state.workbook), 'de-DE'), false)
  const depth = () => editor.history.units[model.id].undo.length
  assert.equal(depth(), 1)
  const edited = toExcelWorkbookView(editor.state.workbook)
  edited.sheets[0].cells.A1.value = '2000.25'; edited.sheets[0].cells.A1.rawValue = 2000.25
  assert.equal(editor.replaceWorkbookFromView(edited, 'de-DE'), true)
  assert.equal(editor.state.workbook.sheets[0].cells['0:0'].display, '2.000,25')
  assert.equal(depth(), 2)
  editor.undo()
  assert.equal(editor.state.workbook.sheets[0].cells['0:0'].display, '1.234,50')
  assert.equal(editor.state.workbook.sheets[0].cells['0:0'].style.font.bold, true)
  editor.redo()
  assert.equal(editor.state.workbook.sheets[0].cells['0:0'].display, '2.000,25')
  assert.equal(editor.state.workbook.sheets[0].cells['0:0'].value, 2000.25)
})
