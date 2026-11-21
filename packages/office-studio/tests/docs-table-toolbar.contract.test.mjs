import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createPinia } from 'pinia'
import { createServer } from 'vite'

let vite
let api
let commands
let models
let controlService
let clipboard
let toolbar
let ribbon
let useEditorStore
beforeAll(async () => {
  vite = await createServer({ cacheDir: 'node_modules/.vite/docs-table-toolbar', optimizeDeps: { noDiscovery: true }, server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  api = await vite.ssrLoadModule('/packages/office-studio/src/index.ts')
  commands = await vite.ssrLoadModule('/packages/office-studio/src/features/document/commands/document-commands.ts')
  models = await vite.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
  controlService = await vite.ssrLoadModule('/packages/office-studio/src/features/document/services/content-controls.ts')
  clipboard = await vite.ssrLoadModule('/packages/office-studio/src/features/document/services/document-clipboard.ts')
  toolbar = await vite.ssrLoadModule('/packages/office-studio/src/components/document/document-toolbar-state.ts')
  ribbon = await vite.ssrLoadModule('/packages/office-studio/src/components/document/document-ribbon.ts')
  ;({ useEditorStore } = await vite.ssrLoadModule('/packages/office-studio/src/stores/editorStore.ts'))
})
afterAll(async () => { await vite?.close() })

function fixture(rows = 3, columns = 3) {
  const model = api.createBlankDocument()
  const before = { ...models.createParagraph('Before'), id: 'before' }
  const after = { ...models.createParagraph('After'), id: 'after' }
  const table = models.createTable(rows, columns)
  table.rows.forEach((row, r) => row.cells.forEach((cell, c) => {
    cell.blocks[0].id = `r${r}c${c}`
    cell.blocks[0].runs = [models.createRun(`R${r}C${c}`)]
    cell.style.fill = '#eef2ff'
  }))
  table.columnWidths = Array.from({ length: columns }, (_, i) => (i + 1) * 100)
  table.headerRowCount = 1
  model.blocks = [before, { ...models.createParagraph(), id: 'table', type: 'table', runs: [], metadata: { table } }, after]
  model.contentControls = []
  return { model, table }
}
function selection(model, row = 0, col = 0, endRow = row, endCol = col) {
  const table = models.tableForBlock(model.blocks.find((block) => block.id === 'table'))
  return {
    anchor: { blockId: 'table', row, col, cellBlockId: table.rows[row].cells[col].blocks[0].id, offset: 0 },
    focus: { blockId: 'table', row: endRow, col: endCol, cellBlockId: table.rows[endRow].cells[endCol].blocks[0].id, offset: 0 },
  }
}
function apply(model, range, id, params = {}) { return commands.executeDocumentCommand(model, range, { id: `doc.command.table-${id}`, ...params }) }
const cellText = (cell) => cell.blocks.flatMap((block) => block.runs).map((run) => run.text).join('')

test('table context follows logical table selection and the contextual toolbar exposes executable structural commands', () => {
  const { model } = fixture()
  const state = toolbar.documentToolbarState(model, selection(model, 0, 1, 1, 2))
  assert.deepEqual(state.table, { blockId: 'table', startRow: 0, endRow: 1, startColumn: 1, endColumn: 2, rowCount: 3, columnCount: 3, hasMergedCells: false })
  const context = { state, canUndo: false, canRedo: false, outline: false, focus: false, paper: 'A4', orientation: 'portrait', zoom: 100, contentControlsAvailable: true }
  const tabs = ribbon.createDocumentRibbon(context)
  const tab = tabs.find((candidate) => candidate.id === 'table')
  assert.equal(tab.contextual, true)
  assert.equal(tab.title, '表格工具')
  const items = tab.groups.flatMap((group) => group.items)
  assert.equal(items.length, 7)
  for (const item of items) assert.ok(ribbon.documentCommandFromRibbon(item.commandId, item.params), item.title)
  assert.equal(ribbon.documentCommandFromRibbon('doc.command.table-insert-row', { position: 'inside' }), undefined)
  const plain = toolbar.documentToolbarState(model, models.collapsedSelection({ blockId: 'before', offset: 0 }))
  assert.equal(plain.table, undefined)
  assert.equal(ribbon.createDocumentRibbon({ ...context, state: plain }).some((candidate) => candidate.id === 'table'), false)
})

test('inserted rows preserve the table style, existing cells and header semantics, while receiving unique empty models', () => {
  const { model, table } = fixture()
  const original = structuredClone(model)
  const result = apply(model, selection(model, 0, 1), 'insert-row', { position: 'before' })
  assert.equal(table.rows.length, 4)
  assert.equal(table.headerRowCount, 2)
  assert.deepEqual(table.rows.slice(1), original.blocks[1].metadata.table.rows)
  assert.ok(table.rows[0].cells.every((cell) => cellText(cell) === '' && cell.style.fill === '#eef2ff'))
  assert.equal(result.focus.row, 0)
  assert.equal(result.focus.col, 1)
  assert.equal(result.focus.cellBlockId, table.rows[0].cells[1].blocks[0].id)
  const ids = table.rows.flatMap((row) => [row.id, ...row.cells.flatMap((cell) => [cell.id, ...cell.blocks.map((block) => block.id)])])
  assert.equal(new Set(ids).size, ids.length)
  assert.deepEqual([model.blocks[0], model.blocks[2]], [original.blocks[0], original.blocks[2]])
})

test('inserted columns update widths and retain existing paragraph identities and control anchors', () => {
  const { model, table } = fixture()
  controlService.createDocumentContentControl(model, { id: 'anchor', type: 'text', blockId: 'r0c2', startOffset: 0, endOffset: 4 })
  const original = structuredClone(table.rows)
  const range = apply(model, selection(model, 1, 1), 'insert-column', { position: 'after' })
  assert.deepEqual(table.columnWidths, [100, 200, 200, 300])
  assert.equal(range.focus.col, 2)
  for (let r = 0; r < 3; r++) {
    assert.deepEqual(table.rows[r].cells[3], original[r].cells[2])
    assert.equal(cellText(table.rows[r].cells[2]), '')
  }
  assert.equal(controlService.contentControlText(model, model.contentControls[0]), 'R0C2')
})

test('clamping a caret after structural undo retains its table cell coordinate when paragraph IDs change', () => {
  const { model } = fixture()
  const old = { blockId: 'table', row: 1, col: 2, cellBlockId: 'removed-paragraph', offset: 999 }
  assert.deepEqual(models.clampPoint(model, old), { blockId: 'table', row: 1, col: 2, cellBlockId: 'r1c2', offset: 4 })
  assert.deepEqual(models.clampPoint(model, { ...old, row: 99, col: 99 }), { blockId: 'table', row: 2, col: 2, cellBlockId: 'r2c2', offset: 4 })
})

test('deleting a row range adjusts repeating headers, removes unlocked anchors, and selects the surviving row', () => {
  const { model, table } = fixture(4, 2)
  table.headerRowCount = 2
  controlService.createDocumentContentControl(model, { id: 'remove', type: 'text', blockId: 'r0c0', startOffset: 0, endOffset: 4 })
  controlService.createDocumentContentControl(model, { id: 'keep', type: 'text', blockId: 'r3c1', startOffset: 0, endOffset: 4, lock: 'sdtLocked' })
  const result = apply(model, selection(model, 0, 0, 1, 1), 'delete-rows')
  assert.equal(table.headerRowCount, 0)
  assert.equal(table.rows.length, 2)
  assert.equal(cellText(table.rows[0].cells[0]), 'R2C0')
  assert.equal(result.focus.cellBlockId, 'r2c0')
  assert.deepEqual(model.contentControls.map((control) => control.id), ['keep'])
})

test('deleting a column range updates widths and keeps the caret inside the remaining grid', () => {
  const { model, table } = fixture(2, 4)
  const result = apply(model, selection(model, 0, 1, 1, 2), 'delete-columns')
  assert.deepEqual(table.columnWidths, [100, 400])
  assert.deepEqual(table.rows.map((row) => row.cells.map(cellText)), [['R0C0', 'R0C3'], ['R1C0', 'R1C3']])
  assert.equal(result.focus.col, 1)
  assert.equal(result.focus.cellBlockId, 'r0c3')
})

test('deleting the last row or column removes the table and places input in the following paragraph', () => {
  for (const id of ['delete-rows', 'delete-columns']) {
    const { model } = fixture(1, 1)
    const result = apply(model, selection(model), id)
    assert.deepEqual(model.blocks.map((block) => block.id), ['before', 'after'])
    assert.deepEqual(result, models.collapsedSelection({ blockId: 'after', offset: 0 }))
  }
})

test('deleting the only table leaves one real editable paragraph', () => {
  const { model } = fixture(1, 1)
  model.blocks = [model.blocks[1]]
  const result = apply(model, selection(model), 'delete')
  assert.equal(model.blocks.length, 1)
  assert.equal(model.blocks[0].type, 'paragraph')
  assert.equal(result.focus.blockId, model.blocks[0].id)
  assert.deepEqual(models.textTargets(model), [{ blockId: model.blocks[0].id }])
})

test('deleting a row, column, or table preflights every locked descendant without partial anchor removal', () => {
  for (const id of ['delete-rows', 'delete-columns', 'delete']) {
    const { model } = fixture(2, 2)
    controlService.createDocumentContentControl(model, { id: 'unlocked', type: 'text', blockId: 'r0c0', startOffset: 0, endOffset: 4 })
    controlService.createDocumentContentControl(model, { id: 'locked', type: 'text', blockId: 'r0c1', startOffset: 0, endOffset: 4, lock: 'controlLocked' })
    const original = structuredClone(model)
    assert.throws(() => apply(model, selection(model, 0, 0, 0, 1), id), /wrapper is locked/)
    assert.deepEqual(model, original)
  }
})

test('table deletion also checks controls in nested table descendants', () => {
  const { model, table } = fixture(2, 2)
  const nested = models.createTable(1, 1)
  nested.rows[0].cells[0].blocks[0].id = 'nested-paragraph'
  const block = { ...models.createParagraph(), type: 'table', runs: [], metadata: { table: nested } }
  table.rows[0].cells[1].blocks.push(block)
  model.contentControls = [{ id: 'nested-lock', type: 'text', blockId: 'nested-paragraph', startOffset: 0, endOffset: 0, lock: 'sdtLocked' }]
  const original = structuredClone(model)
  assert.throws(() => apply(model, selection(model), 'delete-rows'), /wrapper is locked/)
  assert.deepEqual(model, original)
})

test('unsupported merged and irregular grid mutations return explicit errors and leave the complete model intact', () => {
  for (const kind of ['merged', 'irregular']) {
    for (const id of ['insert-row', 'insert-column', 'delete-rows', 'delete-columns']) {
      const { model, table } = fixture()
      if (kind === 'merged') table.rows[1].cells[1].colSpan = 2
      else table.rows[1].cells.pop()
      const original = structuredClone(model)
      assert.throws(() => apply(model, selection(model), id, { position: 'after' }), /合并单元格|网格不规则/)
      assert.deepEqual(model, original)
    }
  }
})

test('a structural action is one document-history item and survives save snapshot reopening', () => {
  const editor = useEditorStore(createPinia())
  const { model } = fixture()
  editor.document.replace({ ...model, id: editor.state.document.id })
  editor.setActiveKind('document')
  const original = JSON.parse(JSON.stringify(editor.document.model))
  const draft = structuredClone(original)
  apply(draft, selection(draft), 'insert-column', { position: 'before' })
  assert.equal(editor.replaceDocumentFromView(draft), true)
  assert.equal(models.tableForBlock(editor.document.model.blocks[1]).rows[0].cells.length, 4)
  editor.undo()
  assert.deepEqual(JSON.parse(JSON.stringify(editor.document.model)), original)
  editor.redo()
  const studio = new api.OfficeStudio({ document: JSON.parse(JSON.stringify(editor.document.model)) })
  const snapshot = api.FOffice.newAPI(studio).getActiveDocument().save()
  assert.equal(models.tableForBlock(snapshot.blocks[1]).rows[0].cells.length, 4)
  studio.dispose()
})

test('source-formatted single-paragraph paste preserves runs inline and adjusts control anchors', () => {
  const model = api.createBlankDocument()
  model.blocks = [{ ...models.createParagraph('AB NAME'), id: 'paragraph' }]
  controlService.createDocumentContentControl(model, { id: 'anchor', type: 'text', blockId: 'paragraph', startOffset: 3, endOffset: 7 })
  const source = models.createParagraph('红色', { bold: true, color: '#dc2626' })
  const result = commands.executeDocumentCommand(model, models.collapsedSelection({ blockId: 'paragraph', offset: 2 }), { id: 'doc.command.paste-blocks', blocks: [source] })
  assert.equal(model.blocks.length, 1)
  assert.equal(models.runText(model.blocks[0].runs), 'AB红色 NAME')
  assert.ok(model.blocks[0].runs.some((run) => run.text === '红色' && run.bold && run.color === '#dc2626'))
  assert.equal(result.focus.offset, 4)
  assert.deepEqual([model.contentControls[0].startOffset, model.contentControls[0].endOffset], [5, 9])
  assert.equal(controlService.contentControlText(model, model.contentControls[0]), 'NAME')
})

test('the ribbon exposes clipboard modes, real character pickers, and only supported content-control creation types', () => {
  const model = api.createBlankDocument()
  const state = toolbar.documentToolbarState(model, models.collapsedSelection({ blockId: model.blocks[0].id, offset: 0 }))
  const items = ribbon.createDocumentRibbon({ state, canUndo: false, canRedo: false, outline: false, focus: false, paper: 'A4', orientation: 'portrait', zoom: 100, contentControlsAvailable: true }).flatMap((tab) => tab.groups.flatMap((group) => group.items))
  assert.equal(items.find((item) => item.id === 'doc-cut').disabled, true)
  assert.equal(items.find((item) => item.id === 'doc-paste').commandId, 'doc.command.paste')
  assert.ok(items.find((item) => item.id === 'doc-paste').children.some((item) => item.commandId === 'doc.command.paste-text'))
  assert.equal(items.find((item) => item.id === 'doc-symbol').params, 'symbol')
  assert.equal(items.find((item) => item.id === 'doc-emoji').params, 'emoji')
  assert.deepEqual(items.find((item) => item.id === 'doc-controls').children.map((item) => item.params), ['text', 'richText', 'checkbox', 'date', 'dropDownList', 'comboBox'])
})

test('clipboard reads preserve HTML when available and fall back to the plain-text browser contract', async () => {
  let textReads = 0
  const reader = {
    read: async () => [{ types: ['text/plain', 'text/html'], getType: async (type) => new Blob([type === 'text/html' ? '<p><b>Source</b></p>' : 'Source']) }],
    readText: async () => { textReads++; return 'Plain' },
  }
  assert.deepEqual(await clipboard.readDocumentClipboard(reader), { text: 'Source', html: '<p><b>Source</b></p>' })
  assert.equal(textReads, 0)
  assert.deepEqual(await clipboard.readDocumentClipboard(reader, true), { text: 'Plain' })
  assert.deepEqual(await clipboard.readDocumentClipboard({ ...reader, read: async () => { throw new Error('No HTML permission') } }), { text: 'Plain' })
  await assert.rejects(() => clipboard.readDocumentClipboard({ readText: async () => { throw new Error('Permission denied') } }), /Permission denied/)
  assert.deepEqual(clipboard.documentPasteCommand({ text: 'Plain' }, true), { id: 'doc.command.insert-text', text: 'Plain' })
  assert.equal(clipboard.documentPasteCommand({ text: '' }, true), undefined)
})
