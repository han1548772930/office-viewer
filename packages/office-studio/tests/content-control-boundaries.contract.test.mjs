import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

let server
let api
let commands
let controls
let modelUtils

beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/content-control-boundaries', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
  commands = await server.ssrLoadModule('/packages/office-studio/src/features/document/commands/document-commands.ts')
  controls = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/content-controls.ts')
  modelUtils = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
})
afterAll(async () => { await server?.close() })

function fixture(text = 'Left NAME right') {
  const model = api.createBlankDocument()
  model.id = 'sdt-boundary-doc'
  model.blocks[0].id = 'paragraph-1'
  model.blocks[0].runs[0].text = text
  model.contentControls = []
  return model
}
function selection(start, end = start, target = { blockId: 'paragraph-1' }) {
  return { anchor: { ...target, offset: start }, focus: { ...target, offset: end } }
}
function add(model, input = {}) {
  return controls.createDocumentContentControl(model, {
    id: 'name', type: 'text', blockId: 'paragraph-1', startOffset: 5, endOffset: 9, ...input,
  })
}
const textOf = (block) => block.runs.map((run) => run.text).join('')

test('content control selection resolves the innermost equal-range child', () => {
  const model = fixture('NAME')
  add(model, { id: 'group', type: 'group', startOffset: 0, endOffset: 4 })
  add(model, { id: 'child', type: 'text', startOffset: 0, endOffset: 4, parentId: 'group' })
  assert.equal(controls.selectedDocumentContentControl(model, 'paragraph-1', 2)?.id, 'child')
  assert.equal(controls.selectedDocumentContentControl(model, 'paragraph-1', 0, 4)?.id, 'child')
  assert.equal(controls.selectedDocumentContentControl(model, 'paragraph-1', 0, 5), undefined)
})

test('a paragraph break before a control moves its anchors to the suffix paragraph', () => {
  const model = fixture()
  add(model)
  commands.executeDocumentCommand(model, selection(5), { id: 'doc.command.insert-text', text: '\n' })
  assert.deepEqual(model.blocks.map(textOf), ['Left ', 'NAME right'])
  assert.deepEqual(controls.contentControlRange(model.contentControls[0]), {
    blockId: model.blocks[1].id, startOffset: 0, endOffset: 4,
  })
  assert.equal(controls.contentControlText(model, model.contentControls[0]), 'NAME')
})

test('multiline replacements preserve controls before and after UTF-16 text and CRLF boundaries', () => {
  const model = fixture('前😀姓名后Z')
  add(model, { id: 'before', startOffset: 0, endOffset: 1 })
  add(model, { startOffset: 3, endOffset: 5 })
  commands.executeDocumentCommand(model, selection(1, 3), { id: 'doc.command.insert-text', text: 'x\r\ny\r\nz' })
  assert.deepEqual(model.blocks.map(textOf), ['前x', 'y', 'z姓名后Z'])
  assert.deepEqual(model.contentControls.map((control) => [control.id, control.blockId, control.startOffset, control.endOffset]), [
    ['before', 'paragraph-1', 0, 1], ['name', model.blocks[2].id, 1, 3],
  ])
  assert.equal(controls.contentControlText(model, model.contentControls[1]), '姓名')
})

test('breaking after a control keeps its paragraph identity, and breaking inside it is atomic', () => {
  const model = fixture()
  add(model)
  commands.executeDocumentCommand(model, selection(9), { id: 'doc.command.insert-text', text: '\n' })
  assert.equal(model.contentControls[0].blockId, 'paragraph-1')
  assert.equal(controls.contentControlText(model, model.contentControls[0]), 'NAME')
  const before = structuredClone(model)
  assert.throws(() => commands.executeDocumentCommand(model, selection(7), { id: 'doc.command.insert-text', text: '\n' }), /cannot split a content control/)
  assert.deepEqual(model, before)
})

test('paragraph splitting updates table-cell control targets without moving them out of their cell', () => {
  const model = fixture()
  const paragraph = model.blocks[0]
  const table = { ...modelUtils.createParagraph(), id: 'table', type: 'table', runs: [], metadata: {
    table: { rows: [{ id: 'row', cells: [{ id: 'cell', blocks: [paragraph], style: {}, rowSpan: 1, colSpan: 1 }] }] },
  } }
  model.blocks = [table]
  add(model)
  const target = { blockId: 'table', row: 0, col: 0, cellBlockId: 'paragraph-1' }
  commands.executeDocumentCommand(model, selection(5, 5, target), { id: 'doc.command.insert-text', text: '\n' })
  const blocks = table.metadata.table.rows[0].cells[0].blocks
  assert.deepEqual(blocks.map(textOf), ['Left ', 'NAME right'])
  assert.equal(model.contentControls[0].blockId, blocks[1].id)
  assert.equal(controls.contentControlText(model, model.contentControls[0]), 'NAME')
})

test('inserting a block moves suffix controls, and inserting one inside a control is rejected before mutation', () => {
  for (const id of ['doc.command.insert-divider', 'doc.command.insert-table', 'doc.command.insert-page-break']) {
    const command = id === 'doc.command.insert-table' ? { id, rows: 1, cols: 1 } : { id }
    const model = fixture()
    add(model)
    commands.executeDocumentCommand(model, selection(5), command)
    assert.equal(controls.contentControlText(model, model.contentControls[0]), 'NAME')
    assert.notEqual(model.contentControls[0].blockId, 'paragraph-1')
    const invalid = fixture()
    add(invalid)
    const before = structuredClone(invalid)
    assert.throws(() => commands.executeDocumentCommand(invalid, selection(7, 8), command), /cannot split a content control/)
    assert.deepEqual(invalid, before)
  }
})

test('the public facade preserves moved anchors through undo, redo, save and reopen', () => {
  const studio = new api.OfficeStudio({ document: fixture('😀NAME') })
  const document = api.FOffice.newAPI(studio).getActiveDocument()
  const control = document.addContentControl({ id: 'name', type: 'text', blockId: 'paragraph-1', startOffset: 2, endOffset: 6 })
  document.insertText(2, 'a\r\nb')
  const moved = control.getRange()
  assert.equal(document.getText(), '😀a\nbNAME')
  assert.deepEqual([moved.startOffset, moved.endOffset], [1, 5])
  assert.equal(control.getText(), 'NAME')
  assert.equal(document.undo(), true)
  assert.deepEqual(control.getRange(), { blockId: 'paragraph-1', startOffset: 2, endOffset: 6 })
  assert.equal(document.redo(), true)
  assert.deepEqual(control.getRange(), moved)
  const snapshot = document.save()
  const reopened = new api.OfficeStudio({ document: snapshot })
  assert.equal(api.FOffice.newAPI(reopened).getActiveDocument().getContentControl('name').getText(), 'NAME')
  const before = structuredClone(document.getModel())
  assert.throws(() => document.insertText(6, '\n'), /cannot split a content control/)
  assert.deepEqual(document.getModel(), before)
  assert.equal(document.undo(), true)
  assert.equal(document.getText(), '😀NAME', 'a rejected edit must not add a history entry')
  studio.dispose()
  reopened.dispose()
})

test('a rejected anchor transformation leaves all earlier ranges unchanged', () => {
  const model = fixture('ABCDEFGHIJ')
  add(model, { id: 'first', startOffset: 1, endOffset: 3 })
  add(model, { id: 'locked', startOffset: 5, endOffset: 8, lock: 'contentLocked' })
  const before = structuredClone(model)
  assert.throws(() => controls.transformContentControlsForBlockEdit(model, 'paragraph-1', 2, 6, 1), /content is locked/)
  assert.deepEqual(model, before)
})

test('the controlLocked alias prevents wrapper deletion through service and ordinary text APIs', () => {
  const model = fixture()
  add(model, { lock: 'controlLocked' })
  const before = structuredClone(model)
  assert.throws(() => controls.removeDocumentContentControl(model, 'name'), /wrapper is locked/)
  assert.throws(() => commands.executeDocumentCommand(model, selection(5, 9), { id: 'doc.command.delete', direction: 1 }), /wrapper is locked/)
  assert.deepEqual(model, before)
  controls.updateDocumentContentControl(model, 'name', { lock: 'unlocked' })
  controls.removeDocumentContentControl(model, 'name')
  assert.deepEqual(model.contentControls, [])
})

test('replacing a container value cannot erase nested controls even when text length is unchanged', () => {
  for (const value of ['ABCD', 'longer value']) {
    const model = fixture('ABCD')
    add(model, { id: 'group', type: 'group', startOffset: 0, endOffset: 4 })
    add(model, { id: 'child', startOffset: 1, endOffset: 3, parentId: 'group' })
    // The same direct replacement guard applies to any imported text-like container with children.
    model.contentControls[0].type = 'richText'
    const before = structuredClone(model)
    assert.throws(() => controls.setDocumentContentControlValue(model, 'group', value), /nested control/)
    assert.deepEqual(model, before)
  }
})

test('a parent and its child migrate together when a paragraph break precedes their ranges', () => {
  const model = fixture('Left NAME right')
  add(model, { id: 'group', type: 'group', startOffset: 5, endOffset: 9 })
  add(model, { id: 'child', startOffset: 6, endOffset: 8, parentId: 'group' })
  commands.executeDocumentCommand(model, selection(5), { id: 'doc.command.insert-text', text: '\n' })
  assert.deepEqual(model.contentControls.map((control) => [control.id, control.blockId, control.startOffset, control.endOffset]), [
    ['group', model.blocks[1].id, 0, 4], ['child', model.blocks[1].id, 1, 3],
  ])
  assert.equal(controls.contentControlText(model, model.contentControls[1]), 'AM')
})

test('unchanged property fields do not prevent unlocking a wrapper from the properties form', () => {
  const model = fixture()
  add(model, { title: 'Customer', placeholder: 'Name', lock: 'sdtLocked' })
  controls.updateDocumentContentControl(model, 'name', { title: 'Customer', placeholder: 'Name', lock: 'unlocked' })
  assert.equal(model.contentControls[0].lock, 'unlocked')
  controls.removeDocumentContentControl(model, 'name')
  assert.deepEqual(model.contentControls, [])
})

test('text and checkbox value commands cannot bypass a locked repeating-section ancestor', () => {
  for (const type of ['text', 'checkbox']) {
    const model = fixture('ABCD')
    add(model, { id: 'section', type: 'repeatingSection', startOffset: 0, endOffset: 4, lock: 'contentLocked' })
    add(model, { id: 'child', type, startOffset: 1, endOffset: 3, parentId: 'section' })
    const before = structuredClone(model)
    const edit = type === 'text'
      ? () => controls.setDocumentContentControlValue(model, 'child', 'new')
      : () => controls.setDocumentContentControlChecked(model, 'child', true)
    assert.throws(edit, /content is locked: section/)
    assert.deepEqual(model, before)
  }
})

test('empty children follow their wrapper when block insertion splits at a shared boundary', () => {
  const model = fixture()
  add(model, { id: 'group', type: 'group', startOffset: 5, endOffset: 9 })
  add(model, { id: 'empty-child', startOffset: 5, endOffset: 5, parentId: 'group' })
  commands.executeDocumentCommand(model, selection(5), { id: 'doc.command.insert-divider' })
  const group = model.contentControls.find((control) => control.id === 'group')
  const child = model.contentControls.find((control) => control.id === 'empty-child')
  assert.equal(child.blockId, group.blockId)
  assert.notEqual(group.blockId, 'paragraph-1')
  assert.deepEqual([child.startOffset, child.endOffset, group.startOffset, group.endOffset], [0, 0, 0, 4])
})

test('filling an empty child grows its empty parent rather than displacing the parent', () => {
  const model = fixture('AB')
  add(model, { id: 'group', type: 'group', startOffset: 1, endOffset: 1 })
  add(model, { id: 'child', startOffset: 1, endOffset: 1, parentId: 'group' })
  controls.setDocumentContentControlValue(model, 'child', 'name')
  assert.equal(textOf(model.blocks[0]), 'AnameB')
  assert.deepEqual(model.contentControls.map((control) => [control.startOffset, control.endOffset]), [[1, 5], [1, 5]])
})
