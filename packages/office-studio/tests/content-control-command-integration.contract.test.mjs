import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// DOC-11 integration drafts. Execution is deferred until the unified validation phase.
let server, api, editor, controls, code, search
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/content-control-command-integration', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
  editor = await server.ssrLoadModule('/packages/office-studio/src/api/document-editor-commands.ts')
  controls = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/content-controls.ts')
  code = await server.ssrLoadModule('/packages/office-studio/src/features/document/code-blocks/code-blocks.ts')
  search = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/document-search.ts')
})
afterAll(async () => { await server?.close() })

const copy = value => JSON.parse(JSON.stringify(value))
const selection = (start, end = start, blockId = 'p') => ({ anchor: { blockId, offset: start }, focus: { blockId, offset: end } })
const text = model => model.blocks.map(block => block.runs.map(run => run.text).join('')).join('\n')
const info = model => model.contentControls.find(control => control.id === 'control')
function fixture(value = 'Name', properties = {}) {
  const model = api.createBlankDocument(); model.id = 'sdt-command-integration'
  model.blocks[0].id = 'p'; model.blocks[0].runs[0].text = value; model.contentControls = []
  controls.createDocumentContentControl(model, { id: 'control', type: 'text', blockId: 'p', startOffset: 0,
    endOffset: value.length, value, metadata: { retainedSource: { part: 'word/document.xml' } }, ...properties })
  return model
}
function open(model) {
  const studio = new api.OfficeStudio({ document: model, currentUserId: 'editor' }); studio.activateDocument(model.id)
  return { studio, document: new api.FDocument(studio, model.id) }
}
const permission = { documentId: 'sdt-command-integration', userId: 'editor' }

test('the shared editor rejects ordinary dropdown edits without mutating its caller-owned candidate', () => {
  const model = fixture('Alpha', { type: 'dropDownList', value: 'A', items: [{ value: 'A', displayText: 'Alpha' }] }), before = copy(model)
  assert.throws(() => editor.applyDocumentEditorCommand(model, { unitId: model.id, selection: selection(0, 5),
    command: { id: 'doc.command.insert-text', text: 'Other' } }, permission), /下拉列表/)
  assert.deepEqual(model, before)
})

test('persistent full-range deletion retains an empty shell and the next explicit inside input retains its identity', () => {
  const { studio, document } = open(fixture('AB', { dataBinding: { xpath: '/name', dirty: false } }))
  const before = copy(document.getModel())
  try {
    const deleted = document.executeCommand({ id: 'doc.command.delete', direction: -1, activeContentControlId: 'control' }, selection(0, 2))
    assert.equal(deleted.changed, true); assert.equal(document.getText(), '')
    assert.equal(info(document.getModel()).id, 'control'); assert.equal(info(document.getModel()).value, '')
    const inserted = document.executeCommand({ id: 'doc.command.insert-text', text: '李😀', activeContentControlId: 'control' }, deleted.selection)
    assert.equal(inserted.changed, true); assert.equal(document.getText(), '李😀')
    assert.equal(info(document.getModel()).startOffset, 0); assert.equal(info(document.getModel()).endOffset, 3)
    assert.equal(info(document.getModel()).value, '李😀'); assert.equal(info(document.getModel()).dataBinding.dirty, true)
    document.undo(); assert.equal(document.getText(), ''); assert.equal(info(document.getModel()).endOffset, 0)
    document.undo(); assert.deepEqual(document.getModel().contentControls, before.contentControls)
    document.redo(); document.redo(); assert.equal(info(document.getModel()).value, '李😀')
  } finally { studio.dispose() }
})

test('a new caret at a nonempty boundary edits outside while an explicit selected control continues inside', () => {
  const outside = fixture('AB'), inside = copy(outside)
  editor.applyDocumentEditorCommand(outside, { unitId: outside.id, selection: selection(2),
    command: { id: 'doc.command.insert-text', text: 'X' } }, permission)
  assert.equal(text(outside), 'ABX'); assert.equal(info(outside).endOffset, 2); assert.equal(info(outside).value, 'AB')
  editor.applyDocumentEditorCommand(inside, { unitId: inside.id, selection: selection(2),
    command: { id: 'doc.command.insert-text', text: 'X', activeContentControlId: 'control' } }, permission)
  assert.equal(text(inside), 'ABX'); assert.equal(info(inside).endOffset, 3); assert.equal(info(inside).value, 'ABX')
  editor.applyDocumentEditorCommand(inside, { unitId: inside.id, selection: selection(3),
    command: { id: 'doc.command.insert-text', text: 'Y', activeContentControlId: 'control' } }, permission)
  assert.equal(info(inside).value, 'ABXY'); assert.equal(info(inside).endOffset, 4)
})

test('empty input does not replace placeholder text or unwrap temporary controls', () => {
  const model = fixture('Name', { showingPlaceholder: true, placeholder: 'Name', temporary: true }), before = copy(model)
  const next = editor.applyDocumentEditorCommand(model, { unitId: model.id, selection: selection(2),
    command: { id: 'doc.command.insert-text', text: '' } }, permission)
  assert.equal(text(model), 'Name'); assert.deepEqual(model.contentControls, before.contentControls)
  assert.deepEqual(next, selection(2))
})

test('BeforeCommandExecute selection reentry preserves the newer caret and prevents stale ordinary input', () => {
  const { studio, document } = open(fixture('AB', { temporary: true }))
  document.setInputSelection(selection(1)); const before = copy(document.getModel())
  const subscription = studio.events.addEvent('BeforeCommandExecute', event => {
    if (event.id === 'doc.command.insert-text') document.setInputSelection(selection(0))
  })
  try {
    assert.throws(() => document.executeCommand({ id: 'doc.command.insert-text', text: 'X' }, selection(1)), /输入选区已更改/)
    assert.deepEqual(document.getModel(), before)
    assert.deepEqual(document.getInputState().location.selection, selection(0))
    assert.equal(studio.editor.history.units[document.id]?.undo.length ?? 0, 0)
  } finally { subscription.dispose(); studio.dispose() }
})

test('find-all replacement rolls back earlier candidates when another control rejects the same batch', () => {
  const model = fixture('A', { type: 'dropDownList', value: 'A', items: [{ value: 'A' }] })
  const other = copy(model.blocks[0]); other.id = 'plain'; other.runs[0].text = 'A'; model.blocks.push(other)
  const before = copy(model), matches = search.findDocumentMatches(model, 'A')
  assert.equal(matches.length, 2)
  assert.throws(() => search.replaceDocumentMatches(model, matches, 'B', permission), /下拉列表/)
  assert.deepEqual(model, before)
})

test('code-source editing updates control values and temporary unwrapping while display-only changes retain the shell', () => {
  const model = api.createBlankDocument(); model.id = 'sdt-command-integration'; model.contentControls = []
  const block = code.createCodeBlock('AB'); block.id = 'p'; model.blocks = [block]
  controls.createDocumentContentControl(model, { id: 'control', type: 'richText', blockId: 'p', startOffset: 0,
    endOffset: 2, value: 'AB', temporary: true })
  assert.equal(code.updateCodeBlock(model, 'p', { lineNumbers: true }), true)
  assert.equal(info(model).temporary, true); assert.equal(info(model).value, 'AB')
  assert.equal(code.updateCodeBlock(model, 'p', { source: 'A😀B' }), true)
  assert.deepEqual(model.contentControls, []); assert.equal(text(model), 'A😀B')
  assert.equal(code.codeBlockData(model.blocks[0]).source, 'A😀B')
})

test('code-block keyboard input uses the same placeholder and typed control finalizer as paragraphs', () => {
  const model = api.createBlankDocument(); model.id = 'sdt-command-integration'; model.contentControls = []
  const block = code.createCodeBlock('Prompt'); block.id = 'p'; model.blocks = [block]
  controls.createDocumentContentControl(model, { id: 'control', type: 'richText', blockId: 'p', startOffset: 0,
    endOffset: 6, value: 'Prompt', placeholder: 'Prompt', showingPlaceholder: true })
  const caret = editor.applyDocumentEditorCommand(model, { unitId: model.id, selection: selection(3),
    command: { id: 'doc.command.insert-text', text: 'let x' } }, permission)
  assert.equal(text(model), 'let x'); assert.equal(code.codeBlockData(model.blocks[0]).source, 'let x')
  assert.equal(info(model).value, 'let x'); assert.equal(info(model).showingPlaceholder, false)
  assert.equal(caret.focus.offset, 5)
})
