import assert from 'node:assert/strict'
import { after, before, test } from 'node:test'
import { createPinia } from 'pinia'
import { createServer } from 'vite'

let server
let commands
let modelUtils
let domain
let useEditorStore

before(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  commands = await server.ssrLoadModule('/packages/office-studio/src/features/document/commands/document-commands.ts')
  modelUtils = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
  domain = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
  ;({ useEditorStore } = await server.ssrLoadModule('/packages/office-studio/src/stores/editorStore.ts'))
})
after(async () => { await server?.close() })

function paragraph(id, text) { return { ...modelUtils.createParagraph(text), id } }
function documentWith(blocks) { return { ...domain.createBlankDocument(), id: 'boundary-doc', blocks } }
const snapshot = (value) => JSON.parse(JSON.stringify(value))
function selection(from, fromOffset, to, toOffset) {
  return { anchor: { blockId: from, offset: fromOffset }, focus: { blockId: to, offset: toOffset } }
}

test('forward and reverse paragraph ranges exclude the next paragraph at offset zero', () => {
  for (const reverse of [false, true]) {
    const model = documentWith([paragraph('one', 'Alpha'), paragraph('two', 'Beta'), paragraph('three', 'Gamma')])
    const range = selection('one', 2, 'two', 0)
    if (reverse) [range.anchor, range.focus] = [range.focus, range.anchor]
    const untouched = structuredClone(model.blocks.slice(1))
    assert.deepEqual(commands.selectedTargets(model, range), [{ blockId: 'one' }])
    commands.executeDocumentCommand(model, range, { id: 'doc.command.format-paragraph', patch: { align: 'center', indent: 2 } })
    assert.equal(model.blocks[0].style.align, 'center')
    assert.equal(model.blocks[0].style.indent, 2)
    assert.deepEqual(model.blocks.slice(1), untouched)
  }
})

test('a positive end offset includes the final paragraph, while a collapsed caret includes only its own paragraph', () => {
  const model = documentWith([paragraph('one', 'Alpha'), paragraph('two', 'Beta'), paragraph('three', 'Gamma')])
  commands.executeDocumentCommand(model, selection('one', 2, 'two', 1), { id: 'doc.command.format-paragraph', patch: { align: 'right' } })
  assert.deepEqual(model.blocks.map((block) => block.style.align), ['right', 'right', 'left'])
  commands.executeDocumentCommand(model, selection('two', 0, 'two', 0), { id: 'doc.command.format-paragraph', patch: { lineHeight: 2 } })
  assert.deepEqual(model.blocks.map((block) => block.style.lineHeight), [1.5, 2, 1.5])
})

test('heading and list commands preserve the first unselected paragraph and its text runs', () => {
  for (const type of ['heading', 'bullet', 'numbered', 'quote']) {
    const model = documentWith([paragraph('one', 'Alpha'), paragraph('two', 'Beta')])
    const untouched = structuredClone(model.blocks[1])
    commands.executeDocumentCommand(model, selection('one', 0, 'two', 0), { id: 'doc.command.set-block-type', type, level: 1 })
    assert.equal(model.blocks[0].type, type)
    assert.deepEqual(model.blocks[1], untouched)
  }
})

test('table-cell paragraph formatting obeys the same half-open boundary', () => {
  const table = paragraph('table', '')
  table.type = 'table'
  table.runs = []
  table.metadata.table = { rows: [{ id: 'row', cells: [{ id: 'cell', blocks: [paragraph('cell-one', 'Alpha'), paragraph('cell-two', 'Beta')], style: {}, rowSpan: 1, colSpan: 1 }] }] }
  const model = documentWith([table, paragraph('after', 'After')])
  const cell = table.metadata.table.rows[0].cells[0]
  const untouched = structuredClone(cell.blocks[1])
  const range = {
    anchor: { blockId: 'table', row: 0, col: 0, cellBlockId: 'cell-one', offset: 1 },
    focus: { blockId: 'table', row: 0, col: 0, cellBlockId: 'cell-two', offset: 0 },
  }
  commands.executeDocumentCommand(model, range, { id: 'doc.command.format-paragraph', patch: { align: 'center' } })
  assert.equal(cell.blocks[0].style.align, 'center')
  assert.deepEqual(cell.blocks[1], untouched)
  assert.equal(model.blocks[1].style.align, 'left')
})

test('character formatting leaves excluded paragraph run IDs intact and collapsed formatting is a no-op', () => {
  const model = documentWith([paragraph('one', 'Alpha'), paragraph('two', 'Beta')])
  const untouched = structuredClone(model.blocks[1])
  commands.executeDocumentCommand(model, selection('one', 1, 'two', 0), { id: 'doc.command.format-text', patch: { bold: true } })
  assert.equal(model.blocks[0].runs.map((run) => run.text).join(''), 'Alpha')
  assert.equal(model.blocks[0].runs.at(-1).bold, true)
  assert.deepEqual(model.blocks[1], untouched)
  const before = structuredClone(model)
  commands.executeDocumentCommand(model, selection('two', 0, 'two', 0), { id: 'doc.command.format-text', patch: { italic: true } })
  assert.deepEqual(model, before)
})

test('replacement through a next-paragraph offset-zero boundary still preserves its complete tail', () => {
  const model = documentWith([paragraph('one', 'Alpha'), paragraph('two', 'Beta'), paragraph('three', 'Gamma')])
  commands.executeDocumentCommand(model, selection('one', 2, 'two', 0), { id: 'doc.command.insert-text', text: 'X' })
  assert.deepEqual(model.blocks.map((block) => [block.id, block.runs.map((run) => run.text).join('')]), [['one', 'AlXBeta'], ['three', 'Gamma']])
})

test('a replacement that relocates the boundary paragraph cannot bypass its content controls', () => {
  const model = documentWith([paragraph('one', 'Alpha'), paragraph('two', 'Beta')])
  model.contentControls = [{ id: 'boundary-control', type: 'text', blockId: 'two', startOffset: 0, endOffset: 4, lock: 'sdtContentLocked' }]
  const before = structuredClone(model)
  assert.throws(() => commands.executeDocumentCommand(model, selection('one', 2, 'two', 0), { id: 'doc.command.insert-text', text: 'X' }), /content control paragraphs/)
  assert.deepEqual(model, before)
})

test('an empty paragraph selected through its paragraph break is included', () => {
  const model = documentWith([paragraph('empty', ''), paragraph('next', 'Next')])
  commands.executeDocumentCommand(model, selection('empty', 0, 'next', 0), { id: 'doc.command.set-block-type', type: 'heading', level: 2 })
  assert.equal(model.blocks[0].type, 'heading')
  assert.equal(model.blocks[0].level, 2)
  assert.equal(model.blocks[1].type, 'paragraph')
})

test('mixed list selections apply one list type to the range and toggle off only when all selected paragraphs match', () => {
  for (const type of ['bullet', 'numbered']) {
    const model = documentWith([paragraph('one', 'Already listed'), paragraph('two', 'Plain paragraph'), paragraph('tail', 'Unselected')])
    model.blocks[0].type = type
    model.blocks[0].level = 2
    model.blocks[2].type = type
    const tail = structuredClone(model.blocks[2])
    const range = selection('one', 0, 'tail', 0)
    commands.executeDocumentCommand(model, range, { id: 'doc.command.set-block-type', type })
    assert.deepEqual(model.blocks.slice(0, 2).map((block) => block.type), [type, type])
    assert.deepEqual(model.blocks.slice(0, 2).map((block) => block.level), [2, 0], 'existing list nesting survives reapplication')
    assert.deepEqual(model.blocks[2], tail)
    commands.executeDocumentCommand(model, range, { id: 'doc.command.set-block-type', type })
    assert.deepEqual(model.blocks.slice(0, 2).map((block) => block.type), ['paragraph', 'paragraph'])
    assert.deepEqual(model.blocks.slice(0, 2).map((block) => block.level), [0, 0])
    assert.deepEqual(model.blocks[2], tail)
  }
})

test('paragraph range commands replay through the shared document history and return to the saved revision', () => {
  const editor = useEditorStore(createPinia())
  editor.document.replace({ ...editor.state.document, blocks: [paragraph('one', 'Alpha'), paragraph('two', 'Beta')] })
  editor.setActiveKind('document')
  const before = snapshot(editor.document.model)
  const draft = snapshot(before)
  commands.executeDocumentCommand(draft, selection('one', 1, 'two', 0), { id: 'doc.command.format-paragraph', patch: { align: 'center' } })
  assert.equal(editor.replaceDocumentFromView(draft), true)
  assert.equal(editor.state.document.modified, true)
  assert.equal(editor.state.document.blocks[0].style.align, 'center')
  assert.deepEqual(snapshot(editor.document.model.blocks[1]), before.blocks[1])
  editor.undo()
  assert.deepEqual(snapshot(editor.document.model), before)
  assert.equal(editor.state.document.modified, false)
  editor.redo()
  assert.equal(editor.state.document.blocks[0].style.align, 'center')
  editor.markSaved()
  const collapsed = snapshot(editor.document.model)
  commands.executeDocumentCommand(collapsed, selection('two', 0, 'two', 0), { id: 'doc.command.format-text', patch: { bold: true } })
  assert.equal(editor.replaceDocumentFromView(collapsed), false)
  assert.equal(editor.state.document.modified, false)
})

test('applying an existing character format or formatting only a paragraph break preserves run identities', () => {
  const model = documentWith([paragraph('one', 'Alpha'), paragraph('two', 'Beta')])
  const before = structuredClone(model)
  commands.executeDocumentCommand(model, selection('one', 1, 'one', 4), { id: 'doc.command.format-text', patch: { bold: false, fontFamily: 'Arial' } })
  assert.deepEqual(model, before)
  commands.executeDocumentCommand(model, selection('one', 5, 'two', 0), { id: 'doc.command.format-text', patch: { bold: true } })
  assert.deepEqual(model, before)
})
