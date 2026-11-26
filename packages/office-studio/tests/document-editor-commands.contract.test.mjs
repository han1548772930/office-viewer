import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

let server, OfficeStudio, FDocument, domain, documentModel, codeBlocks, controls, comments
const studios = new Set()
beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio } = await server.ssrLoadModule('/packages/office-studio/src/api/office-studio.ts'))
  ;({ FDocument } = await server.ssrLoadModule('/packages/office-studio/src/api/facade.ts'))
  domain = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
  documentModel = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
  codeBlocks = await server.ssrLoadModule('/packages/office-studio/src/features/document/code-blocks/code-blocks.ts')
  controls = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/content-controls.ts')
  comments = await server.ssrLoadModule('/packages/office-studio/src/features/document/comments/thread-comments.ts')
})
afterEach(() => { for (const studio of studios) studio.dispose(); studios.clear() })
afterAll(async () => { await server?.close() })

const copy = value => JSON.parse(JSON.stringify(value))
const point = (blockId, offset = 0, coordinates = {}) => ({ blockId, offset, ...coordinates })
const selection = (anchor, focus = anchor) => ({ anchor, focus })
function paragraph(id, text) { return { ...documentModel.createParagraph(text), id } }
function openDocument(blocks, extra = {}) {
  const model = { ...domain.createBlankDocument(), id: 'custom-doc-id', blocks, ...extra, modified: false }
  const studio = new OfficeStudio({ document: model, currentUserId: 'editor' })
  studios.add(studio)
  studio.activateDocument(model.id)
  return { studio, document: new FDocument(studio, model.id) }
}
function history(studio, id) {
  const unit = studio.editor.history.units[id]
  return { undo: unit?.undo.length ?? 0, redo: unit?.redo.length ?? 0, dirty: studio.editor.history.isDirty(id) }
}

test('editor formatting uses one public command and one document history entry, preserving untouched runs through save and reopen', () => {
  const { studio, document } = openDocument([paragraph('first', 'ABCD'), paragraph('second', 'EF')])
  const original = copy(document.getModel())
  const events = []
  studio.events.addEvent('BeforeCommandExecute', event => events.push({ phase: 'before', id: event.id, unitId: event.unitId, kind: event.kind }))
  studio.events.addEvent('CommandExecuted', event => events.push({ phase: 'after', id: event.id, unitId: event.unitId, kind: event.kind }))
  const range = selection(point('first', 1), point('first', 3))
  const result = document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true, color: '#c81e1e' } }, range)
  assert.equal(result.changed, true)
  assert.deepEqual(result.selection, range)
  assert.deepEqual(document.getModel().blocks[0].runs.map(run => [run.text, run.bold, run.color]), [
    ['A', false, '#1f2937'], ['BC', true, '#c81e1e'], ['D', false, '#1f2937'],
  ])
  assert.deepEqual(document.getModel().blocks[1], original.blocks[1])
  assert.deepEqual(history(studio, document.id), { undo: 1, redo: 0, dirty: true })
  assert.deepEqual(events, ['before', 'after'].map(phase => ({ phase, id: 'doc.command.format-text', unitId: document.id, kind: 'document' })))
  assert.equal(document.getModel().modified, true)
  assert.equal(document.undo(), true)
  assert.deepEqual(document.getModel(), original)
  assert.deepEqual(history(studio, document.id), { undo: 0, redo: 1, dirty: false })
  assert.equal(document.redo(), true)
  const saved = document.save()
  assert.equal(saved.modified, false)
  assert.equal(studio.editor.history.isDirty(document.id), false)
  const reopened = new OfficeStudio({ document: copy(saved) })
  studios.add(reopened)
  assert.deepEqual(new FDocument(reopened, saved.id).getModel().blocks, saved.blocks)
})

test('empty edits, edge deletion and current formatting remain no-ops without generating history or changing run ids', () => {
  const { studio, document } = openDocument([paragraph('first', 'ABC')])
  const original = copy(document.getModel())
  assert.equal(original.wordCount, 3)
  for (const [command, range] of [
    [{ id: 'doc.command.insert-text', text: '' }, selection(point('first', 1))],
    [{ id: 'doc.command.delete', direction: -1 }, selection(point('first', 0))],
    [{ id: 'doc.command.delete', direction: 1 }, selection(point('first', 3))],
    [{ id: 'doc.command.format-text', patch: {} }, selection(point('first', 0), point('first', 3))],
    [{ id: 'doc.command.format-text', patch: { bold: false } }, selection(point('first', 0), point('first', 3))],
    [{ id: 'doc.command.format-paragraph', patch: { align: 'left' } }, selection(point('first', 1))],
  ]) {
    assert.equal(document.executeCommand(command, range).changed, false)
    assert.deepEqual(document.getModel(), original)
    assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
  }
})

test('the fixed special-character registration keeps its public events while sharing pending marks, one input transaction and save/reopen', () => {
  const { studio, document } = openDocument([paragraph('first', 'ABCD'), paragraph('tail', 'Untouched')])
  const original = copy(document.getModel()), caret = selection(point('first', 1)), value = '👨‍👩‍👧‍👦e\u0301|a|'
  document.setInputSelection(caret)
  document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true, fontFamily: 'Georgia', color: '#C81E1E' } }, caret)
  const events = []
  for (const [eventName, phase] of [['BeforeCommandExecute', 'before'], ['CommandExecuted', 'after']]) {
    studio.events.addEvent(eventName, event => events.push({ phase, id: event.id, value: event.params.value, unitId: event.unitId, kind: event.kind }))
  }
  const result = document.executeCommand({ id: 'doc.command.insert-text', text: value }, caret, 'doc.command.insert-special-character')
  assert.equal(result.changed, true)
  assert.equal(result.selectionCurrent, true)
  assert.deepEqual(result.selection, selection(point('first', 1 + value.length)))
  assert.deepEqual(document.getInputState().location.selection, result.selection)
  assert.equal(documentModel.runText(document.getModel().blocks[0].runs), 'A' + value + 'BCD')
  const inserted = document.getModel().blocks[0].runs.find(run => run.text === value)
  assert.deepEqual([inserted.bold, inserted.fontFamily, inserted.color], [true, 'Georgia', '#C81E1E'])
  assert.deepEqual(document.getModel().blocks[1], original.blocks[1])
  assert.deepEqual(events, ['before', 'after'].map(phase => ({ phase, id: 'doc.command.insert-special-character', value, unitId: document.id, kind: 'document' })))
  assert.deepEqual(history(studio, document.id), { undo: 1, redo: 0, dirty: true })
  assert.equal(document.undo(), true)
  assert.deepEqual(document.getModel(), original)
  assert.equal(document.redo(), true)
  const saved = document.save(), reopened = openDocument(saved.blocks, saved)
  assert.deepEqual(reopened.document.getModel().blocks, saved.blocks)
})

test('flat special-character values replace the active range while empty, canceled, stale and locked requests preserve history', () => {
  const { studio, document } = openDocument([paragraph('first', 'ABCD')])
  const range = selection(point('first', 1), point('first', 3)), original = copy(document.getModel()), id = 'doc.command.insert-special-character'
  document.setInputSelection(range)
  assert.equal(studio.executeCommandSync(id, { value: '' }), false)
  assert.deepEqual(studio.executeCommandSync(id, { unitId: document.id, command: { id: 'doc.command.insert-text', text: 'ignored' }, selection: range, value: '' }), { changed: false, selection: range })
  assert.throws(() => studio.executeCommandSync(id, { value: 1 }), TypeError)
  assert.throws(() => studio.executeCommandSync(id, { unitId: document.id, command: { id: 'doc.command.format-text', patch: { bold: true } }, selection: range, value: '©' }), TypeError)
  const canceled = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === id) event.cancel = true })
  assert.throws(() => document.executeCommand({ id: 'doc.command.insert-text', text: '©' }, range, id), /canceled/)
  canceled.dispose()
  const moved = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === id) document.setInputSelection(selection(point('first', 4))) })
  assert.throws(() => document.executeCommand({ id: 'doc.command.insert-text', text: '©' }, range, id), /输入选区已更改/)
  moved.dispose()
  assert.deepEqual(document.getModel(), original)
  assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
  document.setInputSelection(range)
  assert.equal(studio.executeCommandSync(id, { value: '|a|' }), true)
  assert.equal(document.getText(), 'A|a|D')
  assert.deepEqual(document.getInputState().location.selection, selection(point('first', 4)))
  assert.deepEqual(history(studio, document.id), { undo: 1, redo: 0, dirty: true })
  const protectedModel = copy(original)
  controls.createDocumentContentControl(protectedModel, { type: 'text', blockId: 'first', startOffset: 1, endOffset: 3, lock: 'contentLocked' })
  const blocked = openDocument(protectedModel.blocks, protectedModel), before = copy(blocked.document.getModel())
  blocked.document.setInputSelection(range)
  assert.throws(() => blocked.studio.executeCommandSync(id, { value: '©' }), /locked/)
  assert.deepEqual(blocked.document.getModel(), before)
  assert.deepEqual(history(blocked.studio, blocked.document.id), { undo: 0, redo: 0, dirty: false })
})

test('public clear uses the document defaults, one history transaction, stable repeat no-ops and persisted defaults', () => {
  const block = paragraph('formatted', 'Hello world')
  block.type = 'heading'; block.level = 3; block.list = { id: 'ordered', preset: 'ORDER_LIST' }; block.style.align = 'center'
  block.runs[0] = { ...block.runs[0], bold: true, italic: true, fontSize: 26, color: '#FF0000', href: 'https://example.org', metadata: { source: { preserved: 'run' } } }
  const defaults = domain.createDocumentDefaultStyle()
  defaults.text.fontFamily = 'Georgia'; defaults.text.eastAsiaFontFamily = '宋体'; defaults.text.fontSize = 15
  defaults.paragraph.align = 'right'; defaults.paragraph.spaceAfter = 12
  const { studio, document } = openDocument([block], { defaultStyle: defaults })
  const original = copy(document.getModel()), commands = []
  studio.events.addEvent('CommandExecuted', event => commands.push([event.id, event.unitId, event.kind]))
  const result = document.executeCommand({ id: 'doc.command.clear-formatting' }, selection(point('formatted', 9)))
  assert.equal(result.changed, true)
  assert.deepEqual(result.selection, selection(point('formatted', 6), point('formatted', 11)))
  const current = copy(document.getModel())
  assert.deepEqual(current.blocks[0].runs.map(run => [run.text, run.bold, run.fontFamily, run.fontSize, run.href]), [
    ['Hello ', true, 'Aptos', 26, 'https://example.org'], ['world', false, 'Georgia', 15, 'https://example.org'],
  ])
  assert.deepEqual(current.blocks[0].style, defaults.paragraph)
  assert.equal(current.blocks[0].list, undefined); assert.equal(current.blocks[0].type, 'paragraph')
  assert.deepEqual(commands, [['doc.command.clear-formatting', document.id, 'document']])
  assert.deepEqual(history(studio, document.id), { undo: 1, redo: 0, dirty: true })
  assert.equal(document.executeCommand({ id: 'doc.command.clear-formatting' }, result.selection).changed, false)
  assert.deepEqual(document.getModel(), current)
  assert.deepEqual(history(studio, document.id), { undo: 1, redo: 0, dirty: true })
  document.undo(); assert.deepEqual(document.getModel(), original)
  document.redo(); assert.deepEqual(document.getModel(), current)
  const saved = document.save(), reopened = new OfficeStudio({ document: copy(saved) })
  studios.add(reopened)
  const reopenedDocument = new FDocument(reopened, saved.id)
  assert.deepEqual(reopenedDocument.getModel().defaultStyle, defaults)
  assert.equal(reopenedDocument.executeCommand({ id: 'doc.command.clear-formatting' }, result.selection).changed, false)
  assert.deepEqual(history(reopened, saved.id), { undo: 0, redo: 0, dirty: false })
  const insertion = document.executeCommand({ id: 'doc.command.insert-text', text: 'new' }, result.selection)
  assert.equal(insertion.changed, true); assert.equal(document.getText(), 'Hello new')
  const run = document.getModel().blocks[0].runs.find(run => run.text === 'new')
  assert.deepEqual([run.bold, run.italic, run.fontFamily, run.fontSize], [false, false, 'Georgia', 15])
})

test('public clear cancellation, permission rejection and outside-word content locks leave model and history intact', () => {
  const block = paragraph('formatted', 'word locked'); block.runs[0].bold = true; block.style.align = 'center'
  const { studio, document } = openDocument([block])
  const range = selection(point('formatted', 0), point('formatted', 4)), original = copy(document.getModel())
  const cancel = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === 'doc.command.clear-formatting') event.cancel = true })
  assert.throws(() => document.executeCommand({ id: 'doc.command.clear-formatting' }, range), /canceled/)
  cancel.dispose()
  assert.deepEqual(document.getModel(), original); assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
  for (const denied of ['paragraph', 'content-control']) {
    const protectedModel = copy(original)
    if (denied === 'paragraph') protectedModel.permissionRules = [{ id: 'deny-clear', scope: 'paragraph', targetId: 'formatted', action: 'edit', effect: 'deny', userIds: [] }]
    else controls.createDocumentContentControl(protectedModel, { type: 'text', blockId: 'formatted', startOffset: 5, endOffset: 11, lock: 'contentLocked' })
    const blocked = openDocument(protectedModel.blocks, protectedModel)
    const before = copy(blocked.document.getModel())
    assert.throws(() => blocked.document.executeCommand({ id: 'doc.command.clear-formatting' }, range))
    assert.deepEqual(blocked.document.getModel(), before)
    assert.deepEqual(history(blocked.studio, blocked.document.id), { undo: 0, redo: 0, dirty: false })
  }
})

test('public clear across body and table rejects a later content lock without publishing earlier changes', () => {
  const first = paragraph('body', 'Body'), table = paragraph('table', '')
  first.type = 'heading'; first.level = 2; first.runs[0].bold = true
  table.type = 'table'; table.runs = []; table.metadata.table = documentModel.createTable(1, 2)
  const earlier = table.metadata.table.rows[0].cells[0].blocks[0], last = table.metadata.table.rows[0].cells[1].blocks[0]
  earlier.runs = [documentModel.createRun('Earlier', { italic: true })]; last.runs = [documentModel.createRun('Later', { underline: true })]
  const { studio, document } = openDocument([first, table], { contentControls: [{ id: 'last-lock', type: 'text', blockId: last.id, startOffset: 3, endOffset: 5, lock: 'contentLocked' }] })
  const before = copy(document.getModel()), commands = []
  studio.events.addEvent('CommandExecuted', event => commands.push(event.id))
  const range = selection(point('body', 0), point('table', 3, { row: 0, col: 1, cellBlockId: last.id }))
  assert.throws(() => document.executeCommand({ id: 'doc.command.clear-formatting' }, range), /locked/)
  assert.deepEqual(document.getModel(), before); assert.deepEqual(commands, [])
  assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
})

test('cancellation and malformed or stale selections do not publish a partial document transaction', () => {
  const { studio, document } = openDocument([paragraph('first', 'ABC')])
  const original = copy(document.getModel())
  let executed = 0
  studio.events.addEvent('CommandExecuted', () => executed++)
  const cancellation = studio.events.addEvent('BeforeCommandExecute', event => { event.cancel = true })
  assert.throws(() => document.executeCommand({ id: 'doc.command.insert-text', text: 'x' }, selection(point('first', 1))), /canceled/)
  cancellation.dispose()
  for (const invalid of [point('first', -1), point('first', 0.5), point('missing', 0), point('first', 0, { row: 0 }), point('first', 0, { col: 0 })]) {
    assert.throws(() => document.executeCommand({ id: 'doc.command.insert-text', text: 'x' }, selection(invalid)))
    assert.deepEqual(document.getModel(), original)
  }
  assert.throws(() => studio.executeCommandSync('doc.command.insert-text', {
    unitId: document.id, command: { id: 'doc.command.delete', direction: -1 }, selection: selection(point('first', 1)),
  }), TypeError)
  assert.equal(executed, 0)
  assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
})

test('a denied paragraph or explicit checklist target is rejected even when another paragraph owns the caret', () => {
  const checklist = paragraph('protected', 'Task')
  checklist.type = 'checklist'; checklist.checklist = { listId: 'tasks', checked: false }
  const { studio, document } = openDocument([paragraph('editable', 'Text'), checklist], {
    permissionRules: [{ id: 'deny', scope: 'paragraph', targetId: 'protected', action: 'edit', effect: 'deny', userIds: [] }],
  })
  const original = copy(document.getModel())
  assert.throws(() => document.executeCommand({ id: 'doc.command.format-paragraph', patch: { align: 'center' } }, selection(point('editable'), point('protected', 2))), /permission denied/)
  assert.throws(() => document.executeCommand({ id: 'doc.command.toggle-checklist-item', blockId: 'protected' }, selection(point('editable'))), /permission denied/)
  assert.deepEqual(document.getModel(), original)
  assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
})

test('invalid mixed paragraph updates leave all paragraph styles and history intact', () => {
  const first = paragraph('first', 'Alpha'), second = paragraph('second', 'Beta')
  second.style.lineHeight = 2
  const { studio, document } = openDocument([first, second])
  const original = copy(document.getModel())
  assert.throws(() => document.executeCommand({ id: 'doc.command.format-paragraph', patch: { align: 'center', lineHeight: 0 } }, selection(point('first'), point('second', 4))), RangeError)
  assert.deepEqual(document.getModel(), original)
  assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
})

test('commands and undo histories remain isolated while editing an inactive document', () => {
  const { studio, document: first } = openDocument([paragraph('first', 'One')])
  const other = studio.createDocument({ id: 'other-doc', text: 'Two' })
  const second = new FDocument(studio, other.id)
  const secondPoint = point(second.getModel().blocks[0].id, 3)
  first.executeCommand({ id: 'doc.command.insert-text', text: '!' }, selection(point('first', 3)))
  studio.activateDocument(second.id)
  first.executeCommand({ id: 'doc.command.insert-text', text: '?' }, selection(point('first', 4)))
  assert.equal(studio.activeDocumentId, second.id)
  assert.equal(first.getText(), 'One!?')
  assert.equal(second.getText(), 'Two')
  second.executeCommand({ id: 'doc.command.insert-text', text: '!' }, selection(secondPoint))
  assert.deepEqual(history(studio, first.id), { undo: 2, redo: 0, dirty: true })
  assert.deepEqual(history(studio, second.id), { undo: 1, redo: 0, dirty: true })
  second.undo()
  assert.equal(second.getText(), 'Two')
  assert.equal(first.getText(), 'One!?')
  first.undo()
  assert.equal(first.getText(), 'One!')
  assert.equal(second.getText(), 'Two')
})

test('flat special-character insertion uses the active table cell and retains its coordinates through history and save/reopen', () => {
  const table = paragraph('table', '')
  table.type = 'table'; table.runs = []; table.metadata.table = documentModel.createTable(1, 1)
  table.metadata.table.rows[0].cells[0].blocks = [paragraph('cell', 'AB')]
  const { studio, document } = openDocument([table]), original = copy(document.getModel()), value = '👩‍💻'
  const coordinates = { row: 0, col: 0, cellBlockId: 'cell' }, range = selection(point('table', 1, coordinates), point('table', 2, coordinates))
  document.setInputSelection(range)
  assert.equal(studio.executeCommandSync('doc.command.insert-special-character', { value }), true)
  assert.equal(documentModel.runText(document.getModel().blocks[0].metadata.table.rows[0].cells[0].blocks[0].runs), 'A' + value)
  assert.deepEqual(document.getInputState().location.selection, selection(point('table', 1 + value.length, coordinates)))
  assert.deepEqual(history(studio, document.id), { undo: 1, redo: 0, dirty: true })
  assert.equal(document.undo(), true)
  assert.deepEqual(document.getModel(), original)
  assert.equal(document.redo(), true)
  const saved = document.save(), reopened = openDocument(saved.blocks, saved)
  assert.deepEqual(reopened.document.getModel().blocks, saved.blocks)
})

test('table navigation returns its new logical caret without a content history entry and rejects stale cell paragraphs', () => {
  const table = paragraph('table', '')
  table.type = 'table'; table.runs = []; table.metadata.table = documentModel.createTable(1, 2)
  const first = table.metadata.table.rows[0].cells[0].blocks[0], second = table.metadata.table.rows[0].cells[1].blocks[0]
  const { studio, document } = openDocument([table])
  const original = copy(document.getModel())
  const result = document.executeCommand({ id: 'doc.command.table-tab', direction: 1 }, selection(point('table', 0, { row: 0, col: 0, cellBlockId: first.id })))
  assert.equal(result.changed, false)
  assert.deepEqual(result.selection, selection(point('table', 0, { row: 0, col: 1, cellBlockId: second.id })))
  assert.throws(() => document.executeCommand({ id: 'doc.command.insert-text', text: 'x' }, selection(point('table', 0, { row: 0, col: 1, cellBlockId: 'removed' }))), RangeError)
  assert.deepEqual(document.getModel(), original)
  assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
})

test('public code editing keeps source, control and comment anchors in one reversible transaction', () => {
  const code = codeBlocks.createCodeBlock('a😀b', 'typescript')
  code.id = 'code'
  const source = { ...domain.createBlankDocument(), id: 'code-doc', blocks: [code] }
  controls.createDocumentContentControl(source, { id: 'code-control', type: 'text', blockId: 'code', startOffset: 3, endOffset: 4 })
  const comment = comments.createDocumentThreadComment(source, selection(point('code', 3), point('code', 4)), 'Review', { documentId: source.id, userId: 'editor', displayName: 'Editor' }, '2026-01-01T00:00:00.000Z')
  const { studio, document } = openDocument(source.blocks, source)
  const original = copy(document.getModel())
  const insertion = document.executeCommand({ id: 'doc.command.insert-text', text: '\n\t' }, selection(point('code', 1)))
  assert.equal(insertion.changed, true)
  assert.equal(document.getModel().blocks.length, 1)
  assert.equal(codeBlocks.codeBlockData(document.getModel().blocks[0]).source, 'a\n\t😀b')
  assert.deepEqual(document.getModel().contentControls.map(control => [control.startOffset, control.endOffset]), [[5, 6]])
  assert.equal(comments.documentThreadCommentAnchorStatus(document.getModel(), document.getModel().comments[comment.id]), 'current')
  document.executeCommand({ id: 'doc.command.delete', direction: 1 }, insertion.selection)
  assert.equal(codeBlocks.codeBlockData(document.getModel().blocks[0]).source, 'a\n\tb')
  assert.deepEqual(history(studio, document.id), { undo: 2, redo: 0, dirty: true })
  document.undo(); document.undo()
  assert.deepEqual(document.getModel(), original)
  document.redo(); document.redo()
  assert.deepEqual(document.getModel().contentControls.map(control => [control.startOffset, control.endOffset]), [[3, 4]])
  assert.equal(comments.documentThreadCommentAnchorStatus(document.getModel(), document.getModel().comments[comment.id]), 'current')
})

test('public code edits enforce content locks without publishing source, anchor or history changes', () => {
  const code = codeBlocks.createCodeBlock('abc'); code.id = 'code'
  const source = { ...domain.createBlankDocument(), blocks: [code] }
  controls.createDocumentContentControl(source, { type: 'text', blockId: 'code', startOffset: 1, endOffset: 2, lock: 'contentLocked' })
  const { studio, document } = openDocument(source.blocks, source)
  const original = copy(document.getModel())
  assert.throws(() => document.executeCommand({ id: 'doc.command.insert-text', text: 'X' }, selection(point('code', 1), point('code', 2))), /locked/)
  assert.deepEqual(document.getModel(), original)
  assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
})

test('legacy flattened insertText remains compatible and workbook events honor a custom document-like workbook id', () => {
  const { studio, document } = openDocument([paragraph('first', 'ABC')])
  assert.equal(document.insertText(1, '!'), true)
  assert.equal(document.getText(), 'A!BC')
  const workbook = studio.createWorkbook({ id: 'document-like-workbook' })
  const seen = []
  studio.events.addEvent('CommandExecuted', event => seen.push(event.kind))
  studio.registerCommand('custom.command.observe', () => true)
  studio.executeCommandSync('custom.command.observe', { unitId: workbook.id })
  studio.executeCommandSync('custom.command.observe', { unitId: document.id })
  assert.deepEqual(seen, ['workbook', 'document'])
})
