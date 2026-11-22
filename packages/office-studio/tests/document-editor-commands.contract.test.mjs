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
