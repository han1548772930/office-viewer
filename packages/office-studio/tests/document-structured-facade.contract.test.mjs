import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Deferred integration drafts exercise the real OfficeStudio command/history boundary.
let server, OfficeStudio, FDocument, domain, modelUtils, bodyTypes
const studios = new Set()
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/document-structured-facade', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio } = await server.ssrLoadModule('/packages/office-studio/src/api/office-studio.ts'))
  ;({ FDocument } = await server.ssrLoadModule('/packages/office-studio/src/api/facade.ts'))
  ;[domain, modelUtils, bodyTypes] = await Promise.all(['domain/model', 'features/document/model/document-model', 'api/document-body-types']
    .map(path => server.ssrLoadModule('/packages/office-studio/src/' + path + '.ts')))
})
afterEach(() => { for (const studio of studios) studio.dispose(); studios.clear() })
afterAll(async () => { await server?.close() })
const copy = value => structuredClone(value)
const paragraph = (id, text = '') => ({ ...modelUtils.createParagraph(text), id })
function openDocument(blocks = [paragraph('body', 'Alpha'), paragraph('tail', 'Omega')], extra = {}) {
  const model = { ...domain.createBlankDocument(), id: 'structured-api', blocks, ...extra, modified: false }
  const studio = new OfficeStudio({ document: model, currentUserId: 'editor' }); studios.add(studio); studio.activateDocument(model.id)
  return { studio, document: new FDocument(studio, model.id) }
}
const history = (studio, id) => ({ undo: studio.editor.history.units[id]?.undo.length ?? 0,
  redo: studio.editor.history.units[id]?.redo.length ?? 0, dirty: studio.editor.history.isDirty(id) })

test('quote facade wraps the requested paragraph, shares commands/history and resolves its latest model after save and reopen', () => {
  const { studio, document } = openDocument(), original = copy(document.getModel()), commands = []
  studio.events.addEvent('CommandExecuted', event => { if (event.kind === 'document') commands.push(event.id) })
  const quote = document.insertQuote(document.getParagraph('body'), { blockId: 'quote' })
  assert.ok(quote); assert.equal(quote.getText(), 'Alpha'); assert.equal(quote.getParagraphs()[0].getId(), 'body')
  assert.equal(document.getQuoteAt(quote.getRange().startIndex + 1).getId(), 'quote')
  assert.equal(document.findQuoteByText('Alp').getId(), 'quote')
  assert.equal(quote.setStyle({ lineColor: '#123456', textColor: '#334455' }), true)
  assert.deepEqual(quote.getStyle(), { lineColor: '#123456', textColor: '#334455' })
  assert.deepEqual(commands, ['doc.command.wrap-structured-block', 'doc.command.update-quote-style'])
  assert.deepEqual(history(studio, document.id), { undo: 2, redo: 0, dirty: true })
  assert.equal(document.undo(), true); assert.notEqual(quote.getStyle().lineColor, '#123456')
  assert.equal(document.redo(), true); assert.equal(quote.getStyle().lineColor, '#123456')
  const saved = document.save(), reopened = openDocument(saved.blocks, saved).document
  assert.equal(reopened.getQuote('quote').getText(), 'Alpha')
  assert.deepEqual(reopened.getQuote('quote').getStyle(), quote.getStyle())
  assert.equal(reopened.getModel().modified, false)
  assert.equal(quote.unwrap(), true); assert.equal(document.getQuote('quote'), null)
  assert.equal(quote.getRange(), null); assert.equal(quote.remove(), false)
  assert.equal(document.getParagraph('body').getText(), 'Alpha')
  assert.equal(document.undo(), true); assert.equal(quote.getRange().blockId, 'quote')
  assert.deepEqual(original.blocks[1], document.getModel().blocks[1])
})

test('callout facade persists independent configuration, real text color and borders without flattening paragraphs', () => {
  const { document } = openDocument()
  const callout = document.insertCallout(document.getParagraph('body'), { blockId: 'callout', config: { icon: '💡', showIcon: false } })
  assert.ok(callout)
  assert.equal(callout.setBorder({ color: '#334455', width: 2, opacity: 0.5, style: 3 }), true)
  assert.equal(callout.setBackgroundColor('#ddeeff'), true)
  assert.equal(callout.setTextColor('#112233'), true)
  assert.deepEqual(callout.getStyle(), { backgroundColor: '#ddeeff', border: { color: '#334455', width: 2, opacity: 0.5, style: 3 }, textColor: '#112233' })
  assert.equal(callout.getConfig().icon, '💡'); assert.equal(callout.getConfig().showIcon, false)
  assert.equal(document.getParagraph('body').getText(), 'Alpha')
  const saved = document.save(), reopened = openDocument(saved.blocks, saved).document
  assert.deepEqual(reopened.getCallout('callout').getConfig(), callout.getConfig())
  assert.deepEqual(reopened.getCallout('callout').getStyle(), callout.getStyle())
})

test('code facade edits actual source paragraphs and code configuration through the same command boundary', () => {
  const { studio, document } = openDocument(), commands = []
  studio.events.addEvent('CommandExecuted', event => { if (event.kind === 'document') commands.push(event.id) })
  const code = document.insertCode(document.getParagraph('body'), { blockId: 'code', config: { language: 'rust', showLineNumbers: true, tabSize: 2, wrap: false } })
  assert.ok(code); assert.equal(code.setText('let a = 1\r\nlet b = 2'), true)
  assert.equal(code.getText(), 'let a = 1\nlet b = 2')
  assert.deepEqual(code.getParagraphs().map(paragraph => paragraph.getText()), ['let a = 1', 'let b = 2'])
  assert.deepEqual(code.getConfig(), { language: 'rust', showLineNumbers: true, tabSize: 2, wrap: false })
  assert.equal(code.updateConfig({ config: { showLineNumbers: false } }), true)
  assert.deepEqual(commands, ['doc.command.wrap-structured-block', 'doc.command.update-code-block', 'doc.command.update-code-block'])
  assert.equal(document.undo(), true); assert.equal(code.getConfig().showLineNumbers, true)
  assert.equal(document.undo(), true); assert.equal(code.getText(), 'Alpha')
  assert.equal(document.redo(), true); assert.equal(code.getText(), 'let a = 1\nlet b = 2')
  const saved = document.save(), reopened = openDocument(saved.blocks, saved).document
  assert.equal(reopened.findCodeByText('let b').getId(), 'code')
  assert.deepEqual(reopened.getCode('code').getConfig(), code.getConfig())
})

test('column facade preserves pt/px units, per-column text/paragraphs, public delimiter ranges and destructive delete semantics', () => {
  const { document } = openDocument()
  const group = document.insertColumnGroup(3, { columnGroupId: 'columns', columnIds: ['left', 'middle', 'right'], gap: 12, widthRatios: [2, 1, 1] })
  assert.ok(group); assert.equal(group.describe().gap, 12)
  assert.equal(modelUtils.findDocumentBlock(document.getModel(), 'columns').columns.gap, 16)
  const left = group.getColumn('left'), middle = group.getColumn(1)
  assert.equal(left.setText('Left\r\nMore'), true)
  const inserted = left.insertParagraph(1, 'Inserted\nExtra')
  assert.equal(inserted.getText(), 'Inserted')
  assert.equal(left.getText(), 'Left\nInserted\nExtra\nMore')
  // TextRange exposes the public data stream, including paragraph sentinels.
  assert.equal(left.getTextRange().getText(), 'Left\rInserted\rExtra\rMore\r')
  assert.deepEqual(left.getParagraphs().map(paragraph => paragraph.getText()), ['Left', 'Inserted', 'Extra', 'More'])
  assert.equal(middle.setText('discard'), true)
  assert.equal(group.setGap(18), true); assert.equal(group.describe().gap, 18)
  assert.equal(modelUtils.findDocumentBlock(document.getModel(), 'columns').columns.gap, 24)
  const range = group.getRange(), stream = document.getBody().dataStream, T = bodyTypes.DOCUMENT_BODY_TOKENS
  assert.equal(stream[range.startOffset], T.groupStart); assert.equal(stream[range.endOffset - 1], T.groupEnd)
  for (const column of range.columns) { assert.equal(stream[column.startOffset], T.columnStart); assert.equal(stream[column.endOffset], T.columnEnd) }
  assert.deepEqual(left.getContentRange(), { startOffset: range.columns[0].startOffset + 1, endOffset: range.columns[0].endOffset, segmentId: '' })
  assert.equal(group.deleteColumn('middle'), true); assert.equal(middle.getRange(), null)
  assert.deepEqual(group.getColumns().map(column => column.getId()), ['left', 'right'])
  assert.equal(left.getText(), 'Left\nInserted\nExtra\nMore')
  const added = group.addColumn('right', 'left', 'new-column')
  assert.equal(added.getId(), 'new-column'); assert.equal(added.getIndex(), 1)
  assert.deepEqual(group.getWidthRatios(), [2, 1, 1])
  const saved = document.save(), reopened = openDocument(saved.blocks, saved).document.getColumnGroup('columns')
  assert.equal(reopened.getColumn('left').getText(), left.getText())
  assert.equal(reopened.describe().gap, 18)
  assert.deepEqual(reopened.getWidthRatios(), group.getWidthRatios())
})

test('invalid insert gaps fall back to the default; invalid gap updates and malformed queries are rejected without mutation', () => {
  for (const gap of [-1, '12', NaN, Infinity, null]) {
    const { document } = openDocument()
    const group = document.insertColumnGroup(2, { columnGroupId: 'columns', gap })
    assert.equal(group.describe().gap, 12)
    const before = copy(document.getModel())
    assert.throws(() => group.setGap(gap), RangeError); assert.deepEqual(document.getModel(), before)
  }
  const { document } = openDocument()
  for (const query of [null, [], 5, { columnGroupId: 7 }, { text: false }]) assert.throws(() => document.findColumnGroups(query), TypeError)
  for (const query of [null, [], 5, { blockId: 7 }, { text: false }]) assert.throws(() => document.findQuotes(query), TypeError)
})

test('a paragraph from another document cannot be wrapped by the facade', () => {
  const first = openDocument(), second = openDocument([paragraph('foreign', 'Foreign')], { id: 'other-doc' })
  const before = copy(first.document.getModel()), foreign = copy(second.document.getModel())
  assert.throws(() => first.document.insertQuote(second.document.getParagraph('foreign')), /另一份文档/)
  assert.deepEqual(first.document.getModel(), before); assert.deepEqual(second.document.getModel(), foreign)
})

test('command cancellation and permissions reject facade structure changes with no history entry', () => {
  const { studio, document } = openDocument(), before = copy(document.getModel())
  const stop = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === 'doc.command.wrap-structured-block') event.cancel = true })
  assert.throws(() => document.insertQuote(document.getParagraph('body'), { blockId: 'quote' }), /canceled/)
  stop.dispose(); assert.deepEqual(document.getModel(), before); assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
  const denied = openDocument(before.blocks, { permissionRules: [{ id: 'deny', scope: 'paragraph', targetId: 'body', action: 'edit', effect: 'deny', userIds: [] }] })
  const snapshot = copy(denied.document.getModel())
  assert.throws(() => denied.document.insertCallout(denied.document.getParagraph('body'), { blockId: 'callout' }))
  assert.deepEqual(denied.document.getModel(), snapshot)
  assert.deepEqual(history(denied.studio, denied.document.id), { undo: 0, redo: 0, dirty: false })
})
