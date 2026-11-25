import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Deferred integration drafts exercise the real OfficeStudio command/history boundary.
let server, OfficeStudio, FDocument, domain, modelUtils, bodyTypes, codeFonts, codeBlocks, contentControls, codec, pagination
const studios = new Set()
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/document-structured-facade', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio } = await server.ssrLoadModule('/packages/office-studio/src/api/office-studio.ts'))
  ;({ FDocument } = await server.ssrLoadModule('/packages/office-studio/src/api/facade.ts'))
  ;[domain, modelUtils, bodyTypes, codeFonts, codeBlocks, contentControls, codec, pagination] = await Promise.all(['domain/model', 'features/document/model/document-model', 'api/document-body-types',
    'features/document/code-blocks/document-code-fonts', 'features/document/code-blocks/code-blocks',
    'features/document/services/content-controls', 'api/document-body-codec', 'features/document/model/document-paragraph-pagination']
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

test('paragraph pagination facade shares the registered settings command and content history', () => {
  const { studio, document } = openDocument(), paragraph = document.getParagraph('body'), commands = []
  studio.events.addEvent('CommandExecuted', event => { if (event.kind === 'document') commands.push(event.id) })
  assert.equal(paragraph.setParagraphStyle({ keepLines: true, keepNext: false, widowControl: true, pageBreakBefore: true }), true)
  assert.deepEqual(pagination.documentParagraphPagination(paragraph.getParagraphStyle()), { keepLines: true, keepNext: false, widowControl: true, pageBreakBefore: true })
  assert.deepEqual(commands, ['doc.command.format-paragraph']); assert.equal(history(studio, document.id).undo, 1)
  assert.equal(document.undo(), true); assert.equal(paragraph.getParagraphStyle().keepLines, undefined)
  assert.equal(document.redo(), true); assert.equal(paragraph.getParagraphStyle().pageBreakBefore, true)
  assert.equal(paragraph.setParagraphStyle({ keepLines: undefined }), false)
})

test('quote Enter and backward merge update the adjacent keepNext rule with one undo item per edit', () => {
  const { studio, document } = openDocument(), quote = document.insertQuote(document.getParagraph('body'), { blockId: 'pagination-quote' })
  const before = copy(document.getModel()), previous = history(studio, document.id)
  document.executeCommand({ id: 'doc.command.insert-text', text: '\nnext' }, { anchor: { blockId: 'body', offset: 5 }, focus: { blockId: 'body', offset: 5 } })
  assert.deepEqual(quote.getParagraphs().map(paragraph => paragraph.getParagraphStyle().keepNext), [true, false])
  assert.equal(history(studio, document.id).undo, previous.undo + 1)
  assert.equal(document.undo(), true); assert.deepEqual(document.getModel().blocks, before.blocks)
  assert.equal(document.redo(), true)
  const last = quote.getParagraphs().at(-1)
  document.executeCommand({ id: 'doc.command.delete', direction: 'backward' }, { anchor: { blockId: last.getId(), offset: 0 }, focus: { blockId: last.getId(), offset: 0 } })
  assert.equal(quote.getParagraphs().length, 1); assert.equal(quote.getParagraphs()[0].getParagraphStyle().keepNext, false)
  assert.equal(document.getParagraph('tail').getParagraphStyle().keepLines, undefined)
})

test('public snapshot save and reopen retain quote pagination and original authorship after unwrap', () => {
  const body = paragraph('body', 'Quote'); Object.assign(body.style, { keepLines: false, keepNext: true, widowControl: false, pageBreakBefore: true })
  const { document } = openDocument([body, paragraph('tail')]), quote = document.insertQuote(document.getParagraph('body'), { blockId: 'pagination-quote' })
  assert.deepEqual(quote.getParagraphs().map(paragraph => [paragraph.getParagraphStyle().keepLines, paragraph.getParagraphStyle().keepNext, paragraph.getParagraphStyle().widowControl]), [[true, false, true]])
  const saved = document.save(), studio = new OfficeStudio({ document: saved, currentUserId: 'editor' }); studios.add(studio); studio.activateDocument(saved.id)
  const reopened = new FDocument(studio, saved.id); assert.equal(reopened.getQuote('pagination-quote').unwrap(), true)
  assert.deepEqual(pagination.documentParagraphPagination(reopened.getParagraph('body').getParagraphStyle()),
    { keepLines: false, keepNext: true, widowControl: false, pageBreakBefore: true })
  const exported = reopened.getBody().paragraphs.find(paragraph => paragraph.paragraphId === 'body').paragraphStyle
  assert.deepEqual([exported.keepLines, exported.keepNext, exported.widowControl, exported.pageBreakBefore], [0, 1, 0, 1])
})

test('same-value author pagination edits survive quote exit and unwrap', () => {
  const { document } = openDocument(), quote = document.insertQuote(document.getParagraph('body'), { blockId: 'pagination-quote' })
  assert.equal(quote.getParagraphs()[0].setParagraphStyle({ keepLines: true, widowControl: false }), true)
  assert.equal(quote.exit(), true); assert.equal(quote.getParagraphs()[0].getParagraphStyle().keepNext, false)
  assert.equal(quote.unwrap(), true)
  assert.equal(document.getParagraph('body').getParagraphStyle().keepLines, true)
  assert.equal(document.getParagraph('body').getParagraphStyle().widowControl, false)
})

test('quote pagination updates reject canceled commands, invalid booleans and whole-paragraph content locks atomically', () => {
  const source = { ...domain.createBlankDocument(), id: 'structured-api', blocks: [paragraph('body', 'Locked span'), paragraph('tail')] }
  contentControls.createDocumentContentControl(source, { id: 'lock', type: 'text', blockId: 'body', startOffset: 2, endOffset: 6, lock: 'contentLocked' })
  const { studio, document } = openDocument(source.blocks, source), before = copy(document.getModel()), previous = history(studio, document.id)
  assert.throws(() => document.getParagraph('body').setParagraphStyle({ keepLines: true }), /locked/)
  assert.deepEqual(document.getModel(), before); assert.deepEqual(history(studio, document.id), previous)
  const editable = openDocument(), original = copy(editable.document.getModel()), previousEditable = history(editable.studio, editable.document.id)
  for (const value of [0, 1, 'yes', null]) assert.throws(() => editable.document.getParagraph('body').setParagraphStyle({ keepNext: value }))
  const stop = editable.studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === 'doc.command.format-paragraph') event.cancel = true })
  try { assert.throws(() => editable.document.getParagraph('body').setParagraphStyle({ pageBreakBefore: true }), /canceled/) } finally { stop.dispose() }
  assert.deepEqual(editable.document.getModel(), original); assert.deepEqual(history(editable.studio, editable.document.id), previousEditable)
})

test('quote to code conversion restores generated rules and undo restores the quote rules', () => {
  const body = paragraph('body', 'Convert'); body.style.keepNext = true; body.style.keepLines = false
  const { document } = openDocument([body, paragraph('tail')]); document.insertQuote(document.getParagraph('body'), { blockId: 'pagination-quote' })
  const selection = { anchor: { blockId: 'body', offset: 0 }, focus: { blockId: 'body', offset: 0 } }
  document.executeCommand({ id: 'doc.command.convert-structured-block', blockId: 'pagination-quote', blockType: 'code' }, selection)
  assert.equal(document.getParagraph('body').getParagraphStyle().keepNext, true); assert.equal(document.getParagraph('body').getParagraphStyle().keepLines, false)
  assert.equal(document.undo(), true); assert.equal(document.getParagraph('body').getParagraphStyle().keepNext, false)
  assert.equal(document.redo(), true); assert.equal(document.getCode('pagination-quote').getText(), 'Convert')
})

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

test('column padding and minimum-height APIs convert pt once, share document history, and survive save/reopen', () => {
  const { studio, document } = openDocument(), commands = []
  studio.events.addEvent('CommandExecuted', event => { if (event.kind === 'document') commands.push(event.id) })
  const group = document.insertColumnGroup(2, { columnGroupId: 'columns' })
  assert.equal(group.getHorizontalPadding(), 6); assert.equal(group.getMinHeight(), 54)
  assert.equal(group.describe().horizontalPadding, 6); assert.equal(group.describe().minHeight, 54)
  const before = copy(document.getModel()), previous = history(studio, document.id).undo
  assert.equal(group.setHorizontalPadding(9), true)
  assert.equal(history(studio, document.id).undo, previous + 1)
  assert.equal(group.setMinHeight(90), true)
  assert.equal(history(studio, document.id).undo, previous + 2)
  assert.deepEqual(commands, ['doc.command.insert-column-group', 'doc.command.update-column-group', 'doc.command.update-column-group'])
  const config = modelUtils.findDocumentBlock(document.getModel(), 'columns').columns
  assert.equal(config.horizontalPadding, 12); assert.equal(config.minHeight, 120)
  assert.equal(document.undo(), true); assert.equal(group.getMinHeight(), 54); assert.equal(group.getHorizontalPadding(), 9)
  assert.equal(document.undo(), true); assert.deepEqual(document.getModel().blocks, before.blocks)
  assert.equal(document.redo(), true); assert.equal(document.redo(), true)
  const saved = document.save(), reopened = openDocument(saved.blocks, saved).document.getColumnGroup('columns')
  assert.equal(reopened.getHorizontalPadding(), 9); assert.equal(reopened.getMinHeight(), 90)
  assert.equal(reopened.describe().horizontalPadding, 9); assert.equal(reopened.describe().minHeight, 90)
  const explicit = openDocument().document.insertColumnGroup(2, { horizontalPadding: 3, minHeight: 36 })
  assert.equal(explicit.getHorizontalPadding(), 3); assert.equal(explicit.getMinHeight(), 36)
})

test('invalid padding and minimum-height insertion or updates do not publish a model or history item', () => {
  for (const field of ['horizontalPadding', 'minHeight']) for (const value of [null, -1, '8', NaN, Infinity]) {
    const { studio, document } = openDocument(), before = copy(document.getModel())
    assert.throws(() => document.insertColumnGroup(2, { columnGroupId: 'columns', [field]: value }), RangeError)
    assert.deepEqual(document.getModel(), before); assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
    const group = document.insertColumnGroup(2, { columnGroupId: 'columns' }), current = copy(document.getModel()), previous = history(studio, document.id)
    assert.throws(() => field === 'horizontalPadding' ? group.setHorizontalPadding(value) : group.setMinHeight(value), RangeError)
    assert.deepEqual(document.getModel(), current); assert.deepEqual(history(studio, document.id), previous)
  }
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

const codeContainer = (id, children, config = {}) => ({ ...paragraph(id), type: 'code', runs: [], children,
  code: { language: 'typescript', showLineNumbers: true, tabSize: 4, wrap: true, ...config } })
const pointSelection = id => ({ anchor: { blockId: id, offset: 0 }, focus: { blockId: id, offset: 0 } })

test('Code font facade formats every run and paragraph default once and persists across body and document reopen', () => {
  const first = paragraph('first', 'A😀B'), second = paragraph('second', 'second'), empty = paragraph('empty')
  first.runs = [modelUtils.createRun('A', { fontFamily: 'Menlo', bold: true }), modelUtils.createRun('😀B', { fontFamily: 'Courier New', color: '#112233' })]
  first.metadata.univerBodyParagraphSource = { version: 1, paragraph: { paragraphStyle: { textStyle: { ff: 'Menlo', authorFontHint: 'retained' } } } }
  const { studio, document } = openDocument([codeContainer('code', [first, second, empty]), paragraph('tail', 'Outside')]), code = document.getCode('code'), commands = []
  studio.events.addEvent('CommandExecuted', event => { if (event.kind === 'document') commands.push(event.id) })
  const before = copy(document.getModel()), previous = history(studio, document.id)
  assert.equal(code.getFontFamily(), 'Menlo'); assert.equal(code.setFontFamily('Consolas'), true)
  assert.equal(history(studio, document.id).undo, previous.undo + 1)
  assert.deepEqual(commands, ['doc.command.set-code-font-family'])
  const current = modelUtils.findDocumentBlock(document.getModel(), 'code')
  for (const child of current.children) {
    assert.equal(child.metadata.documentCodeFontFamily, 'Consolas')
    assert.ok(child.runs.length > 0 && child.runs.every(run => run.fontFamily === 'Consolas' && run.eastAsiaFontFamily === 'Consolas'))
  }
  assert.equal(current.children[0].runs[0].bold, true); assert.equal(current.children[0].runs[1].color, '#112233')
  assert.equal(code.getText(), 'A😀B\nsecond\n'); assert.equal(code.describe().fontFamily, 'Consolas')
  const body = document.getBody()
  assert.ok(body.paragraphs.filter(value => ['first', 'second', 'empty'].includes(value.paragraphId)).every(value => value.paragraphStyle.textStyle.ff === 'Consolas'))
  assert.equal(body.paragraphs.find(value => value.paragraphId === 'first').paragraphStyle.textStyle.authorFontHint, 'retained')
  const decoded = codec.decodeDocumentBody(body, document.getSnapshot()).content.blocks
  assert.equal(codeFonts.documentCodeFontFamily(modelUtils.findDocumentBlock({ blocks: decoded }, 'code')), 'Consolas')
  assert.equal(document.undo(), true); assert.deepEqual(document.getModel().blocks, before.blocks)
  assert.equal(document.redo(), true); assert.equal(code.getFontFamily(), 'Consolas')
  const saved = document.save(), reopened = openDocument(saved.blocks, saved).document
  assert.equal(reopened.getCode('code').getFontFamily(), 'Consolas')
  const at = history(studio, document.id).undo
  assert.equal(code.setFontFamily('Consolas'), false); assert.equal(history(studio, document.id).undo, at)
})

test('Code More patch and source/font properties commit atomically through the registered editor command', () => {
  const { studio, document } = openDocument([codeContainer('code', [paragraph('inside', 'old')]), paragraph('tail')]), code = document.getCode('code'), selection = pointSelection('inside')
  const before = copy(document.getModel()), previous = history(studio, document.id)
  const patch = { language: 'rust', showLineNumbers: false, tabSize: 2, wrap: false }
  assert.equal(document.executeCommand({ id: 'doc.command.update-code-settings', blockId: 'code', patch }, selection).changed, true)
  assert.deepEqual(code.getConfig(), patch); assert.equal(code.getText(), 'old')
  assert.equal(history(studio, document.id).undo, previous.undo + 1)
  assert.equal(document.undo(), true); assert.deepEqual(document.getModel().blocks, before.blocks)
  assert.equal(document.executeCommand({ id: 'doc.command.update-code-block', blockId: 'code', source: 'new\r\nline', patch, fontFamily: 'Menlo' }, selection).changed, true)
  assert.equal(code.getText(), 'new\nline'); assert.deepEqual(code.getConfig(), patch); assert.equal(code.getFontFamily(), 'Menlo')
  assert.equal(history(studio, document.id).undo, previous.undo + 1)
  assert.equal(document.undo(), true); assert.deepEqual(document.getModel().blocks, before.blocks)
  assert.equal(document.redo(), true); assert.equal(code.getText(), 'new\nline'); assert.equal(code.getFontFamily(), 'Menlo')
})

test('unsupported fonts, invalid More settings and canceled font commands publish no model or history', () => {
  const { studio, document } = openDocument([codeContainer('code', [paragraph('inside', 'retained')]), paragraph('tail')]), code = document.getCode('code'), before = copy(document.getModel()), previous = history(studio, document.id)
  assert.throws(() => code.setFontFamily('Arial'), /等宽字体/)
  for (const patch of [{ tabSize: 0 }, { tabSize: 1.5 }, { language: '' }, { wrap: 'yes' }, { showLineNumbers: 1 }]) {
    assert.throws(() => document.executeCommand({ id: 'doc.command.update-code-block', blockId: 'code', source: 'not published', patch, fontFamily: 'Menlo' }, pointSelection('inside')))
    assert.deepEqual(document.getModel(), before); assert.deepEqual(history(studio, document.id), previous)
  }
  const stop = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === 'doc.command.set-code-font-family') event.cancel = true })
  try { assert.throws(() => code.setFontFamily('Menlo'), /canceled/) } finally { stop.dispose() }
  assert.deepEqual(document.getModel(), before); assert.deepEqual(history(studio, document.id), previous)
})

test('Code whole-block font edits honor inner text locks and entity permissions without publishing partial runs', () => {
  for (const locked of [false, true]) {
    const source = { ...domain.createBlankDocument(), id: 'structured-api', blocks: [codeContainer('code', [paragraph('first', 'first'), paragraph('second', 'second')]), paragraph('tail')] }
    if (locked) contentControls.createDocumentContentControl(source, { id: 'locked', type: 'text', blockId: 'second', startOffset: 0, endOffset: 6,
      lock: 'contentLocked' })
    else source.permissionRules = [{ id: 'deny-font', scope: 'entity', targetId: 'code', action: 'edit', effect: 'deny', userIds: [] }]
    const { studio, document } = openDocument(source.blocks, source), before = copy(document.getModel()), previous = history(studio, document.id)
    assert.throws(() => document.getCode('code').setFontFamily('Menlo'))
    assert.deepEqual(document.getModel(), before); assert.deepEqual(history(studio, document.id), previous)
  }
})

test('explicitly registered monospace fonts are usable until disposal while loaded author values remain intact', () => {
  const name = 'Fixture Mono', dispose = codeFonts.registerDocumentCodeFontFamilies([{ value: name, label: 'Fixture 等宽', category: 'monospace' }])
  const { studio, document } = openDocument([codeContainer('code', [paragraph('inside', 'text')]), paragraph('tail')]), code = document.getCode('code')
  try {
    assert.equal(code.setFontFamily(name), true); assert.equal(code.getFontFamily(), name)
    assert.equal(codeFonts.documentCodeFontFamilyOptions(name).find(value => value.value === name).label, 'Fixture 等宽')
  } finally { dispose(); dispose() }
  const before = copy(document.getModel()), previous = history(studio, document.id)
  assert.equal(code.getFontFamily(), name)
  assert.equal(codeFonts.documentCodeFontFamilyOptions(name).find(value => value.value === name).disabled, true)
  assert.throws(() => code.setFontFamily(name), /等宽字体/)
  assert.deepEqual(document.getModel(), before); assert.deepEqual(history(studio, document.id), previous)
  const saved = document.save(), reopened = openDocument(saved.blocks, saved).document.getCode('code')
  assert.equal(reopened.getFontFamily(), name)
  assert.throws(() => codeFonts.registerDocumentCodeFontFamilies([{ value: 'Arial', category: 'sans-serif' }]), /等宽字体/)
})

test('insert Code font option and legacy mixed-run updates use the same real font mutation', () => {
  const inserted = openDocument().document.insertCode({ blockId: 'inserted-code', fontFamily: 'Courier New' })
  assert.ok(inserted); assert.equal(inserted.getFontFamily(), 'Courier New')
  const legacy = codeBlocks.createCodeBlock('AB', ' Vendor.Legacy ')
  legacy.id = 'legacy'; legacy.runs = [modelUtils.createRun('A', { fontFamily: 'Consolas' }), modelUtils.createRun('B', { fontFamily: 'Menlo' })]
  legacy.metadata.codeBlock.vendor = 'retained'
  const { studio, document } = openDocument([legacy, paragraph('tail')]), code = document.getCode('legacy'), before = copy(document.getModel()), previous = history(studio, document.id)
  assert.equal(document.executeCommand({ id: 'doc.command.update-code-block', blockId: 'legacy', fontFamily: 'Consolas', patch: { wrap: false, tabSize: 2 } }, pointSelection('legacy')).changed, true)
  const current = modelUtils.findDocumentBlock(document.getModel(), 'legacy')
  assert.ok(current.runs.every(run => run.fontFamily === 'Consolas' && run.eastAsiaFontFamily === 'Consolas'))
  assert.equal(current.metadata.documentCodeFontFamily, 'Consolas'); assert.equal(current.metadata.codeBlock.vendor, 'retained')
  assert.equal(code.getConfig().language, ' Vendor.Legacy '); assert.equal(code.getConfig().tabSize, 2); assert.equal(code.getConfig().wrap, false)
  assert.equal(history(studio, document.id).undo, previous.undo + 1)
  assert.equal(document.undo(), true); assert.deepEqual(document.getModel().blocks, before.blocks)
  assert.equal(document.redo(), true); assert.equal(code.getFontFamily(), 'Consolas')
})

test('Code to Quote conversion retains identity, source configuration and fonts through undo and body reopen', () => {
  const original = codeContainer('code', [paragraph('inside', 'source')], { language: ' Vendor.Code ', tabSize: 2, wrap: false })
  original.metadata.vendor = { original: true }
  const { studio, document } = openDocument([original, paragraph('tail')]), before = copy(document.getModel()), previous = history(studio, document.id), selection = pointSelection('inside')
  assert.equal(document.executeCommand({ id: 'doc.command.convert-structured-block', blockId: 'code', blockType: 'quote' }, selection).changed, true)
  const quote = document.getQuote('code'); assert.ok(quote); assert.equal(quote.getText(), 'source')
  assert.deepEqual(modelUtils.findDocumentBlock(document.getModel(), 'code').metadata.vendor, { original: true })
  assert.equal(history(studio, document.id).undo, previous.undo + 1)
  assert.equal(document.undo(), true); assert.deepEqual(document.getModel().blocks, before.blocks)
  assert.equal(document.redo(), true)
  const data = codec.encodeDocumentBody({ blocks: document.getModel().blocks, runs: [], paragraphs: [], metadata: {} })
  const decoded = codec.decodeDocumentBody(data.body, data).content.blocks, reopened = openDocument(decoded).document
  assert.equal(reopened.executeCommand({ id: 'doc.command.convert-structured-block', blockId: 'code', blockType: 'code' }, selection).changed, true)
  assert.equal(reopened.getCode('code').getText(), 'source'); assert.deepEqual(reopened.getCode('code').getConfig(), original.code)
  const exported = reopened.getBody().paragraphs.find(value => value.paragraphId === 'inside')
  assert.equal(exported.paragraphStyle.borderLeft, undefined)
})

test('structured conversion rejects non-text Code content and denied targets while identical conversion stays a no-op', () => {
  const table = { ...paragraph('table'), type: 'table', runs: [], metadata: { table: modelUtils.createTable(1, 1) } }
  const inner = table.metadata.table.rows[0].cells[0].blocks[0], quote = { ...paragraph('quote'), type: 'quote', runs: [], children: [table] }
  const { studio, document } = openDocument([quote, paragraph('tail')]), target = modelUtils.textTargets(document.getModel()).find(value => value.cellBlockId === inner.id)
  const selection = { anchor: { ...target, offset: 0 }, focus: { ...target, offset: 0 } }, before = copy(document.getModel()), previous = history(studio, document.id)
  assert.throws(() => document.executeCommand({ id: 'doc.command.convert-structured-block', blockId: 'quote', blockType: 'code' }, selection), /完整的文本段落/)
  assert.deepEqual(document.getModel(), before); assert.deepEqual(history(studio, document.id), previous)
  assert.equal(document.executeCommand({ id: 'doc.command.convert-structured-block', blockId: 'quote', blockType: 'quote' }, selection).changed, false)
  assert.deepEqual(history(studio, document.id), previous)
  const denied = openDocument([codeContainer('code', [paragraph('inside', 'source')]), paragraph('tail')],
    { permissionRules: [{ id: 'deny-convert', scope: 'entity', targetId: 'code', action: 'edit', effect: 'deny', userIds: [] }] })
  const snapshot = copy(denied.document.getModel()), at = history(denied.studio, denied.document.id)
  assert.throws(() => denied.document.executeCommand({ id: 'doc.command.convert-structured-block', blockId: 'code', blockType: 'quote' }, pointSelection('inside')))
  assert.deepEqual(denied.document.getModel(), snapshot); assert.deepEqual(history(denied.studio, denied.document.id), at)
})
