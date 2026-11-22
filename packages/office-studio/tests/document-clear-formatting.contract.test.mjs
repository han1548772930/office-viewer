import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import { collapsedSelection, createParagraph, createRun, createTable, defaultParagraphStyle, runText } from '../src/features/document/model/document-model.ts'
import { documentDefaultStyle } from '../src/features/document/model/document-default-style.ts'
import { officeNumberingMarkerBaseline } from '../src/features/document/model/document-office-numbering-marker.ts'
import { executeDocumentCommand } from '../src/features/document/commands/document-commands.ts'
import { clearDocumentFormattingRuns, documentClearFormattingState } from '../src/features/document/services/document-clear-formatting.ts'
import { createDocumentContentControl } from '../src/features/document/services/content-controls.ts'
import { normalizeDocument as normalizeImportedDocument } from '../src/parser/documentNormalizer.ts'
import { normalizeDocumentHighlight } from '../src/features/document/model/document-highlight.ts'

const point = (block, offset, extra = {}) => ({ blockId: block.id, offset, ...extra })
const range = (anchor, focus = anchor) => ({ anchor, focus })
const context = model => ({ documentId: model.id, userId: 'local-user' })
const clear = (model, selection) => executeDocumentCommand(model, selection, { id: 'doc.command.clear-formatting' })
const defaults = () => ({
  version: 1,
  text: { fontFamily: 'Georgia', eastAsiaFontFamily: '宋体', fontSize: 15, color: 'auto', bold: false, italic: false, underline: false, strike: false, highlight: '#808000' },
  paragraph: { ...defaultParagraphStyle(), align: 'right', indentStart: -12, indentEnd: 9, hanging: 18, lineHeight: 1.2, spaceAfter: 12, spacingRule: 'auto', wordWrap: false, snapToGrid: true },
})

test('caret clear expands its word, applies document defaults, and clears the entire affected paragraph list style', () => {
  const model = createBlankDocument(), block = createParagraph('Hello world', { bold: true, italic: true, underline: true, strike: true, fontSize: 26, color: '#FF0000', highlight: '#00FF00', verticalAlign: 'superscript', href: 'https://example.org' })
  block.type = 'heading'; block.level = 3; block.list = { id: 'numbering', preset: 'ORDER_LIST' }
  block.style.align = 'center'; block.metadata.namedStyleType = 'heading3'; block.metadata.source = { retained: 'original paragraph' }
  block.runs[0].metadata = { source: { retained: 'original run' }, link: 'https://example.org', highlight: '#00FF00', verticalAlign: 'superscript' }
  model.blocks = [block]; model.defaultStyle = defaults()
  const before = structuredClone(block), output = clear(model, collapsedSelection(point(block, 8)))
  assert.deepEqual(output, range(point(block, 6), point(block, 11)))
  assert.equal(runText(block.runs), 'Hello world')
  assert.equal(block.type, 'paragraph'); assert.equal(block.level, 0); assert.equal(block.list, undefined)
  assert.equal(block.metadata.namedStyleType, undefined); assert.deepEqual(block.metadata.source, before.metadata.source)
  assert.deepEqual(block.style, model.defaultStyle.paragraph); assert.notEqual(block.style, model.defaultStyle.paragraph)
  assert.deepEqual(block.runs.map(run => [run.text, run.bold, run.fontFamily, run.highlight, run.verticalAlign, run.href]), [
    ['Hello ', true, 'Aptos', '#00FF00', 'superscript', 'https://example.org'],
    ['world', false, 'Georgia', '#808000', undefined, 'https://example.org'],
  ])
  assert.deepEqual(block.runs[1].metadata.source, before.runs[0].metadata.source)
  assert.equal(block.runs[1].metadata.verticalAlign, undefined)
  executeDocumentCommand(model, output, { id: 'doc.command.insert-text', text: 'replacement' })
  assert.equal(runText(block.runs), 'Hello replacement')
  const inserted = block.runs.find(run => run.text === 'replacement')
  assert.equal(inserted.fontFamily, 'Georgia'); assert.equal(inserted.bold, false); assert.equal(inserted.highlight, '#808000')
})

test('half-open reversed selection clears only intersected paragraphs and ignores a protected next paragraph at zero', () => {
  const model = createBlankDocument(), first = createParagraph('first', { bold: true }), second = createParagraph('second', { italic: true })
  first.type = 'checklist'; first.checklist = { listId: 'tasks', checked: true }
  model.blocks = [first, second]
  model.permissionRules = [{ id: 'deny-next', scope: 'paragraph', targetId: second.id, action: 'edit', effect: 'deny', userIds: [] }]
  const original = structuredClone(second), selection = range(point(second, 0), point(first, 2))
  assert.equal(documentClearFormattingState(model, selection, context(model)).available, true)
  assert.deepEqual(clear(model, selection), selection)
  assert.equal(first.type, 'paragraph'); assert.equal(first.checklist, undefined)
  assert.deepEqual(first.runs.map(run => [run.text, run.bold]), [['fi', true], ['rst', false]])
  assert.deepEqual(second, original)
})

test('zero-length footnote and field entities survive while only selected text runs are split', () => {
  const zeroNote = createRun('', { metadata: { footnoteId: 'note-1' } }), zeroField = createRun('', { metadata: { field: { id: 'field-1' } } })
  const text = createRun('abcdef', { bold: true, href: 'https://example.org', metadata: { source: { preserved: true }, note: 'review' } })
  const result = clearDocumentFormattingRuns([zeroNote, text, zeroField], 2, 4, documentDefaultStyle().text)
  assert.equal(result[0], zeroNote); assert.equal(result.at(-1), zeroField)
  assert.deepEqual(result.slice(1, -1).map(run => [run.text, run.bold, run.href, run.metadata.note]), [
    ['ab', true, 'https://example.org', 'review'], ['cd', false, 'https://example.org', 'review'], ['ef', true, 'https://example.org', 'review'],
  ])
  assert.deepEqual(result[2].metadata.source, text.metadata.source)
})

test('old snapshots use stable defaults and repeated clear does not split runs or change their ids', () => {
  const model = createBlankDocument(), block = createParagraph('already clear')
  delete model.defaultStyle
  model.blocks = [block]
  const before = structuredClone(model), run = block.runs[0], selection = range(point(block, 2), point(block, 7))
  clear(model, selection); clear(model, selection)
  assert.deepEqual(model, before); assert.equal(block.runs[0], run)
  assert.equal(block.runs.length, 1); assert.equal(block.runs[0].eastAsiaFontFamily, undefined)
  const blank = createParagraph('  '); model.blocks = [blank]
  const caret = collapsedSelection(point(blank, 1)), emptyBefore = structuredClone(model)
  assert.equal(documentClearFormattingState(model, caret, context(model)).available, false)
  assert.deepEqual(clear(model, caret), caret); assert.deepEqual(model, emptyBefore)
})

test('new documents persist their default style and keep independent default objects', () => {
  const first = createBlankDocument(), second = createBlankDocument()
  assert.deepEqual(first.defaultStyle, documentDefaultStyle())
  first.defaultStyle.text.fontFamily = 'Changed'; first.defaultStyle.paragraph.align = 'right'
  assert.equal(second.defaultStyle.text.fontFamily, 'Aptos'); assert.equal(second.defaultStyle.paragraph.align, 'left')
})

test('document default colors normalize Office names and bare RGB before reaching Canvas formatting', () => {
  const model = createBlankDocument()
  model.defaultStyle.text.color = 'ABCDEF'; model.defaultStyle.text.highlight = 'darkYellow'
  const resolved = documentDefaultStyle(model)
  assert.equal(resolved.text.color, '#ABCDEF'); assert.equal(resolved.text.highlight, '#808000')
  assert.equal(model.defaultStyle.text.color, 'ABCDEF'); assert.equal(model.defaultStyle.text.highlight, 'darkYellow')
  model.defaultStyle.text.color = 'AUTO'; model.defaultStyle.text.highlight = 'none'
  assert.equal(documentDefaultStyle(model).text.color, 'auto'); assert.equal(documentDefaultStyle(model).text.highlight, undefined)
})

test('a table range clears only its addressed paragraph and removes stale generated table HTML', () => {
  const model = createBlankDocument(), table = createParagraph()
  table.type = 'table'; table.runs = []; table.metadata.table = createTable(1, 2); table.metadata.html = '<table>old generated appearance</table>'
  const first = table.metadata.table.rows[0].cells[0].blocks[0], second = table.metadata.table.rows[0].cells[1].blocks[0]
  first.runs = [createRun('cell', { italic: true })]; second.runs = [createRun('other', { bold: true })]
  model.blocks = [table]
  const original = structuredClone(second), selection = range(point(table, 0, { row: 0, col: 0, cellBlockId: first.id }), point(table, 4, { row: 0, col: 0, cellBlockId: first.id }))
  clear(model, selection)
  assert.equal(first.runs[0].italic, false); assert.deepEqual(second, original); assert.equal(table.metadata.html, undefined)
})

test('paragraph permissions and a content lock outside the selected word reject the complete clear plan', () => {
  for (const denied of ['paragraph', 'content-control']) {
    const model = createBlankDocument(), block = createParagraph('word locked', { bold: true })
    model.blocks = [block]; block.style.align = 'center'
    if (denied === 'paragraph') model.permissionRules = [{ id: 'deny', scope: 'paragraph', targetId: block.id, action: 'edit', effect: 'deny', userIds: [] }]
    else createDocumentContentControl(model, { type: 'text', blockId: block.id, startOffset: 5, endOffset: 11, lock: 'contentLocked' })
    const before = structuredClone(model), selection = range(point(block, 0), point(block, 4))
    const state = documentClearFormattingState(model, selection, context(model))
    assert.equal(state.available, false); assert.ok(state.reason)
    assert.throws(() => clear(model, selection)); assert.deepEqual(model, before)
  }
})

test('invalid document defaults fail before any text, list or paragraph formatting mutation', () => {
  for (const invalidate of [
    style => { style.paragraph.lineHeight = 0 },
    style => { style.paragraph.wordWrap = 'yes' },
    style => { style.paragraph.snapToGrid = 1 },
    style => { style.text.color = 'red' },
    style => { style.text.color = '#FFFFFF\n' },
    style => { style.text.highlight = 'tomato' },
    style => { style.text.highlight = 'constructor' },
    style => { style.text.highlight = '__proto__' },
    style => { style.text.highlight = 42 },
  ]) {
    const model = createBlankDocument(), block = createParagraph('word', { bold: true })
    model.blocks = [block]; model.defaultStyle = defaults(); invalidate(model.defaultStyle)
    const before = structuredClone(model), selection = range(point(block, 0), point(block, 4))
    assert.equal(documentClearFormattingState(model, selection, context(model)).available, false)
    assert.throws(() => clear(model, selection), RangeError); assert.deepEqual(model, before)
  }
})

test('a missing legacy default style uses compatibility defaults while malformed declared defaults fail atomically', () => {
  for (const value of [null, false, 0, { ...defaults(), version: 2 }]) {
    const model = createBlankDocument(), block = createParagraph('word', { bold: true })
    model.blocks = [block]; model.defaultStyle = value
    const before = structuredClone(model), selection = range(point(block, 0), point(block, 4))
    assert.equal(documentClearFormattingState(model, selection, context(model)).available, false)
    assert.throws(() => clear(model, selection), RangeError); assert.deepEqual(model, before)
  }
})

test('a locked cached-label wrapper rejects clear before deleting the label, its comment or its control', () => {
  const model = createBlankDocument(), block = createParagraph('[0] Body', { bold: true })
  block.metadata.officeNumberingMarker = { version: 1, bookmark: 'office_list_v1_locked', bookmarkId: 99, text: '[0] ', sourceNumbering: null,
    xml: '<w:bookmarkStart w:id="99" w:name="office_list_v1_locked"/><w:r><w:t>[0] </w:t></w:r><w:bookmarkEnd w:id="99"/>', baseline: officeNumberingMarkerBaseline(block, '[0] ') }
  model.blocks = [block]
  createDocumentContentControl(model, { id: 'label-wrapper', type: 'richText', blockId: block.id, startOffset: 0, endOffset: 4, lock: 'controlLocked' })
  model.comments = { review: { id: 'review', unitId: model.id, referenceText: '[0]', resolved: false,
    anchor: [{ blockId: block.id, startOffset: 0, endOffset: 3, text: '[0]' }], messages: [],
    createdAt: '2026-11-22T00:00:00Z', updatedAt: '2026-11-22T00:00:00Z' } }
  const before = structuredClone(model), selection = range(point(block, 4), point(block, 8))
  const state = documentClearFormattingState(model, selection, context(model))
  assert.equal(state.available, false); assert.match(state.reason, /wrapper is locked/)
  assert.throws(() => clear(model, selection), /wrapper is locked/); assert.deepEqual(model, before)
})

test('a rejected paragraph inside a cross-body/table clear leaves every earlier paragraph and cached label intact', () => {
  for (const denied of ['paragraph', 'content-control']) {
    const model = createBlankDocument(), first = createParagraph('[0] First', { bold: true }), table = createParagraph()
    first.type = 'heading'; first.level = 2; first.style.align = 'center'
    first.metadata.officeNumberingMarker = { version: 1, bookmark: 'office_list_v1_atomic', bookmarkId: 99, text: '[0] ', sourceNumbering: null,
      xml: '<w:bookmarkStart w:id="99" w:name="office_list_v1_atomic"/><w:r><w:t>[0] </w:t></w:r><w:bookmarkEnd w:id="99"/>', baseline: officeNumberingMarkerBaseline(first, '[0] ') }
    table.type = 'table'; table.runs = []; table.metadata.table = createTable(1, 2); table.metadata.html = '<table>retained until success</table>'
    const earlier = table.metadata.table.rows[0].cells[0].blocks[0], last = table.metadata.table.rows[0].cells[1].blocks[0]
    earlier.runs = [createRun('Earlier', { italic: true })]; last.runs = [createRun('Later', { underline: true })]
    model.blocks = [first, table]
    if (denied === 'paragraph') model.permissionRules = [{ id: 'deny-last', scope: 'paragraph', targetId: last.id, action: 'edit', effect: 'deny', userIds: [] }]
    else createDocumentContentControl(model, { id: 'last-lock', type: 'text', blockId: last.id, startOffset: 3, endOffset: 5, lock: 'contentLocked' })
    const selection = range(point(first, 4), point(table, 3, { row: 0, col: 1, cellBlockId: last.id })), before = structuredClone(model)
    assert.equal(documentClearFormattingState(model, selection, context(model)).available, false)
    assert.throws(() => clear(model, selection)); assert.deepEqual(model, before)
  }
})

test('clearing a cached numbering label remaps comments and controls and keeps zero-length note ids', () => {
  const model = createBlankDocument(), block = createParagraph('[0] Body', { bold: true })
  block.runs = [createRun('', { metadata: { footnoteId: 'before-label' } }), createRun('[0] ', { bold: true }), createRun('', { metadata: { field: { id: 'after-label' } } }), createRun('Body', { bold: true })]
  block.metadata.officeNumberingMarker = { version: 1, bookmark: 'office_list_v1_test', bookmarkId: 99, text: '[0] ', sourceNumbering: null,
    xml: '<w:bookmarkStart w:id="99" w:name="office_list_v1_test"/><w:r><w:t>[0] </w:t></w:r><w:bookmarkEnd w:id="99"/>', baseline: officeNumberingMarkerBaseline(block, '[0] ') }
  model.blocks = [block]
  model.comments = { review: { id: 'review', unitId: model.id, referenceText: 'Body', resolved: false,
    anchor: [{ blockId: block.id, startOffset: 4, endOffset: 8, text: 'Body' }],
    messages: [{ id: 'message', text: 'Review', authorId: 'local-user', authorName: 'Author', createdAt: '2026-11-22T00:00:00Z' }],
    createdAt: '2026-11-22T00:00:00Z', updatedAt: '2026-11-22T00:00:00Z' } }
  createDocumentContentControl(model, { id: 'body-control', type: 'richText', blockId: block.id, startOffset: 4, endOffset: 8 })
  const noteIds = block.runs.filter(run => !run.text.length).map(run => run.id)
  const output = clear(model, range(point(block, 4), point(block, 8)))
  assert.deepEqual(output, range(point(block, 0), point(block, 4)))
  assert.equal(runText(block.runs), 'Body'); assert.equal(block.metadata.officeNumberingMarker, undefined)
  assert.deepEqual(block.runs.filter(run => !run.text.length).map(run => run.id), noteIds)
  assert.deepEqual(model.comments.review.anchor.map(anchor => [anchor.startOffset, anchor.endOffset, anchor.text]), [[0, 4, 'Body']])
  assert.deepEqual(model.contentControls.map(control => [control.startOffset, control.endOffset]), [[0, 4]])
})

test('imported defaults normalize resolved font slots, Office highlights, exact spacing and retained provenance', () => {
  const source = { version: 1, text: { resolvedFontFamilyAscii: 'Georgia', resolvedFontFamilyEastAsia: '宋体', fontSizePt: 15, color: 'auto', italic: true, highlight: 'darkYellow' }, paragraph: { align: 'right', indent: { leftTwips: -180, hangingTwips: 210 }, spacing: { lineTwips: 360, lineRule: 'exact', afterPt: 12 } }, wordWrap: false }
  const model = normalizeImportedDocument({ schemaVersion: 1, parserVersion: 'contract', format: 'docx', fileName: 'defaults.docx', size: 1, model: 'document',
    payload: { modelType: 'document', data: { documentId: 'defaults-doc', defaultStyle: source, blocks: [{ id: 'p', type: 'paragraph', details: { inlines: [{ type: 'text', text: 'word', style: { highlight: 'darkYellow' } }] } }] } },
    diagnostics: { warnings: [], errors: [], stats: {}, capabilities: [] },
  })
  assert.deepEqual([model.defaultStyle.text.fontFamily, model.defaultStyle.text.eastAsiaFontFamily, model.defaultStyle.text.highlight, model.defaultStyle.text.italic], ['Georgia', '宋体', '#808000', true])
  assert.deepEqual([model.defaultStyle.paragraph.lineHeight, model.defaultStyle.paragraph.spacingRule, model.defaultStyle.paragraph.indentStart, model.defaultStyle.paragraph.hanging, model.defaultStyle.paragraph.wordWrap], [24, 'exact', -12, 14, false])
  assert.equal(model.blocks[0].runs[0].metadata.highlight, '#808000')
  assert.equal(model.blocks[0].runs[0].highlight, '#808000')
  assert.deepEqual(model.metadata.source.defaultStyle, source)
  const copy = documentDefaultStyle(model); copy.text.fontFamily = 'Changed'; copy.paragraph.align = 'left'
  assert.equal(model.defaultStyle.text.fontFamily, 'Georgia'); assert.equal(model.defaultStyle.paragraph.align, 'right')
  assert.equal(normalizeDocumentHighlight('none', '#FF0000'), '#FF0000')
  assert.equal(normalizeDocumentHighlight('none'), undefined)
  assert.equal(normalizeDocumentHighlight(undefined, 'abcdef'), '#ABCDEF')
  for (const value of ['constructor', '__proto__', 'toString', '#123', null, 42]) assert.equal(normalizeDocumentHighlight(value), undefined)
})
