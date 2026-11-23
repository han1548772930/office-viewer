import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import { createParagraph, createRun, createTable, runText } from '../src/features/document/model/document-model.ts'
import { documentFieldAutomaticRefreshBlocked, documentTextFieldRanges, documentTextFields, setDocumentTextFields, transformDocumentTextFieldsForBlockEdit } from '../src/features/document/model/document-text-fields.ts'
import { executeDocumentCommand, replaceSelectionText } from '../src/features/document/commands/document-commands.ts'
import { replaceDocumentRange } from '../src/features/document/services/document-facade.ts'
import { canRefreshDocumentStoryFields, canRefreshSelectedDocumentBodyFields, documentBodySelectionHasField, prepareDocumentStoryFieldRefresh } from '../src/features/document/services/document-story-fields.ts'
import { reserveDocumentFieldOpenRefresh } from '../src/features/document/services/document-field-open-refresh.ts'
import { DOCUMENT_STORY_KEYS } from '../src/features/document/model/document-story.ts'
import { EMPTY_SECTION_HEADER_FOOTER, resolveSectionHeaderFooter } from '../src/features/document/services/section-header-footer.ts'
import { deleteRepeatingSectionItem, duplicateRepeatingSectionItem, replacePictureContentControlImage, resetDocumentContentControlPlaceholder } from '../src/features/document/services/content-controls.ts'
import { deleteDocumentNote, insertDocumentNote } from '../src/features/document/notes/document-notes.ts'
import { insertLatexFormula } from '../src/features/document/latex/latex-formula.ts'

// DOC-03 contract drafts. Implementation batch 2026-11-23; execution is deferred.
const copy = value => JSON.parse(JSON.stringify(value))
const context = { documentId: 'body-fields', userId: 'editor' }
const timestamp = new Date(2030, 5, 15, 13, 24, 0).getTime()
const request = { locale: 'en-US', timestamp }
const point = (blockId, offset) => ({ blockId, offset })
const selection = (anchor, focus = anchor) => ({ anchor: { ...anchor }, focus: { ...focus } })
const dateField = (id, startOffset, endOffset, extra = {}) => ({
  id, field: 'DATE', startOffset, endOffset, instruction: 'DATE \\@ "yyyy-MM-dd"', cachedResult: 'old', ...extra,
})
function bodyModel(text = 'old') {
  const block = createParagraph(text); block.id = 'body'
  const model = { ...createBlankDocument(), id: 'body-fields', blocks: [block] }
  setDocumentTextFields(block, [dateField('date', 0, text.length, { metadata: { source: { retained: ['original'] } } })])
  return { model, block }
}

test('one candidate refreshes body and every independent story while mapping controls, comments and selection', () => {
  const { model, block } = bodyModel('📅 old / old')
  block.runs = [createRun('📅 '), createRun('o', { bold: true, color: '#aa0000' }),
    createRun('ld', { italic: true, href: 'https://example.org/date' }), createRun(' / '), createRun('old')]
  setDocumentTextFields(block, [dateField('first', 3, 6), dateField('second', 9, 12, { cachedResult: '2030-06-15' })])
  const boundary = { ...createParagraph(), id: 'section-two', type: 'sectionBreak', sectionBreak: { type: 'nextPage' }, runs: [] }
  model.blocks.push(boundary, createParagraph('Second section'))
  const fields = DOCUMENT_STORY_KEYS.map(story => ({ story, field: 'DATE', offset: 0,
    instruction: 'DATE \\@ "yyyy-MM-dd"', cachedResult: 'old' }))
  model.sectionHeaderFooters = {
    first: { ...EMPTY_SECTION_HEADER_FOOTER, pageFields: copy(fields) },
    'section-two': { ...EMPTY_SECTION_HEADER_FOOTER, storyLinkToPrevious: { header: true }, pageFields: copy(fields) },
  }
  model.contentControls = [{ id: 'wrapper', type: 'richText', blockId: block.id, startOffset: 9, endOffset: 12, lock: 'sdtLocked' }]
  model.comments = { comment: { id: 'comment', anchor: [{ blockId: block.id, startOffset: 9, endOffset: 12, text: 'old' }],
    referenceText: 'old', messages: [], createdAt: '2020-01-01T00:00:00Z', updatedAt: '2020-01-01T00:00:00Z' } }
  const before = copy(model)
  const plan = prepareDocumentStoryFieldRefresh(model, request, context, selection(point('body', 10)))
  assert.equal(plan.updatedFields, 13)
  assert.deepEqual(model, before)
  const candidate = { ...model, blocks: plan.blocks, sectionHeaderFooters: plan.sectionHeaderFooters }
  assert.equal(runText(plan.blocks[0].runs), '📅 2030-06-15 / 2030-06-15')
  assert.deepEqual(documentTextFields(plan.blocks[0]).map(field => [field.id, field.startOffset, field.endOffset, field.cachedResult]),
    [['first', 3, 13, '2030-06-15'], ['second', 16, 26, '2030-06-15']])
  assert.equal(plan.selection.focus.offset, 17)
  assert.deepEqual(plan.contentControls.map(control => [control.id, control.startOffset, control.endOffset, control.lock]), [['wrapper', 16, 26, 'sdtLocked']])
  assert.deepEqual(plan.comments.comment.anchor[0], { blockId: 'body', startOffset: 16, endOffset: 26, text: '2030-06-15' })
  assert.equal(plan.comments.comment.referenceText, '2030-06-15')
  assert.ok(plan.blocks[0].runs.some(run => run.bold && run.color === '#aa0000' && run.text === '2'))
  assert.ok(plan.blocks[0].runs.some(run => run.italic && run.href === 'https://example.org/date' && run.text === '030-06-15'))
  for (const id of ['first', 'section-two']) for (const story of DOCUMENT_STORY_KEYS) {
    assert.equal(resolveSectionHeaderFooter(candidate, id).pageFields.find(field => field.story === story).cachedResult, '2030-06-15')
  }
  assert.equal(prepareDocumentStoryFieldRefresh(candidate, request, context), undefined)
})

test('a caret within a body result selects only that field and leaves stories untouched', () => {
  const { model, block } = bodyModel('old old')
  setDocumentTextFields(block, [dateField('first', 0, 3), dateField('second', 4, 7)])
  model.sectionHeaderFooters = { first: { ...EMPTY_SECTION_HEADER_FOOTER,
    pageFields: [{ story: 'header', field: 'DATE', offset: 0, instruction: 'DATE \\@ "yyyy-MM-dd"', cachedResult: 'old' }] } }
  const range = selection(point('body', 5)), before = copy(model.sectionHeaderFooters)
  assert.equal(documentBodySelectionHasField(model, range), true)
  const plan = prepareDocumentStoryFieldRefresh(model, { ...request, stories: [], bodySelection: range }, context)
  assert.equal(plan.updatedFields, 1)
  assert.equal(runText(plan.blocks[0].runs), 'old 2030-06-15')
  assert.deepEqual(plan.sectionHeaderFooters, before)
  assert.equal(prepareDocumentStoryFieldRefresh(model, { ...request, stories: [] }, context), undefined)
  assert.equal(prepareDocumentStoryFieldRefresh(model, { ...request, stories: [], bodySelection: selection(point('body', 3)) }, context), undefined)
})

test('a multi-paragraph selection ending at offset zero excludes that final paragraph', () => {
  const { model, block } = bodyModel()
  const following = createParagraph('old'); following.id = 'following'
  setDocumentTextFields(following, [dateField('following-date', 0, 3)])
  model.blocks.push(following)
  const range = selection(point(block.id, 0), point(following.id, 0))
  const plan = prepareDocumentStoryFieldRefresh(model, { ...request, stories: [], bodySelection: range }, context)
  assert.equal(plan.updatedFields, 1)
  assert.equal(runText(plan.blocks[0].runs), '2030-06-15')
  assert.equal(runText(plan.blocks[1].runs), 'old')
})

test('refresh skips native locks, content locks, denied fields and outer fields while retaining unknown instructions', () => {
  const { model, block } = bodyModel('old old old old old old')
  setDocumentTextFields(block, [dateField('native-lock', 0, 3, { locked: true }),
    dateField('source-lock', 4, 7, { metadata: { source: { properties: { locked: 1 } } } }),
    dateField('content-lock', 8, 11), dateField('denied', 12, 15),
    dateField('outer', 16, 23, { containsNestedFields: true, cachedResult: 'old old' }),
    dateField('inner', 16, 19), { id: 'unknown', field: 'MERGEFIELD', instruction: 'MERGEFIELD Customer', cachedResult: 'old', startOffset: 20, endOffset: 23 }])
  model.contentControls = [{ id: 'protected', type: 'text', blockId: 'body', startOffset: 8, endOffset: 11, lock: 'contentLocked' }]
  model.permissionRules = [{ id: 'field-denial', scope: 'entity', targetId: 'denied', action: 'edit', effect: 'deny', userIds: ['editor'] }]
  assert.equal(canRefreshDocumentStoryFields(model, context), true)
  const before = copy(model), plan = prepareDocumentStoryFieldRefresh(model, request, context)
  assert.equal(plan.updatedFields, 1)
  assert.equal(runText(plan.blocks[0].runs), 'old old old old 2030-06-15 old')
  const result = documentTextFields(plan.blocks[0])
  for (const id of ['native-lock', 'source-lock', 'content-lock', 'denied', 'outer', 'unknown']) {
    assert.equal(result.find(field => field.id === id).cachedResult, documentTextFields(block).find(field => field.id === id).cachedResult)
  }
  assert.equal(result.find(field => field.id === 'unknown').instruction, 'MERGEFIELD Customer')
  assert.deepEqual(model, before)
  assert.equal(documentFieldAutomaticRefreshBlocked({ locked: false, metadata: { source: { locked: true } } }), false)
})

test('the first result character language takes precedence over field and UI language', () => {
  const { model, block } = bodyModel()
  block.runs[0].metadata = { source: { style: { language: 'fr-FR' } } }
  setDocumentTextFields(block, [dateField('date', 0, 3, { instruction: 'DATE \\@ "MMMM"', language: 'en-US' })])
  const plan = prepareDocumentStoryFieldRefresh(model, { ...request, locale: 'de-DE' }, context)
  assert.equal(runText(plan.blocks[0].runs), 'juin')
  assert.deepEqual(plan.blocks[0].runs[0].metadata.source.style, { language: 'fr-FR' })
})

test('an empty result keeps its chosen language without taking provenance from the following character', () => {
  const { model, block } = bodyModel('Tail')
  block.runs[0].metadata = { source: { text: 'Tail', style: { language: 'fr-FR' } } }
  setDocumentTextFields(block, [dateField('empty', 0, 0, { instruction: 'DATE \\@ "MMMM"', cachedResult: '' })])
  const plan = prepareDocumentStoryFieldRefresh(model, request, context)
  assert.equal(runText(plan.blocks[0].runs), 'JuneTail')
  assert.equal(documentTextFields(plan.blocks[0])[0].language, 'en-US')
  assert.deepEqual(plan.blocks[0].runs.at(-1).metadata.source, { text: 'Tail', style: { language: 'fr-FR' } })
  assert.equal(prepareDocumentStoryFieldRefresh({ ...model, blocks: plan.blocks }, request, context), undefined)
})

test('field refresh keeps a code block source consistent with its run result', () => {
  const { model, block } = bodyModel()
  block.metadata.codeBlock = { source: 'old', language: 'plain', lineNumbers: true }
  const plan = prepareDocumentStoryFieldRefresh(model, request, context)
  assert.equal(plan.blocks[0].metadata.codeBlock.source, '2030-06-15')
  assert.equal(runText(plan.blocks[0].runs), plan.blocks[0].metadata.codeBlock.source)
  assert.equal(block.metadata.codeBlock.source, 'old')
})

test('a nested table field uses its owning section and ancestor table permission', () => {
  const { model } = bodyModel()
  const root = { ...createParagraph(), id: 'root-table', type: 'table', metadata: { table: createTable(1, 1) } }
  const nested = { ...createParagraph(), id: 'nested-table', type: 'table', metadata: { table: createTable(1, 1) } }
  const target = nested.metadata.table.rows[0].cells[0].blocks[0]
  target.id = 'nested-body'; target.runs = [createRun('old')]
  setDocumentTextFields(target, [dateField('nested-field', 0, 3)])
  root.metadata.table.rows[0].cells[0].blocks = [nested]
  model.blocks.push({ ...createParagraph(), id: 'second-section', type: 'sectionBreak', sectionBreak: { type: 'nextPage' }, runs: [] }, root)
  for (const [scope, targetId] of [['section', 'second-section'], ['entity', 'root-table'], ['paragraph', 'root-table']]) {
    model.permissionRules = [{ id: 'protected-container', scope, targetId, action: 'edit', effect: 'deny', userIds: [] }]
    const plan = prepareDocumentStoryFieldRefresh(model, request, context)
    assert.equal(plan.updatedFields, 1)
    const nestedResult = plan.blocks.at(-1).metadata.table.rows[0].cells[0].blocks[0].metadata.table.rows[0].cells[0].blocks[0]
    assert.equal(runText(nestedResult.runs), 'old')
  }
})

test('typing, field-result edits, paragraph split/merge and facade text replacement retain correct ranges', () => {
  const { model, block } = bodyModel('XoldY')
  setDocumentTextFields(block, [dateField('date', 1, 4)])
  replaceSelectionText(model, selection(point('body', 0)), 'Z')
  assert.equal(runText(block.runs), 'ZXoldY')
  assert.deepEqual(documentTextFields(block).map(field => [field.startOffset, field.endOffset, field.cachedResult]), [[2, 5, 'old']])
  const split = replaceSelectionText(model, selection(point('body', 2)), '\n')
  assert.equal(model.blocks.length, 2)
  assert.deepEqual(documentTextFields(model.blocks[1]).map(field => [field.id, field.startOffset, field.endOffset]), [['date', 0, 3]])
  executeDocumentCommand(model, split, { id: 'doc.command.delete', direction: -1 }, context)
  assert.equal(model.blocks.length, 1)
  assert.deepEqual(documentTextFields(model.blocks[0]).map(field => [field.id, field.startOffset, field.endOffset]), [['date', 2, 5]])
  const replaced = replaceDocumentRange(model, 2, 5, 'longer')
  assert.equal(runText(replaced.blocks[0].runs), 'ZXlongerY')
  assert.deepEqual(documentTextFields(replaced.blocks[0]).map(field => [field.id, field.startOffset, field.endOffset, field.cachedResult]), [['date', 2, 8, 'old']])
  const before = copy(replaced)
  assert.throws(() => replaceSelectionText(replaced, selection(point('body', 4)), '\n'), /段落拆分穿过字段/)
  assert.deepEqual(replaced, before)
})

test('rich paragraph paste inserts fresh field identities with shifted result ranges', () => {
  const destination = createParagraph('AB'); destination.id = 'destination'
  const model = { ...createBlankDocument(), id: 'body-fields', blocks: [destination] }
  const source = bodyModel().block, before = copy(source)
  executeDocumentCommand(model, selection(point('destination', 1)), { id: 'doc.command.paste-blocks', blocks: [source] }, context)
  assert.equal(runText(destination.runs), 'AoldB')
  const first = documentTextFields(destination)[0]
  assert.notEqual(first.id, 'date')
  assert.deepEqual([first.startOffset, first.endOffset, first.cachedResult], [1, 4, 'old'])
  executeDocumentCommand(model, selection(point('destination', 4)), { id: 'doc.command.paste-blocks', blocks: [source] }, context)
  const second = documentTextFields(destination)[1]
  assert.notEqual(second.id, first.id)
  assert.deepEqual([second.startOffset, second.endOffset], [4, 7])
  assert.deepEqual(source, before)
})

test('repeating-section duplication and deletion retain fields with independent IDs', () => {
  const { model, block } = bodyModel('AoldZ')
  setDocumentTextFields(block, [dateField('date', 1, 4, { metadata: { source: { retained: ['original'] } } })])
  model.contentControls = [{ id: 'section', type: 'repeatingSection', blockId: 'body', startOffset: 0, endOffset: 5 },
    { id: 'item', type: 'repeatingSectionItem', parentId: 'section', blockId: 'body', startOffset: 0, endOffset: 5 }]
  const duplicate = duplicateRepeatingSectionItem(model, 'item')
  assert.equal(runText(block.runs), 'AoldZAoldZ')
  const fields = documentTextFields(block)
  assert.equal(fields.length, 2); assert.notEqual(fields[0].id, fields[1].id)
  assert.deepEqual(fields.map(field => [field.startOffset, field.endOffset, field.cachedResult]), [[1, 4, 'old'], [6, 9, 'old']])
  fields[1].metadata.source.retained[0] = 'copy'
  assert.deepEqual(fields[0].metadata.source.retained, ['original'])
  deleteRepeatingSectionItem(model, duplicate.id)
  assert.equal(runText(block.runs), 'AoldZ')
  assert.deepEqual(documentTextFields(block).map(field => [field.id, field.startOffset, field.endOffset]), [['date', 1, 4]])
})

test('picture content replacement and placeholder reset move fields with the real suffix', () => {
  const { model, block } = bodyModel('Media old')
  setDocumentTextFields(block, [dateField('date', 6, 9)])
  model.contentControls = [{ id: 'picture', type: 'picture', blockId: 'body', startOffset: 0, endOffset: 5, placeholder: 'Picture' }]
  replacePictureContentControlImage(model, 'picture', { src: 'data:image/png;base64,AA==', alt: 'Sample', width: 32, height: 32 })
  assert.equal(model.blocks.length, 3)
  assert.deepEqual(documentTextFields(model.blocks[2]).map(field => [field.id, field.startOffset, field.endOffset]), [['date', 1, 4]])
  resetDocumentContentControlPlaceholder(model, 'picture')
  assert.equal(model.blocks.length, 1); assert.equal(runText(block.runs), 'Picture old')
  assert.deepEqual(documentTextFields(block).map(field => [field.id, field.startOffset, field.endOffset]), [['date', 8, 11]])
})

test('LaTeX insertion and note insertion/renumber/deletion keep neighboring field ranges', () => {
  const { model, block } = bodyModel()
  for (let index = 0; index < 9; index++) {
    const reference = createParagraph(); reference.id = `note-${index}`
    model.blocks.splice(model.blocks.indexOf(block), 0, reference)
    insertDocumentNote(model, selection(point(reference.id, 0)), 'footnote')
  }
  const inserted = insertDocumentNote(model, selection(point('body', 0)), 'footnote')
  assert.equal(runText(block.runs), '10old')
  assert.deepEqual(documentTextFields(block).map(field => [field.startOffset, field.endOffset]), [[2, 5]])
  deleteDocumentNote(model, 'footnote', inserted.note.id)
  assert.equal(runText(block.runs), 'old')
  assert.deepEqual(documentTextFields(block).map(field => [field.startOffset, field.endOffset]), [[0, 3]])
  insertLatexFormula(model, selection(point('body', 0)), 'x^2', 'inline')
  assert.equal(runText(block.runs), '\uFFFCold')
  assert.deepEqual(documentTextFields(block).map(field => [field.startOffset, field.endOffset]), [[1, 4]])
  insertLatexFormula(model, selection(point('body', 0)), 'y^2', 'block')
  const tail = model.blocks.at(-1)
  assert.equal(runText(tail.runs), '\uFFFCold')
  assert.deepEqual(documentTextFields(tail).map(field => [field.id, field.startOffset, field.endOffset]), [['date', 1, 4]])
})

test('invalid UTF-16, crossing metadata and malformed refresh requests fail before changing the model', () => {
  const { model, block } = bodyModel('📅old')
  const before = copy(model)
  assert.throws(() => transformDocumentTextFieldsForBlockEdit(block, 1, 2, 'x'), /UTF-16/)
  assert.deepEqual(model, before)
  assert.throws(() => setDocumentTextFields(block, [dateField('a', 0, 3), dateField('b', 2, 5)]), /部分交叉/)
  assert.deepEqual(model, before)
  assert.throws(() => prepareDocumentStoryFieldRefresh(model, { ...request, bodySelection: null }, context), /选区/)
  assert.throws(() => prepareDocumentStoryFieldRefresh(model, { ...request, onOpen: 1 }, context), /布尔值/)
  assert.throws(() => prepareDocumentStoryFieldRefresh(model, { ...request, bodySelection: selection(point('missing', 0)) }, context), /UTF-16/)
  assert.deepEqual(model, before)
})

test('paragraph breaks through a result retain its definition and edits in the end paragraph rebase its owner', () => {
  const { model, block } = bodyModel('AoldZ')
  setDocumentTextFields(block, [dateField('date', 1, 4)])
  replaceSelectionText(model, selection(point('body', 2)), '\n')
  const tail = model.blocks[1]
  assert.deepEqual(model.blocks.map(paragraph => runText(paragraph.runs)), ['Ao', 'ldZ'])
  assert.deepEqual(documentTextFields(block).map(field => [field.id, field.startOffset, field.endBlockId, field.endOffset, field.cachedResult]),
    [['date', 1, tail.id, 2, 'old']])
  replaceSelectionText(model, selection(point(tail.id, 1)), 'X')
  assert.equal(documentTextFields(block)[0].endOffset, 3)
  assert.equal(documentTextFields(block)[0].cachedResult, 'old')
  replaceSelectionText(model, selection(point(block.id, 2), point(tail.id, 0)), '')
  assert.equal(model.blocks.length, 1)
  assert.equal(runText(block.runs), 'AolXdZ')
  assert.deepEqual(documentTextFields(block).map(field => [field.startOffset, field.endBlockId, field.endOffset, field.cachedResult]), [[1, undefined, 5, 'old']])
  assert.equal(documentTextFieldRanges(model)[0].absoluteEnd, 5)
})

test('cross-paragraph refresh selects a caret in the last paragraph and preserves SDT, comments, styles and suffix text', () => {
  const first = createParagraph('Aold'); first.id = 'first'
  first.runs = [createRun('A'), createRun('old', { bold: true, color: '#aa0000' })]
  const last = createParagraph('cacheZ'); last.id = 'last'
  last.runs = [createRun('cache', { italic: true, href: 'https://example.org/cache' }), createRun('Z')]
  const model = { ...createBlankDocument(), id: context.documentId, blocks: [first, last] }
  setDocumentTextFields(first, [dateField('cross-date', 1, 5, { endBlockId: last.id, cachedResult: 'old\ncache', metadata: { source: { native: ['begin', 'separate', 'end'] } } })])
  model.contentControls = [
    { id: 'cache-wrapper', type: 'richText', blockId: last.id, startOffset: 0, endOffset: 5, lock: 'sdtLocked' },
    { id: 'suffix-wrapper', type: 'richText', blockId: last.id, startOffset: 5, endOffset: 6, lock: 'sdtLocked' },
  ]
  model.comments = { suffix: { id: 'suffix', anchor: [{ blockId: last.id, startOffset: 5, endOffset: 6, text: 'Z' }],
    referenceText: 'Z', messages: [], createdAt: '2020-01-01T00:00:00Z', updatedAt: '2020-01-01T00:00:00Z' } }
  const range = selection(point(last.id, 2), point(last.id, 6)), before = copy(model)
  assert.equal(documentBodySelectionHasField(model, selection(point(last.id, 2))), true)
  assert.equal(canRefreshSelectedDocumentBodyFields(model, selection(point(last.id, 2)), context), true)
  const plan = prepareDocumentStoryFieldRefresh(model, { ...request, stories: [], bodySelection: range }, context, range)
  assert.equal(plan.updatedFields, 1)
  assert.equal(plan.blocks.length, 1)
  assert.equal(runText(plan.blocks[0].runs), 'A2030-06-15Z')
  const field = documentTextFields(plan.blocks[0])[0]
  assert.deepEqual([field.id, field.startOffset, field.endBlockId, field.endOffset, field.cachedResult], ['cross-date', 1, undefined, 11, '2030-06-15'])
  assert.deepEqual(field.metadata.source, { native: ['begin', 'separate', 'end'] })
  assert.ok(plan.blocks[0].runs.some(run => run.bold && run.color === '#aa0000'))
  assert.ok(plan.blocks[0].runs.some(run => run.italic && run.href === 'https://example.org/cache'))
  assert.deepEqual(plan.selection, selection(point(first.id, 7), point(first.id, 12)))
  assert.deepEqual(plan.contentControls.map(control => [control.id, control.blockId, control.startOffset, control.endOffset, control.lock]),
    [['cache-wrapper', first.id, 1, 11, 'sdtLocked'], ['suffix-wrapper', first.id, 11, 12, 'sdtLocked']])
  assert.deepEqual(plan.comments.suffix.anchor, [{ blockId: first.id, startOffset: 11, endOffset: 12, text: 'Z' }])
  assert.equal(plan.comments.suffix.updatedAt, new Date(timestamp).toISOString())
  assert.deepEqual(model, before)
})

test('nested cross-paragraph definitions follow an inner update without evaluating an unknown enclosing field', () => {
  const { model, block } = bodyModel('Aold'), last = createParagraph('cacheZ'); last.id = 'last'; model.blocks.push(last)
  setDocumentTextFields(block, [
    { id: 'outer', field: 'IF', startOffset: 0, endBlockId: last.id, endOffset: 6, instruction: 'IF retained', cachedResult: 'Aold\ncacheZ' },
    dateField('date', 1, 5, { endBlockId: last.id, cachedResult: 'old\ncache' }),
  ])
  const plan = prepareDocumentStoryFieldRefresh(model, request, context)
  assert.equal(plan.updatedFields, 1)
  const fields = documentTextFields(plan.blocks[0])
  assert.deepEqual(fields.map(field => [field.id, field.endBlockId, field.endOffset, field.cachedResult]),
    [['outer', undefined, 12, 'Aold\ncacheZ'], ['date', undefined, 11, '2030-06-15']])
  assert.equal(documentTextFieldRanges({ blocks: plan.blocks }).length, 2)
})

test('cross-paragraph facade replacement relocates a deleted start into the replacement while preserving its remaining result', () => {
  const { model, block } = bodyModel('ABold'), last = createParagraph('cacheZ'); last.id = 'last'; model.blocks.push(last)
  setDocumentTextFields(block, [dateField('date', 2, 5, { endBlockId: last.id, cachedResult: 'old\ncache' })])
  const next = replaceDocumentRange(model, 1, 8, 'Q\nR')
  assert.deepEqual(next.blocks.map(paragraph => runText(paragraph.runs)), ['AQ', 'RcheZ'])
  const field = documentTextFields(next.blocks[0])[0]
  assert.deepEqual([field.startOffset, field.endBlockId, field.endOffset, field.cachedResult], [1, next.blocks[1].id, 4, 'old\ncache'])
  assert.deepEqual(model.blocks.map(paragraph => runText(paragraph.runs)), ['ABold', 'cacheZ'])
  assert.equal(documentTextFieldRanges(next).length, 1)
})

test('multi-paragraph paste remaps field endpoint IDs and copies provenance independently', () => {
  const first = createParagraph('old'); first.id = 'copied-first'
  const last = createParagraph('cache'); last.id = 'copied-last'
  setDocumentTextFields(first, [dateField('copied-date', 0, 5, { endBlockId: last.id, cachedResult: 'old\ncache', metadata: { source: { retained: ['original'] } } })])
  const host = createParagraph('BeforeAfter'); host.id = 'host'
  const model = { ...createBlankDocument(), id: context.documentId, blocks: [host] }
  executeDocumentCommand(model, selection(point(host.id, 6)), { id: 'doc.command.paste-blocks', blocks: [first, last] }, context)
  assert.deepEqual(model.blocks.map(paragraph => runText(paragraph.runs)), ['Before', 'old', 'cache', 'After'])
  const field = documentTextFields(model.blocks[1])[0]
  assert.notEqual(field.id, 'copied-date')
  assert.equal(field.endBlockId, model.blocks[2].id)
  assert.notEqual(field.endBlockId, last.id)
  field.metadata.source.retained.push('copy-only')
  assert.deepEqual(documentTextFields(first)[0].metadata.source.retained, ['original'])
  assert.equal(documentTextFieldRanges(model).length, 1)
})

test('model-level field validation rejects missing end IDs, surrogate boundaries, reverse order and cross-paragraph partial overlap', () => {
  const blocks = ['AAAA', 'BBBB', 'CCCC', 'DDDD'].map((text, index) => { const block = createParagraph(text); block.id = `p${index}`; return block })
  const model = { ...createBlankDocument(), id: context.documentId, blocks }
  setDocumentTextFields(blocks[0], [dateField('first', 1, 2, { endBlockId: 'p2' })])
  setDocumentTextFields(blocks[1], [dateField('second', 1, 2, { endBlockId: 'p3' })])
  assert.throws(() => documentTextFieldRanges(model), /部分交叉/)
  setDocumentTextFields(blocks[1], [])
  setDocumentTextFields(blocks[0], [dateField('missing', 0, 0, { endBlockId: 'missing' })])
  assert.throws(() => documentTextFieldRanges(model), /末段不存在/)
  setDocumentTextFields(blocks[0], [])
  setDocumentTextFields(blocks[2], [dateField('reverse', 0, 0, { endBlockId: 'p0' })])
  assert.throws(() => documentTextFieldRanges(model), /跨段 UTF-16/)
  setDocumentTextFields(blocks[2], [])
  blocks[1].runs = [createRun('😀')]
  setDocumentTextFields(blocks[0], [dateField('surrogate', 0, 1, { endBlockId: 'p1' })])
  assert.throws(() => documentTextFieldRanges(model), /跨段 UTF-16/)
})

test('cache-only refresh retains partial comment and control ranges and the current body caret', () => {
  const { model, block } = bodyModel('2030-06-15')
  model.contentControls = [{ id: 'partial', type: 'richText', blockId: block.id, startOffset: 2, endOffset: 4, lock: 'sdtLocked' }]
  model.comments = { partial: { id: 'partial', anchor: [{ blockId: block.id, startOffset: 2, endOffset: 4, text: '30' }],
    referenceText: '30', messages: [], createdAt: '2020-01-01T00:00:00Z', updatedAt: '2020-01-01T00:00:00Z' } }
  const caret = selection(point(block.id, 3)), plan = prepareDocumentStoryFieldRefresh(model, request, context, caret)
  assert.equal(plan.updatedFields, 1)
  assert.deepEqual(plan.contentControls, model.contentControls)
  assert.deepEqual(plan.comments, model.comments)
  assert.deepEqual(plan.selection, caret)
  assert.deepEqual(plan.blocks[0].runs, block.runs)
})

test('availability tolerates an out-of-date body selection while command validation and invalid field metadata remain strict', () => {
  const { model } = bodyModel()
  const missing = selection(point('removed', 0))
  assert.equal(documentBodySelectionHasField(model, missing), false)
  assert.equal(canRefreshSelectedDocumentBodyFields(model, missing, context), false)
  assert.throws(() => prepareDocumentStoryFieldRefresh(model, { ...request, bodySelection: missing }, context), /UTF-16/)
  model.blocks[0].metadata.documentTextFields.version = 2
  assert.throws(() => documentBodySelectionHasField(model, selection(point('body', 1))), /版本/)
})

test('field open refresh follows unit and history lifetime across Facades, undo/redo, remounts and independent hosts', () => {
  const host = {}, secondHost = {}
  assert.equal(reserveDocumentFieldOpenRefresh(host, 'first', 1), true)
  assert.equal(reserveDocumentFieldOpenRefresh(host, 'second', 2), true)
  assert.equal(reserveDocumentFieldOpenRefresh(host, 'first', 1), false)
  assert.equal(reserveDocumentFieldOpenRefresh(host, 'first', 1), false)
  assert.equal(reserveDocumentFieldOpenRefresh(host, 'first', 3), true)
  assert.equal(reserveDocumentFieldOpenRefresh(secondHost, 'first', 1), true)
})
