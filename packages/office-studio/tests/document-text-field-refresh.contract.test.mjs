import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import { createParagraph, createRun, createTable, runText } from '../src/features/document/model/document-model.ts'
import { documentFieldAutomaticRefreshBlocked, documentTextFields, setDocumentTextFields, transformDocumentTextFieldsForBlockEdit } from '../src/features/document/model/document-text-fields.ts'
import { executeDocumentCommand, replaceSelectionText } from '../src/features/document/commands/document-commands.ts'
import { replaceDocumentRange } from '../src/features/document/services/document-facade.ts'
import { canRefreshDocumentStoryFields, documentBodySelectionHasField, prepareDocumentStoryFieldRefresh } from '../src/features/document/services/document-story-fields.ts'
import { DOCUMENT_STORY_KEYS } from '../src/features/document/model/document-story.ts'
import { EMPTY_SECTION_HEADER_FOOTER, resolveSectionHeaderFooter } from '../src/features/document/services/section-header-footer.ts'

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

test('a nested table field uses its owning section and ancestor table permission', () => {
  const { model } = bodyModel()
  const root = { ...createParagraph(), id: 'root-table', type: 'table', metadata: { table: createTable(1, 1) } }
  const nested = { ...createParagraph(), id: 'nested-table', type: 'table', metadata: { table: createTable(1, 1) } }
  const target = nested.metadata.table.rows[0].cells[0].blocks[0]
  target.id = 'nested-body'; target.runs = [createRun('old')]
  setDocumentTextFields(target, [dateField('nested-field', 0, 3)])
  root.metadata.table.rows[0].cells[0].blocks = [nested]
  model.blocks.push({ ...createParagraph(), id: 'second-section', type: 'sectionBreak', sectionBreak: { type: 'nextPage' }, runs: [] }, root)
  for (const [scope, targetId] of [['section', 'second-section'], ['entity', 'root-table']]) {
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
