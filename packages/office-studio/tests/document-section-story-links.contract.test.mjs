import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import { createParagraph, createRun, runText } from '../src/features/document/model/document-model.ts'
import { DOCUMENT_STORY_KEYS, documentStoryDefaultStyle, documentStoryFieldCache, documentStoryParagraphs } from '../src/features/document/model/document-story.ts'
import { sectionStoryLinks } from '../src/features/document/model/document-section-stories.ts'
import { applySectionHeaderFooter, EMPTY_SECTION_HEADER_FOOTER, insertSectionPageField, resolveSectionHeaderFooter, sectionStoryLinkedToPrevious } from '../src/features/document/services/section-header-footer.ts'
import { replaceStoryRange, storyBuffer, storyDateField } from '../src/features/document/services/story-editor.ts'
import { prepareDocumentStoryTextFormatting } from '../src/features/document/services/document-story-formatting.ts'
import { prepareDocumentStoryClearFormatting } from '../src/features/document/services/document-clear-formatting.ts'
import { refreshDocumentStoryDateFields, refreshDocumentStoryFieldResult } from '../src/features/document/services/document-story-fields.ts'
import { layoutDocument } from '../src/features/document/layout/document-layout.ts'
import { findDocumentStoryPage } from '../src/features/document/layout/document-story-navigation.ts'

// Deferred service/layout contracts, 2026-11-22. No validation has been run.
const copy = value => JSON.parse(JSON.stringify(value))
const snapshot = (section, story) => ({ text: section[story], content: section.storyContent?.[story], fields: section.pageFields?.filter(field => field.story === story) ?? [] })
const measurer = { measure: text => Array.from(text).length * 8 }
function sourceModel(withLocks = false) {
  const model = createBlankDocument(), defaults = documentStoryDefaultStyle(model)
  model.id = 'section-story-links'; model.page = { width: 320, height: 280, marginTop: 30, marginBottom: 30, marginLeft: 30, marginRight: 30 }
  const first = createParagraph('First'), second = createParagraph('Second')
  const boundary = { ...createParagraph(), id: 'next-section', type: 'sectionBreak', runs: [], sectionBreak: { type: 'nextPage' } }
  model.blocks = [first, boundary, second]
  const section = { ...EMPTY_SECTION_HEADER_FOOTER, differentFirstPage: true, differentOddEven: true, storyContent: {}, pageFields: [] }
  for (const story of DOCUMENT_STORY_KEYS) {
    const prefix = story + '😀', text = prefix + '\nTail'
    section[story] = text
    section.storyContent[story] = {
      runs: [createRun(prefix, { ...defaults.text, bold: true, color: '#ab1234', metadata: { source: { xml: '<retained/>', nested: ['a', 'b'] } } }), createRun('\nTail', { ...defaults.text, italic: true, fontSize: 14 })],
      paragraphs: documentStoryParagraphs(text, defaults.paragraph).map(paragraph => ({ ...paragraph, metadata: { retained: { story } } })),
      ...(withLocks ? { locks: [{ id: story + '-tail', startOffset: prefix.length + 2, endOffset: text.length + 1, lock: 'contentLocked' }] } : {}),
      metadata: { retained: { story, data: [1, 2] } },
    }
    section.pageFields.push({ story, field: 'PAGE', offset: 0, instruction: 'PAGE', cachedResult: '7',
      metadata: { retained: { story }, officeStoryFieldCache: { version: 1, runs: [createRun('7', { ...defaults.text, underline: true, color: '#2345ab' })] } } })
  }
  model.sectionHeaderFooters = { first: section }
  return model
}

for (const story of DOCUMENT_STORY_KEYS) test(`only ${story} unlinks, freezes its rich content, and relinks to the latest source`, () => {
  const model = sourceModel(), before = resolveSectionHeaderFooter(model, 'next-section'), original = snapshot(before, story)
  applySectionHeaderFooter(model, 'next-section', { storyLinkToPrevious: { [story]: false } })
  const unlinked = resolveSectionHeaderFooter(model, 'next-section')
  assert.deepEqual(snapshot(unlinked, story), original)
  assert.equal(unlinked.linkToPrevious, false)
  for (const key of DOCUMENT_STORY_KEYS) assert.equal(sectionStoryLinkedToPrevious(unlinked, key), key !== story)
  const stored = model.sectionHeaderFooters['next-section']
  for (const key of DOCUMENT_STORY_KEYS) if (key !== story) {
    assert.equal(stored[key], '')
    assert.equal(stored.storyContent?.[key], undefined)
    assert.equal(stored.pageFields.some(field => field.story === key), false)
  }
  const first = resolveSectionHeaderFooter(model, 'first')
  applySectionHeaderFooter(model, 'first', replaceStoryRange(first, story, 1, 1, 'New', model))
  assert.deepEqual(snapshot(resolveSectionHeaderFooter(model, 'next-section'), story), original)
  const other = DOCUMENT_STORY_KEYS.find(key => key !== story), previous = resolveSectionHeaderFooter(model, 'first')
  applySectionHeaderFooter(model, 'first', { [other]: previous[other] + ' changed' })
  assert.deepEqual(snapshot(resolveSectionHeaderFooter(model, 'next-section'), other), snapshot(resolveSectionHeaderFooter(model, 'first'), other))
  applySectionHeaderFooter(model, 'next-section', { storyLinkToPrevious: { [story]: true } })
  assert.equal(resolveSectionHeaderFooter(model, 'next-section').linkToPrevious, true)
  assert.deepEqual(snapshot(resolveSectionHeaderFooter(model, 'next-section'), story), snapshot(resolveSectionHeaderFooter(model, 'first'), story))
  assert.equal(model.sectionHeaderFooters['next-section'][story], '')
})

test('legacy partial maps complete their fallback before mixed-state aggregation and later patches', () => {
  const model = sourceModel()
  model.sectionHeaderFooters['next-section'] = { ...EMPTY_SECTION_HEADER_FOOTER, header: 'Local', linkToPrevious: true, storyLinkToPrevious: { header: false } }
  let section = resolveSectionHeaderFooter(model, 'next-section')
  assert.equal(section.header, 'Local'); assert.equal(section.linkToPrevious, false)
  assert.equal(sectionStoryLinkedToPrevious(section, 'footer'), true)
  applySectionHeaderFooter(model, 'next-section', { storyLinkToPrevious: { footer: false } })
  section = resolveSectionHeaderFooter(model, 'next-section')
  assert.equal(sectionStoryLinkedToPrevious(section, 'evenPageFooter'), true)
  assert.equal(sectionStoryLinkedToPrevious(section, 'header'), false)
  assert.equal(sectionStoryLinkedToPrevious(section, 'footer'), false)
  applySectionHeaderFooter(model, 'next-section', { linkToPrevious: true, storyLinkToPrevious: { footer: false } })
  const links = sectionStoryLinks(resolveSectionHeaderFooter(model, 'next-section'))
  assert.equal(links.header, true); assert.equal(links.footer, false); assert.equal(links.firstPageHeader, true)
})

test('inheritance follows each story through several sections and absent stories unlink as empty content', () => {
  const model = sourceModel()
  const boundary = { ...createParagraph(), id: 'third-section', type: 'sectionBreak', runs: [], sectionBreak: { type: 'nextPage' } }
  model.blocks.push(boundary, createParagraph('Third'))
  applySectionHeaderFooter(model, 'next-section', { storyLinkToPrevious: { header: false }, header: 'Second only' })
  assert.equal(resolveSectionHeaderFooter(model, 'third-section').header, 'Second only')
  applySectionHeaderFooter(model, 'third-section', { storyLinkToPrevious: { footer: false } })
  applySectionHeaderFooter(model, 'first', { footer: 'New first footer' })
  assert.notEqual(resolveSectionHeaderFooter(model, 'third-section').footer, 'New first footer')
  assert.equal(resolveSectionHeaderFooter(model, 'next-section').footer, 'New first footer')
  const empty = createBlankDocument(); empty.blocks.push(boundary, createParagraph())
  applySectionHeaderFooter(empty, boundary.id, { storyLinkToPrevious: { evenPageHeader: false } })
  assert.equal(resolveSectionHeaderFooter(empty, boundary.id).evenPageHeader, '')
  assert.equal(sectionStoryLinkedToPrevious(resolveSectionHeaderFooter(empty, boundary.id), 'evenPageHeader'), false)
})

test('unlinked rich runs, paragraphs, source metadata, field caches and locks have independent ownership', () => {
  const model = sourceModel(true), source = copy(model.sectionHeaderFooters.first)
  applySectionHeaderFooter(model, 'next-section', { storyLinkToPrevious: { evenPageHeader: false } })
  const own = model.sectionHeaderFooters['next-section'].storyContent.evenPageHeader
  assert.deepEqual(own, source.storyContent.evenPageHeader)
  assert.notEqual(own, model.sectionHeaderFooters.first.storyContent.evenPageHeader)
  assert.notEqual(own.runs[0].metadata.source, model.sectionHeaderFooters.first.storyContent.evenPageHeader.runs[0].metadata.source)
  const read = resolveSectionHeaderFooter(model, 'next-section')
  read.storyContent.evenPageHeader.runs[0].metadata.source.nested.push('external')
  read.storyContent.evenPageHeader.locks[0].endOffset = 0
  read.pageFields.find(field => field.story === 'evenPageHeader').metadata.officeStoryFieldCache.runs[0].color = '#000000'
  assert.deepEqual(model.sectionHeaderFooters.first, source)
  assert.deepEqual(own, source.storyContent.evenPageHeader)
})

test('linked content and field edits are rejected atomically while unrelated unlinked edits accept echoed stories', () => {
  const model = sourceModel(), before = copy(model)
  assert.throws(() => applySectionHeaderFooter(model, 'next-section', { header: 'Shadow' }), /Unlink/)
  assert.deepEqual(model, before)
  applySectionHeaderFooter(model, 'next-section', { storyLinkToPrevious: { header: false } })
  const section = resolveSectionHeaderFooter(model, 'next-section')
  applySectionHeaderFooter(model, 'next-section', replaceStoryRange(section, 'header', 1, 1, 'Own', model))
  insertSectionPageField(model, 'next-section', 'header', 'NUMPAGES', { offset: 0 })
  const edited = copy(model)
  assert.throws(() => insertSectionPageField(model, 'next-section', 'footer', 'NUMPAGES'), /Unlink/)
  assert.deepEqual(model, edited)
  assert.deepEqual(snapshot(resolveSectionHeaderFooter(model, 'next-section'), 'footer'), snapshot(resolveSectionHeaderFooter(model, 'first'), 'footer'))
})

test('invalid maps, first-section association and a locked relink publish no partial section changes', () => {
  const model = sourceModel(), before = copy(model)
  for (const patch of [{ storyLinkToPrevious: null }, { storyLinkToPrevious: [] }, { storyLinkToPrevious: { mystery: true } }, { storyLinkToPrevious: { header: 1 } }, { unexpected: true }]) {
    assert.throws(() => applySectionHeaderFooter(model, 'next-section', patch)); assert.deepEqual(model, before)
  }
  for (const patch of [{ linkToPrevious: true }, ...DOCUMENT_STORY_KEYS.map(story => ({ storyLinkToPrevious: { [story]: true } }))]) {
    assert.throws(() => applySectionHeaderFooter(model, 'first', patch), /first section/); assert.deepEqual(model, before)
  }
  const defaults = documentStoryDefaultStyle(model), text = 'Protected'
  model.sectionHeaderFooters['next-section'] = { ...EMPTY_SECTION_HEADER_FOOTER, footer: text, storyContent: { footer: {
    runs: [createRun(text, defaults.text)], paragraphs: documentStoryParagraphs(text, defaults.paragraph), locks: [{ id: 'protected', startOffset: 0, endOffset: text.length, lock: 'contentLocked' }],
  } } }
  const locked = copy(model)
  assert.throws(() => applySectionHeaderFooter(model, 'next-section', { storyLinkToPrevious: { header: true, footer: true } }), /锁定|内容锁/)
  assert.deepEqual(model, locked)
})

test('format and clear commands use the requested story association rather than the mixed aggregate', () => {
  const model = sourceModel(), context = { documentId: model.id, userId: 'editor', defaultAllow: true }
  applySectionHeaderFooter(model, 'next-section', { storyLinkToPrevious: { header: false } })
  const request = { sectionId: 'next-section', story: 'header', start: 1, end: 3 }
  const plan = prepareDocumentStoryTextFormatting(model, request, { fontSize: 19 }, context)
  assert.ok(plan.patch)
  applySectionHeaderFooter(model, 'next-section', plan.patch)
  assert.ok(prepareDocumentStoryClearFormatting(model, request, context))
  assert.throws(() => prepareDocumentStoryTextFormatting(model, { ...request, story: 'footer' }, { fontSize: 19 }, context), /链接/)
  assert.throws(() => prepareDocumentStoryClearFormatting(model, { ...request, story: 'footer' }, context), /链接/)
  assert.equal(resolveSectionHeaderFooter(model, 'first').storyContent.header.runs[0].fontSize, 9)
})

test('known date refresh preserves mixed cache runs, fields in other stories and source metadata', () => {
  const model = sourceModel(), defaults = documentStoryDefaultStyle(model)
  const old = storyDateField('time', 'en-US', new Date(2026, 10, 22, 12, 34, 56))
  const field = { ...old, story: 'header', offset: 0, metadata: { retained: { values: [1, 2] }, officeStoryFieldCache: { version: 1, runs: [
    createRun('12:', { ...defaults.text, bold: true, color: '#aa0000' }), createRun('34:', { ...defaults.text, italic: true, color: '#00aa00' }), createRun('56', { ...defaults.text, underline: true, color: '#0000aa' }),
  ] } } }
  const before = copy(field), refreshed = refreshDocumentStoryFieldResult(field, '98:76:54', model)
  assert.equal(runText(documentStoryFieldCache(refreshed).runs), '98:76:54')
  assert.deepEqual(documentStoryFieldCache(refreshed).runs.map(run => [run.text, run.color]), [['98:', '#aa0000'], ['76:', '#00aa00'], ['54', '#0000aa']])
  assert.deepEqual(field, before)
  refreshed.metadata.retained.values.push(3); assert.deepEqual(field, before)
  for (const text of ['12:34:56!', '12:34:5', '12:34:😀']) {
    const result = refreshDocumentStoryFieldResult(field, text, model)
    assert.equal(runText(documentStoryFieldCache(result).runs), text)
  }
  applySectionHeaderFooter(model, 'first', { pageFields: [field, { ...copy(field), story: 'footer' }] })
  applySectionHeaderFooter(model, 'next-section', { storyLinkToPrevious: { header: false } })
  const section = resolveSectionHeaderFooter(model, 'next-section'), footer = copy(section.pageFields.find(field => field.story === 'footer'))
  const fields = refreshDocumentStoryDateFields(section, 'header', 'en-US', new Date(2026, 10, 22, 13, 35, 57), model)
  applySectionHeaderFooter(model, 'next-section', { pageFields: fields })
  assert.equal(resolveSectionHeaderFooter(model, 'next-section').pageFields.find(field => field.story === 'header').cachedResult, '13:35:57')
  assert.deepEqual(resolveSectionHeaderFooter(model, 'next-section').pageFields.find(field => field.story === 'footer'), footer)
  assert.throws(() => refreshDocumentStoryDateFields(section, 'footer', 'en-US', new Date(), model), /链接/)
})

test('a cached field result protected by a story content lock rejects the whole section patch', () => {
  const model = createBlankDocument(), defaults = documentStoryDefaultStyle(model), text = 'Label'
  const field = { story: 'header', field: 'PAGE', offset: 0, instruction: 'PAGE', cachedResult: '7', metadata: {
    retained: { source: 'locked-field' }, officeStoryFieldCache: { version: 1, runs: [createRun('7', { ...defaults.text, bold: true, color: '#123456' })] },
  } }
  model.sectionHeaderFooters = { first: { ...EMPTY_SECTION_HEADER_FOOTER, header: text, pageFields: [field], storyContent: { header: {
    runs: [createRun(text, defaults.text)], paragraphs: documentStoryParagraphs(text, defaults.paragraph),
    locks: [{ id: 'protected-field', startOffset: 0, endOffset: 1, lock: 'contentLocked' }],
  } } } }
  assert.equal(storyBuffer(resolveSectionHeaderFooter(model, 'first'), 'header').value[0], '\uFFFC')
  const before = copy(model), refreshed = refreshDocumentStoryFieldResult(field, '8', model)
  assert.throws(() => applySectionHeaderFooter(model, 'first', { differentFirstPage: true, pageFields: [refreshed] }), /锁定|内容锁/)
  assert.deepEqual(model, before)
  assert.deepEqual(field, before.sectionHeaderFooters.first.pageFields[0])
  applySectionHeaderFooter(model, 'first', { differentFirstPage: true, pageFields: [refreshDocumentStoryFieldResult(field, '7', model)] })
  assert.equal(resolveSectionHeaderFooter(model, 'first').differentFirstPage, true)
  assert.deepEqual(resolveSectionHeaderFooter(model, 'first').pageFields[0], before.sectionHeaderFooters.first.pageFields[0])
})

test('story navigation skips parity fillers and finds the matching section and logical-page variant', () => {
  const model = sourceModel()
  model.blocks[1].sectionBreak = { type: 'evenPage', setup: { pageNumberStart: 3 } }
  applySectionHeaderFooter(model, 'next-section', { differentFirstPage: false })
  const layout = layoutDocument(model, measurer), blank = layout.pages.findIndex(page => page.isParityBlank)
  assert.ok(blank > 0)
  assert.notEqual(findDocumentStoryPage(layout, { area: 'header', preferredPage: blank }), blank)
  const header = findDocumentStoryPage(layout, { area: 'header', sectionId: 'next-section', story: 'evenPageHeader', preferredPage: blank })
  assert.ok(header > blank)
  assert.equal(layout.pageHeaderFooters[header].headerStory, 'evenPageHeader')
  assert.equal(layout.pages[header].pageNumber, 4)
  assert.equal(findDocumentStoryPage(layout, { area: 'footer', sectionId: 'missing' }), undefined)
})
