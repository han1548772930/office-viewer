import assert from 'node:assert/strict'
import { test } from 'vitest'
import { normalizeDocument } from '../src/parser/documentNormalizer.ts'
import { documentTextFieldRanges } from '../src/features/document/model/document-text-fields.ts'
import { documentTocHeadingTarget } from '../src/features/document/toc/document-toc.ts'

// Parser-to-domain drafts; execution and real DOCX samples remain in the
// unified validation phase. These DTOs do not prove file-format roundtrips.
function paragraph(id, text, details = {}) {
  return { id, type: 'paragraph', content: text, details: { inlines: [{ type: 'text', text }], ...details } }
}
function fixture() {
  const first = paragraph('entry-one', 'Chapter One\t99', { documentTextFields: { version: 1, fields: [
    { id: 'native-toc', field: 'TOC', instruction: 'TOC \\o "1-3" \\h \\z \\u',
      startOffset: 0, endBlockId: 'entry-two', endOffset: 13, cachedResult: 'Chapter One\t99\nChapter Two\t7', containsNestedFields: true },
    { id: 'page-one', field: 'PAGEREF', instruction: 'PAGEREF "_TocOne" \\h', startOffset: 12, endOffset: 14, cachedResult: '99' },
  ] } })
  const second = paragraph('entry-two', 'Chapter Two\t7', { documentTextFields: { version: 1, fields: [
    { id: 'page-two', field: 'PAGEREF', instruction: 'PAGEREF _TocTwo \\h', startOffset: 12, endOffset: 13, cachedResult: '7' },
  ] } })
  const headings = [paragraph('heading-one', 'Chapter One', { headingLevel: 1, bookmarks: [{ name: '_TocOne', childIndex: 0 }] }),
    paragraph('heading-two', 'Chapter Two', { headingLevel: 2, bookmarks: [{ name: '_TocTwo', childIndex: 0 }] })]
  return { schemaVersion: 1, parserVersion: 'contract-draft', format: 'docx', fileName: 'native-toc.docx', size: 128,
    model: 'document', diagnostics: {}, payload: { modelType: 'document', data: { documentId: 'native-toc-import', blocks: [first, second, ...headings] } } }
}
const body = model => model.blocks.map(block => block.runs.map(run => run.text).join('')).join('\n')

test('body ranges and heading bookmarks are restored before native TOC context is derived', () => {
  const source = fixture(), before = structuredClone(source), model = normalizeDocument(source)
  assert.deepEqual(source, before)
  assert.equal(model.modified, false)
  assert.equal(body(model), 'Chapter One\t99\nChapter Two\t7\nChapter One\nChapter Two')
  assert.deepEqual(model.metadata.nativeTocHydration, {
    recognized: [{ fieldId: 'native-toc', tocId: 'native-toc', blockIds: ['entry-one', 'entry-two'] }], retained: [], unmapped: [],
  })
  const first = model.blocks[0], second = model.blocks[1]
  assert.equal(first.metadata.documentToc.headingId, 'heading-one')
  assert.equal(first.metadata.documentToc.pageNumber, 99)
  assert.equal(second.metadata.documentToc.headingId, 'heading-two')
  assert.equal(second.metadata.documentToc.pageNumber, 7)
  assert.equal(first.metadata.documentToc.native.fieldId, 'native-toc')
  assert.equal(documentTocHeadingTarget(model, first), 'heading-one')
  assert.deepEqual(documentTextFieldRanges(model).map(range => range.field),
    source.payload.data.blocks.slice(0, 2).flatMap(block => block.details.documentTextFields.fields))
  assert.equal(model.metadata.source, source.payload.data)
})

test('unsupported TOC instructions keep native text and fields and report full update separately from pages', () => {
  const source = fixture(), field = source.payload.data.blocks[0].details.documentTextFields.fields[0]
  field.instruction += ' \\t "Custom,1"'
  const model = normalizeDocument(source), entry = model.metadata.nativeTocHydration.recognized[0]
  assert.match(entry.fullUpdateUnavailableReason, /尚未接入/)
  assert.equal(entry.pageUpdateUnavailableReason, undefined)
  assert.equal(model.blocks[0].metadata.documentToc.native.instruction, field.instruction)
  assert.equal(documentTextFieldRanges(model)[0].field.instruction, field.instruction)
  assert.equal(model.modified, false)
})

test('a bookmark inside a heading is retained without inventing a paragraph-start target', () => {
  const source = fixture()
  source.payload.data.blocks[2].details.bookmarks[0].childIndex = 1
  const model = normalizeDocument(source), entry = model.metadata.nativeTocHydration.recognized[0]
  assert.match(entry.pageUpdateUnavailableReason, /段首标题书签/)
  assert.equal(model.blocks[0].metadata.documentToc.headingId, undefined)
  assert.equal(documentTextFieldRanges(model).find(range => range.field.id === 'page-one').field.cachedResult, '99')
  assert.equal(model.modified, false)
})

test('partial native result ranges remain preserved and are reported without assigning a TOC context', () => {
  const source = fixture()
  source.payload.data.blocks[0].details.documentTextFields.fields[0].startOffset = 1
  const model = normalizeDocument(source), report = model.metadata.nativeTocHydration
  assert.equal(report.recognized.length, 0)
  assert.match(report.unmapped[0].reason, /完整的顶层段落范围/)
  assert.equal(model.blocks[0].metadata.documentToc, undefined)
  assert.equal(documentTextFieldRanges(model)[0].field.startOffset, 1)
  assert.equal(model.modified, false)
})

test('invalid cross-paragraph field IDs fail normalization before native context can disguise them', () => {
  const source = fixture()
  source.payload.data.blocks[0].details.documentTextFields.fields[0].endBlockId = 'missing'
  assert.throws(() => normalizeDocument(source), /末段不存在/)
  assert.equal(source.payload.data.blocks[0].details.documentTextFields.fields[0].endBlockId, 'missing')
})

test('documents without native TOC fields have an explicit empty hydration report', () => {
  const source = fixture()
  for (const block of source.payload.data.blocks) delete block.details.documentTextFields
  const model = normalizeDocument(source)
  assert.deepEqual(model.metadata.nativeTocHydration, { recognized: [], retained: [], unmapped: [] })
  assert.ok(model.blocks.every(block => block.metadata.documentToc === undefined))
  assert.equal(model.modified, false)
})
