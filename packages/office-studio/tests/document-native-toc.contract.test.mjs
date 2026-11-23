import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import { collapsedSelection, createParagraph, runText } from '../src/features/document/model/document-model.ts'
import { documentTextFields, documentTextFieldRanges, setDocumentTextFields } from '../src/features/document/model/document-text-fields.ts'
import { hydrateDocumentTocsFromNativeFields } from '../src/features/document/toc/document-native-toc.ts'
import { applyDocumentTocMutation, buildDocumentTocBlocks, buildDocumentTocInstruction, DEFAULT_DOCUMENT_TOC_OPTIONS,
  documentTocHeadingTarget, documentTocIdForSelection, documentTocOptionsAtBlock } from '../src/features/document/toc/document-toc.ts'
import { documentTocActionState, prepareDocumentTocAction } from '../src/features/document/toc/document-toc-command.ts'

// Deferred native-range/model contracts. No DOCX, browser, or test execution is claimed.
function nativeFixture(instruction = 'TOC \\o "1-3" \\h \\z \\u') {
  const model = createBlankDocument(); model.id = 'native-toc-document'
  const one = createParagraph('Alpha'), two = createParagraph('Beta')
  one.id = 'one'; one.type = 'heading'; one.level = 1
  two.id = 'two'; two.type = 'heading'; two.level = 2
  one.metadata.source = { details: { bookmarks: [{ name: '_TocAlpha', childIndex: 0 }] } }
  two.metadata.source = { details: { bookmarks: [{ name: '_TocBeta', childIndex: 0 }] } }
  const first = createParagraph('Alpha\t9'), second = createParagraph('Beta\t99')
  first.id = 'native-row-one'; second.id = 'native-row-two'
  for (const [index, block] of [first, second].entries()) block.metadata.source = { details: { styleId: 'TOC' + (index + 1),
    style: { tabStops: [{ alignment: 'right', leader: 'hyphen', positionTwips: 7200 }] } } }
  setDocumentTextFields(first, [
    { id: 'native-toc', field: 'TOC', instruction, cachedResult: 'Alpha\t9\nBeta\t99', startOffset: 0, endBlockId: second.id, endOffset: 7, containsNestedFields: true,
      metadata: { source: { format: 'docx', parentFieldId: null, instructionXml: '<w:r><w:instrText>TOC</w:instrText></w:r>' } } },
    { id: 'page-one', field: 'PAGEREF', instruction: 'PAGEREF "_TocAlpha" \\h', cachedResult: '9', startOffset: 6, endOffset: 7,
      metadata: { source: { parentFieldId: 'native-toc' } } },
  ])
  setDocumentTextFields(second, [{ id: 'page-two', field: 'PAGEREF', instruction: 'PAGEREF _TocBeta \\h', cachedResult: '99', startOffset: 5, endOffset: 7,
    metadata: { source: { parentFieldId: 'native-toc' } } }])
  model.blocks = [first, second, one, two]; model.modified = false
  return { model, first, second, one, two, selection: collapsedSelection({ blockId: first.id, offset: 6 }),
    context: { documentId: model.id, userId: 'local-user' } }
}

function withoutContext(model) {
  const copy = structuredClone(model)
  for (const block of copy.blocks) delete block.metadata.documentToc
  return copy
}

test('native TOC hydration derives stable context and bookmark navigation without editing text, fields or source', () => {
  const { model, first, second, selection } = nativeFixture(), source = structuredClone(model)
  const report = hydrateDocumentTocsFromNativeFields(model)
  assert.deepEqual(report, { recognized: [{ fieldId: 'native-toc', tocId: 'native-toc', blockIds: [first.id, second.id] }], retained: [], unmapped: [] })
  assert.deepEqual(withoutContext(model), source)
  assert.equal(model.modified, false)
  assert.equal(documentTocIdForSelection(model, selection), 'native-toc')
  assert.equal(documentTocHeadingTarget(model, first), 'one'); assert.equal(documentTocHeadingTarget(model, second), 'two')
  assert.deepEqual(documentTocOptionsAtBlock(first), { ...DEFAULT_DOCUMENT_TOC_OPTIONS, tabLeader: 'dashes' })
  const metadata = first.metadata, fields = documentTextFields(first)
  assert.deepEqual(hydrateDocumentTocsFromNativeFields(model), report)
  assert.equal(first.metadata, metadata); assert.equal(documentTextFields(first), fields)
})

test('native page-only and full updates keep original field IDs and use current heading IDs for layout pages', () => {
  const { model, first, second, one, selection, context } = nativeFixture()
  hydrateDocumentTocsFromNativeFields(model)
  const source = structuredClone(model), page = prepareDocumentTocAction(model, selection, 'native-toc', 'pageNumbersOnly', { one: 2, two: 3 }, context, 0)
  assert.deepEqual(model, source)
  applyDocumentTocMutation(model, page)
  assert.equal(runText(first.runs), 'Alpha\t2'); assert.equal(runText(second.runs), 'Beta\t3')
  assert.deepEqual(documentTextFieldRanges(model).map(range => range.field.id), ['native-toc', 'page-one', 'page-two'])
  assert.equal(documentTextFields(first)[0].cachedResult, 'Alpha\t2\nBeta\t3')
  one.runs[0].text = 'Revised Alpha'
  const full = prepareDocumentTocAction(model, page.selection, 'native-toc', 'all', { one: 4, two: 5 }, context, 0)
  applyDocumentTocMutation(model, full)
  assert.equal(runText(first.runs), 'Revised Alpha\t4'); assert.equal(runText(second.runs), 'Beta\t5')
  assert.deepEqual(documentTextFieldRanges(model).map(range => range.field.id), ['native-toc', 'page-one', 'page-two'])
  assert.equal(documentTextFields(first)[0].instruction, source.blocks[0].metadata.documentTextFields.fields[0].instruction)
})

test('native outline ranges starting above one rebuild only the specified heading levels', () => {
  const { model, first, second, one, two, selection, context } = nativeFixture('TOC \\o "2-4" \\h \\z \\u')
  one.level = 2; two.level = 4
  const omitted = createParagraph('Top-level title'); omitted.id = 'omitted'; omitted.type = 'heading'; omitted.level = 1; model.blocks.push(omitted)
  hydrateDocumentTocsFromNativeFields(model)
  assert.deepEqual(documentTocOptionsAtBlock(first), { ...DEFAULT_DOCUMENT_TOC_OPTIONS, levels: 4, minimumLevel: 2, tabLeader: 'dashes' })
  assert.equal(buildDocumentTocInstruction({ ...DEFAULT_DOCUMENT_TOC_OPTIONS, levels: 4, minimumLevel: 2 }), 'TOC \\o "2-4" \\h \\z \\u')
  applyDocumentTocMutation(model, prepareDocumentTocAction(model, selection, 'native-toc', 'all', { one: 2, two: 3 }, context, 0))
  assert.deepEqual(model.blocks.filter(block => block.metadata.documentToc).map(block => block.metadata.documentToc.headingId), ['one', 'two'])
  assert.equal(first.id, 'native-row-one'); assert.equal(second.id, 'native-row-two')
})

test('native custom sources keep page-only and deletion available while full rebuilding reports unsupported semantics', () => {
  const { model, first, selection, context } = nativeFixture('TOC \\t "Custom,1" \\h')
  const report = hydrateDocumentTocsFromNativeFields(model), source = structuredClone(model)
  assert.ok(report.recognized[0].fullUpdateUnavailableReason)
  assert.equal(documentTocActionState(model, selection, 'native-toc', 'pageNumbersOnly', { one: 2, two: 3 }, context).available, true)
  const unavailable = documentTocActionState(model, selection, 'native-toc', 'all', { one: 2, two: 3 }, context)
  assert.equal(unavailable.available, false); assert.equal(unavailable.reason, report.recognized[0].fullUpdateUnavailableReason)
  assert.throws(() => prepareDocumentTocAction(model, selection, 'native-toc', 'all', { one: 2, two: 3 }, context), error => error.message === unavailable.reason)
  assert.deepEqual(model, source)
  assert.equal(documentTocActionState(model, selection, 'native-toc', 'delete', {}, context).available, true)
  applyDocumentTocMutation(model, prepareDocumentTocAction(model, selection, 'native-toc', 'delete', {}, context, 0))
  assert.deepEqual(model.blocks.map(block => block.id), ['one', 'two']); assert.equal(documentTextFieldRanges(model).length, 0)
  assert.equal(first.metadata.documentTextFields.fields[0].instruction, source.blocks[0].metadata.documentTextFields.fields[0].instruction)
})

test('partial native TOC results and nontext ranges remain unchanged and explicitly unmapped', () => {
  for (const form of ['partial', 'pageBreak']) {
    const { model, first } = nativeFixture()
    if (form === 'partial') documentTextFields(first)[0].startOffset = 1
    else model.blocks.splice(1, 0, { ...createParagraph(), id: 'break', type: 'pageBreak', runs: [] })
    const source = structuredClone(model), report = hydrateDocumentTocsFromNativeFields(model)
    assert.equal(report.recognized.length, 0); assert.equal(report.unmapped.length, 1)
    assert.ok(report.unmapped[0].reason); assert.deepEqual(model, source)
  }
})

test('matching display text cannot replace a missing, duplicate or non-start bookmark target', () => {
  for (const form of ['missing', 'duplicate', 'insideHeading']) {
    const { model, first, one, two, selection, context } = nativeFixture()
    if (form === 'missing') one.metadata.source.details.bookmarks = []
    else if (form === 'duplicate') two.metadata.source.details.bookmarks.push({ name: '_TocAlpha', childIndex: 0 })
    else one.metadata.source.details.bookmarks[0].childIndex = 1
    const source = structuredClone(model), report = hydrateDocumentTocsFromNativeFields(model)
    assert.equal(first.metadata.documentToc.headingId, undefined); assert.equal(documentTocHeadingTarget(model, first), undefined)
    assert.ok(report.recognized[0].pageUpdateUnavailableReason)
    assert.equal(documentTocActionState(model, selection, 'native-toc', 'pageNumbersOnly', { one: 2, two: 3 }, context).available, false)
    assert.equal(documentTocActionState(model, selection, 'native-toc', 'all', { one: 2, two: 3 }, context).available, false)
    assert.equal(documentTocActionState(model, selection, 'native-toc', 'delete', {}, context).available, true)
    assert.deepEqual(withoutContext(model), source)
  }
})

test('native page-only updates skip an unresolvable locked page reference and retain that field result', () => {
  const { model, first, second, selection, context } = nativeFixture()
  const page = documentTextFields(second)[0]; page.locked = true; page.instruction = 'PAGEREF missing \\p'
  const report = hydrateDocumentTocsFromNativeFields(model)
  assert.equal(report.recognized[0].pageUpdateUnavailableReason, undefined)
  const locked = structuredClone(page), plan = prepareDocumentTocAction(model, selection, 'native-toc', 'pageNumbersOnly', { one: 2 }, context, 0)
  applyDocumentTocMutation(model, plan)
  assert.equal(runText(first.runs), 'Alpha\t2'); assert.equal(runText(second.runs), 'Beta\t99')
  assert.deepEqual(documentTextFields(second)[0], locked)
  assert.equal(documentTocActionState(model, plan.selection, 'native-toc', 'all', { one: 2, two: 3 }, context).available, false)
  assert.equal(documentTocActionState(model, plan.selection, 'native-toc', 'delete', {}, context).available, false)
})

test('changed native instruction invalidates derived update options until the hook reparses current source', () => {
  const { model, first, selection, context } = nativeFixture()
  hydrateDocumentTocsFromNativeFields(model)
  documentTextFields(first)[0].instruction = 'TOC \\o "1-2" \\h'
  const source = structuredClone(model)
  for (const action of ['all', 'pageNumbersOnly']) {
    const state = documentTocActionState(model, selection, 'native-toc', action, { one: 2, two: 3 }, context)
    assert.equal(state.available, false); assert.match(state.reason, /指令已改变/)
  }
  assert.deepEqual(model, source)
  assert.equal(documentTocActionState(model, selection, 'native-toc', 'delete', {}, context).available, true)
  hydrateDocumentTocsFromNativeFields(model)
  assert.equal(documentTocOptionsAtBlock(first).levels, 2)
  assert.equal(documentTocActionState(model, selection, 'native-toc', 'all', { one: 2, two: 3 }, context).available, true)
})

test('instruction-only TOCs and invalid field boundaries do not publish derived context', () => {
  const instructionOnly = nativeFixture()
  documentTextFields(instructionOnly.first)[0].metadata.source.nestedInInstruction = true
  const source = structuredClone(instructionOnly.model), report = hydrateDocumentTocsFromNativeFields(instructionOnly.model)
  assert.equal(report.recognized.length, 0); assert.match(report.unmapped[0].reason, /字段指令中/)
  assert.deepEqual(instructionOnly.model, source)
  const invalid = nativeFixture(); invalid.first.runs[0].text = '😀Alpha\t9'
  documentTextFields(invalid.first)[1].startOffset = 1
  const malformed = structuredClone(invalid.model)
  assert.throws(() => hydrateDocumentTocsFromNativeFields(invalid.model), /UTF-16/)
  assert.deepEqual(invalid.model, malformed)
})

test('native relative nested-field provenance remains preserved and blocks update while whole-range deletion stays available', () => {
  const { model, first, selection, context } = nativeFixture()
  const outer = documentTextFields(first)[0]
  outer.metadata.source.nestedFields = [{ field: 'PAGEREF', offset: 6, cachedResult: '9' }]
  const source = structuredClone(model)
  hydrateDocumentTocsFromNativeFields(model)
  for (const action of ['all', 'pageNumbersOnly']) {
    const state = documentTocActionState(model, selection, 'native-toc', action, { one: 2, two: 3 }, context)
    assert.equal(state.available, false); assert.match(state.reason, /嵌套字段/)
  }
  assert.equal(documentTocActionState(model, selection, 'native-toc', 'delete', {}, context).available, true)
  assert.deepEqual(withoutContext(model), source)
})

test('native hydration retains existing app-created TOC metadata and does not manufacture hyperlink behavior', () => {
  const { model, first, selection } = nativeFixture('TOC \\o "1-3"')
  hydrateDocumentTocsFromNativeFields(model)
  assert.equal(documentTocIdForSelection(model, selection), 'native-toc'); assert.equal(documentTocHeadingTarget(model, first), undefined)
  const generated = buildDocumentTocBlocks(model, { ...DEFAULT_DOCUMENT_TOC_OPTIONS, format: 'modern' }, { one: 2, two: 3 }, 'app-toc', [], 'Contents')
  model.blocks = [...generated, ...model.blocks.filter(block => block.type === 'heading')]
  const source = structuredClone(model), report = hydrateDocumentTocsFromNativeFields(model)
  assert.equal(report.recognized.length, 0); assert.equal(report.retained.length, 1); assert.equal(report.retained[0].tocId, 'app-toc')
  assert.deepEqual(model, source)
})
