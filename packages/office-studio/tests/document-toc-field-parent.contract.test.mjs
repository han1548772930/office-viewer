import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import { collapsedSelection, createParagraph, runText } from '../src/features/document/model/document-model.ts'
import { documentTextFields, documentTextFieldRanges, setDocumentTextFields } from '../src/features/document/model/document-text-fields.ts'
import { applyDocumentTocMutation, applyDocumentTocPageNumberUpdate, buildDocumentTocBlocks, DEFAULT_DOCUMENT_TOC_OPTIONS,
  prepareDocumentTocDeletion, prepareDocumentTocFullUpdate, prepareDocumentTocPageNumberUpdate } from '../src/features/document/toc/document-toc.ts'
import { reconcileDocumentTocFieldParents } from '../src/features/document/toc/document-toc-edit.ts'

// Deferred TOC/Office source-parent contracts, 2026-11-24. Writer and test execution remain pending.
function fixture(pages = { one: 9, two: 99 }) {
  const model = createBlankDocument(); model.id = 'toc-field-parent'
  const one = createParagraph('Alpha'), two = createParagraph('Beta')
  one.id = 'one'; one.type = 'heading'; one.level = 1
  two.id = 'two'; two.type = 'heading'; two.level = 2
  model.blocks = [one, two]
  const entries = buildDocumentTocBlocks(model, DEFAULT_DOCUMENT_TOC_OPTIONS, pages, 'toc-parent', [], 'Contents', false)
  model.blocks.unshift(...entries)
  const outer = documentTextFields(entries[0]).find(field => field.field === 'TOC')
  outer.metadata = { source: { parentFieldId: null, retainedXml: '<w:fldChar/>' } }
  return { model, entries, outer, selection: collapsedSelection({ blockId: entries[0].id, offset: 0 }) }
}

test('full TOC rebuilding moves a surviving PAGEREF out of a removed intermediate native field and into its actual outer TOC', () => {
  const { model, entries, outer, selection } = fixture(), first = entries[0]
  const page = documentTextFields(first).find(field => field.field === 'PAGEREF')
  page.metadata.source = { ...page.metadata.source, parentFieldId: 'old-link', retained: { origin: 'native-page' } }
  setDocumentTextFields(first, [...documentTextFields(first), { id: 'old-link', field: 'HYPERLINK', instruction: 'HYPERLINK \\l one',
    cachedResult: 'ha\t9', startOffset: 3, endOffset: 7, containsNestedFields: true, metadata: { source: { parentFieldId: outer.id } } }])
  model.blocks.find(block => block.id === 'one').runs[0].text = 'Revised Alpha'
  const before = structuredClone(model), plan = prepareDocumentTocFullUpdate(model, 'toc-parent', { one: 12, two: 13 }, selection, 0)
  assert.ok(plan?.changed); assert.deepEqual(model, before)
  applyDocumentTocMutation(model, plan)
  const fields = documentTextFieldRanges(model), updated = fields.find(range => range.field.id === page.id)
  assert.equal(fields.some(range => range.field.id === 'old-link'), false)
  assert.equal(updated.field.metadata.source.parentFieldId, outer.id)
  assert.equal(updated.field.metadata.source.nestedInInstruction, false)
  assert.deepEqual(updated.field.metadata.source.retained, { origin: 'native-page' })
  assert.equal(fields.find(range => range.field.id === outer.id).field.metadata.source.parentFieldId, null)
  assert.equal(runText(updated.block.runs), 'Revised Alpha\t12')
})

test('full TOC rebuilding restores an outside owner before binding generated page results to its original field identity', () => {
  const { model, entries, outer, selection } = fixture(), prefix = createParagraph('Prefix'); prefix.id = 'prefix'
  const last = entries.at(-1), pageIds = documentTextFieldRanges(model).filter(range => range.field.field === 'PAGEREF').map(range => range.field.id)
  setDocumentTextFields(entries[0], documentTextFields(entries[0]).filter(field => field.id !== outer.id))
  setDocumentTextFields(prefix, [{ ...structuredClone(outer), startOffset: 0, endBlockId: last.id, endOffset: runText(last.runs).length }])
  model.blocks.unshift(prefix)
  model.blocks.find(block => block.id === 'one').runs[0].text = 'Revised Alpha'
  const before = structuredClone(model), plan = prepareDocumentTocFullUpdate(model, 'toc-parent', { one: 12, two: 13 }, selection, 0)
  assert.ok(plan?.changed); assert.deepEqual(model, before)
  applyDocumentTocMutation(model, plan)
  const fields = documentTextFieldRanges(model), outers = fields.filter(range => range.field.field === 'TOC')
  const pages = fields.filter(range => range.field.field === 'PAGEREF')
  assert.deepEqual(outers.map(range => range.field.id), [outer.id])
  assert.equal(outers[0].block.id, prefix.id)
  assert.equal(outers[0].field.endBlockId, last.id)
  assert.equal(outers[0].field.endOffset, runText(model.blocks.find(block => block.id === last.id).runs).length)
  assert.equal(outers[0].field.metadata.source.parentFieldId, null)
  assert.equal(outers[0].field.metadata.source.retainedXml, '<w:fldChar/>')
  assert.equal(outers[0].field.containsNestedFields, true)
  assert.equal(outers[0].field.cachedResult, 'Prefix\nRevised Alpha\t12\nBeta\t13')
  assert.deepEqual(pages.map(range => range.field.id), pageIds)
  for (const page of pages) {
    assert.equal(page.field.metadata.source.parentFieldId, outer.id)
    assert.equal(page.field.metadata.source.nestedInInstruction, false)
  }
  assert.equal(runText(model.blocks[0].runs), 'Prefix')
})

test('page-only creation installs each new page result under the existing TOC and marks its real nested results', () => {
  const { model, outer, selection } = fixture({}), before = structuredClone(model)
  assert.equal(documentTextFieldRanges(model).filter(range => range.field.field === 'PAGEREF').length, 0)
  const plan = prepareDocumentTocPageNumberUpdate(model, 'toc-parent', { one: 12, two: 13 }, selection, 0)
  assert.ok(plan?.changed); assert.deepEqual(model, before)
  applyDocumentTocPageNumberUpdate(model, plan)
  const fields = documentTextFieldRanges(model), pages = fields.filter(range => range.field.field === 'PAGEREF')
  assert.equal(pages.length, 2)
  for (const page of pages) {
    assert.equal(page.field.metadata.source.parentFieldId, outer.id)
    assert.equal(page.field.metadata.source.nestedInInstruction, false)
    assert.equal(page.field.endOffset - page.field.startOffset, 2)
  }
  assert.equal(fields.find(range => range.field.id === outer.id).field.containsNestedFields, true)
})

test('TOC deletion removes its descendant fields and keeps the explicit root and surrounding surviving children', () => {
  const { model, entries, outer, selection } = fixture(), prefix = createParagraph('Prefix'); prefix.id = 'prefix'
  setDocumentTextFields(prefix, [
    { id: 'root', field: 'UNKNOWN', startOffset: 0, endBlockId: 'two', endOffset: 4, cachedResult: 'retained', metadata: { source: { parentFieldId: null } } },
    { id: 'prefix-child', field: 'UNKNOWN', startOffset: 0, endOffset: 6, metadata: { source: { parentFieldId: 'root' } } },
  ])
  outer.metadata.source.parentFieldId = 'root'; model.blocks.unshift(prefix)
  const before = structuredClone(model), plan = prepareDocumentTocDeletion(model, 'toc-parent', selection, 0)
  assert.ok(plan?.changed); assert.deepEqual(model, before)
  applyDocumentTocMutation(model, plan)
  const fields = documentTextFieldRanges(model)
  assert.deepEqual(fields.map(range => range.field.id), ['root', 'prefix-child'])
  assert.equal(fields[0].field.metadata.source.parentFieldId, null)
  assert.equal(fields[1].field.metadata.source.parentFieldId, 'root')
  assert.ok(entries.every(entry => !model.blocks.some(block => block.id === entry.id)))
})

function graphFixture() {
  const model = createBlankDocument(), paragraph = createParagraph('Text'); paragraph.id = 'graph'
  setDocumentTextFields(paragraph, [
    { id: 'root', field: 'UNKNOWN', startOffset: 0, endOffset: 4, metadata: { source: { parentFieldId: null, retained: true } } },
    { id: 'wrapper', field: 'UNKNOWN', startOffset: 0, endOffset: 4, metadata: { source: { parentFieldId: 'root' } } },
    { id: 'child', field: 'UNKNOWN', startOffset: 0, endOffset: 0, metadata: { source: { parentFieldId: 'wrapper', retained: 'empty' } } },
    { id: 'explicit-root', field: 'UNKNOWN', startOffset: 0, endOffset: 0, metadata: { source: { parentFieldId: null } } },
    { id: 'no-provenance', field: 'UNKNOWN', startOffset: 0, endOffset: 0, metadata: { retained: 'unknown' } },
  ])
  model.blocks = [paragraph]
  return model
}

test('empty and equal-range native descendants use a surviving explicit ancestor, or null after its removal, without inferring unknown provenance', () => {
  const before = graphFixture()
  for (const removeRoot of [false, true]) {
    const candidate = structuredClone(before), paragraph = candidate.blocks[0]
    setDocumentTextFields(paragraph, documentTextFields(paragraph).filter(field => field.id !== 'wrapper' && (!removeRoot || field.id !== 'root')))
    reconcileDocumentTocFieldParents(before, candidate)
    const fields = documentTextFields(paragraph), child = fields.find(field => field.id === 'child')
    assert.equal(child.metadata.source.parentFieldId, removeRoot ? null : 'root')
    assert.equal(child.metadata.source.retained, 'empty'); assert.deepEqual([child.startOffset, child.endOffset], [0, 0])
    assert.equal(fields.find(field => field.id === 'explicit-root').metadata.source.parentFieldId, null)
    assert.equal(Object.hasOwn(fields.find(field => field.id === 'no-provenance').metadata, 'source'), false)
    assert.equal(runText(paragraph.runs), 'Text')
  }
  assert.equal(documentTextFields(before.blocks[0]).find(field => field.id === 'child').metadata.source.parentFieldId, 'wrapper')
})

test('a retained ancestor that no longer contains a descendant is removed from that descendant provenance', () => {
  const before = graphFixture(), candidate = structuredClone(before), paragraph = candidate.blocks[0]
  setDocumentTextFields(paragraph, documentTextFields(paragraph).filter(field => field.id !== 'wrapper').map(field => field.id === 'root' ? { ...field, startOffset: 1 } : field))
  reconcileDocumentTocFieldParents(before, candidate)
  assert.equal(documentTextFields(paragraph).find(field => field.id === 'child').metadata.source.parentFieldId, null)
})

test('cycles between same-position empty native fields reject the candidate before source metadata is changed', () => {
  const model = createBlankDocument(), paragraph = createParagraph(); paragraph.id = 'empty-graph'
  setDocumentTextFields(paragraph, [
    { id: 'a', field: 'UNKNOWN', startOffset: 0, endOffset: 0, metadata: { source: { parentFieldId: 'b' } } },
    { id: 'b', field: 'UNKNOWN', startOffset: 0, endOffset: 0, metadata: { source: { parentFieldId: 'a' } } },
  ])
  model.blocks = [paragraph]
  const candidate = structuredClone(model), preserved = structuredClone(candidate)
  assert.throws(() => reconcileDocumentTocFieldParents(model, candidate), /循环/)
  assert.deepEqual(candidate, preserved)
})
