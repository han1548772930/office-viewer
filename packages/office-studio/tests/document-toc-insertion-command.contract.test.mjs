import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import { cloneDocument, collapsedSelection, createParagraph, createTable, runText } from '../src/features/document/model/document-model.ts'
import { executeDocumentCommand } from '../src/features/document/commands/document-commands.ts'
import { documentTocInsertionState, prepareDocumentTocInsertion } from '../src/features/document/toc/document-toc-command.ts'
import { documentTocIdAtBlock } from '../src/features/document/toc/document-toc.ts'
import { documentTextFieldRanges, setDocumentTextFields } from '../src/features/document/model/document-text-fields.ts'
import { documentContentControlCommandIsTextInput } from '../src/features/document/services/content-control-editing.ts'

// Unexecuted insertion/ACL/field contracts. Runtime and Office validation remain deferred.
function fixture() {
  const model = createBlankDocument()
  model.id = 'toc-insertion-contract'
  const body = { ...createParagraph('BeforeAfter'), id: 'body' }
  const first = { ...createParagraph('First heading'), id: 'heading-first', type: 'heading', level: 1 }
  const second = { ...createParagraph('Second heading'), id: 'heading-second', type: 'heading', level: 2 }
  model.blocks = [body, first, second]
  return { model, selection: collapsedSelection({ blockId: 'body', offset: 6 }), context: { documentId: model.id, userId: 'editor' } }
}
const command = (patch = {}) => ({ id: 'doc.command.insert-table-of-contents',
  pageNumbers: { 'heading-first': 4, 'heading-second': 19 }, ...patch })
const tocBlocks = model => model.blocks.filter(block => documentTocIdAtBlock(block))

test('the fixed insertion command creates editable TOC/PAGEREF fields, splits the real caret and preserves heading identities', () => {
  const { model, selection, context } = fixture(), source = cloneDocument(model)
  const next = executeDocumentCommand(model, selection, command(), context), blocks = tocBlocks(model)
  assert.equal(documentContentControlCommandIsTextInput(command()), true)
  assert.equal(blocks.length, 2)
  assert.deepEqual(blocks.map(block => block.metadata.documentToc.role), ['entry', 'entry'])
  assert.deepEqual(blocks.map(block => runText(block.runs)), ['First heading\t4', 'Second heading\t19'])
  assert.equal(runText(model.blocks[0].runs), 'Before')
  assert.equal(runText(model.blocks.find(block => block.id === next.focus.blockId).runs), 'After')
  assert.deepEqual(model.blocks.filter(block => block.type === 'heading'), source.blocks.filter(block => block.type === 'heading'))
  const ranges = documentTextFieldRanges(model), outer = ranges.find(range => range.field.field === 'TOC')
  const pages = ranges.filter(range => range.field.field === 'PAGEREF')
  assert.equal(outer.field.endBlockId, blocks[1].id)
  assert.equal(outer.field.containsNestedFields, true)
  assert.deepEqual(pages.map(range => range.field.cachedResult), ['4', '19'])
  assert.ok(pages.every(range => range.field.metadata.source.parentFieldId === outer.field.id))
  assert.ok(blocks.every(block => block.runs.some(run => run.href === '#' + block.metadata.documentToc.headingId)))
})

test('automatic insertion adds only its explicit trimmed title while custom insertion omits the title', () => {
  for (const [title, count] of [[undefined, 0], ['', 0], ['   ', 0], ['  目录  ', 1]]) {
    const { model, selection, context } = fixture()
    executeDocumentCommand(model, selection, command({ title }), context)
    const titles = tocBlocks(model).filter(block => block.metadata.documentToc.role === 'title')
    assert.equal(titles.length, count)
    if (count) assert.equal(runText(titles[0].runs), '目录')
  }
})

test('custom depth can include headings above the automatic level-three range', () => {
  const { model, selection, context } = fixture()
  model.blocks[1].level = 4; model.blocks[2].level = 5
  const before = cloneDocument(model)
  assert.throws(() => executeDocumentCommand(model, selection, command(), context), /没有.*标题/)
  assert.deepEqual(model, before)
  executeDocumentCommand(model, selection, command({ levels: 4 }), context)
  assert.deepEqual(tocBlocks(model).map(block => block.metadata.documentToc.headingId), ['heading-first'])
})

test.each([[0, 1], [1.9, 1], [50, 2]])('finite heading depth %s follows the fixed clamped/truncated contract', (levels, entries) => {
  const { model, selection, context } = fixture()
  executeDocumentCommand(model, selection, command({ levels }), context)
  assert.equal(tocBlocks(model).length, entries)
})

test('missing layout labels retain unavailable page caches and never read a heading ID through Object.prototype', () => {
  const { model, selection, context } = fixture()
  model.blocks[1].id = 'toString'
  executeDocumentCommand(model, selection, command({ pageNumbers: {} }), context)
  assert.deepEqual(tocBlocks(model).map(block => runText(block.runs)), ['First heading', 'Second heading'])
  assert.equal(documentTextFieldRanges(model).filter(range => range.field.field === 'PAGEREF').length, 0)
})

test.each([
  { levels: Infinity }, { levels: '3' }, { showPageNumbers: 'true' }, { rightAlignPageNumbers: 1 },
  { tabLeader: 'hyphen' }, { format: 'unsupported' }, { title: 5 }, { pageNumbers: null },
  { pageNumbers: [] }, { pageNumbers: { 'heading-first': 0 } }, { pageNumbers: { 'heading-first': NaN } },
])('malformed insertion options %j reject before changing document data', patch => {
  const { model, selection, context } = fixture(), before = cloneDocument(model)
  assert.throws(() => executeDocumentCommand(model, selection, command(patch), context))
  assert.deepEqual(model, before)
})

test.each(['modern', 'range', 'table', 'nested', 'existing-toc'])('insertion rejects the %s context using the same capability and command checks', kind => {
  const { model, context } = fixture()
  let selection = collapsedSelection({ blockId: 'body', offset: 2 })
  if (kind === 'modern') model.metadata = { documentFlavor: 2, documentFlavorContract: 'univer' }
  else if (kind === 'range') selection = { anchor: { blockId: 'body', offset: 0 }, focus: { blockId: 'body', offset: 2 } }
  else if (kind === 'table') {
    const table = createTable(1, 1), cell = table.rows[0].cells[0].blocks[0]
    model.blocks.unshift({ ...createParagraph(), id: 'table', type: 'table', runs: [], metadata: { table } })
    selection = collapsedSelection({ blockId: 'table', row: 0, col: 0, cellBlockId: cell.id, offset: 0 })
  } else if (kind === 'nested') {
    const child = { ...createParagraph('Nested'), id: 'nested' }
    model.blocks.unshift({ ...createParagraph(), type: 'quote', children: [child] })
    selection = collapsedSelection({ blockId: child.id, offset: 0 })
  } else model.blocks[0].metadata.documentToc = { id: 'existing', role: 'entry', headingId: 'heading-first' }
  const before = cloneDocument(model)
  assert.equal(documentTocInsertionState(model, selection, context).available, false)
  assert.throws(() => executeDocumentCommand(model, selection, command(), context))
  assert.deepEqual(model, before)
})

test.each(['edit', 'view'])('insertion respects %s permissions on the target or copied source heading', action => {
  const { model, selection, context } = fixture()
  model.permissionRules = [{ id: 'deny', scope: 'paragraph', targetId: action === 'view' ? 'heading-first' : 'body',
    action, effect: 'deny', userIds: [] }]
  const before = cloneDocument(model)
  assert.throws(() => executeDocumentCommand(model, selection, command(), context))
  assert.deepEqual(model, before)
})

test('a locked field at the insertion point and an auxiliary host are rejected before structural edits', () => {
  const { model, selection, context } = fixture()
  setDocumentTextFields(model.blocks[0], [{ id: 'locked-date', field: 'DATE', instruction: 'DATE', startOffset: 3, endOffset: 8, locked: true }])
  const before = cloneDocument(model)
  assert.equal(documentTocInsertionState(model, selection, context).available, false)
  assert.throws(() => executeDocumentCommand(model, selection, command(), context), /锁定字段/)
  assert.deepEqual(model, before)
  assert.throws(() => executeDocumentCommand(model, selection, { ...command(), host: { kind: 'story', sectionId: 'first', story: 'header' } }, context))
  assert.deepEqual(model, before)
})

test('a detached insertion plan preserves the source and an empty heading set rejects instead of creating a no-op entry', () => {
  const { model, selection, context } = fixture(), before = cloneDocument(model)
  const blocks = prepareDocumentTocInsertion(model, selection, command(), context)
  blocks[0].runs[0].text = 'Changed detached row'
  assert.deepEqual(model, before)
  model.blocks = [model.blocks[0]]
  const empty = cloneDocument(model)
  assert.throws(() => executeDocumentCommand(model, selection, command(), context), /没有.*标题/)
  assert.deepEqual(model, empty)
})
