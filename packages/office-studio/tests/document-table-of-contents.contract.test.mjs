import assert from 'node:assert/strict'
import { test } from 'vitest'
import {
  buildDocumentTocBlocks,
  collectDocumentTocHeadings,
  deleteDocumentToc,
  documentTocHeadingTarget,
  documentTocIdForSelection,
  documentTocOptionsAtBlock,
  documentTocPageNumbers,
  supportsDocumentTableOfContents,
  updateDocumentToc,
} from '../src/features/document/toc/document-toc.ts'
import { executeDocumentCommand } from '../src/features/document/commands/document-commands.ts'
import { createOfficeMessages, supportedOfficeLocales } from '../src/i18n/office-i18n.ts'
import { createDocumentRibbon, documentTocContextMenuItems } from '../src/components/document/document-ribbon.ts'
import { collapsedSelection, createRun, normalizeDocument, runText } from '../src/features/document/model/document-model.ts'
import { saveDocumentSnapshot } from '../src/features/document/services/document-facade.ts'
import { visibleRibbonTabs } from '../src/platform/ribbon/ribbon-schema.ts'

const options = {
  levels: 3,
  showPageNumbers: true,
  rightAlignPageNumbers: true,
  tabLeader: 'dots',
  format: 'fromTemplate',
}

function block(id, text, type = 'heading', level = 1) {
  return { id, type, level, runs: [{ id: `${id}-run`, text }], style: { align: 'left', indent: 0, lineHeight: 1.5, spaceBefore: 0, spaceAfter: 8 }, metadata: {} }
}

function model(blocks) {
  return { kind: 'document', id: 'doc-1', name: 'Document', blocks, page: {}, modified: false, wordCount: 0 }
}

test('collects eligible heading paragraphs without including body text or blank headings', () => {
  const document = model([
    block('body', 'Body', 'paragraph'),
    block('one', ' Chapter 1 '),
    block('three', 'Section', 'heading', 3),
    block('four', 'Excluded', 'heading', 4),
    block('blank', '  '),
  ])

  assert.deepEqual(collectDocumentTocHeadings(document, 3), [
    { id: 'one', text: 'Chapter 1', level: 1 },
    { id: 'three', text: 'Section', level: 3 },
  ])
})

test('collects numbered heading labels and honors configured outline depth', () => {
  const numbered = block('numbered', 'Overview')
  numbered.metadata.listLabel = '2.4.'
  const document = model([
    numbered,
    block('deep', 'Appendix', 'heading', 4),
  ])

  assert.deepEqual(collectDocumentTocHeadings(document, 3), [
    { id: 'numbered', text: '2.4. Overview', level: 1 },
  ])
  assert.deepEqual(collectDocumentTocHeadings(document, 4).map(({ id }) => id), ['numbered', 'deep'])
})

test('normalizes directory levels and persisted options without accepting invalid source heading levels', () => {
  const document = model([
    block('one', 'Chapter 1'),
    block('two', 'Chapter 2', 'heading', 2),
    block('two', 'Duplicate identity', 'heading', 2),
    block('invalid', 'Invalid', 'heading', Number.NaN),
    block('fraction', 'Fraction', 'heading', 1.5),
  ])
  assert.deepEqual(collectDocumentTocHeadings(document, Number.NaN).map(({ id }) => id), ['one', 'two'])
  const toc = buildDocumentTocBlocks(document, { ...options, levels: 1.9, showPageNumbers: false })
  assert.deepEqual(toc.filter((entry) => entry.metadata.documentToc?.role === 'entry').map((entry) => entry.metadata.documentToc.headingId), ['one'])
  assert.deepEqual(documentTocOptionsAtBlock(toc[0]), { ...options, levels: 1, showPageNumbers: false, rightAlignPageNumbers: false })
})

test('provides translated TOC controls for every supported workbench locale', () => {
  const messages = createOfficeMessages()
  for (const locale of supportedOfficeLocales) {
    assert.ok(messages[locale]['document.toc.insert'], `${locale} insert label is missing`)
    assert.ok(messages[locale]['document.toc.contents-title'], `${locale} automatic title is missing`)
    assert.ok(messages[locale]['document.toc.update-pages'], `${locale} update label is missing`)
    assert.ok(messages[locale]['document.toc.preview-heading'], `${locale} preview label is missing`)
  }
})

test('builds editable paragraph entries with source IDs, page mapping, nesting and configuration metadata', () => {
  const document = model([block('one', 'Chapter 1'), block('two', 'Section 1.1', 'heading', 2)])
  const toc = buildDocumentTocBlocks(document, options, { one: 2, two: 4 }, 'toc-a', [], 'Contents')

  assert.equal(toc.length, 3)
  assert.equal(toc[0].runs[0].text, 'Contents')
  assert.equal(toc[1].runs[0].text, 'Chapter 1\t2')
  assert.equal(toc[2].runs[0].text, 'Section 1.1\t4')
  assert.equal(toc[2].style.indent, 18)
  assert.deepEqual(toc[2].metadata.documentToc, {
    id: 'toc-a', role: 'entry', headingId: 'two', headingText: 'Section 1.1', level: 2, pageNumber: 4, options,
  })
})

test('resolves TOC entry hyperlinks to existing heading blocks and ignores stale targets', () => {
  const document = model([block('heading-one', 'Chapter 1')])
  const [entry] = buildDocumentTocBlocks(document, options, {}, 'toc-a', [], 'Contents', false)

  assert.equal(documentTocHeadingTarget(document, entry), 'heading-one')
  entry.metadata.documentToc.headingId = 'removed-heading'
  assert.equal(documentTocHeadingTarget(document, entry), undefined)
  assert.equal(documentTocHeadingTarget(document, block('body', 'Body', 'paragraph')), undefined)
})

test('updates only page numbers without changing entry IDs or heading text', () => {
  const document = model([block('one', 'Chapter 1')])
  const toc = buildDocumentTocBlocks(document, options, { one: 1 }, 'toc-a')
  document.blocks.splice(0, 0, ...toc)
  const entryId = document.blocks[1].id

  assert.equal(updateDocumentToc(document, 'toc-a', 'pageNumbersOnly', { one: 6 }), true)
  const entry = document.blocks[1]
  assert.equal(entry.id, entryId)
  assert.equal(runText(entry.runs), 'Chapter 1\t6')
  assert.equal(entry.metadata.documentToc.headingText, 'Chapter 1')
  assert.equal(entry.metadata.documentToc.headingId, 'one')
  assert.equal(entry.metadata.documentToc.pageNumber, 6)

  const refreshedRuns = entry.runs
  assert.equal(updateDocumentToc(document, 'toc-a', 'pageNumbersOnly', { one: 6 }), true)
  assert.equal(entry.runs, refreshedRuns)
})

test('appends the first page number to unnumbered entries without truncating multiword labels', () => {
  for (const rightAlignPageNumbers of [true, false]) {
    const document = model([block('one', 'Chapter 1')])
    const [entry] = buildDocumentTocBlocks(document, { ...options, rightAlignPageNumbers }, {}, 'toc-a', [], 'Contents', false)
    document.blocks.unshift(entry)
    const separator = rightAlignPageNumbers ? '\t' : ' '

    assert.equal(updateDocumentToc(document, 'toc-a', 'pageNumbersOnly', { one: 6 }), true)
    assert.equal(runText(entry.runs), `Chapter 1${separator}6`)
    assert.equal(updateDocumentToc(document, 'toc-a', 'pageNumbersOnly', { one: 14 }), true)
    assert.equal(runText(entry.runs), `Chapter 1${separator}14`)
  }
})

test('rebuilds all entries in place and keeps matching paragraph identities', () => {
  const document = model([block('one', 'Chapter 1'), block('two', 'Old section', 'heading', 2)])
  const toc = buildDocumentTocBlocks(document, options, { one: 1, two: 2 }, 'toc-a')
  document.blocks.splice(1, 0, ...toc)
  const previousEntryId = document.blocks[2].id
  document.blocks[4].runs[0].text = 'New section'

  assert.equal(updateDocumentToc(document, 'toc-a', 'all', { one: 3, two: 5 }), true)
  const entries = document.blocks.filter((candidate) => candidate.metadata.documentToc?.role === 'entry')
  assert.deepEqual(entries.map((entry) => entry.runs[0].text), ['Chapter 1\t3', 'New section\t5'])
  assert.equal(entries[0].id, previousEntryId)
  assert.equal(document.blocks.find((candidate) => candidate.id === 'one')?.id, 'one')
})

test('full update removes deleted headings and preserves the cached directory when no headings remain', () => {
  const firstHeading = block('one', 'Chapter 1')
  const secondHeading = block('two', 'Chapter 2')
  const document = model([firstHeading, secondHeading])
  document.blocks.splice(1, 0, ...buildDocumentTocBlocks(document, options, { one: 1, two: 2 }, 'toc-a'))
  const originalEntryIds = document.blocks
    .filter((entry) => entry.metadata.documentToc?.role === 'entry')
    .map((entry) => entry.id)

  document.blocks = document.blocks.filter((entry) => entry.id !== 'one')
  assert.equal(updateDocumentToc(document, 'toc-a', 'all', { two: 4 }), true)
  assert.deepEqual(document.blocks.filter((entry) => entry.metadata.documentToc?.role === 'entry').map((entry) => entry.metadata.documentToc.headingId), ['two'])
  assert.ok(originalEntryIds.includes(document.blocks.find((entry) => entry.metadata.documentToc?.role === 'entry').id))

  document.blocks = document.blocks.filter((entry) => entry.id !== 'two')
  const before = structuredClone(document)
  assert.equal(updateDocumentToc(document, 'toc-a', 'all'), false)
  assert.deepEqual(document, before)
})

test('page-only refresh preserves edited labels and run styling, and does not partially mutate unresolved references', () => {
  const document = model([block('one', 'Chapter 1'), block('two', 'Chapter 2')])
  const toc = buildDocumentTocBlocks(document, options, { one: 1, two: 2 }, 'toc-a')
  document.blocks.unshift(...toc)
  toc[1].runs = [createRun('Edited label', { bold: true }), createRun('\t'), createRun('1', { italic: true, color: '#ff0000' })]
  const before = structuredClone(document)

  assert.equal(updateDocumentToc(document, 'toc-a', 'pageNumbersOnly', { one: 8 }), false)
  assert.deepEqual(document, before)
  assert.equal(updateDocumentToc(document, 'toc-a', 'pageNumbersOnly', { one: 8, two: 9 }), true)
  assert.equal(toc[1].runs.map((run) => run.text).join(''), 'Edited label\t8')
  assert.equal(toc[1].runs[0].bold, true)
  assert.equal(toc[1].runs.at(-1).italic, true)
  assert.equal(toc[1].runs.at(-1).color, '#ff0000')
})

test('snapshot save and normalized reopen preserve directory options, entry identities and live heading targets', () => {
  const document = model([block('one', 'Chapter 1')])
  const custom = { ...options, levels: 5, tabLeader: 'dashes', format: 'modern' }
  const toc = buildDocumentTocBlocks(document, custom, { one: 2 }, 'toc-saved', [], '目录')
  document.blocks.unshift(...toc)
  const reopened = normalizeDocument(saveDocumentSnapshot(document))
  const entry = reopened.blocks.find((candidate) => candidate.metadata.documentToc?.role === 'entry')

  assert.equal(entry.id, toc[1].id)
  assert.deepEqual(entry.metadata.documentToc.options, custom)
  assert.equal(documentTocHeadingTarget(reopened, entry), 'one')
  assert.equal(updateDocumentToc(reopened, 'toc-saved', 'pageNumbersOnly', { one: 3 }), true)
  assert.equal(entry.runs.map((run) => run.text).join(''), 'Chapter 1\t3')
  assert.equal(document.blocks.find((candidate) => candidate.id === entry.id).runs.map((run) => run.text).join(''), 'Chapter 1\t2')
})

test('document commands delete the full TOC as one history-model mutation and ignore unknown IDs', () => {
  const document = model([block('one', 'Chapter 1')])
  const toc = buildDocumentTocBlocks(document, options, {}, 'toc-a')
  document.blocks.splice(0, 0, ...toc)
  const selection = { anchor: { blockId: toc[1].id, offset: 0 }, focus: { blockId: toc[1].id, offset: 0 } }

  executeDocumentCommand(document, selection, { id: 'doc.command.delete-table-of-contents', tocId: 'toc-a' })
  assert.deepEqual(document.blocks.map((entry) => entry.id), ['one'])
  assert.equal(deleteDocumentToc(document, 'missing'), false)
})

test('document commands cannot mutate another TOC or bypass a paragraph-level deny rule', () => {
  const document = model([block('one', 'Chapter 1')])
  const firstToc = buildDocumentTocBlocks(document, options, {}, 'toc-a')
  const secondToc = buildDocumentTocBlocks(document, options, {}, 'toc-b')
  document.blocks.splice(0, 0, ...firstToc, ...secondToc)
  const selection = { anchor: { blockId: firstToc[1].id, offset: 0 }, focus: { blockId: firstToc[1].id, offset: 0 } }

  assert.throws(() => executeDocumentCommand(document, selection, {
    id: 'doc.command.delete-table-of-contents', tocId: 'toc-b',
  }), /selection inside the target table/)

  document.permissionRules = [{
    id: 'deny-toc-row', scope: 'paragraph', targetId: firstToc[1].id, action: 'edit', effect: 'deny', userIds: [],
  }]
  assert.throws(() => executeDocumentCommand(document, selection, {
    id: 'doc.command.delete-table-of-contents', tocId: 'toc-a',
  }), (error) => error.code === 'DOCUMENT_PERMISSION_DENIED')
})

test('Insert > CELL exposes automatic and custom entries only for eligible traditional-body selections', () => {
  const document = model([block('heading', 'Chapter')])
  const ribbon = (insertDisabled, available = true, selectedId) => createDocumentRibbon({
    state: { marks: {}, mixedMarks: [], paragraph: { align: 'left', indent: 0 }, hasSelection: false },
    canUndo: false, canRedo: false, outline: false, focus: false, fullscreen: false,
    paper: 'A4', orientation: 'portrait', zoom: 100, contentControlsAvailable: false,
    tableOfContents: { available, insertDisabled, selectedId, editDisabled: false },
  })
  const insertTab = ribbon(false).find((tab) => tab.id === 'insert')
  const insertGroup = insertTab.groups.find((group) => group.id === 'doc-insert-cell')
  const item = insertGroup.items.find((candidate) => candidate.id === 'doc-insert-table-of-contents')
  assert.equal(item.order, 7)
  assert.deepEqual(item.children.map((child) => child.params), [{ mode: 'automatic' }, { mode: 'custom' }])
  assert.equal(item.disabled, false)
  assert.equal(ribbon(true).find((tab) => tab.id === 'insert').groups.find((group) => group.id === 'doc-insert-cell').items[0].disabled, true)
  assert.equal(visibleRibbonTabs(ribbon(true, false)).find((tab) => tab.id === 'insert').groups.some((group) => group.id === 'doc-insert-cell'), false)
})

test('contextual TOC tab and Layout context entries dispatch update and delete with the selected ID', () => {
  const document = model([block('heading', 'Chapter')])
  const tocBlocks = buildDocumentTocBlocks(document, options, {}, 'toc-active')
  document.blocks.splice(0, 0, ...tocBlocks)
  const selection = collapsedSelection({ blockId: tocBlocks[1].id, offset: 0 })
  const tocId = documentTocIdForSelection(document, selection)
  assert.equal(tocId, 'toc-active')

  const tabs = createDocumentRibbon({
    state: { marks: {}, mixedMarks: [], paragraph: { align: 'left', indent: 0 }, hasSelection: false },
    canUndo: false, canRedo: false, outline: false, focus: false, fullscreen: false,
    paper: 'A4', orientation: 'portrait', zoom: 100, contentControlsAvailable: false,
    tableOfContents: { available: true, insertDisabled: true, selectedId: tocId, editDisabled: false },
  })
  const contextual = tabs.find((tab) => tab.id === 'table-of-contents')
  assert.equal(contextual.contextual, true)
  assert.deepEqual(contextual.groups[0].items.map((item) => [item.commandId, item.params]), [
    ['doc.operation.update-table-of-contents', { tocId: 'toc-active' }],
    ['doc.command.delete-table-of-contents', { tocId: 'toc-active' }],
  ])
  assert.deepEqual(documentTocContextMenuItems(tocId, false).map((item) => item.commandId), [
    'doc.operation.update-table-of-contents',
    'doc.command.delete-table-of-contents',
  ])
  assert.equal(documentTocIdForSelection(document, {
    anchor: { blockId: tocBlocks[1].id, offset: 0 },
    focus: { blockId: 'heading', offset: 1 },
  }), undefined)
})

test('the automatic Insert action inserts generated blocks through the standard document command', () => {
  const document = model([block('heading', 'Chapter')])
  const insertionPoint = collapsedSelection({ blockId: 'heading', offset: 0 })
  const generated = buildDocumentTocBlocks(document, options, { heading: 1 }, 'toc-inserted', [], 'Contents')

  executeDocumentCommand(document, insertionPoint, { id: 'doc.command.paste-blocks', blocks: generated })
  assert.deepEqual(document.blocks.map((entry) => entry.metadata.documentToc?.role ?? 'body'), ['body', 'title', 'entry', 'body'])
  assert.equal(document.blocks[2].metadata.documentToc.headingId, 'heading')
})

test('traditional flavor gating and page-number mapping follow layout heading IDs and one-based page labels', () => {
  const document = model([])
  assert.equal(supportsDocumentTableOfContents(document), true)
  document.metadata = { documentFlavor: 'modern' }
  assert.equal(supportsDocumentTableOfContents(document), false)
  document.metadata = { documentStyle: { documentFlavor: 2 } }
  assert.equal(supportsDocumentTableOfContents(document), false)
  assert.deepEqual(documentTocPageNumbers([
    { id: 'first', page: 0 },
    { id: 'second', page: 2 },
    { id: 'invalid', page: -1 },
  ]), { first: 1, second: 3 })
})
