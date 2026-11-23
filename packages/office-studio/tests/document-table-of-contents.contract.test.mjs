import assert from 'node:assert/strict'
import { test } from 'vitest'
import {
  buildDocumentTocBlocks,
  buildDocumentTocInstruction,
  collectDocumentTocHeadings,
  deleteDocumentToc,
  documentTocHeadingTarget,
  documentTocIdForSelection,
  documentTocOptionsAtBlock,
  documentTocPageNumbers,
  supportsDocumentTableOfContents,
  applyDocumentTocPageNumberUpdate,
  applyDocumentTocMutation,
  prepareDocumentTocFullUpdate,
  prepareDocumentTocDeletion,
  prepareDocumentTocPageNumberUpdate,
  updateDocumentToc,
} from '../src/features/document/toc/document-toc.ts'
import { executeDocumentCommand } from '../src/features/document/commands/document-commands.ts'
import { documentTocActionState, prepareDocumentTocAction } from '../src/features/document/toc/document-toc-command.ts'
import { createOfficeMessages, supportedOfficeLocales } from '../src/i18n/office-i18n.ts'
import { createDocumentRibbon, documentTocContextMenuItems } from '../src/components/document/document-ribbon.ts'
import { collapsedSelection, createRun, normalizeDocument, runText } from '../src/features/document/model/document-model.ts'
import { saveDocumentSnapshot } from '../src/features/document/services/document-facade.ts'
import { visibleRibbonTabs } from '../src/platform/ribbon/ribbon-schema.ts'
import { documentTextFieldRanges, documentTextFields } from '../src/features/document/model/document-text-fields.ts'

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
  assert.equal(toc[2].style.indent, 0)
  assert.equal(toc[2].style.indentStart, 18)
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
  executeDocumentCommand(document, { anchor: { blockId: toc[1].id, offset: 0 }, focus: { blockId: toc[1].id, offset: 9 } }, {
    id: 'doc.command.insert-text', text: 'Edited label', marks: { bold: true },
  })
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

test('page-only preparation preserves native fields, mixed result styles and suffix anchors before one publication', () => {
  const document = model([block('one', 'Alpha'), block('two', 'Beta')])
  const entries = buildDocumentTocBlocks(document, options, { one: 99, two: 7 }, 'toc-native', [], 'Contents', false)
  const [first, second] = entries
  first.runs = [createRun('Alpha\t'), createRun('9', { bold: true, href: 'https://example.test/page' }),
    createRun('9', { italic: true, color: '#ff0000' }), createRun('!', { underline: true })]
  first.metadata.documentTextFields = { version: 1, fields: [
    { id: 'toc-field', field: 'TOC', instruction: 'TOC \\o "1-3" \\h', cachedResult: 'old', startOffset: 0, endOffset: 6, endBlockId: second.id },
    { id: 'page-one', field: 'PAGEREF', instruction: 'PAGEREF one \\h', cachedResult: '99', startOffset: 6, endOffset: 8,
      metadata: { source: { native: { retained: true } } } },
  ] }
  second.metadata.documentTextFields = { version: 1, fields: [
    { id: 'page-two', field: 'PAGEREF', cachedResult: '7', startOffset: 5, endOffset: 6, metadata: { source: { properties: { locked: 1 } } } },
  ] }
  document.blocks.unshift(...entries)
  document.contentControls = [
    { id: 'wrapper', type: 'group', blockId: first.id, startOffset: 0, endOffset: 9, lock: 'sdtLocked' },
    { id: 'page-control', type: 'richText', parentId: 'wrapper', blockId: first.id, startOffset: 6, endOffset: 8, lock: 'unlocked' },
    { id: 'suffix-control', type: 'text', parentId: 'wrapper', blockId: first.id, startOffset: 8, endOffset: 9, lock: 'unlocked' },
  ]
  document.comments = { suffix: { id: 'suffix', messages: [], anchor: [{ blockId: first.id, startOffset: 8, endOffset: 9, text: '!' }],
    referenceText: '!', updatedAt: '2000-01-01T00:00:00.000Z' } }
  const selection = { anchor: { blockId: first.id, offset: 9 }, focus: { blockId: first.id, offset: 7 } }
  const before = structuredClone(document), timestamp = new Date(2030, 5, 15, 13).getTime()
  const plan = prepareDocumentTocPageNumberUpdate(document, 'toc-native', { one: 3 }, selection, timestamp)
  assert.equal(plan.changed, true)
  assert.deepEqual(document, before)
  assert.deepEqual(plan.selection, { anchor: { blockId: first.id, offset: 8 }, focus: { blockId: first.id, offset: 7 } })
  applyDocumentTocPageNumberUpdate(document, plan)
  assert.equal(document.blocks[0], first)
  assert.equal(runText(first.runs), 'Alpha\t3!')
  assert.equal(runText(second.runs), 'Beta\t7')
  assert.equal(first.runs.find(run => run.text === '3').bold, true)
  assert.equal(first.runs.find(run => run.text === '3').href, 'https://example.test/page')
  assert.deepEqual(documentTextFields(first).find(field => field.id === 'page-one'), {
    ...before.blocks[0].metadata.documentTextFields.fields[1], endOffset: 7, cachedResult: '3',
  })
  assert.equal(documentTextFields(first).find(field => field.id === 'toc-field').cachedResult, 'Alpha\t3!\nBeta\t7')
  assert.deepEqual(document.contentControls.map(control => [control.id, control.startOffset, control.endOffset]),
    [['wrapper', 0, 8], ['page-control', 6, 7], ['suffix-control', 7, 8]])
  assert.deepEqual(document.comments.suffix.anchor, [{ blockId: first.id, startOffset: 7, endOffset: 8, text: '!' }])
  assert.equal(document.comments.suffix.updatedAt, new Date(timestamp).toISOString())
  documentTextFields(first).find(field => field.id === 'page-one').metadata.source.native.retained = false
  assert.equal(before.blocks[0].metadata.documentTextFields.fields[1].metadata.source.native.retained, true)
  assert.equal(documentTextFieldRanges(document).length, 3)
})

test('page and outer cache repairs retain runs, partial controls, comment anchors and the selected character', () => {
  const document = model([block('one', 'Alpha')])
  const [entry] = buildDocumentTocBlocks(document, options, { one: 3 }, 'toc-cache', [], 'Contents', false)
  entry.runs = [createRun('Alpha\t', { bold: true }), createRun('3', { italic: true })]
  entry.metadata.documentTextFields = { version: 1, fields: [
    { id: 'outer-cache', field: 'TOC', cachedResult: 'stale', startOffset: 0, endOffset: 7 },
    { id: 'page-cache', field: 'PAGEREF', cachedResult: 'stale', startOffset: 6, endOffset: 7 },
  ] }
  document.blocks.unshift(entry)
  document.contentControls = [{ id: 'partial', type: 'text', blockId: entry.id, startOffset: 6, endOffset: 7, lock: 'unlocked' }]
  document.comments = { partial: { id: 'partial', messages: [], anchor: [{ blockId: entry.id, startOffset: 6, endOffset: 7, text: '3' }], referenceText: '3', updatedAt: 'old' } }
  const runs = entry.runs, controls = document.contentControls, comments = document.comments
  const selection = { anchor: { blockId: entry.id, offset: 6 }, focus: { blockId: entry.id, offset: 7 } }
  const plan = prepareDocumentTocPageNumberUpdate(document, 'toc-cache', { one: 3 }, selection)
  assert.equal(plan.changed, true)
  assert.deepEqual(plan.selection, selection)
  applyDocumentTocPageNumberUpdate(document, plan)
  assert.equal(entry.runs, runs)
  assert.equal(document.contentControls, controls)
  assert.equal(document.comments, comments)
  assert.equal(documentTextFields(entry).find(field => field.id === 'page-cache').cachedResult, '3')
  assert.equal(documentTextFields(entry).find(field => field.id === 'outer-cache').cachedResult, 'Alpha\t3')
  assert.equal(prepareDocumentTocPageNumberUpdate(document, 'toc-cache', { one: 3 }, selection).changed, false)
})

test('native locked TOCs and unresolved references preserve the complete source candidate', () => {
  const document = model([block('one', 'Alpha'), block('two', 'Beta')])
  const entries = buildDocumentTocBlocks(document, options, { one: 2, two: 4 }, 'toc-locked', [], 'Contents', false)
  entries[0].metadata.documentTextFields = { version: 1, fields: [{ id: 'native-lock', field: 'TOC', cachedResult: 'old',
    startOffset: 0, endOffset: 6, endBlockId: entries[1].id, metadata: { source: { locked: '1' } } }] }
  document.blocks.unshift(...entries)
  const before = structuredClone(document)
  assert.equal(prepareDocumentTocPageNumberUpdate(document, 'toc-locked', { one: 9, two: 10 }), undefined)
  assert.deepEqual(document, before)
  entries[0].metadata.documentTextFields.fields[0].locked = false
  const unlocked = structuredClone(document)
  assert.equal(prepareDocumentTocPageNumberUpdate(document, 'toc-locked', { one: 9 }), undefined)
  assert.deepEqual(document, unlocked)
})

test('page-only updates preserve inline note and formula entities in native and fallback page ranges', () => {
  const entities = [{ note: { id: 'note-1' } }, { noteReference: { noteId: 'note-1' } }, { latexFormula: { latex: 'x', mode: 'inline', version: 1 } }]
  for (const metadata of entities) for (const native of [false, true]) for (const empty of [false, true]) {
    const document = model([block('one', 'Alpha'), block('two', 'Beta')])
    const [unsafe, safe] = buildDocumentTocBlocks(document, options, { one: 99, two: 4 }, 'toc-entities', [], 'Contents', false)
    unsafe.runs = [createRun('Alpha\t'), createRun('9'), createRun(empty ? '' : '9', { metadata }), ...(empty ? [createRun('9')] : [])]
    delete unsafe.metadata.documentTextFields
    if (native) unsafe.metadata.documentTextFields = { version: 1, fields: [{ id: 'native-page', field: 'PAGEREF', instruction: 'PAGEREF one', cachedResult: '99', startOffset: 6, endOffset: 8 }] }
    const owner = block('entity-toc-owner', 'Prefix', 'paragraph')
    owner.metadata.documentTextFields = { version: 1, fields: [{ id: 'entity-toc', field: 'TOC', startOffset: 0, endOffset: runText(safe.runs).length, endBlockId: safe.id, cachedResult: 'old' }] }
    document.blocks.unshift(owner, unsafe, safe)
    const preserved = structuredClone(unsafe), runs = unsafe.runs
    const plan = prepareDocumentTocPageNumberUpdate(document, 'toc-entities', { two: 11 })
    assert.equal(plan.changed, true)
    applyDocumentTocPageNumberUpdate(document, plan)
    assert.deepEqual(unsafe, preserved)
    assert.equal(unsafe.runs, runs)
    assert.equal(runText(safe.runs), 'Beta\t11')
  }
})

test('generated directories persist a TOC result with nested PAGEREF definitions and canonical switches', () => {
  const document = model([block('one', 'Alpha'), block('two', 'Beta')])
  const toc = buildDocumentTocBlocks(document, options, { one: 3, two: 9 }, 'toc-generated')
  document.blocks.unshift(...toc)
  const fields = documentTextFieldRanges(document), outer = fields.find(range => range.field.field === 'TOC')
  assert.equal(fields.length, 3)
  assert.equal(documentTextFields(toc[0]).length, 0)
  assert.equal(outer.block.id, toc[1].id); assert.equal(outer.endParagraph.block.id, toc[2].id)
  assert.equal(outer.field.cachedResult, 'Alpha\t3\nBeta\t9')
  assert.equal(outer.field.instruction, 'TOC \\o "1-3" \\h \\z \\u')
  assert.deepEqual(fields.filter(range => range.field.field === 'PAGEREF').map(range => [range.field.instruction, range.field.cachedResult]),
    [['PAGEREF "one" \\h', '3'], ['PAGEREF "two" \\h', '9']])
  assert.equal(buildDocumentTocInstruction({ ...options, showPageNumbers: false }), 'TOC \\o "1-3" \\h \\z \\u \\n')
  assert.equal(buildDocumentTocInstruction({ ...options, rightAlignPageNumbers: false }), 'TOC \\o "1-3" \\h \\z \\u \\p " "')
})

test('the first page append grows the outer endpoint and installs only the actual numeric result ranges', () => {
  for (const rightAlignPageNumbers of [false, true]) {
    const document = model([block('one', 'Alpha'), block('two', 'Beta')])
    const entries = buildDocumentTocBlocks(document, { ...options, rightAlignPageNumbers }, {}, 'toc-first-page', [], 'Contents', false)
    document.blocks.unshift(...entries)
    const outerId = documentTextFields(entries[0]).find(field => field.field === 'TOC').id
    const plan = prepareDocumentTocPageNumberUpdate(document, 'toc-first-page', { one: 12, two: 4 })
    applyDocumentTocPageNumberUpdate(document, plan)
    const fields = documentTextFieldRanges(document), outer = fields.find(range => range.field.id === outerId)
    const separator = rightAlignPageNumbers ? '\t' : ' '
    assert.deepEqual(entries.map(entry => runText(entry.runs)), [`Alpha${separator}12`, `Beta${separator}4`])
    assert.equal(outer.field.endOffset, 6); assert.equal(outer.field.endBlockId, entries[1].id)
    assert.equal(outer.field.cachedResult, `Alpha${separator}12\nBeta${separator}4`)
    assert.deepEqual(fields.filter(range => range.field.field === 'PAGEREF').map(range => [range.field.startOffset, range.field.endOffset, range.field.cachedResult]), [[6, 8, '12'], [5, 6, '4']])
    const runs = entries.map(entry => entry.runs)
    assert.equal(prepareDocumentTocPageNumberUpdate(document, 'toc-first-page', { one: 12, two: 4 }).changed, false)
    assert.deepEqual(entries.map(entry => entry.runs), runs)
  }
})

test('a full candidate preserves matching fields and paragraphs while removing deleted entries and mapping their anchors', () => {
  const document = model([block('one', 'Alpha'), block('two', 'Beta')])
  const toc = buildDocumentTocBlocks(document, options, { one: 2, two: 3 }, 'toc-full')
  document.blocks.unshift(...toc)
  const first = toc[1], retained = toc[2]
  const outerId = documentTextFields(first).find(field => field.field === 'TOC').id
  const retainedPageId = documentTextFields(retained).find(field => field.field === 'PAGEREF').id
  const removedPageId = documentTextFields(first).find(field => field.field === 'PAGEREF').id
  document.contentControls = [
    { id: 'removed-wrapper', type: 'group', blockId: first.id, startOffset: 0, endOffset: 7, lock: 'unlocked' },
    { id: 'removed-child', type: 'text', parentId: 'removed-wrapper', blockId: first.id, startOffset: 6, endOffset: 7 },
    { id: 'retained-wrapper', type: 'group', blockId: retained.id, startOffset: 0, endOffset: 6, lock: 'sdtLocked' },
    { id: 'retained-page', type: 'richText', parentId: 'retained-wrapper', blockId: retained.id, startOffset: 5, endOffset: 6, lock: 'unlocked' },
  ]
  document.comments = {
    removed: { id: 'removed', messages: [{ id: 'reply', text: 'Keep the thread' }], anchor: [{ blockId: first.id, startOffset: 0, endOffset: 5, text: 'Alpha' }], referenceText: 'Alpha', updatedAt: 'old' },
    retained: { id: 'retained', messages: [], anchor: [{ blockId: retained.id, startOffset: 5, endOffset: 6, text: '3' }], referenceText: '3', updatedAt: 'old' },
  }
  document.blocks = document.blocks.filter(entry => entry.id !== 'one')
  document.blocks.push(block('three', 'Gamma'))
  const source = structuredClone(document), timestamp = Date.UTC(2030, 5, 15)
  const selection = { anchor: { blockId: retained.id, offset: 6 }, focus: { blockId: first.id, offset: 2 } }
  const plan = prepareDocumentTocFullUpdate(document, 'toc-full', { two: 10, three: 5 }, selection, timestamp)
  assert.equal(plan.changed, true); assert.deepEqual(document, source)
  assert.deepEqual(plan.selection, { anchor: { blockId: retained.id, offset: 7 }, focus: { blockId: retained.id, offset: 0 } })
  applyDocumentTocMutation(document, plan)
  const entries = document.blocks.filter(entry => entry.metadata.documentToc?.role === 'entry')
  assert.deepEqual(entries.map(entry => entry.metadata.documentToc.headingId), ['two', 'three'])
  assert.equal(entries[0], retained); assert.equal(runText(retained.runs), 'Beta\t10')
  const fields = documentTextFieldRanges(document)
  assert.equal(fields.length, 3); assert.equal(fields.some(range => range.field.id === removedPageId), false)
  assert.equal(fields.find(range => range.field.id === outerId).block.id, retained.id)
  assert.equal(fields.find(range => range.field.id === outerId).field.cachedResult, 'Beta\t10\nGamma\t5')
  assert.equal(fields.find(range => range.field.id === retainedPageId).field.cachedResult, '10')
  assert.deepEqual(document.contentControls.map(control => [control.id, control.startOffset, control.endOffset]), [['retained-wrapper', 0, 7], ['retained-page', 5, 7]])
  assert.deepEqual(document.comments.removed.anchor, [{ blockId: retained.id, startOffset: 0, endOffset: 0, text: '' }])
  assert.equal(document.comments.removed.messages[0].text, 'Keep the thread')
  assert.deepEqual(document.comments.retained.anchor, [{ blockId: retained.id, startOffset: 5, endOffset: 7, text: '10' }])
  assert.equal(document.comments.retained.updatedAt, new Date(timestamp).toISOString())
  const runs = retained.runs
  assert.equal(prepareDocumentTocFullUpdate(document, 'toc-full', { two: 10, three: 5 }, plan.selection, timestamp).changed, false)
  assert.equal(retained.runs, runs)
})

test('full rebuild keeps page controls, comments, formatting and backward selection separate from a renamed heading', () => {
  const document = model([block('one', 'Alpha')])
  const [entry] = buildDocumentTocBlocks(document, options, { one: 9 }, 'toc-renamed', [], 'Contents', false)
  entry.runs = [createRun('Alpha\t', { fontFamily: 'Georgia', italic: true }), createRun('9', { bold: true, color: '#ab1234' })]
  document.blocks.unshift(entry)
  const sourceLength = runText(entry.runs).length, oldPageStart = sourceLength - 1
  document.contentControls = [
    { id: 'wrapper', type: 'group', blockId: entry.id, startOffset: 0, endOffset: sourceLength, lock: 'sdtLocked' },
    { id: 'heading', type: 'richText', parentId: 'wrapper', blockId: entry.id, startOffset: 0, endOffset: 5 },
    { id: 'page', type: 'richText', parentId: 'wrapper', blockId: entry.id, startOffset: oldPageStart, endOffset: sourceLength },
  ]
  document.comments = {
    heading: { id: 'heading', messages: [], anchor: [{ blockId: entry.id, startOffset: 0, endOffset: 5, text: 'Alpha' }], referenceText: 'Alpha', updatedAt: 'old' },
    page: { id: 'page', messages: [], anchor: [{ blockId: entry.id, startOffset: oldPageStart, endOffset: sourceLength, text: '9' }], referenceText: '9', updatedAt: 'old' },
  }
  const label = 'Revised heading'
  document.blocks.find(candidate => candidate.id === 'one').runs[0].text = label
  const selection = { anchor: { blockId: entry.id, offset: sourceLength }, focus: { blockId: entry.id, offset: oldPageStart } }
  const source = structuredClone(document), start = label.length + 1, end = start + 2
  const plan = prepareDocumentTocFullUpdate(document, 'toc-renamed', { one: 12 }, selection, 0)
  assert.deepEqual(document, source)
  assert.deepEqual(plan.selection, { anchor: { blockId: entry.id, offset: end }, focus: { blockId: entry.id, offset: start } })
  applyDocumentTocMutation(document, plan)
  assert.equal(runText(entry.runs), `${label}\t12`)
  assert.deepEqual(document.contentControls.map(({ id, startOffset, endOffset }) => ({ id, startOffset, endOffset })), [
    { id: 'wrapper', startOffset: 0, endOffset: end }, { id: 'heading', startOffset: 0, endOffset: label.length }, { id: 'page', startOffset: start, endOffset: end },
  ])
  assert.deepEqual(document.comments.page.anchor, [{ blockId: entry.id, startOffset: start, endOffset: end, text: '12' }])
  assert.equal(document.comments.heading.referenceText, label)
  const numberRun = entry.runs.find(run => run.text === '12')
  assert.equal(numberRun.bold, true); assert.equal(numberRun.color, '#ab1234')
  assert.equal(entry.runs.filter(run => run.text !== '12').every(run => run.fontFamily === 'Georgia'), true)
  assert.equal(prepareDocumentTocFullUpdate(document, 'toc-renamed', { one: 12 }, plan.selection, 0).changed, false)
})

test('deleting a directory preserves an enclosing unknown field and anchors comments on surviving text', () => {
  const document = model([block('one', 'Alpha'), block('two', 'Beta')])
  const toc = buildDocumentTocBlocks(document, options, { one: 1, two: 2 }, 'toc-delete', [], 'Contents', false)
  const prefix = block('prefix', 'Prefix', 'paragraph'), tail = block('tail', 'Tail', 'paragraph')
  prefix.metadata.documentTextFields = { version: 1, fields: [{ id: 'enclosing', field: 'UNKNOWN', instruction: 'UNKNOWN retained', cachedResult: 'original', startOffset: 0, endOffset: 4, endBlockId: tail.id }] }
  document.blocks.unshift(prefix, ...toc, tail)
  document.contentControls = [{ id: 'removed-control', type: 'text', blockId: toc[0].id, startOffset: 0, endOffset: 5, lock: 'unlocked' },
    { id: 'tail-control', type: 'text', blockId: tail.id, startOffset: 0, endOffset: 4, lock: 'contentLocked' }]
  document.comments = { removed: { id: 'removed', messages: [], anchor: [{ blockId: toc[0].id, startOffset: 0, endOffset: 5, text: 'Alpha' }], referenceText: 'Alpha', updatedAt: 'old' },
    tail: { id: 'tail', messages: [], anchor: [{ blockId: tail.id, startOffset: 0, endOffset: 4, text: 'Tail' }], referenceText: 'Tail', updatedAt: 'old' } }
  const before = structuredClone(document), selection = collapsedSelection({ blockId: toc[1].id, offset: 3 })
  const plan = prepareDocumentTocDeletion(document, 'toc-delete', selection, Date.UTC(2030, 5, 15))
  assert.deepEqual(document, before); assert.deepEqual(plan.selection, collapsedSelection({ blockId: tail.id, offset: 0 }))
  applyDocumentTocMutation(document, plan)
  assert.equal(document.blocks.includes(prefix), true); assert.equal(document.blocks.includes(tail), true)
  assert.deepEqual(documentTextFieldRanges(document).map(range => range.field.id), ['enclosing'])
  assert.equal(documentTextFields(prefix)[0].cachedResult, 'original')
  assert.equal(documentTextFields(prefix)[0].endBlockId, tail.id)
  assert.deepEqual(document.contentControls.map(control => control.id), ['tail-control'])
  assert.deepEqual(document.comments.removed.anchor, [{ blockId: tail.id, startOffset: 0, endOffset: 0, text: '' }])
  assert.deepEqual(document.comments.tail, before.comments.tail)
})

test('deleting the final directory keeps a blank input paragraph and rejects locked wrappers atomically', () => {
  for (const lock of ['unlocked', 'controlLocked']) {
    const document = model([block('one', 'Alpha')])
    const [entry] = buildDocumentTocBlocks(document, options, { one: 1 }, 'toc-only', [], 'Contents', false)
    document.blocks = [entry]
    document.contentControls = [{ id: 'wrapper', type: 'group', blockId: entry.id, startOffset: 0, endOffset: 7, lock }]
    document.comments = { removed: { id: 'removed', messages: [], anchor: [{ blockId: entry.id, startOffset: 0, endOffset: 5, text: 'Alpha' }], referenceText: 'Alpha', updatedAt: 'old' } }
    const before = structuredClone(document)
    if (lock === 'controlLocked') {
      assert.throws(() => prepareDocumentTocDeletion(document, 'toc-only'), /wrapper is locked/)
      assert.deepEqual(document, before)
    } else {
      const plan = prepareDocumentTocDeletion(document, 'toc-only', collapsedSelection({ blockId: entry.id, offset: 3 }))
      applyDocumentTocMutation(document, plan)
      assert.equal(document.blocks.length, 1); assert.equal(runText(document.blocks[0].runs), '')
      assert.equal(document.blocks[0].metadata.documentToc, undefined); assert.equal(documentTextFieldRanges(document).length, 0)
      assert.deepEqual(plan.selection, collapsedSelection({ blockId: document.blocks[0].id, offset: 0 }))
      assert.deepEqual(document.comments.removed.anchor, [{ blockId: document.blocks[0].id, startOffset: 0, endOffset: 0, text: '' }])
    }
  }
})

test('full and deletion commands enforce field ACLs even when an unchanged control shares an entity ID', () => {
  for (const action of ['pageNumbersOnly', 'all', 'delete']) {
    const document = model([block('one', 'Alpha')])
    const [entry] = buildDocumentTocBlocks(document, options, { one: 2 }, 'toc-acl', [], 'Contents', false)
    const body = block('body', 'Other', 'paragraph'); document.blocks.unshift(entry, body)
    const pageId = documentTextFields(entry).find(field => field.field === 'PAGEREF').id
    document.contentControls = [{ id: pageId, type: 'text', blockId: body.id, startOffset: 0, endOffset: 5 }]
    document.permissionRules = [{ id: 'deny-page', scope: 'entity', targetId: pageId, action: 'edit', effect: 'deny', userIds: [] }]
    const source = structuredClone(document), selection = collapsedSelection({ blockId: entry.id, offset: 2 })
    assert.throws(() => executeDocumentCommand(document, selection, action === 'delete'
      ? { id: 'doc.command.delete-table-of-contents', tocId: 'toc-acl' }
      : { id: 'doc.command.update-table-of-contents', tocId: 'toc-acl', mode: action, pageNumbers: { one: 3 } }), error => error.code === 'DOCUMENT_PERMISSION_DENIED')
    assert.deepEqual(document, source)
  }
})

test('page updates check permissions for a native outer field owner outside the TOC rows', () => {
  const document = model([block('one', 'Alpha')])
  const [entry] = buildDocumentTocBlocks(document, options, { one: 2 }, 'toc-permission', [], 'Contents', false)
  const owner = block('owner', 'Prefix', 'paragraph')
  owner.metadata.documentTextFields = { version: 1, fields: [{ id: 'external-toc', field: 'TOC', cachedResult: 'old',
    startOffset: 0, endOffset: 7, endBlockId: entry.id }] }
  document.blocks.unshift(owner, entry)
  document.permissionRules = [{ id: 'deny-owner', scope: 'paragraph', targetId: owner.id, action: 'edit', effect: 'deny', userIds: [] }]
  const before = structuredClone(document)
  assert.throws(() => executeDocumentCommand(document, collapsedSelection({ blockId: entry.id, offset: 6 }), {
    id: 'doc.command.update-table-of-contents', tocId: 'toc-permission', mode: 'pageNumbersOnly', pageNumbers: { one: 3 },
  }), error => error.code === 'DOCUMENT_PERMISSION_DENIED')
  assert.deepEqual(document, before)
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
  }), /选区位于目标目录内/)

  document.permissionRules = [{
    id: 'deny-toc-row', scope: 'paragraph', targetId: firstToc[1].id, action: 'edit', effect: 'deny', userIds: [],
  }]
  assert.throws(() => executeDocumentCommand(document, selection, {
    id: 'doc.command.delete-table-of-contents', tocId: 'toc-a',
  }), (error) => error.code === 'DOCUMENT_PERMISSION_DENIED')
})

test('TOC capabilities distinguish unresolved pages, valid no-ops, locked children and source-preserving rejections', () => {
  const document = model([block('one', 'Alpha')])
  const [entry] = buildDocumentTocBlocks(document, options, { one: 2 }, 'toc-capability', [], 'Contents', false)
  document.blocks.unshift(entry)
  const selected = collapsedSelection({ blockId: entry.id, offset: 0 }), context = { documentId: document.id, userId: 'local-user' }
  const initial = structuredClone(document)
  assert.equal(documentTocActionState(document, selected, 'toc-capability', 'pageNumbersOnly', { one: 2 }, context).available, true)
  assert.equal(prepareDocumentTocAction(document, selected, 'toc-capability', 'pageNumbersOnly', { one: 2 }, context).changed, false)
  assert.equal(documentTocActionState(document, selected, 'toc-capability', 'all', {}, context).available, false)
  assert.equal(documentTocActionState(document, selected, 'toc-capability', 'delete', {}, context).available, true)
  assert.deepEqual(document, initial)
  const page = documentTextFields(entry).find(field => field.field === 'PAGEREF')
  page.locked = true
  const locked = structuredClone(document)
  for (const action of ['pageNumbersOnly', 'all', 'delete']) {
    const state = documentTocActionState(document, selected, 'toc-capability', action, { one: 9 }, context)
    assert.equal(state.available, false)
    assert.ok(state.reason)
    assert.throws(() => prepareDocumentTocAction(document, selected, 'toc-capability', action, { one: 9 }, context), error => error.message === state.reason)
    assert.deepEqual(document, locked)
  }
})

test('TOC capability checks include changed field entity permissions before a command publishes', () => {
  const document = model([block('one', 'Alpha')])
  const [entry] = buildDocumentTocBlocks(document, options, { one: 2 }, 'toc-capability-acl', [], 'Contents', false)
  document.blocks.unshift(entry)
  const page = documentTextFields(entry).find(field => field.field === 'PAGEREF')
  document.permissionRules = [{ id: 'deny-page', scope: 'entity', targetId: page.id, action: 'edit', effect: 'deny', userIds: [] }]
  const selected = collapsedSelection({ blockId: entry.id, offset: 0 }), context = { documentId: document.id, userId: 'local-user' }
  const source = structuredClone(document)
  for (const action of ['pageNumbersOnly', 'all', 'delete']) {
    assert.equal(documentTocActionState(document, selected, 'toc-capability-acl', action, { one: 9 }, context).available, false)
    assert.throws(() => prepareDocumentTocAction(document, selected, 'toc-capability-acl', action, { one: 9 }, context), error => error.code === 'DOCUMENT_PERMISSION_DENIED')
    assert.deepEqual(document, source)
  }
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

test('TOC Ribbon and Layout menu keep update/delete capabilities and reasons independent', () => {
  for (const [update, remove] of [[{ available: false, reason: 'Pages are pending' }, { available: true }],
    [{ available: true }, { available: false, reason: 'Wrapper cannot be removed' }]]) {
    const tabs = createDocumentRibbon({
      state: { marks: {}, mixedMarks: [], paragraph: { align: 'left', indent: 0 }, hasSelection: false },
      canUndo: false, canRedo: false, outline: false, focus: false, fullscreen: false,
      paper: 'A4', orientation: 'portrait', zoom: 100, contentControlsAvailable: false,
      tableOfContents: { available: true, insertDisabled: true, selectedId: 'selected-toc', editDisabled: true,
        updateDisabled: !update.available, deleteDisabled: !remove.available, updateReason: update.reason, deleteReason: remove.reason },
    })
    const ribbon = tabs.find(tab => tab.id === 'table-of-contents').groups[0].items
    const menu = documentTocContextMenuItems('selected-toc', { update, delete: remove })
    for (const items of [ribbon, menu]) {
      assert.deepEqual(items.map(item => item.disabled), [!update.available, !remove.available])
      assert.deepEqual(items.map(item => item.disabledReason), [update.reason, remove.reason])
      assert.deepEqual(items.map(item => item.params), [{ tocId: 'selected-toc' }, { tocId: 'selected-toc' }])
    }
  }
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

test('TOC page labels follow a section restart and skip parity blanks or unavailable layout pages', () => {
  const headings = [{ id: 'first', page: 0 }, { id: 'blank', page: 1 }, { id: 'restart', page: 2 },
    { id: 'following', page: 3 }, { id: 'pending', page: 4 }, { id: 'invalid', page: -1 }]
  const pages = [{ pageNumber: 9, isParityBlank: false }, { pageNumber: 10, isParityBlank: true },
    { pageNumber: 1, isParityBlank: false }, { pageNumber: 2, isParityBlank: false }]
  assert.deepEqual(documentTocPageNumbers(headings, pages), { first: 9, restart: 1, following: 2 })
  assert.deepEqual(documentTocPageNumbers(headings, []), {})
  for (const pageNumber of [undefined, 0, -1, 1.5, NaN, Infinity]) {
    assert.deepEqual(documentTocPageNumbers([{ id: 'heading', page: 0 }], [{ pageNumber }]), {})
  }
})
