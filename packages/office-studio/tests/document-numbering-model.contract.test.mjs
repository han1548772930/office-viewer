import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import { normalizeDocument } from '../src/parser/documentNormalizer.ts'
import { createParagraph, createTable } from '../src/features/document/model/document-model.ts'
import { blockListLevel, formatListOrdinal, paragraphListMarkers } from '../src/features/document/model/document-lists.ts'
import { blockHasNumbering, blockNumberingDefinition, blockNumberingId, numberingDefinitionFromParts } from '../src/features/document/model/document-numbering.ts'
import { officeNumberingMarkerBaseline, officeNumberingModelContract, officeNumberingPrefixContract, officeNumberingPrefixLength } from '../src/features/document/model/document-office-numbering-marker.ts'
import { canContinueDocumentNumbering, canRestartDocumentNumbering, documentTabCommand, executeDocumentCommand } from '../src/features/document/commands/document-commands.ts'
import { documentThreadCommentAnchorStatus } from '../src/features/document/comments/thread-comments.ts'

function envelope(blocks, numberingDefinitions, extra = {}) {
  return {
    schemaVersion: 1, parserVersion: 'contract', format: 'docx', fileName: 'format.docx', size: 100,
    model: 'document', payload: { modelType: 'document', data: { documentId: 'format-doc', blocks, numberingDefinitions, ...extra } },
    diagnostics: { warnings: [], errors: [], stats: {}, capabilities: [] },
  }
}

function parsed(id, style, extra = {}) {
  return { id, type: 'text', content: id, details: { style, inlines: [{ type: 'text', text: id }], ...extra } }
}

function definition(levels, overrides = {}) {
  return numberingDefinitionFromParts({ abstractNumId: 7, levels }, { numId: 9, abstractNumId: 7, ...overrides })
}

function listed(id, level, numbering) {
  return { ...createParagraph(id), id, type: 'numbered', level, metadata: { numberingListId: 'shared', numberingDefinition: numbering } }
}

test('paragraph normalization reads twips, line rules, signed indentation and explicit zero spacing', () => {
  const model = normalizeDocument(envelope([
    parsed('exact', { indent: { leftTwips: -150, rightTwips: 300, firstLineTwips: 180, hangingTwips: 60 }, spacing: { beforeTwips: 0, afterTwips: 0, lineTwips: 360, lineRule: 'exact' }, wordWrap: false, snapToGrid: true }),
    parsed('auto', { spacing: { lineTwips: 180, lineRule: 'auto' } }),
    parsed('minimum', { spacing: { lineTwips: 300, lineRule: 'atLeast' } }),
    parsed('legacy', { indent: { leftPt: 12, right: 90, firstLinePt: 6 }, spacing: { beforePt: 3, after: 0, lineMultiplier: 2 } }),
  ]))
  assert.deepEqual(model.blocks[0].style, { align: 'left', indent: 0, indentStart: -10, indentEnd: 20, indentFirstLine: 12, hanging: 4, spacingRule: 'exact', lineHeight: 24, spaceBefore: 0, spaceAfter: 0, wordWrap: false, snapToGrid: true })
  assert.equal(model.blocks[1].style.lineHeight, 0.75)
  assert.equal(model.blocks[1].style.spacingRule, 'auto')
  assert.equal(model.blocks[2].style.lineHeight, 20)
  assert.equal(model.blocks[2].style.spacingRule, 'atLeast')
  assert.equal(model.blocks[3].style.indentStart, 16)
  assert.equal(model.blocks[3].style.indentEnd, 6)
  assert.equal(model.blocks[3].style.indentFirstLine, 8)
  assert.equal(model.blocks[3].style.spaceBefore, 4)
  assert.equal(model.blocks[3].style.spaceAfter, 0)
})

test('numbered headings keep independent heading and list levels and invisible numbering stays a list', () => {
  const numberingDefinitions = {
    abstracts: [{ abstractNumId: 7, levels: [{ ilvl: 0, start: 1, format: 'decimal', text: '%1.' }, { ilvl: 1, start: 1, format: 'none', text: '%2.' }] }],
    instances: [{ numId: 9, abstractNumId: 7, levelStartOverrides: { 0: 0 } }],
  }
  const model = normalizeDocument(envelope([
    { ...parsed('heading', { styleId: 'Heading2', headingLevel: 2 }, { headingLevel: 2, numbering: { numId: 9, ilvl: 0 } }), type: 'heading' },
    parsed('invisible', {}, { numbering: { numId: 9, ilvl: 1 } }),
  ], numberingDefinitions))
  const heading = model.blocks[0]
  assert.equal(heading.type, 'heading')
  assert.equal(heading.level, 2)
  assert.equal(blockHasNumbering(heading), true)
  assert.equal(blockListLevel(heading).index, 0)
  assert.equal(heading.metadata.numberingLevel, 0)
  assert.equal(model.blocks[1].type, 'numbered')
  const markers = paragraphListMarkers(model.blocks)
  assert.equal(markers.get('heading').marker, '0.')
  assert.equal(markers.get('invisible').marker, '')
})

test('one list instance continues across body paragraphs and separate table cells', () => {
  const numbering = definition([{ ilvl: 0, start: 1, format: 'decimal', text: '%1.' }], { levelStartOverrides: { 0: 0 } })
  const first = listed('first', 0, numbering)
  const table = { ...createParagraph(), type: 'table', metadata: { table: createTable(1, 2) } }
  table.metadata.table.rows[0].cells[0].blocks = [listed('cell-a', 0, numbering)]
  table.metadata.table.rows[0].cells[1].blocks = [listed('cell-b', 0, numbering)]
  const markers = paragraphListMarkers([first, table, listed('last', 0, numbering)])
  assert.deepEqual(['first', 'cell-a', 'cell-b', 'last'].map(id => markers.get(id).marker), ['0.', '1.', '2.', '3.'])
})

test('zero starts apply on first use and nested levels restart from their abstract start later', () => {
  const numbering = definition([
    { ilvl: 0, start: 1, format: 'decimal', text: '%1.' },
    { ilvl: 1, start: 3, format: 'decimal', text: '%1.%2.', lvlRestart: 1 },
  ], { levelStartOverrides: { 1: 0 } })
  const blocks = [0, 1, 1, 0, 1].map((level, index) => listed(`p-${index}`, level, numbering))
  const markers = paragraphListMarkers(blocks)
  assert.deepEqual(blocks.map(block => markers.get(block.id).marker), ['1.', '1.0.', '1.1.', '2.', '2.3.'])
})

test('effective overrides retain omitted level fields and markers keep their own character formatting', () => {
  const numbering = definition([{ ilvl: 0, start: 2, format: 'decimalZero', text: '%1)', suffix: 'space', indent: { leftTwips: -150, rightTwips: -30, hangingTwips: 300 }, runStyle: { fontFamilyAscii: 'Georgia', fontSizePt: 18, bold: true, color: 'auto' } }], { levelOverrides: [{ ilvl: 0, text: '[%1]' }] })
  const marker = paragraphListMarkers([listed('custom', 0, numbering)]).get('custom')
  assert.equal(marker.marker, '[02]')
  assert.equal(marker.fontFamily, 'Georgia')
  assert.equal(marker.fontSize, 18)
  assert.equal(marker.color, 'auto')
  assert.equal(marker.bold, true)
  assert.equal(marker.inset, -10)
  assert.equal(marker.markerIndent, -30)
  assert.equal(marker.rightIndent, -2)
  assert.equal(marker.gap, 4)
})

test('number formatting preserves the zero boundary and Office alphabetic overflow', () => {
  assert.equal(formatListOrdinal(0, 'decimalZero', 4), '0')
  assert.equal(formatListOrdinal(3, 'decimal', 4), '0003')
  assert.equal(formatListOrdinal(27, 'lowerLetter'), 'aa')
  assert.equal(formatListOrdinal(28, 'upperLetter'), 'BB')
  assert.equal(formatListOrdinal(14, 'upperRoman'), 'XIV')
})

function consecutiveDefinition() {
  const numbering = definition([
    { ilvl: 0, start: 1, format: 'decimal', text: '%1.' },
    { ilvl: 1, start: 3, format: 'decimal', text: '%1.%2.' },
  ], { consecutiveListId: 9, levelStartOverrides: { 0: 0 } })
  numbering.abstract.consecutiveNumbering = true
  return numbering
}

const documentWith = blocks => ({ ...createBlankDocument(), id: 'numbering-command-document', blocks })
const caret = block => ({ anchor: { blockId: block.id, offset: 0 }, focus: { blockId: block.id, offset: 0 } })

test('restarting a consecutive instance detaches its future paragraphs from the original counter', () => {
  const numbering = consecutiveDefinition()
  const blocks = [0, 1, 0, 1].map((level, index) => listed(`consecutive-${index}`, level, numbering))
  const model = documentWith(blocks)
  const runs = blocks.map(block => block.runs)
  assert.deepEqual(blocks.map(block => paragraphListMarkers(blocks).get(block.id).marker), ['0.', '0.1.', '2.', '2.3.'])
  executeDocumentCommand(model, caret(blocks[1]), { id: 'doc.command.restart-list-numbering', start: 7 })
  const markers = paragraphListMarkers(model.blocks)
  assert.deepEqual(blocks.map(block => markers.get(block.id).marker), ['0.', '0.7.', '8.', '8.9.'])
  assert.equal(blockNumberingId(blocks[0]), 'shared')
  const restarted = blockNumberingId(blocks[1])
  assert.notEqual(restarted, 'shared')
  for (const block of blocks.slice(1)) {
    assert.equal(blockNumberingId(block), restarted)
    assert.equal(blockNumberingDefinition(block).instance.consecutiveListId, restarted)
  }
  assert.equal(blockNumberingDefinition(blocks[0]).instance.consecutiveListId, 9)
  blocks.forEach((block, index) => assert.equal(block.runs, runs[index]))
})

test('the first selected nested consecutive level honors its explicit start instead of another level override', () => {
  const numbering = consecutiveDefinition()
  numbering.instance.levelStartOverrides['1'] = 7
  const blocks = [listed('nested-first', 1, numbering), listed('root-next', 0, numbering)]
  const markers = paragraphListMarkers(blocks)
  assert.equal(markers.get('nested-first').marker, '0.7.')
  assert.equal(markers.get('root-next').marker, '8.')
})

test('restart and Tab use a numbered heading list level while keeping heading depth and text intact', () => {
  const numbering = consecutiveDefinition()
  const first = { ...listed('heading-first', 0, numbering), type: 'heading', level: 2, metadata: { numberingListId: 'shared', numberingDefinition: numbering, numberingLevel: 0 } }
  const nested = { ...listed('heading-nested', 0, numbering), type: 'heading', level: 4, metadata: { numberingListId: 'shared', numberingDefinition: numbering, numberingLevel: 1 } }
  const last = { ...listed('quote-last', 0, numbering), type: 'quote', metadata: { numberingListId: 'shared', numberingDefinition: numbering, numberingLevel: 0 } }
  const model = documentWith([first, nested, last])
  const contents = model.blocks.map(block => JSON.stringify(block.runs))
  assert.equal(canRestartDocumentNumbering(model, caret(first)), true)
  assert.deepEqual(documentTabCommand(model, caret(nested), false), { id: 'doc.command.change-list-nesting-level', direction: 1 })
  executeDocumentCommand(model, caret(first), { id: 'doc.command.restart-list-numbering', start: 10 })
  assert.deepEqual(model.blocks.map(block => [block.type, block.level, block.metadata.numberingLevel]), [['heading', 2, 0], ['heading', 4, 1], ['quote', 0, 0]])
  const markers = paragraphListMarkers(model.blocks)
  assert.deepEqual(model.blocks.map(block => markers.get(block.id).marker), ['10.', '10.11.', '12.'])
  assert.deepEqual(model.blocks.map(block => JSON.stringify(block.runs)), contents)
})

test('already shared consecutive groups make continue numbering a no-op even with separate instance IDs', () => {
  const firstDefinition = consecutiveDefinition(), nextDefinition = consecutiveDefinition()
  nextDefinition.instance.numId = 10
  const first = listed('primary', 0, firstDefinition)
  const second = { ...listed('secondary', 1, nextDefinition), metadata: { numberingDefinition: nextDefinition, numberingListId: 'secondary-instance' } }
  const model = documentWith([first, second]), before = JSON.stringify(model)
  assert.equal(canContinueDocumentNumbering(model, caret(second)), false)
  executeDocumentCommand(model, caret(second), { id: 'doc.command.continue-list-numbering' })
  assert.equal(JSON.stringify(model), before)
})

test('continuing a different numbered heading instance preserves each target numbering depth', () => {
  const numbering = definition([{ ilvl: 0, start: 1, format: 'decimal', text: '%1.' }, { ilvl: 1, start: 1, format: 'decimal', text: '%1.%2.' }])
  const first = listed('prior-list', 0, numbering)
  const heading = { ...listed('later-heading', 0, numbering), type: 'heading', level: 3, metadata: { numberingDefinition: numbering, numberingListId: 'later-list', numberingLevel: 1 } }
  const model = documentWith([first, heading])
  assert.equal(canContinueDocumentNumbering(model, caret(heading)), true)
  executeDocumentCommand(model, caret(heading), { id: 'doc.command.continue-list-numbering' })
  assert.equal(blockNumberingId(heading), blockNumberingId(first))
  assert.equal(heading.level, 3)
  assert.equal(heading.metadata.numberingLevel, 1)
  assert.equal(paragraphListMarkers(model.blocks).get(heading.id).marker, '1.1.')
})

test('a protected later numbered heading rejects a restart before any instance is changed', () => {
  const numbering = consecutiveDefinition(), first = listed('editable-number', 0, numbering)
  const heading = { ...listed('protected-heading', 0, numbering), type: 'heading', level: 2, metadata: { numberingDefinition: numbering, numberingListId: 'shared', numberingLevel: 1 } }
  const model = documentWith([first, heading])
  model.permissionRules = [{ id: 'locked-heading', scope: 'paragraph', targetId: heading.id, action: 'edit', effect: 'deny', userIds: ['local-user'] }]
  const context = { documentId: model.id, userId: 'local-user' }, before = JSON.stringify(model)
  assert.equal(canRestartDocumentNumbering(model, caret(first), context), false)
  assert.throws(() => executeDocumentCommand(model, caret(first), { id: 'doc.command.restart-list-numbering', start: 7 }, context))
  assert.equal(JSON.stringify(model), before)
})

function withCachedMarker(block, text) {
  block.metadata.officeNumberingMarker = {
    version: 1, bookmark: 'office_list_v1_fixture', bookmarkId: 99, text,
    xml: `<w:bookmarkStart w:id="99" w:name="office_list_v1_fixture"/><w:r><w:t>${text}</w:t></w:r><w:bookmarkEnd w:id="99"/>`,
    baseline: officeNumberingMarkerBaseline(block, text),
  }
  return block
}

test('imported numbering marker baselines copy the numbering and prefix instead of following subsequent edits', () => {
  const raw = parsed('cached', {}, {
    numbering: { numId: 9, ilvl: 0 },
    inlines: [{ type: 'text', text: '[0] ', style: { color: '#123456', fontFamilyAscii: 'Georgia', bold: true } }, { type: 'text', text: 'Item', style: {} }],
    officeNumberingMarker: { version: 1, bookmark: 'office_list_v1_fixture', bookmarkId: 99, text: '[0] ', xml: '<w:bookmarkStart w:id="99" w:name="office_list_v1_fixture"/><w:r><w:t>[0] </w:t></w:r><w:bookmarkEnd w:id="99"/>' },
  })
  raw.content = '[0] Item'
  const model = normalizeDocument(envelope([raw], {
    abstracts: [{ abstractNumId: 7, levels: [{ ilvl: 0, start: 0, format: 'decimal', text: '%1.' }] }],
    instances: [{ numId: 9, abstractNumId: 7 }],
  }))
  const block = model.blocks[0], baseline = block.metadata.officeNumberingMarker.baseline
  assert.deepEqual(baseline.numbering, officeNumberingModelContract(block))
  const before = JSON.stringify(baseline)
  block.metadata.numberingDefinition.levels[0].text = '(%1)'
  block.runs[0].color = '#654321'
  raw.details.numbering.numId = 44
  assert.equal(JSON.stringify(baseline), before)
  assert.equal(baseline.prefix[0].properties.color, '#123456')
  assert.equal(baseline.numbering.numbering.numId, 9)
})

test('unchanged cached labels survive run resegmentation and body editing but edited label text is rejected', () => {
  const block = withCachedMarker({ ...createParagraph('03. Body'), id: 'resegmented' }, '03. ')
  const run = block.runs[0], baseline = block.metadata.officeNumberingMarker.baseline
  block.runs = [{ ...run, id: 'first-fragment', text: '0' }, { ...run, id: 'next-fragment', text: '3. ' }, { ...run, id: 'new-body', text: 'Changed body', italic: true }]
  assert.deepEqual(officeNumberingPrefixContract(block, '03. '), baseline.prefix)
  assert.deepEqual(officeNumberingModelContract(block), baseline.numbering)
  assert.equal(officeNumberingPrefixLength(block), 4)
  block.runs[1].text = '4. '
  assert.equal(officeNumberingPrefixLength(block), undefined)
})

test('cached-prefix extraction respects UTF-16 boundaries and never consumes note or LaTeX objects', () => {
  const emoji = createParagraph('😀 body')
  assert.equal(officeNumberingPrefixContract(emoji, '\uD83D'), undefined)
  assert.equal(officeNumberingPrefixContract(emoji, '😀')[0].text, '😀')
  for (const metadata of [{ note: { type: 'footnote', noteId: 'n1' } }, { latexFormula: { latex: 'x', mode: 'inline' } }]) {
    const block = withCachedMarker(createParagraph('1. Item'), '1. '), run = block.runs[0]
    block.runs = [{ ...run, text: '', metadata }, { ...run, text: '1. Item' }]
    assert.equal(officeNumberingPrefixLength(block), undefined)
    block.runs = [{ ...run, text: '1', metadata }, { ...run, text: '. Item' }]
    assert.equal(officeNumberingPrefixLength(block), undefined)
  }
})

function anchoredThread(id, block, startOffset, endOffset) {
  const text = block.runs.map(run => run.text).join('').slice(startOffset, endOffset)
  return { id, unitId: 'numbering-command-document', resolved: false, referenceText: text,
    anchor: [{ blockId: block.id, startOffset, endOffset, text }],
    messages: [{ id: `${id}-message`, text: 'Keep this comment.', authorId: 'local-user', authorName: 'Author', createdAt: '2026-11-23T00:00:00Z' }],
    createdAt: '2026-11-23T00:00:00Z', updatedAt: '2026-11-23T00:00:00Z' }
}

test('explicit numbering replacement removes one cached prefix and transforms selection, comments and body controls together', () => {
  const block = withCachedMarker({ ...createParagraph('03. Body'), id: 'replace-label' }, '03. ')
  const model = documentWith([block])
  model.comments = { crossing: anchoredThread('crossing', block, 2, 7), suffix: anchoredThread('suffix', block, 6, 8), label: anchoredThread('label', block, 1, 3) }
  model.contentControls = [{ id: 'body-control', type: 'richText', blockId: block.id, startOffset: 4, endOffset: 8, lock: 'unlocked' }]
  const next = executeDocumentCommand(model, { anchor: { blockId: block.id, offset: 8 }, focus: { blockId: block.id, offset: 2 } }, { id: 'doc.command.change-list-type', listType: 'ORDER_LIST' })
  assert.equal(block.runs.map(run => run.text).join(''), 'Body')
  assert.equal(block.type, 'numbered')
  assert.equal(block.metadata.officeNumberingMarker, undefined)
  assert.deepEqual(next, { anchor: { blockId: block.id, offset: 4 }, focus: { blockId: block.id, offset: 0 } })
  assert.deepEqual(model.comments.crossing.anchor, [{ blockId: block.id, startOffset: 0, endOffset: 3, text: 'Bod' }])
  assert.deepEqual(model.comments.suffix.anchor, [{ blockId: block.id, startOffset: 2, endOffset: 4, text: 'dy' }])
  assert.deepEqual(model.comments.label.anchor, [{ blockId: block.id, startOffset: 0, endOffset: 0, text: '' }])
  assert.equal(model.comments.label.messages[0].text, 'Keep this comment.')
  assert.equal(documentThreadCommentAnchorStatus(model, model.comments.label), 'current')
  assert.equal(model.contentControls[0].startOffset, 0)
  assert.equal(model.contentControls[0].endOffset, 4)
})

test('parser normalization preserves explicit point comment anchors in body and nested table paragraphs', () => {
  const body = { ...parsed('body', {}), content: 'Alpha', sourceRefs: [{ nodeId: 'source-body' }], details: { inlines: [{ type: 'text', text: 'Alpha' }] } }
  const nested = { ...parsed('nested-paragraph', {}), content: 'A😀B', sourceRefs: [{ nodeId: 'source-nested' }], details: { inlines: [{ type: 'text', text: 'A😀B' }] } }
  const table = { id: 'outer-table', type: 'table', children: [{ id: 'outer-cell', type: 'table_cell', details: {}, children: [
    { id: 'nested-table', type: 'table', children: [{ id: 'nested-cell', type: 'table_cell', details: {}, children: [nested] }] },
  ] }] }
  const comments = [
    { id: 'body-point', text: 'Deleted list label', targets: [{ nodeId: 'source-body', story: 'body', start: 0, end: 0 }] },
    { id: 'nested-point', text: 'Point after emoji', targets: [{ nodeId: 'source-nested', story: 'body', start: 3, end: 3 }] },
    { id: 'nested-range', text: 'Emoji range', targets: [{ nodeId: 'nested-paragraph', story: 'body', start: 1, end: 3 }] },
  ]
  const model = normalizeDocument(envelope([body, table], undefined, { comments }))
  assert.deepEqual(model.comments['docx-comment-body-point'].anchor, [{ blockId: 'body', startOffset: 0, endOffset: 0, text: '' }])
  assert.deepEqual(model.comments['docx-comment-nested-point'].anchor, [{ blockId: 'nested-paragraph', startOffset: 3, endOffset: 3, text: '' }])
  assert.deepEqual(model.comments['docx-comment-nested-range'].anchor, [{ blockId: 'nested-paragraph', startOffset: 1, endOffset: 3, text: '😀' }])
  assert.equal(model.comments['docx-comment-body-point'].referenceText, '')
  assert.equal(documentThreadCommentAnchorStatus(model, model.comments['docx-comment-body-point']), 'current')
  assert.equal(documentThreadCommentAnchorStatus(model, model.comments['docx-comment-nested-point']), 'current')
})

test('parser normalization rejects unknown, reversed, implicit, fractional and split-surrogate comment anchors', () => {
  const paragraph = { ...parsed('body', {}), content: 'A😀B', details: { inlines: [{ type: 'text', text: 'A😀B' }] } }
  const invalid = [
    { nodeId: 'body', start: 2, end: 2 }, { nodeId: 'body', start: 1, end: 2 },
    { nodeId: 'body', start: 3, end: 1 }, { nodeId: 'body', start: -1, end: 0 },
    { nodeId: 'body', start: 0.5, end: 1 }, { nodeId: 'body', start: 0 },
    { nodeId: 'body', end: 4 }, { nodeId: 'body', start: 5, end: 5 },
    { nodeId: 'missing', start: 0, end: 0 }, { nodeId: 'body', story: 'header', start: 0, end: 0 },
  ]
  const comments = invalid.map((target, index) => ({ id: String(index), text: 'Preserved unanchored message', targets: [target] }))
  const model = normalizeDocument(envelope([paragraph], undefined, { comments }))
  for (const comment of Object.values(model.comments)) {
    assert.deepEqual(comment.anchor, [])
    assert.equal(comment.messages[0].text, 'Preserved unanchored message')
    assert.equal(documentThreadCommentAnchorStatus(model, comment), 'orphaned')
  }
})

test('a locked later content control rejects a multi-paragraph label replacement before the first paragraph changes', () => {
  const first = withCachedMarker({ ...createParagraph('1. First'), id: 'cached-first' }, '1. ')
  const later = withCachedMarker({ ...createParagraph('2. Later'), id: 'cached-later' }, '2. ')
  const model = documentWith([first, later])
  model.comments = { first: anchoredThread('first', first, 3, 8) }
  model.contentControls = [
    { id: 'movable', type: 'richText', blockId: first.id, startOffset: 3, endOffset: 8, lock: 'unlocked' },
    { id: 'locked-later', type: 'richText', blockId: later.id, startOffset: 1, endOffset: 7, lock: 'contentLocked' },
  ]
  const before = JSON.stringify(model), originalRuns = first.runs, originalControls = model.contentControls
  assert.throws(() => executeDocumentCommand(model, { anchor: { blockId: first.id, offset: 3 }, focus: { blockId: later.id, offset: 7 } }, { id: 'doc.command.change-list-type', listType: 'ORDER_LIST' }), /locked/i)
  assert.equal(JSON.stringify(model), before)
  assert.equal(first.runs, originalRuns)
  assert.equal(model.contentControls, originalControls)
})

test('empty body note references and their identities remain after cached-label replacement', () => {
  const block = { ...createParagraph('4. Body'), id: 'body-note' }, run = block.runs[0]
  const noteRun = { ...run, id: 'footnote-reference', text: '', metadata: { note: { type: 'footnote', noteId: 'n1' } } }
  block.runs = [{ ...run, id: 'label', text: '4. ' }, noteRun, { ...run, id: 'body', text: 'Body' }]
  withCachedMarker(block, '4. ')
  const model = documentWith([block])
  model.footnotes = { n1: { id: 'n1', text: 'Note body', referenceBlockId: block.id, referenceRunId: noteRun.id } }
  const next = executeDocumentCommand(model, { anchor: { blockId: block.id, offset: 5 }, focus: { blockId: block.id, offset: 5 } }, { id: 'doc.command.change-list-type', listType: 'ORDER_LIST' })
  assert.deepEqual(block.runs.map(run => [run.id, run.text]), [[noteRun.id, ''], ['body', 'Body']])
  assert.deepEqual(block.runs[0].metadata.note, { type: 'footnote', noteId: 'n1' })
  assert.equal(model.footnotes.n1.referenceRunId, noteRun.id)
  assert.equal(next.focus.offset, 2)
})

test('explicit replacement retains a cached prefix edited into ordinary text instead of deleting the edited text', () => {
  const block = withCachedMarker({ ...createParagraph('[0] Item'), id: 'edited-label' }, '[0] ')
  block.runs[0].text = '[review] Item'
  const model = documentWith([block])
  const selection = { anchor: { blockId: block.id, offset: 12 }, focus: { blockId: block.id, offset: 12 } }
  assert.deepEqual(executeDocumentCommand(model, selection, { id: 'doc.command.change-list-type', listType: 'ORDER_LIST' }), selection)
  assert.equal(block.runs.map(run => run.text).join(''), '[review] Item')
  assert.equal(block.metadata.officeNumberingMarker, undefined)
})
