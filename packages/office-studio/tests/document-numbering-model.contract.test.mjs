import assert from 'node:assert/strict'
import { test } from 'vitest'
import { normalizeDocument } from '../src/parser/documentNormalizer.ts'
import { createParagraph, createTable } from '../src/features/document/model/document-model.ts'
import { blockListLevel, formatListOrdinal, paragraphListMarkers } from '../src/features/document/model/document-lists.ts'
import { blockHasNumbering, numberingDefinitionFromParts } from '../src/features/document/model/document-numbering.ts'

function envelope(blocks, numberingDefinitions) {
  return {
    schemaVersion: 1, parserVersion: 'contract', format: 'docx', fileName: 'format.docx', size: 100,
    model: 'document', payload: { modelType: 'document', data: { documentId: 'format-doc', blocks, numberingDefinitions } },
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
