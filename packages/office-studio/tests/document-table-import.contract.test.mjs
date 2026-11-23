import assert from 'node:assert/strict'
import { test } from 'vitest'
import { normalizeDocument } from '../src/parser/documentNormalizer.ts'
import { documentTableGrid } from '../src/features/document/model/document-table-grid.ts'
import { tableForBlock, textTargets } from '../src/features/document/model/document-model.ts'

// Parser DTO -> domain drafts. Real DOCX, Canvas, IME and performance validation
// are separate and deferred to the user's unified validation phase.
const wordNamespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const property = tag => ({ version: 1, propertiesXml: `<w:${tag}/>`, namespaces: { w: wordNamespace } })
const paragraph = (id, text = '') => ({ id, type: 'paragraph', content: text,
  details: { kind: 'paragraph', inlines: [{ type: 'text', text }] }, children: [] })
const cell = (id, row, column, text = '', details = {}) => ({ id, type: 'table_cell', content: text,
  details: { rowIndex: row, columnIndex: column, rowSpan: 1, columnSpan: 1, ...details }, children: [paragraph(`${id}-paragraph`, text)] })
function rawTable() {
  return { id: 'native-table', type: 'table', content: 'A\tB\n\tC', details: {
    rowCount: 2, columnCount: 2, style: { columnWidthsTwips: [900, 1200], alignment: 'right', layout: 'fixed' },
    rowStyles: [{ heightTwips: 450, heightRule: 'exact', cantSplit: true }, {}],
    nativeTableProperties: { version: 1, nativeProperties: property('tblPr'), rows: [
      { nativeProperties: property('trPr'), cells: [0, 1].map(columnIndex => ({ columnIndex, nativeProperties: property('tcPr') })) },
      { nativeProperties: property('trPr'), cells: [0, 1].map(columnIndex => ({ columnIndex, nativeProperties: property('tcPr') })) },
    ] },
  }, children: [cell('a', 0, 0, 'A', { rowSpan: 2 }), cell('b', 0, 1, 'B'),
    cell('a-continued', 1, 0, '', { style: { vMergeContinuation: true } }), cell('c', 1, 1, 'C')] }
}
function envelope(table) {
  return { schemaVersion: 1, parserVersion: 'draft-contract', format: 'docx', fileName: 'table.docx', size: 128,
    model: 'document', diagnostics: {}, payload: { modelType: 'document', data: { documentId: 'document', blocks: [table] } } }
}
function authoredSource(nativeMatches = true, topologyMatches = true) {
  return { version: 1, tableId: 'authored-table', nativeMatches, topologyMatches, data: {
    columnWidths: [60.025, 79.975], headerRowCount: 1, titleRow: true,
    columnTypes: { 1: { type: 'dropdown', dropdownOptions: [{ id: 'chosen', label: 'B', color: '#ff0000' }] } },
    style: { align: 'center', layout: 'auto', fill: '#fff2cc', customSetting: 'table' },
    rows: [
      { id: 'authored-first', height: 30.025, heightRule: 'exact', allowSplit: false,
        style: { heightTwips: 451, isHeader: true, customSetting: 'row' }, cells: [
          { id: 'authored-a', columnIndex: 0, rowSpan: 2, colSpan: 1, style: { fill: '#ff0000', customSetting: 'a' } },
          { id: 'authored-b', columnIndex: 1, rowSpan: 1, colSpan: 1, style: { fill: '#00ff00', selectedDropdownOptionId: 'chosen' } },
        ] },
      { id: 'authored-last', heightRule: 'auto', allowSplit: true, cells: [
        { id: 'authored-c', columnIndex: 1, rowSpan: 1, colSpan: 1, style: { verticalAlign: 'bottom' } },
      ] },
    ],
  } }
}
const read = raw => {
  const model = normalizeDocument(envelope(raw)), block = model.blocks[0]
  return { model, block, table: tableForBlock(block), grid: documentTableGrid(tableForBlock(block)) }
}

test('logical columns differ from stored cell indices and normalization keeps the parser input unchanged', () => {
  const raw = rawTable(), before = structuredClone(raw), { model, table, grid } = read(raw)
  assert.deepEqual(raw, before)
  assert.equal(model.modified, false)
  assert.equal(grid.cellAt(1, 0).cell.id, 'a')
  assert.equal(table.rows[1].cells.length, 1)
  assert.equal(table.rows[1].cells[0].columnIndex, 1)
  assert.deepEqual(table.columnWidths, [60, 80])
  assert.equal(table.rows[0].height, 30)
  assert.equal(table.rows[0].heightRule, 'exact')
  assert.equal(table.rows[0].allowSplit, false)
  const target = textTargets(model).find(target => target.cellBlockId === 'c-paragraph')
  assert.equal(target.col, 0)
  assert.equal(target.row, 1)
})

test('fully covered rows remain actual empty anchor rows and missing legacy column indices are inferred', () => {
  const raw = rawTable()
  raw.children = [cell('span', 0, 0, 'Full', { rowSpan: 2, columnSpan: 2 })]
  delete raw.children[0].details.columnIndex
  const { table, grid } = read(raw)
  assert.equal(table.rows.length, 2)
  assert.deepEqual(table.rows[1].cells, [])
  assert.equal(table.rows[0].cells[0].columnIndex, 0)
  assert.equal(grid.cellAt(1, 1).cell.id, 'span')
})

test('Rust canonical horizontal members retain all body targets and member provenance without using vertical offsets', () => {
  const raw = rawTable(), source = { version: 1, members: [
    { cellId: 'left', columnIndex: 0, columnSpan: 1, blockIds: ['left-text'], nativeProperties: property('tcPr') },
    { cellId: 'right', columnIndex: 1, columnSpan: 1, blockIds: ['right-text'], nativeProperties: property('tcPr') },
  ] }
  raw.children = [cell('combined', 0, 0, 'A\nB😀', { rowSpan: 2, columnSpan: 2 })]
  raw.children[0].children = [paragraph('left-text', 'A'), paragraph('right-text', 'B😀')]
  raw.details.nativeTableProperties.rows[0].cells = [{ columnIndex: 0, columnSpan: 2,
    nativeProperties: { ...property('tcPr'), sourceHorizontalMerge: source, rewriteProperties: ['tcW', 'tcBorders'] } }]
  raw.details.nativeTableProperties.rows[1].cells = [{ columnIndex: 0, columnSpan: 2, nativeProperties: property('tcPr') }]
  const { model, table, grid } = read(raw), owner = table.rows[0].cells[0]
  assert.equal(grid.cellAt(1, 1).cell.id, 'combined')
  assert.deepEqual(owner.blocks.map(block => block.id), ['left-text', 'right-text'])
  assert.equal(textTargets(model).filter(target => ['left-text', 'right-text'].includes(target.cellBlockId)).length, 2)
  assert.deepEqual(owner.style.nativeProperties.rewriteProperties, ['tcW', 'tcBorders'])
  assert.deepEqual(owner.style.nativeHorizontalMergeSource.rows[0], { rowId: table.rows[0].id, members: source.members })
  assert.equal(owner.style.nativeContinuationContent, undefined)
})

test('matching native settings restore precise authored sizes IDs header metadata and dropdown option identity', () => {
  const raw = rawTable(); raw.details.officeTable = authoredSource()
  const { block, table } = read(raw)
  assert.equal(block.id, 'authored-table')
  assert.equal(table.rows[0].id, 'authored-first')
  assert.equal(table.rows[1].id, 'authored-last')
  assert.equal(table.rows[0].cells[0].id, 'authored-a')
  assert.deepEqual(table.columnWidths, [60.025, 79.975])
  assert.equal(table.rows[0].height, 30.025)
  assert.equal(table.headerRowCount, 1)
  assert.equal(table.titleRow, true)
  assert.equal(table.rows[0].cells[1].style.selectedDropdownOptionId, 'chosen')
  assert.equal(table.columnTypes[1].dropdownOptions[0].id, 'chosen')
  const continuation = table.rows[0].cells[0].style.nativeContinuations[1]
  assert.equal(continuation.rowId, 'authored-last')
  assert.equal(continuation.ownerCellId, 'authored-a')
  assert.equal(continuation.colSpan, 1)
})

test('external native geometry replaces authored geometry while non-native author settings remain', () => {
  const raw = rawTable(); raw.details.officeTable = authoredSource(false)
  const { table } = read(raw)
  assert.deepEqual(table.columnWidths, [60, 80])
  assert.equal(table.style.align, 'right')
  assert.equal(table.style.customSetting, 'table')
  assert.equal(table.style.fill, undefined)
  assert.equal(table.rows[0].height, 30)
  assert.equal(table.rows[0].style.heightTwips, 450)
  assert.equal(table.rows[0].style.isHeader, undefined)
  assert.equal(table.rows[0].style.customSetting, 'row')
  assert.equal(table.rows[0].cells[0].style.fill, undefined)
  assert.equal(table.rows[0].cells[0].style.customSetting, 'a')
  assert.equal(table.rows[0].cells[1].style.selectedDropdownOptionId, 'chosen')
})

test('changed topology retains stale source without assigning its cells widths or column types to the new grid', () => {
  const raw = rawTable(); raw.details.officeTable = authoredSource(false, false)
  const { block, table } = read(raw)
  assert.equal(block.id, 'authored-table')
  assert.equal(table.rows[0].cells[0].id, 'a')
  assert.equal(table.columnTypes, undefined)
  assert.equal(table.titleRow, undefined)
  assert.deepEqual(table.columnWidths, [60, 80])
  assert.deepEqual(block.metadata.officeTableStaleSource, raw.details.officeTable)
})

test('synthetic DOCX cell tails are removed without removing an authored empty paragraph', () => {
  const raw = rawTable(), empty = raw.children[1]
  empty.content = ''; empty.children = [paragraph('authored-empty')]
  empty.children.push(paragraph('native-tail'))
  raw.details.nativeTableProperties.rows[0].cells[1].syntheticTail = true
  const { table } = read(raw)
  assert.deepEqual(table.rows[0].cells[1].blocks.map(block => block.id), ['authored-empty'])
})

test('independent native continuation text becomes current domain blocks with stable row and owner identities', () => {
  const raw = rawTable(); raw.details.officeTable = authoredSource()
  raw.children[2].content = 'Continuation😀'
  raw.children[2].children = [paragraph('continuation-paragraph', 'Continuation😀')]
  const { table } = read(raw), source = table.rows[0].cells[0].style.nativeContinuationContent[1]
  assert.equal(source.version, 1)
  assert.equal(source.rowId, 'authored-last')
  assert.equal(source.ownerCellId, 'authored-a')
  assert.equal(source.colSpan, 1)
  assert.equal(source.blocks[0].id, 'continuation-paragraph')
  assert.equal(source.blocks[0].runs.map(run => run.text).join(''), 'Continuation😀')
  assert.deepEqual(source.source, raw.children[2])
})

test('native text cleared externally cannot be resurrected by an old authored continuation mapping', () => {
  const raw = rawTable(); raw.details.officeTable = authoredSource()
  raw.details.officeTable.data.rows[0].cells[0].style.nativeContinuationContent = { 1: {
    version: 1, ownerCellId: 'authored-a', rowId: 'authored-last', colSpan: 1,
    blocks: [{ id: 'stale-p', type: 'paragraph', runs: [{ text: 'Stale text' }], metadata: {} }],
  } }
  const { table } = read(raw)
  assert.equal(table.rows[0].cells[0].style.nativeContinuationContent, undefined)
})

test('illegal overlaps holes spans budgets and malformed native source mappings fail normalization', () => {
  const cases = [
    raw => { raw.children[1].details.columnIndex = 0 },
    raw => { raw.children = [cell('hole', 0, 1, 'Hole'), cell('hole-2', 1, 1, 'Hole')] },
    raw => { raw.children[0].details.rowSpan = 3 },
    raw => { raw.details.rowCount = 10_001 },
    raw => { raw.details.style.columnWidthsTwips = [900, 0] },
    raw => { raw.details.nativeTableProperties.rows[0].cells.push({ columnIndex: 0, nativeProperties: property('tcPr') }) },
    raw => { delete raw.details.nativeTableProperties.nativeProperties.namespaces },
    raw => { raw.details.officeTable = authoredSource(true, false) },
    raw => { raw.details.officeTable = authoredSource(); raw.details.officeTable.data.rows[0].cells[0].colSpan = 2 },
  ]
  for (const mutate of cases) { const raw = rawTable(); mutate(raw); assert.throws(() => read(raw)) }
})
