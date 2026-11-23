import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import { createParagraph, createTable, runText, tableForBlock } from '../src/features/document/model/document-model.ts'
import { layoutDocument, caretRect, hitTestDocument } from '../src/features/document/layout/document-layout.ts'
import { measureDocumentTable, insertDocumentTablePageGap } from '../src/features/document/layout/document-table-measurement.ts'
import { hitTestDocumentTable, documentTableSelectionAtHit, documentTableFragmentRows } from '../src/features/document/layout/document-table-interaction.ts'

// DOC-06 layout drafts. No execution, device evidence or performance claim accompanies this file.
const measurer = { measure: text => Array.from(text).length * 8 }
const copy = value => JSON.parse(JSON.stringify(value))
const paragraph = (id, text) => {
  const block = { ...createParagraph(text, { fontSize: 12 }), id }
  Object.assign(block.style, { spaceBefore: 0, spaceAfter: 0, lineHeight: 20, spacingRule: 'exact', wordWrap: true })
  return block
}
function tableBlock(id, rows, columns) {
  const table = createTable(rows, columns)
  table.style = { layout: 'fixed', cellMargin: { top: 4, right: 4, bottom: 4, left: 4 } }
  table.columnWidths = Array.from({ length: columns }, () => 100)
  table.rows.forEach((row, r) => {
    row.id = `${id}-row-${r}`
    row.cells.forEach((cell, c) => { cell.id = `${id}-cell-${r}-${c}`; cell.blocks = [paragraph(`${id}-p-${r}-${c}`, `Row${r}Col${c}`)] })
  })
  return { ...paragraph(id, ''), type: 'table', runs: [], metadata: { table } }
}
function documentWith(blocks, page = {}) {
  return { ...createBlankDocument(), id: 'table-layout', blocks,
    page: { width: 300, height: 160, marginLeft: 20, marginRight: 20, marginTop: 20, marginBottom: 20, ...page } }
}
const rendered = lines => lines.flatMap(line => line.glyphs.map(glyph => glyph.text)).join('')

test('fixed column widths, table alignment and cell margins affect one shared layout geometry', () => {
  const block = tableBlock('table', 1, 2), table = tableForBlock(block)
  table.columnWidths = [70, 130]; table.style.width = 200; table.style.align = 'right'
  const model = documentWith([block]), original = copy(model), layout = layoutDocument(model, measurer)
  const first = layout.tableFragments.find(fragment => fragment.tableId === 'table')
  assert.deepEqual(first.columnWidths, [70, 130])
  assert.equal(first.width, 200)
  assert.equal(first.x + first.width, layout.pages[0].content.x + layout.pages[0].content.width)
  assert.equal(layout.cells[0].width, 70)
  assert.equal(layout.lines[0].x, first.x + 4)
  assert.deepEqual(model, original)
})

test('exact rows retain physical height and clip overflow while preserving all source text', () => {
  const block = tableBlock('table', 1, 1), table = tableForBlock(block), paragraph = table.rows[0].cells[0].blocks[0]
  paragraph.runs[0].text = 'LONG'.repeat(25)
  table.rows[0].height = 28; table.rows[0].heightRule = 'exact'
  const original = copy(block), measured = measureDocumentTable(block, 260, new Map(), measurer)
  assert.equal(measured.height, 28)
  assert.equal(measured.cells[0].height, 28)
  assert.ok(measured.lines.length > 1)
  assert.equal(rendered(measured.lines), runText(paragraph.runs))
  assert.equal(measured.lines.every(line => line.clip.y + line.clip.height <= 28), true)
  const layout = layoutDocument(documentWith([block]), measurer)
  assert.ok(layout.lines.length < measured.lines.length)
  assert.equal(layout.lines.every(line => line.clip.y + line.clip.height <= layout.cells[0].y + layout.cells[0].height), true)
  assert.deepEqual(block, original)
})

test('a spanning cell grows adjustable rows and keeps exact rows at their declared size', () => {
  const block = tableBlock('table', 3, 1), table = tableForBlock(block), owner = table.rows[0].cells[0]
  owner.rowSpan = 3; owner.blocks[0].runs[0].text = 'Many words '.repeat(10)
  table.rows[1].cells = []; table.rows[2].cells = []
  table.rows[0].height = 25; table.rows[0].heightRule = 'exact'
  table.rows[1].heightRule = 'auto'; table.rows[2].heightRule = 'atLeast'
  const measured = measureDocumentTable(block, 260, new Map(), measurer)
  assert.equal(measured.rowHeights[0], 25)
  assert.ok(measured.rowHeights[1] > 34); assert.ok(measured.rowHeights[2] > 34)
  assert.equal(measured.cells[0].height, measured.rowHeights.reduce((sum, height) => sum + height, 0))
})

test('nested tables retain their own text targets and take priority in table hit testing', () => {
  const outer = tableBlock('outer', 1, 1), inner = tableBlock('inner', 2, 2)
  tableForBlock(outer).columnWidths = [230]; tableForBlock(inner).columnWidths = [90, 90]
  tableForBlock(outer).rows[0].cells[0].blocks = [paragraph('outer-before', 'Before'), inner, paragraph('outer-after', 'After')]
  const model = documentWith([outer], { height: 300 }), original = copy(model), layout = layoutDocument(model, measurer)
  const nestedCell = layout.cells.find(cell => cell.tableId === 'inner' && cell.row === 0 && cell.columnIndex === 0)
  assert.equal(nestedCell.depth, 1)
  const hit = hitTestDocumentTable(layout, nestedCell.x + nestedCell.width / 2, nestedCell.y + nestedCell.height / 2)
  assert.equal(hit.fragment.tableId, 'inner')
  const selected = documentTableSelectionAtHit(model, hit, { anchor: { blockId: 'outer', row: 0, col: 0, cellBlockId: 'outer-before', offset: 0 }, focus: { blockId: 'outer', row: 0, col: 0, cellBlockId: 'outer-before', offset: 0 } })
  assert.equal(selected.tableSelection.tableId, 'inner')
  const line = layout.lines.find(line => line.target.cellBlockId === 'inner-p-0-0')
  assert.equal(line.target.blockId, 'inner')
  const caret = caretRect(layout, { ...line.target, offset: line.start })
  assert.equal(hitTestDocument(layout, caret.x + 1, caret.y + 1).cellBlockId, 'inner-p-0-0')
  assert.deepEqual(model, original)
})

test('repeated headers use the same model IDs on later pages and do not duplicate source content', () => {
  const block = tableBlock('table', 12, 2), table = tableForBlock(block)
  table.headerRowCount = 1
  table.rows.forEach(row => { row.height = 34; row.heightRule = 'exact'; row.allowSplit = false })
  const model = documentWith([block]), original = copy(model), layout = layoutDocument(model, measurer)
  assert.ok(layout.pages.length >= 3)
  const repeated = layout.tableFragments.filter(fragment => fragment.repeatedHeader)
  assert.ok(repeated.length >= 2)
  assert.equal(repeated.every(fragment => fragment.sourceStart === 0), true)
  assert.equal(layout.cells.filter(cell => cell.repeatedHeader).every(cell => cell.row === 0 && ['table-cell-0-0', 'table-cell-0-1'].includes(cell.cellId)), true)
  const originalText = table.rows.flatMap(row => row.cells).map(cell => runText(cell.blocks[0].runs)).join('')
  assert.equal(rendered(layout.lines.filter(line => !line.repeatedHeader)), originalText)
  const repeatedLine = layout.lines.find(line => line.repeatedHeader && line.page > 0)
  const point = { ...repeatedLine.target, offset: repeatedLine.start, pageAffinity: repeatedLine.page }
  const caret = caretRect(layout, point)
  assert.ok(caret.y >= layout.pages[repeatedLine.page].y)
  assert.deepEqual(model, original)
})

test('the measured repeated-header closure includes a rowSpan without rewriting imported header flags', () => {
  const block = tableBlock('table', 4, 2), table = tableForBlock(block), owner = table.rows[0].cells[0]
  owner.rowSpan = 2; table.rows[1].cells.splice(0, 1); table.headerRowCount = 1
  const original = copy(block), measured = measureDocumentTable(block, 260, new Map(), measurer)
  assert.equal(measured.headerRowCount, 2)
  assert.equal(measured.headerHeight, measured.rowStarts[2])
  assert.equal(table.headerRowCount, 1)
  assert.deepEqual(block, original)
})

test('a tall splitting row places whole visible lines on pages and keeps all text addressable', () => {
  const block = tableBlock('table', 1, 1), table = tableForBlock(block)
  table.rows[0].allowSplit = true
  table.rows[0].cells[0].blocks[0].runs[0].text = 'ABCDEFGH'.repeat(30)
  const model = documentWith([block]), layout = layoutDocument(model, measurer)
  assert.ok(layout.pages.length > 1)
  assert.equal(rendered(layout.lines), runText(table.rows[0].cells[0].blocks[0].runs))
  assert.equal(layout.lines.every(line => line.clip.y <= line.y && line.clip.y + line.clip.height >= line.y + line.height - 1e-6), true)
  const fragments = layout.tableFragments.filter(fragment => !fragment.repeatedHeader)
  assert.equal(fragments[0].continuedAfter, true)
  assert.equal(fragments.at(-1).continuedBefore, true)
  assert.equal(fragments.every(fragment => documentTableFragmentRows(fragment).every(row => row.contentHeight > 0)), true)
})

test('pagination gaps change derived row geometry while resize retains the measured content height', () => {
  const block = tableBlock('table', 1, 1), table = tableForBlock(block)
  table.rows[0].cells[0].blocks[0].runs[0].text = 'LONG'.repeat(20)
  const original = copy(block), measured = measureDocumentTable(block, 260, new Map(), measurer), row = measured.tables[0].rows[0]
  const height = row.contentHeight, oldHeight = measured.height, cut = measured.lines[1].y
  insertDocumentTablePageGap(measured, cut, 12)
  assert.equal(measured.height, oldHeight + 12)
  assert.equal(row.contentHeight, height)
  assert.equal(row.height, height + 12)
  assert.deepEqual(block, original)
})
