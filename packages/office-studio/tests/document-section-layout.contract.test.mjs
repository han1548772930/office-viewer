import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import { createParagraph, createTable } from '../src/features/document/model/document-model.ts'
import { applyDocumentSectionSetupPatch, documentEffectiveSectionSetup } from '../src/features/document/services/document-settings.ts'
import { applySectionHeaderFooter } from '../src/features/document/services/section-header-footer.ts'
import { caretRect, documentStoryLayoutLines, hitTestDocument, layoutDocument, PAGE_GAP, PAGE_INSET } from '../src/features/document/layout/document-layout.ts'
import { DocumentSectionFlow } from '../src/features/document/layout/document-section-flow.ts'
import { sceneFromDocument, sceneOffset } from '../src/features/document/rendering/document-scene.ts'

// Deferred drafts, 2026-11-22: do not treat these assertions as executed evidence.
const measurer = { measure: text => Array.from(text).length * 10 }
const copy = value => JSON.parse(JSON.stringify(value))
function paragraph(text, options = {}) {
  const block = createParagraph(text, { fontSize: 12 })
  Object.assign(block.style, { spaceBefore: 0, spaceAfter: 0, lineHeight: 20, spacingRule: 'exact', wordWrap: true }, options.style)
  if (options.type) block.type = options.type
  return block
}
const marker = type => ({ ...paragraph(''), type, runs: [] })
const section = (type, setup = {}) => ({ ...marker('sectionBreak'), sectionBreak: { type, setup } })
function documentWith(blocks, setup = {}, page = {}) {
  return { ...createBlankDocument(), id: 'section-layout-doc', blocks,
    page: { width: 240, height: 140, marginLeft: 20, marginRight: 20, marginTop: 20, marginBottom: 20, ...page }, sectionSetup: setup }
}
const renderedText = lines => lines.flatMap(line => line.glyphs.map(glyph => glyph.text)).join('')

test('automatic overflow fills columns in reading order before creating another physical page', () => {
  const block = paragraph('ABCDEFGHI'.repeat(11)), model = documentWith([block], { columnCount: 2, columnGap: 20 })
  const before = copy(model), layout = layoutDocument(model, measurer)
  assert.equal(layout.pages.length, 2)
  assert.deepEqual(layout.lines.map(line => [line.page, line.column]), [
    ...Array.from({ length: 5 }, () => [0, 0]), ...Array.from({ length: 5 }, () => [0, 1]), [1, 0],
  ])
  assert.equal(renderedText(layout.lines), block.runs[0].text)
  assert.deepEqual(layout.pages[0].sections[0].columns.map(column => column.width), [90, 90])
  assert.deepEqual(layout.lines.slice(0, 5).map(line => line.y), [48, 68, 88, 108, 128])
  assert.equal(layout.lines[5].x, layout.pages[0].content.x + 110)
  assert.deepEqual(layout.pageContent.map(page => page.lines.length), [10, 1])
  assert.deepEqual(model, before)
})

test('explicit column and page boundaries have distinct effects in a multi-column section', () => {
  const blocks = [paragraph('Left'), marker('columnBreak'), paragraph('Right'), marker('columnBreak'), paragraph('Next'), marker('pageBreak'), paragraph('Page')]
  const layout = layoutDocument(documentWith(blocks, { columnCount: 2, columnGap: 20 }), measurer)
  assert.deepEqual(layout.lines.map(line => [line.page, line.column]), [[0, 0], [0, 1], [1, 0], [2, 0]])
  assert.equal(layout.pages[0].sections[0].hasExplicitColumnBreak, true)
  assert.equal(layout.pages[2].pageNumber, 3)
})

test('NEXT_COLUMN shares a compatible page, retains the section top and skips preceding columns', () => {
  const first = paragraph('Top'), continuous = section('continuous', { columnCount: 2, columnGap: 20 })
  const middle = paragraph('First column'), next = section('nextColumn', { columnCount: 2, columnGap: 20 }), last = paragraph('Second column')
  const layout = layoutDocument(documentWith([first, continuous, middle, next, last]), measurer)
  assert.equal(layout.pages.length, 1)
  assert.equal(layout.pages[0].sectionId, 'first')
  assert.deepEqual(layout.pages[0].sections.map(part => part.sectionId), ['first', continuous.id, next.id])
  const previousPart = layout.pages[0].sections[1], nextPart = layout.pages[0].sections[2]
  assert.equal(nextPart.y, previousPart.y)
  assert.equal(nextPart.columns[0].full, true)
  assert.equal(nextPart.columns[0].usedHeight, 0)
  const line = layout.lines.find(line => line.target.blockId === last.id)
  assert.equal(line.column, 1)
  assert.equal(line.y, previousPart.y)
  assert.equal(line.x, nextPart.columns[1].x)
})

test('NEXT_COLUMN creates a new page when the last column is exhausted or page geometry differs', () => {
  for (const setup of [{ columnCount: 2, columnGap: 20 }, { columnCount: 2, columnGap: 20, marginLeft: 30 }, { width: 300, height: 160, columnCount: 3 }]) {
    const next = section('nextColumn', setup), after = paragraph('After')
    const layout = layoutDocument(documentWith([paragraph('Left'), marker('columnBreak'), paragraph('Right'), next, after], { columnCount: 2, columnGap: 20 }), measurer)
    assert.equal(layout.pages.length, 2)
    assert.equal(layout.lines.at(-1).page, 1)
    assert.equal(layout.lines.at(-1).column, 0)
    assert.equal(layout.pages[1].sectionId, next.id)
  }
})

test('continuous changes preserve the physical page and rewrap the suffix after margins change on overflow', () => {
  const boundary = section('continuous', { marginLeft: 60, marginRight: 60 })
  const block = paragraph('中😀'.repeat(100), { style: { indentStart: 10, indentFirstLine: 20 } })
  const model = documentWith([paragraph('Before'), boundary, block]), layout = layoutDocument(model, measurer)
  assert.equal(layout.pages[0].sections[1].sectionId, boundary.id)
  assert.equal(layout.pages[0].content.width, 200)
  assert.equal(layout.pages[1].content.width, 120)
  const lines = layout.lines.filter(line => line.target.blockId === block.id)
  assert.equal(lines[0].x, layout.pages[0].content.x + 30)
  const overflow = lines.find(line => line.page === 1)
  assert.equal(overflow.x, layout.pages[1].content.x + 10)
  assert.ok(overflow.glyphs.length <= 11)
  assert.equal(renderedText(lines), block.runs[0].text)
  const boundaries = new Set([0, block.runs[0].text.length, ...Array.from(new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(block.runs[0].text), part => part.index)])
  for (const line of lines) {
    assert.ok(boundaries.has(line.start)); assert.ok(boundaries.has(line.end))
    for (const glyph of line.glyphs) { assert.ok(boundaries.has(glyph.offset)); assert.ok(boundaries.has(glyph.end)) }
    assert.ok(line.x + line.width <= layout.pages[line.page].content.x + layout.pages[line.page].content.width + 1e-6)
  }
  const caret = caretRect(layout, { blockId: block.id, offset: overflow.glyphs[2].offset })
  assert.equal(hitTestDocument(layout, caret.x + 1, caret.y + 2).offset, overflow.glyphs[2].offset)
})

test('continuous column geometry changes balance complete text and retain the final paragraph spacing', () => {
  const before = paragraph('A'.repeat(54), { style: { spaceAfter: 6 } }), boundary = section('continuous', { columnCount: 1 }), after = paragraph('Following')
  const layout = layoutDocument(documentWith([before, boundary, after], { columnCount: 2, columnGap: 20 }, { height: 220 }), measurer)
  const lines = layout.lines.filter(line => line.target.blockId === before.id), part = layout.pages[0].sections[0]
  assert.deepEqual(lines.map(line => line.column), [0, 0, 0, 1, 1, 1])
  assert.deepEqual(part.columns.map(column => column.usedHeight), [60, 66])
  assert.equal(part.usedHeight, 66)
  assert.equal(layout.lines.at(-1).y, layout.pages[0].content.y + 66)
  assert.equal(renderedText(lines), before.runs[0].text)
})

test('continuous sections keep reading-order columns when geometry is unchanged and start a new page when full', () => {
  const before = paragraph('A'.repeat(54)), next = section('continuous', { columnCount: 2, columnGap: 20 }), after = paragraph('Following')
  const roomy = layoutDocument(documentWith([before, next, after], { columnCount: 2, columnGap: 20 }, { height: 220 }), measurer)
  assert.deepEqual(roomy.lines.filter(line => line.target.blockId === before.id).map(line => line.column), [0, 0, 0, 0, 0, 0])
  assert.equal(roomy.lines.at(-1).y, roomy.pages[0].content.y + 120)
  const full = layoutDocument(documentWith([before, next, after], { columnCount: 2, columnGap: 20 }), measurer)
  assert.equal(full.pages.length, 2)
  assert.equal(full.lines.at(-1).page, 1)
})

test('tables remain indivisible rows and use the active section column width after reflow', () => {
  const table = { ...paragraph(''), type: 'table', runs: [], metadata: { table: createTable(1, 2) } }
  const next = section('continuous', { marginLeft: 60, marginRight: 60 })
  table.metadata.table.rows[0].cells[0].blocks = [paragraph('A'.repeat(40))]
  table.metadata.table.rows[0].cells[1].blocks = [paragraph('B'.repeat(40))]
  const layout = layoutDocument(documentWith([paragraph('Before'), next, table]), measurer)
  assert.equal(layout.cells.length, 2)
  assert.equal(layout.cells[0].page, 1)
  assert.equal(layout.cells[0].width + layout.cells[1].width, 120)
  assert.equal(layout.cells[0].height, layout.cells[1].height)
  assert.ok(layout.cells[0].height > layout.pages[1].content.height)
  assert.equal(layout.pages.length, 2)
})

test('mixed physical page sizes are centered and share scene, caret and hit-test transforms', () => {
  const first = paragraph('A'), next = section('continuous', { width: 420, height: 240, orientation: 'landscape', marginLeft: 30 }), after = paragraph('Z')
  const model = documentWith([first, next, after]), layout = layoutDocument(model, measurer)
  assert.equal(layout.width, 420 + PAGE_INSET * 2)
  assert.equal(layout.pages[0].x, PAGE_INSET + 90)
  assert.equal(layout.pages[1].x, PAGE_INSET)
  assert.equal(layout.pages[1].y, layout.pages[0].y + 140 + PAGE_GAP)
  assert.equal(layout.height, layout.pages[1].y + 240 + PAGE_INSET)
  assert.equal(layout.pages[1].orientation, 'landscape')
  const point = { blockId: first.id, offset: 0 }, caret = caretRect(layout, point)
  assert.deepEqual(hitTestDocument(layout, caret.x + 1, caret.y + 2), point)
  const viewport = { width: 800, height: 800, scrollLeft: 12, scrollTop: 10, zoom: 150 }, offset = sceneOffset(layout, viewport)
  const scene = sceneFromDocument(model, layout, viewport, { selection: { anchor: point, focus: point }, focused: false, caretVisible: false, composing: false, images: new Map() })
  const command = scene.commands.find(command => command.kind === 'text' && command.text === 'A')
  assert.equal(command.x, offset.x + layout.lines[0].glyphs[0].x * 1.5)
  assert.equal(command.clip.x, offset.x + layout.pages[0].x * 1.5)
})

test('page restart and continuation drive PAGE fields while NUMPAGES counts physical pages', () => {
  const second = section('nextPage', { pageNumberStart: 2 }), third = section('nextPage')
  const model = documentWith([paragraph('A'), marker('pageBreak'), paragraph('B'), second, paragraph('C'), marker('pageBreak'), paragraph('D'), third, paragraph('E')], { pageNumberStart: 4 })
  applySectionHeaderFooter(model, 'first', { header: 'P', footer: 'T', pageFields: [{ story: 'header', field: 'PAGE', offset: 1 }, { story: 'footer', field: 'NUMPAGES', offset: 1 }] })
  const layout = layoutDocument(model, measurer)
  assert.deepEqual(layout.pages.map(page => page.pageNumber), [4, 5, 2, 3, 4])
  assert.deepEqual(layout.pageHeaderFooters.map(story => story.header), ['P4', 'P5', 'P2', 'P3', 'P4'])
  assert.deepEqual(layout.pageHeaderFooters.map(story => story.footer), ['T5', 'T5', 'T5', 'T5', 'T5'])
})

test('continuous restarts keep the shared page label and propagate the latest restart to subsequent pages', () => {
  const restart = section('continuous', { pageNumberStart: 7 }), continuation = section('continuous')
  const model = documentWith([paragraph('A'), restart, paragraph('B'), continuation, paragraph('C'), marker('pageBreak'), paragraph('D')], {}, { height: 240 })
  const layout = layoutDocument(model, measurer)
  assert.deepEqual(layout.pages.map(page => [page.pageNumber, page.pageNumberStart]), [[1, 1], [8, 7]])
  assert.equal(layout.pages[0].sectionId, 'first')
  assert.equal(layout.pages[1].sectionId, continuation.id)
  assert.equal(layout.pages[1].sectionPageIndex, 1)
})

test('first-page and even-page stories select by section position and logical numbering', () => {
  const model = documentWith([paragraph('A'), marker('pageBreak'), paragraph('B'), marker('pageBreak'), paragraph('C')], { pageNumberStart: 4 })
  applySectionHeaderFooter(model, 'first', { header: 'Odd', firstPageHeader: 'First', evenPageHeader: 'Even', differentFirstPage: true, differentOddEven: true })
  const layout = layoutDocument(model, measurer)
  assert.deepEqual(layout.pageHeaderFooters.map(story => [story.headerStory, story.header]), [['firstPageHeader', 'First'], ['header', 'Odd'], ['evenPageHeader', 'Even']])
})

test('parity filler pages contain no body, header/footer, caret or editable story', () => {
  const next = section('evenPage', { pageNumberStart: 3 }), model = documentWith([paragraph('A'), next, paragraph('B')])
  applySectionHeaderFooter(model, 'first', { header: 'Odd', evenPageHeader: 'Even', differentOddEven: true })
  const layout = layoutDocument(model, measurer)
  assert.deepEqual(layout.pages.map(page => [page.pageNumber, page.isParityBlank]), [[1, false], [3, true], [4, false]])
  assert.equal(layout.pages[1].sections.length, 0)
  assert.deepEqual(layout.pageContent[1], { lines: [], cells: [], images: [], dividers: [], columnSeparators: [] })
  assert.equal(layout.pageHeaderFooters[1].header, '')
  assert.deepEqual(documentStoryLayoutLines(layout, { page: 1, area: 'header' }), [])
  assert.equal(hitTestDocument(layout, layout.pages[1].content.x, layout.pages[1].content.y), undefined)
  assert.equal(layout.pageHeaderFooters[2].header, 'Even')
})

test('an explicit page boundary is reused by a section without suppressing required parity filler', () => {
  const next = section('nextPage', { pageNumberStart: 5 }), model = documentWith([paragraph('A'), marker('pageBreak'), next, paragraph('B')])
  applySectionHeaderFooter(model, 'first', { differentOddEven: true })
  const layout = layoutDocument(model, measurer)
  assert.equal(layout.pages.length, 3)
  assert.equal(layout.pages[1].isParityBlank, true)
  assert.equal(layout.pages[2].pageNumber, 5)
  const ordinary = layoutDocument(documentWith([paragraph('A'), marker('pageBreak'), section('nextPage'), paragraph('B')]), measurer)
  assert.equal(ordinary.pages.length, 2)
  const consecutive = layoutDocument(documentWith([paragraph('A'), marker('pageBreak'), marker('pageBreak'), paragraph('B')]), measurer)
  assert.equal(consecutive.pages.length, 3)
})

test('column separators use shared Canvas geometry and clip to the physical page', () => {
  const block = paragraph('A'.repeat(60)), model = documentWith([block], { columnCount: 2, columnGap: 20, columnSeparator: 'line' })
  const layout = layoutDocument(model, measurer), separator = layout.columnSeparators[0], page = layout.pages[0]
  assert.equal(separator.x, page.content.x + 100)
  assert.equal(separator.y, page.content.y)
  assert.equal(separator.height, 100)
  assert.equal(layout.pageContent[0].columnSeparators[0], separator)
  const point = { blockId: block.id, offset: 0 }, viewport = { width: layout.width, height: layout.height, scrollLeft: 0, scrollTop: 0, zoom: 100 }
  const scene = sceneFromDocument(model, layout, viewport, { selection: { anchor: point, focus: point }, focused: false, caretVisible: false, composing: false, images: new Map() })
  const command = scene.commands.find(command => command.kind === 'line' && command.x1 === separator.x && command.x2 === separator.x && command.y1 === separator.y)
  assert.equal(command.y2, separator.y + separator.height)
  assert.equal(command.clip.x, page.x)
})

test('effective section geometry rejects invalid combinations before publishing any model change', () => {
  const next = section('nextPage', { columnCount: 2, columnGap: 150 }), model = documentWith([paragraph('A'), next, paragraph('B')])
  const before = copy(model)
  assert.throws(() => applyDocumentSectionSetupPatch(model, 'first', { width: 180 }), /positive content width/)
  assert.deepEqual(model, before)
  assert.throws(() => applyDocumentSectionSetupPatch(model, next.id, { marginTop: 70, marginBottom: 70 }), /positive finite content area/)
  assert.deepEqual(model, before)
  assert.throws(() => applyDocumentSectionSetupPatch(model, 'first', { unsupported: 1 }), /Unsupported section/)
  assert.deepEqual(model, before)
  assert.equal(documentEffectiveSectionSetup(model, next.id).pageNumberStart, undefined)
  const legacy = documentWith([paragraph('A')], { width: 280, marginLeft: 30 })
  applyDocumentSectionSetupPatch(legacy, 'first', { width: 260, marginLeft: 25 })
  assert.equal(documentEffectiveSectionSetup(legacy, 'first').page.width, 260)
  assert.equal(documentEffectiveSectionSetup(legacy, 'first').page.marginLeft, 25)
})

test('an oversized indivisible layout item is placed once in a fresh full-height column', () => {
  const flow = new DocumentSectionFlow(documentWith([paragraph('A')], { columnCount: 2, columnGap: 20 }))
  flow.consume(20)
  flow.ensureRoom(500)
  assert.equal(flow.pages.length, 1)
  assert.equal(flow.column.index, 1)
  flow.consume(500)
  flow.ensureRoom(20)
  assert.equal(flow.pages.length, 2)
  assert.equal(flow.column.index, 0)
})
