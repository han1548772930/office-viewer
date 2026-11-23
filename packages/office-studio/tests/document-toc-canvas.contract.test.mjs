import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import { collapsedSelection, createParagraph, normalizeDocument, runText } from '../src/features/document/model/document-model.ts'
import { documentTextFields, setDocumentTextFields } from '../src/features/document/model/document-text-fields.ts'
import { caretRect, hitTestDocument, layoutDocument } from '../src/features/document/layout/document-layout.ts'
import { sceneFromDocument, sceneOffset } from '../src/features/document/rendering/document-scene.ts'
import { buildDocumentTocBlocks, DEFAULT_DOCUMENT_TOC_OPTIONS } from '../src/features/document/toc/document-toc.ts'
import { documentTocHighlightLayout, documentTocPageTab } from '../src/features/document/toc/document-toc-layout.ts'

// Deferred geometry/Canvas scene contracts, 2026-11-24. No test or pixel validation has run.
const measurer = { measure: text => [...text].reduce((width, character) => width + (character === '.' ? 2 : character === '-' ? 4 : character === '_' ? 6 : character === '·' ? 3 : 8), 0) }
const close = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-6, `${actual} must equal ${expected}`)

function fixture({ labels = ['Alpha', 'Beta'], levels = [1, 2], pages = [9, 12], width = 440, height = 360, options = {}, columns = 1 } = {}) {
  const source = createBlankDocument(); source.id = 'toc-canvas-contract'
  source.page = { width, height, marginLeft: 40, marginRight: 40, marginTop: 40, marginBottom: 40 }
  source.sectionSetup = { columnCount: columns, columnGap: 24 }
  const headings = labels.map((label, index) => {
    const block = createParagraph(label); block.id = `heading-${index}`; block.type = 'heading'; block.level = levels[index] ?? 1
    return block
  })
  source.blocks = headings
  const toc = buildDocumentTocBlocks(source, { ...DEFAULT_DOCUMENT_TOC_OPTIONS, ...options }, Object.fromEntries(headings.map((block, index) => [block.id, pages[index] ?? index + 1])), 'canvas-toc', [], 'Contents')
  source.blocks = [...toc, ...headings]
  const model = normalizeDocument(source), entries = model.blocks.filter(block => block.metadata.documentToc?.role === 'entry')
  return { model, entries, selection: collapsedSelection({ blockId: entries[0].id, offset: 0 }) }
}

function pageTab(layout, entry) {
  for (const line of layout.lines.filter(line => line.target.blockId === entry.id)) {
    const glyph = line.glyphs.find(glyph => glyph.tab)
    if (glyph) return { line, glyph }
  }
  throw new Error('TOC entry has no laid out right tab')
}

function columnFor(layout, line) {
  return layout.pages[line.page].sections.find(section => section.fragment === line.sectionFragment).columns[line.column ?? 0]
}

function pageDigits(layout, entry) {
  const field = documentTextFields(entry).find(field => field.field === 'PAGEREF')
  return layout.lines.filter(line => line.target.blockId === entry.id).flatMap(line => line.glyphs)
    .filter(glyph => glyph.offset >= field.startOffset && glyph.end <= field.endOffset)
}

function scene(model, layout, selection, viewport = {}, extra = {}) {
  return sceneFromDocument(model, layout, { width: 900, height: 700, scrollLeft: 0, scrollTop: 0, zoom: 100, ...viewport },
    { selection, focused: true, caretVisible: true, composing: false, images: new Map(), tocPrimaryColor: '#125a9c', ...extra })
}

for (const [leader, character] of [['dots', '.'], ['dashes', '-'], ['underline', '_'], ['none', '']]) {
  test(`TOC ${leader} draws inside one tab gap, with page digits and both caret endpoints at their real positions`, () => {
    const { model, entries, selection } = fixture({ options: { tabLeader: leader } }), before = structuredClone(model)
    const layout = layoutDocument(model, measurer), { line, glyph } = pageTab(layout, entries[0]), digits = pageDigits(layout, entries[0])
    assert.equal(glyph.text, '\t'); assert.equal(glyph.offset, 5); assert.equal(glyph.end, 6)
    assert.equal(glyph.tab.leader, leader); assert.ok(glyph.width > 32)
    assert.ok(glyph.tab.leaderWidth >= 0 && glyph.tab.leaderWidth <= glyph.width)
    close(digits[0].x, glyph.x + glyph.width)
    const column = columnFor(layout, line)
    close(digits.at(-1).x + digits.at(-1).width, column.x + column.width)
    assert.deepEqual(hitTestDocument(layout, glyph.x + glyph.width / 4, line.y + line.height / 2), { blockId: entries[0].id, offset: 5 })
    assert.deepEqual(hitTestDocument(layout, glyph.x + glyph.width * 3 / 4, line.y + line.height / 2), { blockId: entries[0].id, offset: 6 })
    close(caretRect(layout, { blockId: entries[0].id, offset: 5 }).x, glyph.x)
    close(caretRect(layout, { blockId: entries[0].id, offset: 6 }).x, digits[0].x)
    const rendered = scene(model, layout, selection), offset = sceneOffset(layout, { width: 900, height: 700, scrollLeft: 0, scrollTop: 0, zoom: 100 })
    const decoration = rendered.commands.find(command => command.kind === 'text' && command.x === glyph.x + offset.x && command.y === line.baseline)
    if (character) {
      assert.ok(glyph.tab.leaderText.length); assert.ok([...glyph.tab.leaderText].every(value => value === character))
      assert.equal(decoration.text, glyph.tab.leaderText); close(decoration.maxWidth, glyph.width)
    } else {
      assert.equal(glyph.tab.leaderText, ''); assert.equal(glyph.tab.leaderWidth, 0); assert.equal(decoration, undefined)
    }
    assert.equal(rendered.commands.some(command => command.kind === 'text' && command.text === '    '), false)
    assert.deepEqual(model, before)
  })
}

test('TOC levels use pixel indents while page numbers share the same column right edge', () => {
  const { model, entries } = fixture(), layout = layoutDocument(model, measurer)
  const one = pageTab(layout, entries[0]), two = pageTab(layout, entries[1])
  assert.equal(entries[1].style.indent, 0); assert.equal(entries[1].style.indentStart, 18)
  close(two.line.x - one.line.x, 18)
  const first = pageDigits(layout, entries[0]).at(-1), second = pageDigits(layout, entries[1]).at(-1)
  close(first.x + first.width, second.x + second.width)
  assert.ok(two.line.x < columnFor(layout, two.line).x + 24)
})

test('a long or nearly full title wraps before its right tab without overlap or synthetic text offsets', () => {
  for (const title of ['ABCDEFGHIJKLMN', 'Long heading '.repeat(8), '目录😀标题'.repeat(12)]) {
    const { model, entries } = fixture({ labels: [title], pages: [123], width: 200 })
    entries[0].style.wordWrap = true
    const layout = layoutDocument(model, measurer), lines = layout.lines.filter(line => line.target.blockId === entries[0].id)
    assert.ok(lines.length > 1)
    const glyphs = lines.flatMap(line => line.glyphs)
    assert.equal(glyphs.map(glyph => glyph.text).join(''), runText(entries[0].runs))
    for (let index = 1; index < glyphs.length; index++) assert.equal(glyphs[index - 1].end, glyphs[index].offset)
    for (const line of lines) {
      const column = columnFor(layout, line)
      assert.ok(line.x >= column.x); assert.ok(line.width <= column.width + 1e-6)
      for (let index = 1; index < line.glyphs.length; index++) close(line.glyphs[index].x, line.glyphs[index - 1].x + line.glyphs[index - 1].width)
    }
    const { line, glyph } = pageTab(layout, entries[0]), digits = pageDigits(layout, entries[0])
    assert.equal(glyph.end, glyph.offset + 1); close(digits[0].x, glyph.x + glyph.width)
    close(digits.at(-1).x + digits.at(-1).width, columnFor(layout, line).x + columnFor(layout, line).width)
    assert.equal(hitTestDocument(layout, glyph.x + glyph.width * 3 / 4, line.y + line.height / 2).offset, glyph.end)
  }
})

test('leader measurement cannot paint past the aligned page digits when repeated glyphs have different advance', () => {
  const { model, entries } = fixture({ labels: ['Alpha'] })
  const nonLinear = { measure: text => /^\.+$/.test(text) ? text.length * 2 + text.length ** 2 / 100 : measurer.measure(text) }
  const layout = layoutDocument(model, nonLinear), { glyph } = pageTab(layout, entries[0])
  assert.ok(glyph.tab.leaderText.length > 0)
  assert.ok(nonLinear.measure(glyph.tab.leaderText) <= glyph.width)
  close(glyph.tab.leaderWidth, nonLinear.measure(glyph.tab.leaderText))
  close(pageDigits(layout, entries[0])[0].x, glyph.x + glyph.width)
})

test('a native right stop uses twips from the column origin, including a retained tab that disagrees with instruction options', () => {
  for (const [native, leader, character] of [['dot', 'dots', '.'], ['hyphen', 'dashes', '-'], ['underscore', 'underline', '_'], ['middleDot', 'middleDots', '·'], ['none', 'none', '']]) {
    const { model, entries } = fixture({ labels: ['Alpha'], options: { rightAlignPageNumbers: false } }), entry = entries[0]
    entry.runs[0].text = 'Alpha\t9'; entry.style.indentStart = 24; entry.style.align = 'right'
    entry.metadata.source = { details: { style: { tabStops: [{ alignment: 'right', leader: native, positionTwips: 4800 }], retained: 'original' } } }
    const before = structuredClone(model), page = documentTocPageTab(entry)
    assert.deepEqual(page, { offset: 5, stops: [{ right: 320, leader }] })
    const layout = layoutDocument(model, measurer), { line, glyph } = pageTab(layout, entry), digits = pageDigits(layout, entry)
    close(line.x, columnFor(layout, line).x + 24)
    close(digits.at(-1).x + digits.at(-1).width, columnFor(layout, line).x + 320)
    assert.equal(glyph.tab.leader, leader)
    assert.ok([...glyph.tab.leaderText].every(value => value === character))
    assert.deepEqual(model, before)
  }
})

test('unsupported native stop data is retained and does not turn into a configured dot leader', () => {
  for (const stop of [{ alignment: 'right', leader: 'heavy', positionTwips: 4800 }, { alignment: 'left', leader: 'dot', positionTwips: 4800 },
    { alignment: 'right', leader: 'dot', positionTwips: -1 }, { alignment: 'right', leader: 'dot', positionTwips: 4800.5 },
    { alignment: 'right', leader: 'dot', positionTwips: 4800, clear: true }]) {
    const { model, entries } = fixture({ labels: ['Alpha'] }), entry = entries[0]
    entry.metadata.source = { details: { style: { tabStops: [stop] }, xml: '<w:tabs/>' } }
    const before = structuredClone(model)
    assert.equal(documentTocPageTab(entry), undefined)
    assert.ok(layoutDocument(model, measurer).lines.filter(line => line.target.blockId === entry.id).flatMap(line => line.glyphs).every(glyph => !glyph.tab))
    assert.deepEqual(model, before)
  }
})

test('TOC highlighting covers the entire directory and has separate outlines for physical pages and columns', () => {
  const { model, selection } = fixture({ labels: Array.from({ length: 32 }, (_, index) => `Entry ${index}`), height: 240, columns: 2 })
  const layout = layoutDocument(model, measurer), highlight = documentTocHighlightLayout(model, layout, selection)
  const ids = new Set(model.blocks.filter(block => block.metadata.documentToc).map(block => block.id))
  const lines = layout.lines.filter(line => ids.has(line.target.blockId))
  assert.equal(highlight.tocId, 'canvas-toc'); assert.equal(highlight.ranges.length, lines.length)
  assert.ok(new Set(lines.map(line => line.page)).size > 1); assert.deepEqual([...new Set(lines.map(line => line.column))].sort(), [0, 1])
  assert.equal(highlight.outlines.length, new Set(lines.map(line => `${line.page}:${line.sectionFragment}:${line.column}`)).size)
  for (const line of lines) {
    const range = highlight.ranges.find(range => range.page === line.page && range.x === line.x && range.y === line.y)
    assert.ok(range); close(range.width, Math.max(2, line.width)); close(range.height, line.height)
    assert.ok(highlight.outlines.some(outline => outline.page === line.page && outline.x <= line.x - 4 && outline.y <= line.y - 4
      && outline.x + outline.width >= line.x + line.width + 4 && outline.y + outline.height >= line.y + line.height + 4))
  }
  assert.equal(documentTocHighlightLayout(model, layout, collapsedSelection({ blockId: 'heading-0', offset: 0 })), undefined)
  assert.equal(documentTocHighlightLayout(model, layout, { anchor: selection.anchor, focus: { blockId: 'heading-0', offset: 1 } }), undefined)
  assert.equal(documentTocHighlightLayout(model, layout, collapsedSelection({ ...selection.anchor, row: 0, col: 0, cellBlockId: 'cell' })), undefined)
})

test('a half-open backward selection ending at the next body paragraph still highlights only its TOC', () => {
  const { model, entries } = fixture(), layout = layoutDocument(model, measurer)
  const selection = { anchor: { blockId: 'heading-0', offset: 0 }, focus: { blockId: entries[0].id, offset: 0 } }
  assert.equal(documentTocHighlightLayout(model, layout, selection).tocId, 'canvas-toc')
})

test('TOC scene layers inherit the primary token, viewport clipping and zoom and disappear during story editing', () => {
  const { model, entries, selection } = fixture(), before = structuredClone(model), layout = layoutDocument(model, measurer)
  const viewport = { width: 900, height: 160, scrollLeft: 10, scrollTop: 50, zoom: 150 }
  const rendered = scene(model, layout, selection, viewport), highlight = documentTocHighlightLayout(model, layout, selection), off = sceneOffset(layout, viewport)
  const fills = rendered.commands.filter(command => command.kind === 'rect' && command.fill === 'rgba(18, 90, 156, 0.08)')
  const outlines = rendered.commands.filter(command => command.kind === 'rect' && command.stroke === 'rgba(18, 90, 156, 0.8)')
  assert.ok(fills.length > 0); assert.ok(outlines.length > 0)
  for (const fill of fills) {
    const range = highlight.ranges.find(range => Math.abs(fill.x - (off.x + range.x * 1.5)) < 1e-6 && Math.abs(fill.y - (off.y + range.y * 1.5)) < 1e-6)
    assert.ok(range); close(fill.width, range.width * 1.5); close(fill.height, range.height * 1.5)
    const page = layout.pages[range.page]
    assert.deepEqual(fill.clip, { x: off.x + page.x * 1.5, y: off.y + page.y * 1.5, width: page.width * 1.5, height: page.height * 1.5 })
  }
  const { line, glyph } = pageTab(layout, entries[0])
  const textIndex = rendered.commands.findIndex(command => command.kind === 'text' && command.text === glyph.tab.leaderText && command.y === off.y + line.baseline * 1.5)
  assert.ok(textIndex > rendered.commands.indexOf(fills[0]))
  assert.ok(rendered.commands.indexOf(outlines[0]) > textIndex)
  const caretIndex = rendered.commands.findIndex(command => command.kind === 'rect' && command.fill === '#24569c')
  assert.ok(caretIndex > rendered.commands.indexOf(outlines[0]))
  const changed = scene(model, layout, selection, viewport, { tocPrimaryColor: '#b42850' })
  assert.ok(changed.commands.some(command => command.kind === 'rect' && command.fill === 'rgba(180, 40, 80, 0.08)'))
  const story = scene(model, layout, selection, viewport, { story: { location: { page: 0, area: 'header' }, selection: { start: 0, end: 0 } } })
  assert.equal(story.commands.some(command => command.kind === 'rect' && (command.fill.endsWith(', 0.08)') || command.stroke?.endsWith(', 0.8)'))), false)
  assert.equal(scene(model, layout, selection, { scrollTop: layout.height + 1000 }).commands.some(command => command.kind === 'rect' && command.stroke?.endsWith(', 0.8)')), false)
  assert.deepEqual(model, before)
})

test('TOC highlight rejects invalid UTF-16 carets and never fills a parity blank', () => {
  const { model, entries, selection } = fixture({ labels: ['😀Title'] })
  const intro = createParagraph('Intro'); intro.id = 'intro'
  const section = { ...createParagraph(), id: 'toc-section', type: 'sectionBreak', runs: [], sectionBreak: { type: 'oddPage' } }
  model.blocks.unshift(intro, section)
  const layout = layoutDocument(model, measurer), highlight = documentTocHighlightLayout(model, layout, selection)
  assert.ok(layout.pages.some(page => page.isParityBlank))
  assert.ok(highlight.ranges.every(range => !layout.pages[range.page].isParityBlank))
  assert.equal(documentTocHighlightLayout(model, layout, collapsedSelection({ blockId: entries[0].id, offset: 1 })), undefined)
})
