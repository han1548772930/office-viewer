import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import { createParagraph, createRun, createTable, runText } from '../src/features/document/model/document-model.ts'
import { setDocumentImageAtom } from '../src/features/document/images/document-image-atom.ts'
import { documentImageWrapBands } from '../src/features/document/layout/document-image-layout.ts'
import { documentPhysicalLineFragments, documentPhysicalLineNavigation, emptyLayout, layoutDocument } from '../src/features/document/layout/document-layout.ts'
import { measureDocumentImageTextRow } from '../src/features/document/layout/document-paragraph-layout.ts'
import { positionLine } from '../src/features/document/layout/document-text-layout.ts'
import { measureDocumentBlockCollections } from '../src/features/document/layout/document-block-measurement.ts'
import { measureDocumentTable } from '../src/features/document/layout/document-table-measurement.ts'
import { DocumentSectionFlow } from '../src/features/document/layout/document-section-flow.ts'
import { sceneFromDocument } from '../src/features/document/rendering/document-scene.ts'

// Deferred DOC-07 physical-row/atom/pagination contracts; no layout test or renderer has run.
const source = 'data:image/png;base64,AA=='
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-6, `${actual} != ${expected}`)
const measurer = { measure: text => Array.from(text).length * 8 }
const paragraph = (id, runs = [createRun('')]) => {
  const block = { ...createParagraph(), id, runs }
  Object.assign(block.style, { spaceBefore: 0, spaceAfter: 0, spacingRule: 'exact', lineHeight: 20, wordWrap: true, widowControl: false })
  return block
}
const atom = (id, patch = {}) => {
  const run = createRun()
  setDocumentImageAtom(run, id, { src: source, alt: id, width: 40, height: 30, wrappingStyle: 'inline', ...patch })
  return run
}
const document = blocks => ({ ...createBlankDocument(), id: 'inline-layout', blocks,
  page: { width: 420, height: 420, marginLeft: 40, marginRight: 40, marginTop: 40, marginBottom: 40 } })

test('inline images share the text baseline under auto, atLeast and exact spacing and send flip flags to the shared scene', () => {
  for (const spacingRule of ['auto', 'atLeast', 'exact']) {
    const block = paragraph('p', [createRun('L'), atom('small', { height: 24, flipX: true }), createRun('M', { fontSize: 22 }),
      atom('large', { height: 60, flipY: true }), createRun('R')])
    Object.assign(block.style, { spacingRule, lineHeight: spacingRule === 'auto' ? 1.2 : 20 })
    const model = document([block]), layout = layoutDocument(model, measurer), line = layout.lines.find(line => line.target.blockId === 'p')
    assert.equal(layout.images.length, 2)
    for (const image of layout.images) near(image.y + image.height, line.baseline)
    assert.equal(line.glyphs.filter(glyph => glyph.imageAtom).length, 2)
    const scene = sceneFromDocument(model, layout, { width: 800, height: 800, scrollLeft: 0, scrollTop: 0, zoom: 100 }, {
      selection: { anchor: { blockId: 'p', offset: 0 }, focus: { blockId: 'p', offset: 0 } }, focused: true, caretVisible: false,
      composing: false, images: new Map([[source, { complete: true, naturalWidth: 1, naturalHeight: 1 }]]) })
    const commands = scene.commands.filter(command => command.kind === 'image')
    assert.equal(commands.length, 2); assert.equal(commands[0].flipX, true); assert.equal(commands[1].flipY, true)
    assert.equal(scene.commands.some(command => command.kind === 'text' && command.text.includes('\uFFFC')), false)
  }
})

test('square wrapping consumes both intervals in reading order with one baseline and no repeated UTF-16 suffix', () => {
  const obstacle = { x: 80, y: 0, width: 70, height: 100, wrappingStyle: 'wrapSquare', wrapSide: 'both' }
  const bands = documentImageWrapBands({ x: 0, y: 0, width: 300, height: 200 }, 0, 20, [obstacle])
  assert.deepEqual(bands.map(band => [band.x, band.width]), [[0, 80], [150, 150]])
  const text = '😀' + 'a'.repeat(70), block = paragraph('p', [createRun(text)])
  const row = measureDocumentImageTextRow(block, undefined, 300, bands, measurer)
  assert.equal(row.fragments.length, 2); assert.equal(row.fragments[0].raw.start, 0)
  assert.equal(row.fragments[1].raw.start, row.fragments[0].raw.end)
  assert.equal(row.fragments.flatMap(fragment => fragment.raw.glyphs).map(glyph => glyph.text).join(''), text.slice(0, row.nextOffset))
  assert.equal(row.complete, false); assert.ok(row.fragments.every(fragment => fragment.raw.start !== 1 && fragment.raw.end !== 1))
  const placed = row.fragments.map(fragment => positionLine(fragment.raw, { blockId: 'p' }, fragment.band.x, 0, fragment.band.width, row.height, 0, 'left', false))
  near(placed[0].baseline, placed[1].baseline)
})

test('an explicit line break ends the physical row and a trailing break retains exactly one empty row', () => {
  const block = paragraph('p', [createRun('😀a\u2028bc\u2028')]), bands = [{ x: 0, y: 0, width: 80, height: 20 }, { x: 120, y: 0, width: 80, height: 20 }]
  const first = measureDocumentImageTextRow(block, undefined, 200, bands, measurer)
  assert.equal(first.fragments.length, 1); assert.equal(first.nextOffset, 4); assert.equal(first.complete, false)
  const second = measureDocumentImageTextRow(block, undefined, 200, bands, measurer, undefined, undefined, 1, first.nextOffset)
  assert.equal(second.fragments.length, 1); assert.equal(second.nextOffset, 7); assert.equal(second.complete, false)
  const last = measureDocumentImageTextRow(block, undefined, 200, bands, measurer, undefined, undefined, 2, second.nextOffset)
  assert.equal(last.fragments.length, 1); assert.equal(last.fragments[0].raw.glyphs.length, 0)
  assert.equal(last.nextOffset, 7); assert.equal(last.complete, true)
})

test('Home and End use all fragments and vertical navigation skips the current physical row within its host', () => {
  const visual = (start, end, x, y, rowKey, fragmentIndex = 0) => {
    const run = createRun('a'.repeat(end - start)), raw = { start, end, width: (end - start) * 8, fontHeight: 14.67,
      glyphs: Array.from(run.text, (text, index) => ({ text, offset: start + index, end: start + index + 1, width: 8, run, font: '11px Aptos' })) }
    return { ...positionLine(raw, { blockId: 'p' }, x, y, raw.width, 20, 0, 'left', false), rowKey, fragmentIndex }
  }
  const previous = visual(0, 2, 0, 0, 'previous'), left = visual(2, 6, 0, 20, 'split'), right = visual(6, 8, 120, 20, 'split', 1)
  const next = visual(8, 12, 0, 40, 'next'), foreign = { ...right, host: { kind: 'note', noteKind: 'footnote', noteId: 'other' } }
  const layout = { ...emptyLayout(), lines: [previous, left, right, next, foreign] }, point = { blockId: 'p', offset: 7 }
  assert.deepEqual(documentPhysicalLineFragments(layout, right), [left, right])
  assert.equal(documentPhysicalLineNavigation(layout, point, 'home').offset, 2)
  assert.equal(documentPhysicalLineNavigation(layout, point, 'end').offset, 8)
  assert.equal(documentPhysicalLineNavigation(layout, point, 'down', 10).offset, 9)
  assert.equal(documentPhysicalLineNavigation(layout, { blockId: 'p', offset: 10 }, 'up', 135).offset, 8)
})

test('table and auxiliary measurements retain atom source addresses and advance a split row once', () => {
  const blocks = [paragraph('p', [atom('float', { width: 70, height: 60, wrappingStyle: 'wrapSquare',
    positionH: { relativeFrom: 1, posOffset: 80 }, positionV: { relativeFrom: 1, posOffset: 0 } }), createRun('a'.repeat(65)), atom('inline')])]
  const options = { width: 300, markers: new Map(), measurer, measureTable: measureDocumentTable, tableCellId: 'cell',
    target: blockId => ({ blockId: 'table', cellId: 'cell', cellBlockId: blockId }) }
  const measured = measureDocumentBlockCollections([blocks], options), fragments = measured.lines.filter(line => line.rowKey)
  assert.ok(fragments.length >= 2); near(fragments[0].y, fragments[1].y)
  assert.ok(measured.lines.every(line => line.target.cellBlockId === 'p' && line.tableCellId === 'cell'))
  assert.deepEqual(measured.images.find(image => image.blockId === 'float').anchor.target, options.target('p'))
  const inline = measured.images.find(image => image.blockId === 'inline')
  assert.equal(inline.anchor.startOffset, runText(blocks[0].runs).length - 1)
  assert.ok(measured.lines.at(-1).end === runText(blocks[0].runs).length)
  const table = createTable(1, 1); table.rows[0].cells[0].blocks = blocks
  const nested = measureDocumentTable({ ...paragraph('table'), type: 'table', runs: [], metadata: { table } }, 320, new Map(), measurer)
  assert.ok(nested.images.every(image => image.tableCellId)); assert.equal(nested.images.find(image => image.blockId === 'inline').anchor.target.cellBlockId, 'p')
})

test('paragraph indents trim wrap intervals once and preserve the same inset after the obstacle ends', () => {
  const block = paragraph('p', [atom('float', { width: 120, height: 40, wrappingStyle: 'wrapSquare', wrapSide: 'right',
    positionH: { relativeFrom: 1, posOffset: 0 }, positionV: { relativeFrom: 1, posOffset: 0 } }), createRun('a'.repeat(110))])
  Object.assign(block.style, { indentStart: 20, indentEnd: 15, indentFirstLine: 10 })
  const measured = measureDocumentBlockCollections([[block]], { width: 240, markers: new Map(), measurer, measureTable: measureDocumentTable })
  const first = measured.lines[0], after = measured.lines.find(line => line.y >= 40)
  assert.equal(first.x, 120); assert.equal(after.x, 20)
  assert.ok(first.glyphs.every(glyph => glyph.x + glyph.width <= 225 + 1e-6))
  assert.ok(measured.lines.every(line => line.glyphs.every(glyph => glyph.x + glyph.width <= 225 + 1e-6)))
})

test('wrap overflow moves to the next column on the same page and recalculates its bands', () => {
  const model = document([paragraph('p')]); model.sectionSetup = { columnCount: 2, columnGap: 20 }
  const flow = new DocumentSectionFlow(model), original = flow.column
  flow.consume(original.height - 30)
  const obstacle = { x: original.x, y: flow.cursor, width: original.width, height: 40, wrappingStyle: 'wrapTopAndBottom', wrapSide: 'both' }
  flow.addImageObstacle(obstacle)
  const bands = flow.imageWrapBands(20, 10, 5)
  assert.equal(flow.page.index, 0); assert.equal(flow.column.index, 1)
  assert.equal(bands[0].x, flow.column.x + 10); assert.equal(bands[0].y, flow.column.y)
  assert.equal(bands[0].width, flow.column.width - 15)
})

test('a topBottom float taller than the page retains one anchor and makes finite text pagination progress', () => {
  const block = paragraph('p', [atom('tall', { width: 340, height: 900, wrappingStyle: 'wrapTopAndBottom',
    positionH: { relativeFrom: 1, posOffset: 0 }, positionV: { relativeFrom: 1, posOffset: 0 } }), createRun('a'.repeat(160))])
  const model = document([block]), before = structuredClone(model), layout = layoutDocument(model, measurer)
  assert.equal(layout.images.filter(image => image.blockId === 'tall').length, 1)
  assert.ok(layout.pages.length > 1 && layout.pages.length < 8)
  assert.equal(layout.images.find(image => image.blockId === 'tall').page, 0)
  assert.equal(layout.lines.at(-1).end, runText(block.runs).length)
  assert.ok(layout.lines.every(line => [line.x, line.y, line.height, line.baseline].every(Number.isFinite)))
  assert.deepEqual(model, before)
})
