import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import { createParagraph, createTable } from '../src/features/document/model/document-model.ts'
import { numberingDefinitionFromParts } from '../src/features/document/model/document-numbering.ts'
import { caretRect, hitTestDocument, layoutDocument } from '../src/features/document/layout/document-layout.ts'
import { sceneFromDocument } from '../src/features/document/rendering/document-scene.ts'

const measurer = { measure: text => Array.from(text).length * 10 }
const documentWith = blocks => ({ ...createBlankDocument(), id: 'layout-doc', blocks, page: { width: 200, height: 400, marginLeft: 20, marginRight: 20, marginTop: 40, marginBottom: 40 } })

test('first-line, hanging and right indentation use different wrapping bounds for later lines', () => {
  const block = createParagraph('A'.repeat(26))
  Object.assign(block.style, { indentStart: 15, indentFirstLine: 25, hanging: 5, indentEnd: 10, wordWrap: true })
  const layout = layoutDocument(documentWith([block]), measurer), page = layout.pages[0]
  assert.deepEqual(layout.lines.map(line => line.end), [11, 24, 26])
  assert.equal(layout.lines[0].x, page.content.x + 35)
  assert.equal(layout.lines[1].x, page.content.x + 15)
  for (const line of layout.lines) assert.ok(line.x + line.width <= page.content.x + page.content.width - 10)
  const caret = caretRect(layout, { blockId: block.id, offset: 12 })
  assert.equal(caret.x, layout.lines[1].glyphs[1].x)
  assert.equal(hitTestDocument(layout, caret.x + 1, caret.y + 2).offset, 12)
})

test('negative first-line positioning remains visible inside the page margin', () => {
  const block = createParagraph('Margin')
  Object.assign(block.style, { indentStart: -10, hanging: 10 })
  const model = documentWith([block]), layout = layoutDocument(model, measurer)
  const line = layout.lines[0]
  assert.equal(line.x, layout.pages[0].content.x - 20)
  const selection = { anchor: { blockId: block.id, offset: 0 }, focus: { blockId: block.id, offset: 0 } }
  const scene = sceneFromDocument(model, layout, { width: layout.width, height: 450, scrollLeft: 0, scrollTop: 0, zoom: 100 }, { selection, focused: false, caretVisible: false, composing: false, images: new Map() })
  const first = scene.commands.find(command => command.kind === 'text' && command.text === 'M')
  assert.equal(first.x, line.x)
  assert.equal(first.clip.x, layout.pages[0].x)
})

test('automatic fractional, exact and minimum line spacing retain their distinct units', () => {
  const heights = (spacingRule, lineHeight) => {
    const block = createParagraph('Line', { fontSize: 12 })
    Object.assign(block.style, { spacingRule, lineHeight })
    return layoutDocument(documentWith([block]), measurer).lines[0]
  }
  assert.equal(heights('auto', 0.75).height, 12)
  assert.equal(heights('exact', 8).height, 8)
  assert.equal(heights('atLeast', 8).height, 16)
  assert.equal(heights('atLeast', 20).height, 20)
  assert.equal(heights('auto', 2).glyphLineHeight, 16)
})

test('wordWrap controls breaks within words and oversized words still break at grapheme boundaries', () => {
  const block = createParagraph('aa bbbb cc')
  const model = documentWith([block])
  model.page.marginLeft = model.page.marginRight = 70
  const texts = () => layoutDocument(model, measurer).lines.map(line => line.glyphs.map(glyph => glyph.text).join(''))
  assert.deepEqual(texts(), ['aa ', 'bbbb ', 'cc'])
  block.style.wordWrap = true
  assert.deepEqual(texts(), ['aa bbb', 'b cc'])
  block.style.wordWrap = false
  block.runs[0].text = 'abcdefgh'
  assert.deepEqual(texts(), ['abcdef', 'gh'])
  block.runs[0].text = '中文中文中文中文'
  assert.deepEqual(texts(), ['中文中文中文', '中文'])
})

test('list layout shares body counters with cells and uses marker font, color and image dimensions', () => {
  const numbering = numberingDefinitionFromParts({ abstractNumId: 7, levels: [{ ilvl: 0, start: 1, format: 'decimal', text: '%1.', indent: { leftTwips: 420, hangingTwips: 420 }, runStyle: { fontFamilyAscii: 'Georgia', fontSizePt: 9, color: 'auto', bold: true } }] }, { numId: 9, abstractNumId: 7 })
  const listed = id => ({ ...createParagraph(id), id, type: 'numbered', metadata: { numberingDefinition: numbering, numberingListId: 'shared' } })
  const table = createTable(1, 2)
  table.metadata.table.rows[0].cells[0].blocks = [listed('cell-a')]
  table.metadata.table.rows[0].cells[1].blocks = [listed('cell-b')]
  const model = documentWith([listed('first'), table, listed('last')]), layout = layoutDocument(model, measurer)
  assert.deepEqual(layout.lines.filter(line => line.marker !== undefined).map(line => line.marker), ['1.', '2.', '3.', '4.'])
  assert.match(layout.lines[0].markerStyle.font, /^bold 12px Georgia/u)
  assert.equal(layout.lines[0].markerStyle.color, 'auto')
  const selection = { anchor: { blockId: 'first', offset: 0 }, focus: { blockId: 'first', offset: 0 } }
  const scene = sceneFromDocument(model, layout, { width: 520, height: 850, scrollLeft: 0, scrollTop: 0, zoom: 200 }, { selection, focused: false, caretVisible: false, composing: false, images: new Map() })
  const marker = scene.commands.find(command => command.kind === 'text' && command.text === '1.')
  assert.match(marker.font, /^bold 24px Georgia/u)
  assert.equal(marker.color, '#1f2937')
  const picture = structuredClone(numbering)
  picture.levels[0].pictureBullet = { src: 'data:image/png;base64,AA==', widthPx: 9, heightPx: 13 }
  const imageModel = documentWith([{ ...listed('picture'), metadata: { numberingDefinition: picture, numberingListId: 'image-list' } }])
  const imageLayout = layoutDocument(imageModel, measurer)
  assert.equal(imageLayout.lines[0].marker, '')
  assert.equal(imageLayout.images[0].width, 9)
  assert.equal(imageLayout.images[0].height, 13)
})
