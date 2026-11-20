import assert from 'node:assert/strict'
import { test } from 'node:test'
import { activeOutlineIdForBlock, activeOutlineIdInViewport, buildDocumentOutline, documentOutlineEnabled, documentOutlineOverlapsTable, documentOutlinePlacement, nextOutlineFocusId } from '../src/features/document/services/document-outline.ts'

function block(id, text, type = 'heading', level = 1) {
  return { id, type, level, runs: [{ id: `${id}-run`, text }], style: {}, metadata: {} }
}

function model(blocks, name = '测试文档', metadata) {
  return { kind: 'document', name, blocks, page: {}, modified: false, wordCount: 0, metadata }
}

test('builds an immutable outline from current stable paragraph IDs, preserving levels and duplicate text', () => {
  const source = model([
    block('normal', '正文', 'paragraph'),
    block('first', '章一'),
    block('third', '同名标题', 'heading', 3),
    block('other', '同名标题', 'heading', 2),
    block('blank', ' \t\n '),
    block('last', ' 跨\n行\t标题 ', 'heading', 6),
  ])
  const before = structuredClone(source)
  const items = buildDocumentOutline(source)
  assert.equal(items[0].isTitle, true)
  assert.equal(items[0].blockId, undefined)
  assert.equal(items[0].text, '测试文档')
  assert.deepEqual(items.slice(1).map(({ id, blockId, text, level, depth, blockIndex }) => ({ id, blockId, text, level, depth, blockIndex })), [
    { id: 'first', blockId: 'first', text: '章一', level: 1, depth: 0, blockIndex: 1 },
    { id: 'third', blockId: 'third', text: '同名标题', level: 3, depth: 2, blockIndex: 2 },
    { id: 'other', blockId: 'other', text: '同名标题', level: 2, depth: 1, blockIndex: 3 },
    { id: 'last', blockId: 'last', text: '跨 行 标题', level: 6, depth: 5, blockIndex: 5 },
  ])
  assert.deepEqual(source, before)
})

test('rebuild follows edits, removal and reordering instead of stale source styles', () => {
  const formerHeading = block('normal', '已改正文', 'paragraph')
  formerHeading.metadata = { source: { details: { style: { headingLevel: 1, namedStyleType: 'TITLE' } } } }
  const source = model([formerHeading, block('two', '第二章'), block('one', '第一章')])
  assert.deepEqual(buildDocumentOutline(source, { includeDocumentTitle: false }).map(({ id }) => id), ['two', 'one'])
  source.blocks[1].runs[0].text = '改名后的第二章'
  source.blocks.splice(2, 1)
  assert.deepEqual(buildDocumentOutline(source, { includeDocumentTitle: false }).map(({ id, text }) => ({ id, text })), [{ id: 'two', text: '改名后的第二章' }])
})

test('normalizes malformed levels locally, skips invalid/repeated IDs and avoids a title ID collision', () => {
  const source = model([
    block('__document_outline_title__', '真实段落', 'heading', 99),
    block('zero', '零', 'heading', 0),
    block('nan', '非数值', 'heading', Number.NaN),
    block('fraction', '小数', 'heading', 2.8),
    block('zero', '重复ID'),
    block('', '无ID'),
  ])
  const items = buildDocumentOutline(source)
  assert.equal(items[0].id, '__document_outline_title___')
  assert.deepEqual(items.slice(1).map(({ level }) => level), [6, 1, 1, 2])
  assert.equal(source.blocks[0].level, 99)
  assert.deepEqual(buildDocumentOutline(model([], ' ')), [])
})

test('explicit outline configuration overrides modern/traditional defaults without looking at page dimensions', () => {
  assert.equal(documentOutlineEnabled(model([])), false)
  assert.equal(documentOutlineEnabled(model([], '', { documentFlavor: 'traditional' })), false)
  assert.equal(documentOutlineEnabled(model([], '', { documentFlavor: 1 })), false)
  assert.equal(documentOutlineEnabled(model([], '', { documentFlavor: 'MODERN' })), true)
  assert.equal(documentOutlineEnabled(model([], '', { documentStyle: { documentFlavor: 2 } })), true)
  assert.equal(documentOutlineEnabled(model([], '', { source: { documentStyle: { documentFlavor: 'modern' } } })), true)
  assert.equal(documentOutlineEnabled(model([], '', { documentFlavor: 'traditional' }), true), true)
  assert.equal(documentOutlineEnabled(model([], '', { documentFlavor: 'modern' }), false), false)
})

test('maps a body paragraph to its preceding section and document title before the first heading', () => {
  const source = model([block('intro', '导语', 'paragraph'), block('first', '第一章'), block('body', '正文', 'paragraph'), block('second', '第二章')])
  const items = buildDocumentOutline(source)
  assert.equal(activeOutlineIdForBlock(items, source, 'intro'), items[0].id)
  assert.equal(activeOutlineIdForBlock(items, source, 'first'), 'first')
  assert.equal(activeOutlineIdForBlock(items, source, 'body'), 'first')
  assert.equal(activeOutlineIdForBlock(items, source, 'second'), 'second')
  assert.equal(activeOutlineIdForBlock(items, source, 'deleted'), undefined)
  assert.equal(activeOutlineIdForBlock(buildDocumentOutline(source, { includeDocumentTitle: false }), source, 'intro'), undefined)
})

test('viewport tracking prefers a visible heading, then a preceding heading, and honors exact edges', () => {
  const source = model([block('intro', '导语', 'paragraph'), block('first', '第一章'), block('body', '正文', 'paragraph'), block('second', '第二章')])
  const items = buildDocumentOutline(source)
  const bounds = [
    { blockId: 'intro', top: 10, bottom: 30 },
    { blockId: 'first', top: 40, bottom: 60 },
    { blockId: 'body', top: 70, bottom: 170 },
    { blockId: 'second', top: 180, bottom: 200 },
  ]
  assert.equal(activeOutlineIdInViewport(items, bounds, 0, 30), items[0].id)
  assert.equal(activeOutlineIdInViewport(items, bounds, 70, 180), 'first')
  assert.equal(activeOutlineIdInViewport(items, bounds, 70, 190), 'second')
  assert.equal(activeOutlineIdInViewport(items, bounds, 60, 150), 'first')
  assert.equal(activeOutlineIdInViewport(items, bounds, 200, 300), undefined)
  assert.equal(activeOutlineIdInViewport(items, bounds, 80, 80), undefined)
  assert.equal(activeOutlineIdInViewport(items, bounds, Number.NaN, 300), undefined)
})

test('keyboard movement clamps at boundaries and remains separate from activation', () => {
  const source = model([block('one', '一'), block('two', '二'), block('three', '三')])
  const items = buildDocumentOutline(source, { includeDocumentTitle: false })
  assert.equal(nextOutlineFocusId(items, 'two', 'ArrowDown'), 'three')
  assert.equal(nextOutlineFocusId(items, 'three', 'ArrowDown'), 'three')
  assert.equal(nextOutlineFocusId(items, 'one', 'ArrowUp'), 'one')
  assert.equal(nextOutlineFocusId(items, 'three', 'Home'), 'one')
  assert.equal(nextOutlineFocusId(items, 'one', 'End'), 'three')
  assert.equal(nextOutlineFocusId(items, 'removed', 'ArrowDown'), 'one')
  assert.equal(nextOutlineFocusId(items, 'two', 'Enter'), undefined)
  assert.equal(nextOutlineFocusId([], undefined, 'Home'), undefined)
})

test('viewport tracking ignores zero-height or invalid heading bounds between visible paragraphs', () => {
  const source = model([
    block('first', '第一章'), block('one', '正文一', 'paragraph'), block('hidden', '无效标题'),
    block('two', '正文二', 'paragraph'), block('bad', '无效标题二'),
  ])
  const items = buildDocumentOutline(source)
  const bounds = [
    { blockId: 'first', top: 0, bottom: 20 },
    { blockId: 'one', top: 30, bottom: 80 },
    { blockId: 'hidden', top: 80, bottom: 80 },
    { blockId: 'two', top: 90, bottom: 140 },
    { blockId: 'bad', top: Number.NaN, bottom: Number.NaN },
  ]
  assert.equal(activeOutlineIdInViewport(items, bounds, 30, 140), 'first')
  assert.equal(activeOutlineIdInViewport(items, bounds, 90, 140), 'first')
  assert.equal(activeOutlineIdInViewport(items, bounds.slice(2, 3), 0, 300), undefined)
})

test('placement accounts for physical page width, editor zoom, padding and exact available space', () => {
  const input = { availableWidth: 1098, pageWidth: 794, zoom: 100 }
  assert.equal(documentOutlinePlacement(input), 'side-bar')
  assert.equal(documentOutlinePlacement({ ...input, availableWidth: 1097 }), 'float')
  assert.equal(documentOutlinePlacement({ ...input, zoom: 120 }), 'float')
  assert.equal(documentOutlinePlacement({ ...input, pageWidth: 1123 }), 'float')
  assert.equal(documentOutlinePlacement({ ...input, pageWidth: 1123, zoom: 60 }), 'side-bar')
  assert.equal(documentOutlinePlacement({ ...input, availableWidth: 0 }), 'float')
  assert.equal(documentOutlinePlacement({ ...input, zoom: Number.NaN }), 'float')
  assert.equal(documentOutlinePlacement({ ...input, availableWidth: 1097, outlineWidth: Number.NaN, canvasPadding: Number.NaN }), 'float')
})

test('wide table overlap uses the visible scrolled and clipped content rather than its empty inset', () => {
  const outline = { x: 16, y: 16, width: 256, height: 184 }
  const bounds = { x: 420, y: 100, width: 1920, height: 180 }
  const clip = { x: 16, y: 0, width: 1408, height: 1000 }
  assert.equal(documentOutlineOverlapsTable(outline, { bounds, clip, scrollLeft: 0 }), false)
  assert.equal(documentOutlineOverlapsTable(outline, { bounds, clip, scrollLeft: 148 }), false)
  assert.equal(documentOutlineOverlapsTable(outline, { bounds, clip, scrollLeft: 200 }), true)
  assert.equal(documentOutlineOverlapsTable(outline, { bounds, clip: { ...clip, x: 300 }, scrollLeft: 400 }), false)
  assert.equal(documentOutlineOverlapsTable(outline, { bounds: { ...bounds, y: 200 }, clip, scrollLeft: 400 }), false)
  assert.equal(documentOutlineOverlapsTable(outline, { bounds: { ...bounds, y: -180 }, clip, scrollLeft: 400 }), false)
  assert.equal(documentOutlineOverlapsTable(outline, { bounds: { ...bounds, x: 100 } }), true)
  assert.equal(documentOutlineOverlapsTable(outline, { bounds: { ...bounds, x: Number.NaN } }), false)
})

test('placement floats only for an actual visible wide-table collision and can dock again after scroll', () => {
  const input = {
    availableWidth: 1440, pageWidth: 794, zoom: 100,
    outlineBounds: { x: 16, y: 16, width: 256, height: 184 },
    tables: [{ bounds: { x: 420, y: 100, width: 1920, height: 180 }, scrollLeft: 200 }],
  }
  assert.equal(documentOutlinePlacement(input), 'float')
  assert.equal(documentOutlinePlacement({ ...input, tables: [{ ...input.tables[0], scrollLeft: 0 }] }), 'side-bar')
  assert.equal(documentOutlinePlacement({ ...input, tables: [{ ...input.tables[0], bounds: { ...input.tables[0].bounds, y: 201 } }] }), 'side-bar')
})
