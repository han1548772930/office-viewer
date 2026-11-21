import assert from 'node:assert/strict'
import { test } from 'vitest'
import { documentRectScrollOffset, documentViewportLocation, zoomedDocumentScrollOffset } from '../src/features/document/layout/document-viewport.ts'

const layout = {
  width: 850, height: 2350,
  pages: [{ index: 0, y: 28, height: 1123 }, { index: 1, y: 1175, height: 1123 }],
  lines: [
    { y: 100, height: 22, target: { blockId: 'heading-1' } },
    { y: 122, height: 22, target: { blockId: 'paragraph-1' } },
    { y: 1250, height: 22, target: { blockId: 'heading-2' } },
  ],
}
const viewport = { width: 1000, height: 600, scrollLeft: 0, scrollTop: 0, zoom: 100 }

test('outline viewport bounds use logical coordinates, zoom and horizontal page centering', () => {
  const original = structuredClone(layout)
  assert.deepEqual(documentViewportLocation(layout, viewport), {
    pageIndex: 0, firstVisibleBlockId: 'heading-1', top: 0, bottom: 600, left: -75, right: 925,
  })
  const location = documentViewportLocation(layout, { ...viewport, width: 500, scrollLeft: 20, scrollTop: 2450, zoom: 200 })
  assert.deepEqual(location, { pageIndex: 1, firstVisibleBlockId: 'heading-2', top: 1225, bottom: 1525, left: 10, right: 260 })
  assert.deepEqual(layout, original)
})

test('a hidden viewport reports no visible paragraph or active page', () => {
  const location = documentViewportLocation(layout, { ...viewport, height: 0 })
  assert.equal(location.pageIndex, undefined)
  assert.equal(location.firstVisibleBlockId, undefined)
})

test('visible line bounds are strict at page edges and do not confuse paragraph start with visibility', () => {
  assert.equal(documentViewportLocation(layout, { ...viewport, scrollTop: 110, height: 20 }).firstVisibleBlockId, 'heading-1')
  assert.equal(documentViewportLocation(layout, { ...viewport, scrollTop: 122, height: 20 }).firstVisibleBlockId, 'paragraph-1')
  assert.equal(documentViewportLocation(layout, { ...viewport, scrollTop: 1160, height: 10 }).firstVisibleBlockId, undefined)
})

test('zoom preserves the reading position rather than jumping to an off-screen text selection', () => {
  const offset = zoomedDocumentScrollOffset(layout, { ...viewport, width: 500, scrollLeft: 100, scrollTop: 900 }, 150)
  assert.deepEqual(offset, { left: 275, top: 1500 })
  const restored = zoomedDocumentScrollOffset(layout, { ...viewport, width: 500, scrollLeft: offset.left, scrollTop: offset.top, zoom: 150 }, 100)
  assert.deepEqual(restored, { left: 100, top: 900 })
})

test('zoom respects centered narrow documents, explicit anchors and scroll limits', () => {
  assert.deepEqual(zoomedDocumentScrollOffset(layout, viewport, 50), { left: 0, top: 0 })
  assert.deepEqual(zoomedDocumentScrollOffset(layout, { ...viewport, width: 500, scrollLeft: 100, scrollTop: 900 }, 200, { x: 20, y: 40 }), { left: 220, top: 1840 })
})

test('caret visibility scrolls only the minimum distance and retains a visible caret', () => {
  assert.deepEqual(documentRectScrollOffset(layout, viewport, { x: 100, y: 100, width: 1, height: 20 }), { left: 0, top: 0 })
  assert.deepEqual(documentRectScrollOffset(layout, { ...viewport, width: 500 }, { x: 650, y: 900, width: 1, height: 20 }), { left: 179, top: 348 })
  assert.deepEqual(documentRectScrollOffset(layout, { ...viewport, scrollTop: 1700 }, { x: 100, y: 1500, width: 1, height: 20 }), { left: 0, top: 1472 })
})

test('oversized targets and document edges stay within the scrollable content', () => {
  const offset = documentRectScrollOffset(layout, viewport, { x: 0, y: 2300, width: 850, height: 400 })
  assert.deepEqual(offset, { left: 0, top: 1750 })
})
