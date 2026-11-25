import assert from 'node:assert/strict'
import { test } from 'vitest'
import { clampSheetRichTextEditorScroll, stepSheetRichTextEditorDragScroll } from '../src/rendering/spreadsheet/rich-text-editor-scroll.ts'

// Draft only. No test runner, browser or performance verification has been executed.
const extent = () => ({ clip: { x: 20, y: 30, width: 100, height: 80 }, contentWidth: 600, contentHeight: 480 })

test('pointer positions inside the visible clip keep scroll still while all results stay within content extents', () => {
  const frame = extent(), current = { left: 200, top: 150 }
  for (const pointer of [{ x: 20, y: 30 }, { x: 120, y: 110 }, { x: 60, y: 50 }]) {
    assert.deepEqual(stepSheetRichTextEditorDragScroll(frame, current, pointer, 16), current)
  }
  assert.deepEqual(clampSheetRichTextEditorScroll(frame, { left: -10, top: 1000 }), { left: 0, top: 400 })
  assert.deepEqual(clampSheetRichTextEditorScroll({ ...frame, contentWidth: 50, contentHeight: 40 }, current), { left: 0, top: 0 })
})

test('captured pointer scrolling has symmetric directions, distance-based speed and a finite screen-speed cap', () => {
  const frame = extent(), current = { left: 200, top: 150 }
  const right = stepSheetRichTextEditorDragScroll(frame, current, { x: 130, y: 120 }, 16)
  const left = stepSheetRichTextEditorDragScroll(frame, current, { x: 10, y: 20 }, 16)
  assert.ok(Math.abs(right.left - current.left - 3.2) < 1e-9)
  assert.ok(Math.abs(right.top - current.top - 3.2) < 1e-9)
  assert.ok(Math.abs(current.left - left.left - 3.2) < 1e-9)
  assert.ok(Math.abs(current.top - left.top - 3.2) < 1e-9)
  const near = stepSheetRichTextEditorDragScroll(frame, current, { x: 121, y: 50 }, 16)
  const far = stepSheetRichTextEditorDragScroll(frame, current, { x: 10000, y: 50 }, 16)
  assert.ok(near.left > current.left && near.left < right.left)
  assert.ok(Math.abs(far.left - current.left - 19.2) < 1e-9)
  assert.equal(far.top, current.top)
})

test('a stationary pointer advances every frame until the bound and can reverse away from that bound', () => {
  const frame = extent(), pointer = { x: 300, y: 300 }
  let scroll = { left: 0, top: 0 }
  const first = stepSheetRichTextEditorDragScroll(frame, scroll, pointer, 16)
  const second = stepSheetRichTextEditorDragScroll(frame, first, pointer, 16)
  assert.ok(second.left > first.left && second.top > first.top)
  for (let at = 0; at < 100; at++) scroll = stepSheetRichTextEditorDragScroll(frame, scroll, pointer, 16)
  assert.deepEqual(scroll, { left: 500, top: 400 })
  assert.deepEqual(stepSheetRichTextEditorDragScroll(frame, scroll, pointer, 16), scroll)
  const reverse = stepSheetRichTextEditorDragScroll(frame, scroll, { x: -100, y: -100 }, 16)
  assert.ok(reverse.left < scroll.left && reverse.top < scroll.top)
})

test('scroll follows elapsed frame time while late frames receive at most 64 milliseconds of movement', () => {
  const frame = extent(), pointer = { x: 140, y: 130 }, current = { left: 100, top: 100 }
  const one = stepSheetRichTextEditorDragScroll(frame, current, pointer, 16)
  const two = stepSheetRichTextEditorDragScroll(frame, one, pointer, 16)
  const sameDuration = stepSheetRichTextEditorDragScroll(frame, current, pointer, 32)
  assert.ok(Math.abs(two.left - sameDuration.left) < 1e-9)
  assert.ok(Math.abs(two.top - sameDuration.top) < 1e-9)
  assert.deepEqual(stepSheetRichTextEditorDragScroll(frame, current, pointer, 10000),
    stepSheetRichTextEditorDragScroll(frame, current, pointer, 64))
})

test('invalid timing, hidden clips and malformed coordinates cannot create nonfinite view scroll', () => {
  const frame = extent(), current = { left: 20, top: 30 }, pointer = { x: 200, y: 200 }
  for (const elapsed of [0, -1, NaN, Infinity]) assert.deepEqual(stepSheetRichTextEditorDragScroll(frame, current, pointer, elapsed), current)
  assert.deepEqual(stepSheetRichTextEditorDragScroll(frame, current, { x: NaN, y: 200 }, 16), current)
  assert.deepEqual(stepSheetRichTextEditorDragScroll(frame, current, { x: 200, y: Infinity }, 16), current)
  for (const clip of [{ ...frame.clip, width: 0 }, { ...frame.clip, height: -1 }, { ...frame.clip, x: Infinity }]) {
    assert.deepEqual(stepSheetRichTextEditorDragScroll({ ...frame, clip }, current, pointer, 16), current)
  }
  assert.deepEqual(clampSheetRichTextEditorScroll({ ...frame, contentWidth: Infinity, contentHeight: NaN }, { left: Infinity, top: -1 }), { left: 0, top: 0 })
})

