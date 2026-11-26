import assert from 'node:assert/strict'
import { test } from 'vitest'
import { getRecordValueKey } from '../src/domain/record-value.ts'
import { assertSlicerCanvasWindow, readSlicerCanvasWindow, readSlicerCanvasMetadata,
  copySlicerCanvasMetadata } from '../src/features/slicer/slicer-canvas-data-validation.ts'

// Deferred DTO contracts. These tests exercise retained presentation validation,
// not a second membership authority or any production computation fallback.
const item = (value, label = String(value.type === 'blank' ? '' : value.value), extra = {}) =>
  ({ key: getRecordValueKey(value), value, label, hasData: true, ...extra })
const metadata = () => ({ fieldId: 'field', name: 'Date', ready: true, timeline: { minDate: -2.25, maxDate: 45000.5, interactive: true } })
const invalid = error => error?.code === 'SLICER_CANVAS_INVALID_RESULT'
const budget = error => error?.code === 'SLICER_CANVAS_MEMORY_LIMIT'

test('all six raw types remain distinct and Canvas captions retain exact UTF-16', () => {
  const items = [{ type: 'blank' }, { type: 'string', value: '1' }, { type: 'number', value: 1 }, { type: 'date', value: 1 },
    { type: 'boolean', value: true }, { type: 'error', value: '#N/A' }].map((value, index) => item(value, index === 1 ? '\ud800\0😀' : `caption-${index}`))
  const result = readSlicerCanvasWindow({ total: 6, items }, 0, 6, 10000)
  assert.deepEqual(result.window, { total: 6, items }); assert.equal(new Set(result.window.items.map(item => item.key)).size, 6)
  assert.equal(result.window.items[1].label, '\ud800\0😀'); assert.equal(result.bytes > 0, true)
})

test('different captions may share the same typed key while checked and partial flags stay distinct', () => {
  const raw = { type: 'date', value: 100 }
  const items = [item(raw, '100', { selected: false, partial: true }), item(raw, '1900-04-09', { selected: true, partial: false })]
  assert.deepEqual(readSlicerCanvasWindow({ total: 2, items }, 0, 2, 10000).window.items, items)
})

test('a validated member array owns fresh values and preserves absent optional flags', () => {
  const raw = { total: 1, items: [item({ type: 'string', value: 'source' })] }
  const result = readSlicerCanvasWindow(raw, 0, 2, 10000)
  result.window.items[0].value.value = 'changed'; result.window.items[0].label = 'changed'
  assert.equal(raw.items[0].value.value, 'source'); assert.equal(raw.items[0].label, 'source')
  assert.equal(Object.hasOwn(result.window.items[0], 'selected'), false); assert.equal(Object.hasOwn(result.window.items[0], 'partial'), false)
})

test.each([
  null, { total: -1, items: [] }, { total: NaN, items: [] }, { total: Number.MAX_SAFE_INTEGER + 1, items: [] },
  { total: 2, items: null }, { total: 2, items: [item({ type: 'number', value: 1 })] },
  { total: 2, items: Array(2) }, { total: 1, items: [item({ type: 'number', value: 1 }), item({ type: 'number', value: 2 })] },
  { total: 1, items: [item({ type: 'number', value: 1 }, 'one', { key: 'wrong' })] },
  { total: 1, items: [item({ type: 'number', value: 1 }, 'one', { hasData: 1 })] },
  { total: 1, items: [item({ type: 'number', value: 1 }, 'one', { selected: 'yes' })] },
  { total: 1, items: [item({ type: 'number', value: 1 }, 'one', { selected: true, partial: true })] },
  { total: 1, items: [{ key: '[]', value: { type: 'number', value: Infinity }, label: 'x', hasData: true }] },
  { total: 1, items: [{ key: '[]', value: { type: 'blank', value: null }, label: 'x', hasData: true }] },
  { total: 2, items: [item({ type: 'number', value: 1 }, 'same'), item({ type: 'number', value: 2 }, 'same')] },
])('malformed member window %j is rejected with the result code', input => {
  assert.throws(() => readSlicerCanvasWindow(input, 0, 2, 10000), invalid)
})

test('tail windows, empty sources and reads after the source end have complete lengths', () => {
  assert.deepEqual(readSlicerCanvasWindow({ total: 3, items: [item({ type: 'number', value: 2 })] }, 2, 2, 10000).window.items.length, 1)
  assert.deepEqual(readSlicerCanvasWindow({ total: 0, items: [] }, 0, 2, 10000).window, { total: 0, items: [] })
  assert.deepEqual(readSlicerCanvasWindow({ total: 3, items: [] }, 9, 2, 10000).window, { total: 3, items: [] })
})

test('metadata bounds, date grouping and nested copies retain fractional and negative serial dates', () => {
  const input = { ...metadata(), dateGroup: { type: 3, dateType: 'date', excelLike: { unit: 'day', autoStart: true, autoEnd: false } } }
  const result = readSlicerCanvasMetadata(input, 10000), copied = copySlicerCanvasMetadata(result.metadata)
  copied.timeline.minDate = 0; copied.dateGroup.excelLike.autoEnd = true
  assert.deepEqual(result.metadata, input); assert.equal(input.timeline.minDate, -2.25)
  assert.equal(input.dateGroup.excelLike.autoEnd, false)
  const empty = { fieldId: 'field', name: '', ready: true, timeline: { interactive: false } }
  assert.deepEqual(readSlicerCanvasMetadata(empty, 10000).metadata, empty)
})

test.each([
  null, { ...metadata(), fieldId: '' }, { ...metadata(), name: 1 }, { ...metadata(), ready: false },
  { ...metadata(), timeline: { interactive: true } }, { ...metadata(), timeline: { minDate: 2, interactive: true } },
  { ...metadata(), timeline: { minDate: 2, maxDate: 1, interactive: true } },
  { ...metadata(), timeline: { minDate: NaN, maxDate: 3, interactive: true } },
  { ...metadata(), timeline: { minDate: 0, maxDate: 3, interactive: 'yes' } },
  { ...metadata(), dateGroup: { type: 2, dateType: 'month' } },
  { ...metadata(), dateGroup: { type: 3, dateType: 'day' } },
  { ...metadata(), dateGroup: { type: 3, dateType: 'month', excelLike: { unit: 'year', autoStart: true, autoEnd: true } } },
  { ...metadata(), dateGroup: { type: 3, dateType: 'month', excelLike: { unit: 'month', autoStart: 'yes', autoEnd: true } } },
])('malformed source metadata %j is rejected without a fabricated empty timeline', input => {
  assert.throws(() => readSlicerCanvasMetadata(input, 10000), invalid)
})

test('member and metadata byte admission includes long text and array/object overhead', () => {
  const long = 'x'.repeat(1000)
  assert.throws(() => readSlicerCanvasWindow({ total: 1, items: [item({ type: 'string', value: long })] }, 0, 2, 500), budget)
  assert.throws(() => readSlicerCanvasWindow({ total: 0, items: [] }, 0, 2, 47), budget)
  assert.throws(() => readSlicerCanvasMetadata({ ...metadata(), name: long }, 500), budget)
})

test.each([[-1, 2], [0, 0], [0, 9], [0.5, 2], [Number.MAX_SAFE_INTEGER, 2], [0, Infinity]])(
  'invalid request start=%s count=%s is rejected before allocation', (start, count) => {
    assert.throws(() => assertSlicerCanvasWindow(start, count, 8), error => error?.code === 'SLICER_CANVAS_INVALID_WINDOW')
  })

test('negative zero uses the canonical raw key without changing fractional date values', () => {
  const items = [item({ type: 'number', value: -0 }, 'zero'), item({ type: 'date', value: -2.25 }, 'fraction')]
  const result = readSlicerCanvasWindow({ total: 2, items }, 0, 2, 10000).window
  assert.equal(Object.is(result.items[0].value.value, -0), false); assert.equal(result.items[1].value.value, -2.25)
})
