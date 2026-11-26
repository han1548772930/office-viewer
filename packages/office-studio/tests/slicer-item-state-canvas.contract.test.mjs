import assert from 'node:assert/strict'
import { test } from 'vitest'
import { slicerControlItemStates } from '../src/features/slicer/slicer-control-selection.ts'
import { drawingAffinePoint } from '../src/domain/drawing-affine.ts'
import { toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { SpreadsheetGeometry } from '../src/rendering/spreadsheet/layout.ts'
import { hitTestSlicerObjects, slicerObjectDrawCommands, slicerObjectLayouts } from '../src/features/slicer/slicer-renderer.ts'
import { slicerDrawingWorkbook } from './fixtures/slicer-drawing-model-fixture.mjs'

// Deferred contracts for the real shared item planner and Canvas geometry; none have run.
const items = ['1', 'true', 'Other']
const input = (index, extra = {}) => ({ index, value: items[index], ctrlKey: false, metaKey: false,
  shiftKey: false, multiSelect: false, ...extra })

test.each(['ctrlKey', 'metaKey', 'multiSelect'])('%s toggle preserves another caption partial membership', modifier => {
  const state = { checkedItems: ['Other'], partialItems: ['1', 'true'] }, before = structuredClone(state)
  assert.deepEqual(slicerControlItemStates(items, state, input(2, { [modifier]: true })),
    { checkedItems: [], partialItems: ['1', 'true'] })
  assert.deepEqual(slicerControlItemStates(items, state, input(0, { [modifier]: true })),
    { checkedItems: ['1', 'Other'], partialItems: ['true'] })
  assert.deepEqual(state, before)
})

test('ordinary selection and inclusive Shift replace touched captions with full member state', () => {
  const state = { checkedItems: ['Other'], partialItems: ['1', 'true'] }
  assert.deepEqual(slicerControlItemStates(items, state, input(1)), { checkedItems: ['true'], partialItems: [] })
  for (const [anchorIndex, index] of [[0, 2], [2, 0]]) assert.deepEqual(
    slicerControlItemStates(items, state, input(index, { anchorIndex, shiftKey: true, ctrlKey: true })),
    { checkedItems: items, partialItems: [] })
})

test('unchecking the last item preserves an empty filter and obsolete captions cannot resurrect', () => {
  assert.deepEqual(slicerControlItemStates(items, { checkedItems: ['1'], partialItems: [] }, input(0, { multiSelect: true })),
    { checkedItems: [], partialItems: [] })
  assert.deepEqual(slicerControlItemStates(items, { checkedItems: ['removed'], partialItems: ['removed', '1'] }, input(2, { multiSelect: true })),
    { checkedItems: ['Other'], partialItems: ['1'] })
})

test.each([{ index: -1 }, { index: 3 }, { index: .5 }, { value: 'removed' }, { anchorIndex: -1 }, { anchorIndex: Infinity }])
  ('stale index/value/anchor %j cannot generate an item state', change => {
    assert.throws(() => slicerControlItemStates(items, { checkedItems: ['Other'], partialItems: ['1'] }, input(0, change)), /项目|锚点/)
  })

test('contradictory full and partial state is rejected before any selection is returned', () => {
  assert.throws(() => slicerControlItemStates(items, { checkedItems: ['1'], partialItems: ['1'] }, input(2)), /同时/)
})

test.each([false, true])('legacy/affine=%s layout exposes distinct full, partial and empty item states without changing hit geometry', affine => {
  const workbook = slicerDrawingWorkbook({ affine }), model = workbook.slicers.find(control => control.id === 's')
  model.selectedItems = ['1', 'true']
  const geometry = new SpreadsheetGeometry(toExcelWorkbookView(workbook).sheets[0])
  const full = slicerObjectLayouts(geometry, [{ model, items }], false)[0]
  const partial = slicerObjectLayouts(geometry, [{ model, items, itemState: { checkedItems: ['true'], partialItems: ['1'] } }], false)[0]
  assert.deepEqual(partial.items.map(item => [item.value, item.selected, item.partial]),
    [['1', false, true], ['true', true, false], ['Other', false, false]])
  assert.deepEqual(partial.items.map(item => item.rect), full.items.map(item => item.rect))
  const item = partial.items[0], point = drawingAffinePoint(partial.matrix, item.rect.x + item.rect.width / 2, item.rect.y + item.rect.height / 2)
  assert.deepEqual(hitTestSlicerObjects([partial], point.x, point.y), { type: 'item', id: 's', value: '1', index: 0 })
  const canvas = slicerObjectDrawCommands([partial], { controlUI: 'canvas' })
  const ant = slicerObjectDrawCommands([partial], { controlUI: 'antd' })
  assert.ok(canvas.some(command => command.kind === 'line' && command.width === 2))
  assert.ok(!ant.some(command => command.kind === 'text' && command.text === '1'))
})
