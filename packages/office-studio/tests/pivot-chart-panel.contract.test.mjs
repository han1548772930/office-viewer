import assert from 'node:assert/strict'
import { test } from 'vitest'
import { addressToPosition, createPivotChartDraft, pivotChartConfig, pivotChartPatch } from '../src/features/pivot/services/pivot-chart-panel-state.ts'

test('pivot chart draft maps linked chart settings and creates a valid facade update patch', () => {
  const chart = {
    id: 'pivot-chart-1', sheetId: 'sheet-1', pivotTableId: 'pivot-1', type: 'line',
    source: { range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 }, orientation: 'columns' },
    position: { x: 480, y: 24 }, size: { width: 720, height: 360 }, title: 'Monthly sales',
    legend: { visible: false, position: 'right' }, categoryField: 0, valueFields: [2],
    xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, revision: 1,
  }
  const draft = createPivotChartDraft(chart)
  assert.equal(draft.position, 'F2')
  assert.equal(draft.pivotTableId, 'pivot-1')
  draft.title = 'Revenue by month'
  draft.legendVisible = true
  draft.legendPosition = 'top'
  const patch = pivotChartPatch(draft)
  assert.equal(patch.type, 'line')
  assert.equal(patch.title, 'Revenue by month')
  assert.deepEqual(patch.position, { x: 480, y: 24 })
  assert.deepEqual(patch.legend, { visible: true, position: 'top' })
  assert.equal(patch.pivotTableId, 'pivot-1')
})

test('pivot chart position rejects invalid spreadsheet addresses', () => {
  assert.deepEqual(addressToPosition('AA12'), { x: 2496, y: 264 })
  for (const address of ['', 'A0', '1A', 'A1:B2']) assert.throws(() => addressToPosition(address), /valid cell address/)
})

test('pivot chart settings preserve all supported chart types', () => {
  const types = ['bar', 'line', 'area', 'pie', 'doughnut', 'stacked-column']
  for (const type of types) {
    const draft = createPivotChartDraft(undefined, 'pivot-1')
    draft.type = type
    assert.equal(pivotChartConfig(draft).type, type)
    assert.equal(pivotChartPatch(draft).type, type)
  }
})
