import assert from 'node:assert/strict'
import { test } from 'vitest'
import { addressToPosition, createPivotChartDraft, pivotChartConfig, pivotChartPatch } from '../src/features/pivot/services/pivot-chart-panel-state.ts'
import { chartPlacementGeometry } from '../src/features/charts/chart-placement.ts'
import { CHART_TYPE_CATALOG } from '../src/features/charts/chart-catalog.ts'

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
  for (const address of ['', 'A0', '1A', 'A1:B2']) assert.throws(() => addressToPosition(address), /有效的放置单元格/)
})

test('pivot chart settings preserve all supported chart types', () => {
  const types = [...new Set(['bar', 'columnStacked', ...CHART_TYPE_CATALOG.map(item => item.value)])]
  for (const type of types) {
    const draft = createPivotChartDraft(undefined, 'pivot-1')
    draft.type = type
    assert.equal(pivotChartConfig(draft).type, type)
    assert.equal(pivotChartPatch(draft).type, type)
  }
})

test('pivot placement uses real hidden and resized axes and clamps the default address for a small worksheet', () => {
  const workbook = { sheets: [{ id: 'sheet', name: 'Small', cells: {}, rowCount: 1, columnCount: 3,
    defaultColumnWidth: 80, columnWidths: { 0: 120 }, hiddenColumns: [1], defaultRowHeight: 30 }] }
  const geometry = chartPlacementGeometry(workbook, 'sheet')
  const draft = createPivotChartDraft(undefined, 'pivot-1', geometry)
  assert.equal(draft.position, 'C1')
  assert.deepEqual(pivotChartConfig(draft, geometry).position, { x: 120, y: 0 })
  draft.position = 'B1'
  assert.deepEqual(pivotChartPatch(draft, geometry).position, { x: 120, y: 0 })
  assert.throws(() => addressToPosition('D1', geometry), /边界/)
})

test('pivot settings preserve precise dragged coordinates until the placement address changes', () => {
  const chart = { id: 'chart', sheetId: 'sheet', pivotTableId: 'pivot', type: 'line3D', source: { range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 }, orientation: 'columns' },
    position: { x: 487.5, y: 25.25 }, size: { width: 500, height: 300 }, legend: { visible: true, position: 'bottom' }, categoryField: 0, valueFields: [1],
    options: { threeD: { view: { rotY: 42 } }, depthAxis: { title: 'Region' }, series: { 1: { fillOpacity: 0.35 } } }, xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, revision: 2 }
  const draft = createPivotChartDraft(chart)
  assert.equal(draft.position, 'F2')
  assert.deepEqual(pivotChartPatch(draft).position, chart.position)
  assert.deepEqual(pivotChartPatch(draft).options, chart.options)
  draft.position = 'G3'
  assert.deepEqual(pivotChartPatch(draft).position, { x: 576, y: 48 })
})
