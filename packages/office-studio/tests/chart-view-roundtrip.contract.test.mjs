import assert from 'node:assert/strict'
import { test } from 'node:test'
import { applyExcelWorkbookView, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { createBlankWorkbook } from '../src/domain/model.ts'

function chartModel() {
  return {
    id: 'chart-sales',
    sheetId: 'sheet-sales',
    type: 'line',
    source: { range: { startRow: 0, startCol: 0, endRow: 2, endCol: 1 }, orientation: 'columns' },
    position: { x: 144, y: 200 },
    size: { width: 480, height: 300 },
    title: 'Sales',
    legend: { visible: true, position: 'bottom' },
    categoryField: 0,
    valueFields: [1],
    xAxis: { scaleType: 'linear' },
    yAxis: { scaleType: 'log', logBase: 10 },
    revision: 2,
  }
}

function workbook() {
  const model = createBlankWorkbook()
  model.sheets[0].id = 'sheet-sales'
  model.sheets[0].rowCount = 10
  model.sheets[0].columnCount = 5
  model.charts = [chartModel()]
  return model
}

test('Excel view preserves chart data bindings, style and object layout across edits', () => {
  const original = workbook()
  const view = toExcelWorkbookView(original)
  assert.equal(view.charts?.[0]?.title, 'Sales')
  assert.deepEqual(view.charts?.[0]?.source.range, { startRow: 0, startCol: 0, endRow: 2, endCol: 1 })

  view.charts[0].position.x = 320
  view.charts[0].size.height = 360
  const next = applyExcelWorkbookView(original, view)

  assert.equal(next.charts?.[0]?.position.x, 320)
  assert.equal(next.charts?.[0]?.size.height, 360)
  assert.equal(next.charts?.[0]?.yAxis.logBase, 10)
  assert.equal(original.charts?.[0]?.position.x, 144)
  assert.notEqual(next.charts?.[0], original.charts?.[0])
})

test('legacy views without a charts field do not erase workbook chart objects', () => {
  const original = workbook()
  const view = toExcelWorkbookView(original)
  delete view.charts
  const next = applyExcelWorkbookView(original, view)
  assert.equal(next.charts?.[0]?.id, 'chart-sales')
})
