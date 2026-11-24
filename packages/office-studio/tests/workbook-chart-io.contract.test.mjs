import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// Draft only. Run with the remaining IO contracts during unified validation.
let server, imports, source
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[imports, source] = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/parser/workbookCharts.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/charts/chart-source-compute.ts'),
  ])
})
afterAll(async () => { await server?.close() })
function sheet(id = 'data', name = 'Data') {
  return { id, name, visible: true, cells: {}, mergedRanges: [], rowCount: 100, columnCount: 20,
    defaultRowHeight: 20, defaultColumnWidth: 100, metadata: {} }
}
function authored() {
  return { type: 'line', source: { sheetName: 'Data', orientation: 'columns',
    range: { startRow: 0, startCol: 0, endRow: 2, endCol: 2 } }, categoryField: 0, valueFields: [2],
    options: { mapping: { categoryIndexes: [0, 1] }, plotVisibleOnly: true } }
}

test('native ordinals and custom-error vectors remap to original source field identities', () => {
  const model = sheet(), raw = { index: 0, kind: 'combination', series: [
    { categories: "'Data'!$B$3:$B$4", values: "'Data'!$D$3:$D$4", options: { type: 'column' },
      errorBars: [{ seriesIndex: 0, direction: 'y', type: 'plus', valueType: 'cust', plusFormula: "'Data'!$E$3:$E$4" }] },
    { categories: "'Data'!$B$3:$B$4", values: "'Data'!$C$3:$C$4", options: { type: 'line', rightAxis: true },
      trendlines: [{ seriesIndex: 1, type: 'linear' }] },
  ] }
  const result = imports.normalizeWorkbookCharts([{ charts: [raw] }], [model], [])
  assert.deepEqual(result.unresolved, [])
  const chart = result.charts[0]
  assert.deepEqual(chart.source.range, { startRow: 1, startCol: 1, endRow: 3, endCol: 4 })
  assert.deepEqual(chart.valueFields, [2, 1])
  assert.equal(chart.options.series[2].type, 'column')
  assert.equal(chart.options.series[1].rightAxis, true)
  assert.equal(chart.options.trendlines[0].seriesIndex, 1)
  assert.equal(chart.options.errorBars[0].seriesIndex, 2)
  assert.equal(chart.options.errorBars[0].plusIndex, 3)
})

test('two-cell editAs absolute keeps cell coordinates and sparse EMU geometry', () => {
  const model = sheet(); model.rowHeights = { 0: 30 }; model.columnWidths = { 0: 150 }
  const raw = { index: 0, authored: authored(), drawing: { behavior: 'absolute', fromCell: { row: 2, col: 2 },
    toCell: { row: 5, col: 5 }, fromOffsetEmu: [5 * 9525, 7 * 9525], toOffsetEmu: [15 * 9525, 17 * 9525] } }
  const result = imports.normalizeWorkbookCharts([{ charts: [raw] }], [model], [])
  assert.deepEqual(result.unresolved, [])
  assert.deepEqual(result.charts[0].position, { x: 255, y: 57 })
  assert.deepEqual(result.charts[0].size, { width: 310, height: 70 })
})

test('authored helper ownership requires the matching hidden sheet and A1 marker', () => {
  const model = sheet(), helper = sheet('helper', '__ov_chart_1'); helper.visible = false
  helper.cells['0:0'] = { value: 'office-studio:chart-helper:v1:original-key' }
  const raw = { index: 0, authored: { ...authored(), helperSheets: [{ name: '__ov_chart_1', key: 'original-key' }] } }
  const result = imports.normalizeWorkbookCharts([{ charts: [raw] }, {}], [model, helper], [])
  assert.deepEqual(result.unresolved, [])
  assert.deepEqual(helper.metadata.officeChartHelper, { version: 1, chartId: result.charts[0].id, key: 'original-key' })
  const wrong = sheet('wrong', '__ov_chart_2'); wrong.visible = false; wrong.cells['0:0'] = { value: 'ordinary data' }
  imports.normalizeWorkbookCharts([{ charts: [{ index: 0, authored: { ...authored(), helperSheets: [{ name: wrong.name, key: 'original-key' }] } }] }, {}], [model, wrong], [])
  assert.equal(wrong.metadata.officeChartHelper, undefined)
})

test('literal or disjoint native objects retain a concrete unresolved reason', () => {
  const model = sheet()
  for (const raw of [
    { kind: 'line', series: [{ categories: undefined, values: undefined }] },
    { kind: 'line', series: [{ categories: "'Data'!$A$2:$A$3", values: "'Data'!$B$3:$B$4" }] },
  ]) {
    const result = imports.normalizeWorkbookCharts([{ charts: [raw] }], [model], [])
    assert.equal(result.charts.length, 0)
    assert.equal(result.unresolved.length, 1)
    assert.match(result.unresolved[0].reason, /UNBOUND|DISJOINT/)
  }
})

test('successful imports retain deduplicated native differences beside the editable chart and source DTO', () => {
  const model = sheet(), raw = { index: 4, authored: authored(),
    unsupportedReasons: ['CHART_NATIVE_GRADIENT_PATH_UNMAPPED', 'CHART_NATIVE_PATTERNED_LINE_STROKE', 'CHART_NATIVE_GRADIENT_PATH_UNMAPPED'] }
  model.metadata.source = { charts: [raw] }
  const result = imports.normalizeWorkbookCharts([{ charts: [raw] }], [model], [])
  assert.deepEqual(result.unresolved, [])
  assert.equal(result.charts.length, 1)
  assert.deepEqual(model.metadata.officeChartDifferences, { [result.charts[0].id]: { index: 4,
    reasons: ['CHART_NATIVE_GRADIENT_PATH_UNMAPPED', 'CHART_NATIVE_PATTERNED_LINE_STROKE'] } })
  assert.equal(model.metadata.source.charts[0], raw)
  imports.normalizeWorkbookCharts([{ charts: [{ index: 4, authored: authored() }] }], [model], [])
  assert.equal(Object.hasOwn(model.metadata, 'officeChartDifferences'), false)
  assert.equal(model.metadata.source.charts[0], raw)
})

test('invalid diagnostics produce an unresolved import and never publish partial difference records', () => {
  for (const reasons of [null, 'CHART_NATIVE_GRADIENT_PATH_UNMAPPED', {}, new Array(1), [''], [7], ['CHART_NATIVE_GRADIENT_PATH_UNMAPPED', null]]) {
    const model = sheet(); model.metadata.officeChartDifferences = { old: { index: 0, reasons: ['STALE'] } }
    const result = imports.normalizeWorkbookCharts([{ charts: [{ index: 0, authored: authored(), unsupportedReasons: reasons }] }], [model], [])
    assert.deepEqual(result.charts, [])
    assert.deepEqual(result.unresolved, [{ sheetId: model.id, index: 0, reason: 'CHART_IMPORT_DIAGNOSTICS_INVALID' }])
    assert.equal(model.metadata.officeChartDifferences, undefined)
  }
})

test('visible-only defaults share the source reader and preserve field indexes in both orientations', async () => {
  const model = sheet()
  model.rowCount = 3; model.columnCount = 3; model.hiddenRows = [1]; model.hiddenColumns = [2]
  for (let row = 0; row < 3; row++) for (let column = 0; column < 3; column++) model.cells[row + ':' + column] = { value: row * 10 + column }
  const book = { kind: 'workbook', id: 'book', sheets: [model] }
  const base = { sheetId: 'data', source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 2, endCol: 2 } } }
  const columns = await source.readChartRecordsAsync(book, base)
  assert.deepEqual(columns.visibleFields, [0, 1]); assert.deepEqual(columns.records, [[20, 21, 22]])
  const rows = await source.readChartRecordsAsync(book, { ...base, source: { ...base.source, orientation: 'rows' } })
  assert.deepEqual(rows.visibleFields, [0, 2]); assert.deepEqual(rows.records, [[1, 11, 21]])
  const all = await source.readChartRecordsAsync(book, { ...base, options: { plotVisibleOnly: false } })
  assert.equal(all.visibleFields, undefined); assert.equal(all.records.length, 2)
})
