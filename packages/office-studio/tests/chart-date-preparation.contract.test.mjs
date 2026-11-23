import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { chartDateSerial } from '../src/features/charts/chart-date-axis.ts'
import { prepareChartComputeInputs } from '../src/features/charts/chart-input-preparation.ts'
import { estimateChartDataMemory } from '../src/features/charts/chart-compute-memory.ts'
import { readChartRecordsAsync } from '../src/features/charts/chart-source-compute.ts'
import { projectChartRecords } from '../src/features/charts/chart-record-projection.ts'
import { prepareChartStatisticsTask } from '../src/features/charts/chart-statistics-compute.ts'

// Drafts for the unified verification phase; these have not been executed.
function chart(options = {}) {
  return { id: 'date-chart', sheetId: 'data', type: 'line', revision: 0,
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 3, endCol: 1 } },
    categoryField: 0, valueFields: [1], position: { x: 0, y: 0 }, size: { width: 400, height: 280 },
    legend: { visible: true, position: 'bottom' }, xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' },
    options: { mapping: { useDateAxis: true }, ...options } }
}

test('IO and workbench date preparation use the same epoch, pseudo-date and strict-text fixture', () => {
  const fixtures = JSON.parse(readFileSync(new URL('../../../tests/fixtures/charts/date-serials.json', import.meta.url), 'utf8'))
  for (const fixture of fixtures) assert.equal(chartDateSerial(fixture.value, fixture.date1904), fixture.serial, JSON.stringify(fixture))
})

test('equivalent raw dates survive category aggregation and feed elapsed serials to Worker trends', async () => {
  const source = { date1904: true, categories: ['Day one', 'Day one', 'Day three'], categoryValues: [0, '1904-01-01', '1904-01-03'],
    fields: ['Date', 'Value'], records: [[0, 2], ['1904-01-01', 3], ['1904-01-03', 9]],
    series: [{ field: 1, name: 'Value', values: [2, 3, 9], xValues: [100, 200, 300] }] }
  const model = chart({ aggregation: { aggregate: true }, trendlines: [{ seriesIndex: 1, type: 'linear' }] })
  const data = await prepareChartComputeInputs(source, model)
  assert.equal(data.date1904, true)
  assert.deepEqual(data.categoryValues, [0, '1904-01-03'])
  assert.deepEqual(data.series[0].values, [5, 9])
  assert.deepEqual(data.series[0].xValues, [0, 2])
  assert.equal(data.records, undefined)
  const plan = await prepareChartStatisticsTask(data, model)
  assert.equal(plan.task.params.rowCount, 2)
  const xColumn = plan.task.params.series[0].xColumn
  assert.deepEqual([...plan.task.values.slice(xColumn * 2, (xColumn + 1) * 2)], [0, 2])
  assert.deepEqual([...plan.task.states.slice(xColumn * 2, (xColumn + 1) * 2)], [1, 1])
})

test('Top-N keeps source coordinates and a synthetic Others has no one-row record or date identity', async () => {
  const source = { categories: ['A', 'B', 'C'], categoryValues: [59, 60, 61], records: [[59, 2], [60, 9], [61, 3]],
    series: [{ field: 1, name: 'Value', values: [2, 9, 3], xValues: [10, 20, 30], sizeValues: [1, 2, 3] }] }
  const options = { mapping: { useDateAxis: false }, aggregation: { topN: { count: 2, includeOthers: true, othersLabel: 'Rest' } } }
  const output = await prepareChartComputeInputs(source, chart(options))
  assert.deepEqual(output.categories, ['B', 'C', 'Rest'])
  assert.deepEqual(output.categoryValues, [60, 61, 'Rest'])
  assert.deepEqual(output.series[0].xValues, [20, 30, null])
  assert.deepEqual(output.series[0].sizeValues, [2, 3, null])
  assert.equal(output.records, undefined)
  await assert.rejects(() => prepareChartComputeInputs(source, chart({ ...options, mapping: { useDateAxis: true } })), /无法作为日期轴值/)
})

test('conflicting coordinates and dates cannot masquerade as one aggregated observation', async () => {
  const source = { categories: ['Same', 'Same'], categoryValues: [59, 60],
    series: [{ field: 1, name: 'Value', values: [1, 2], xValues: [10, 20] }] }
  const output = await prepareChartComputeInputs(source, chart({ mapping: { useDateAxis: false }, aggregation: { aggregate: true } }))
  assert.deepEqual(output.series[0].xValues, [null])
  assert.deepEqual(output.categoryValues, ['Same'])
  await assert.rejects(() => prepareChartComputeInputs(source, chart({ aggregation: { aggregate: true } })), /无法作为日期轴值/)
})

test('source projection carries the workbook date system and its raw dates in both orientations', async () => {
  for (const orientation of ['columns', 'rows']) {
    const matrix = orientation === 'columns' ? [['Date', 'Value'], ['1904-01-01', 2], [1, 4]]
      : [['Date', '1904-01-01', 1], ['Value', 2, 4]]
    const workbook = createBlankWorkbook()
    workbook.date1904 = true
    const sheet = workbook.sheets[0]
    sheet.id = 'data'; sheet.rowCount = matrix.length; sheet.columnCount = matrix[0].length
    sheet.cells = Object.fromEntries(matrix.flatMap((row, r) => row.map((value, c) => [r + ':' + c, createCell(value)])))
    const model = chart()
    model.source = { orientation, range: { startRow: 0, startCol: 0, endRow: matrix.length - 1, endCol: matrix[0].length - 1 } }
    const records = await readChartRecordsAsync(workbook, model)
    assert.equal(records.date1904, true)
    const data = await prepareChartComputeInputs(projectChartRecords(records, model), model)
    assert.deepEqual(data.categoryValues, ['1904-01-01', 1])
    assert.deepEqual(data.categoryDateValues, [0, 1])
  }
})

test('date category lineage and serial coordinates participate in chart admission memory estimates', async () => {
  const base = { categories: ['A'], series: [{ field: 1, name: 'Value', values: [1] }] }
  const baseline = await estimateChartDataMemory(base)
  const decorated = await estimateChartDataMemory({ ...base, categoryValues: ['1904-01-01'], categoryDateValues: [0] })
  assert.ok(decorated >= baseline + 32 + 16 + 32 + 20 + 32 + 16)
})
