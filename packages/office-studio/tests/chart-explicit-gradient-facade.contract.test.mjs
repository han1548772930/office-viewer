import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// P16 integration drafts. A server starts only in the deferred validation phase.
let server, api
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})
afterAll(async () => { await server?.close() })

function gradient() {
  return api.chartLinearGradient({ start: { x: -0.2, y: 0 }, end: { x: 1.2, y: 1 },
    stops: [{ offset: 1, color: 'blue', opacity: 0.25 }, { offset: 0, color: 'rgba(255, 0, 0, 0.6)' }] })
}
function setup() {
  const studio = new api.OfficeStudio(), workbook = api.FOffice.newAPI(studio).getActiveWorkbook(), sheet = workbook.getActiveSheet()
  sheet.getRange('A1:C4').setValues([['Month', 'North', 'South'], ['Jan', 10, 15], ['Feb', 20, 18], ['Mar', 30, 24]])
  return { studio, workbook, sheet }
}

test('public gradient construction and metadata-only Builder descriptions detach colors without evaluating source records', () => {
  const { studio, sheet } = setup()
  try {
    const source = gradient(), before = structuredClone(source), builder = sheet.newChart('column').setSource('A1:C4').setValueFields([2, 1]).setAllSeriesStyle({ color: source })
    source.start.x = 0
    source.stops[0].opacity = 1
    const first = builder.describe({ fields: { 0: 'Month', 1: 'North', 2: 'South' } }), second = builder.describe()
    assert.equal(first.id, 'chart-draft')
    assert.equal(second.id, first.id)
    assert.deepEqual(first.series[1].color, before)
    assert.equal(first.series[2].name, 'South')
    first.series[1].color.stops[0].color = 'purple'
    assert.deepEqual(builder.build().options.allSeriesStyle.color, before)
    assert.deepEqual(builder.describe().series[1].color, before)
    const snapshot = builder.build()
    assert.throws(() => builder.setAllSeriesStyle({ color: { type: 'linear-gradient', start: { x: 0, y: 0 } } }), TypeError)
    assert.deepEqual(builder.build(), snapshot)
    assert.throws(() => sheet.newChart('scatter').setAllSeriesStyle({ color: gradient() }), TypeError)
    assert.throws(() => sheet.newChart('line').setSeries(1, { dataPoints: { 0: { color: gradient() } } }), TypeError)
  } finally { studio.dispose() }
})

test('live source colors, nullable reset and undo/redo preserve detached descriptions and internal saved content', async () => {
  const { studio, workbook, sheet } = setup()
  let reopened
  try {
    const color = gradient(), chart = await sheet.insertChart(sheet.newChart('column').setSource('A1:C4').setValueFields([1, 2]).setAutoGradientFill(false).build())
    await chart.setAllSeriesStyle({ color, fillOpacity: 0.5 })
    const options = chart.getOptions(), description = chart.getDescription(), data = chart.getData(), commands = chart.getRenderCommands()
    options.allSeriesStyle.color.stops[0].color = 'lime'
    description.series[1].color.end.y = 0
    color.stops[1].color = 'yellow'
    assert.deepEqual(chart.getOptions().allSeriesStyle.color, gradient())
    assert.deepEqual(chart.getDescription().series[1].color, gradient())
    await chart.setAllSeriesStyle({ color: null })
    assert.equal(chart.getOptions().allSeriesStyle.color, undefined)
    workbook.undo()
    assert.deepEqual(chart.getDescription().series[1].color, gradient())
    workbook.redo()
    assert.equal(chart.getDescription().series[1].color, undefined)
    workbook.undo()
    const snapshot = JSON.parse(JSON.stringify(workbook.save())), serialized = JSON.stringify(snapshot.charts)
    assert.ok(serialized.includes('linear-gradient'))
    for (const field of ['colorStops', 'authored', 'bounds', 'global']) assert.equal(serialized.includes(`"${field}"`), false)
    reopened = new api.OfficeStudio({ workbook: snapshot })
    const restored = api.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet().getChartById(chart.getChartId())
    assert.deepEqual(restored.getOptions().allSeriesStyle.color, gradient())
    assert.deepEqual(restored.getDescription().series, chart.getDescription().series)
    assert.deepEqual(restored.getData(), data)
    assert.deepEqual(restored.getRenderCommands(), commands)
  } finally { reopened?.dispose(); studio.dispose() }
})

test('type switches retain inactive point authoring while new solid-only writes reject without creating history', async () => {
  const { studio, workbook, sheet } = setup()
  try {
    const color = gradient(), chart = await sheet.insertChart(sheet.newChart('column').setSource('A1:C4').setValueFields([1])
      .setSeries(1, { color: 'red', dataPoints: { 0: { color } } }).build())
    await chart.update({ type: 'line' })
    const before = chart.getInfo(), contents = JSON.stringify(workbook.save())
    assert.deepEqual(chart.getOptions().series[1].dataPoints[0].color, color)
    assert.deepEqual(chart.getDescription().series[1].dataPoints[0].color, color)
    await assert.rejects(async () => { await chart.setSeries(1, { dataPoints: { 0: { color } } }) }, TypeError)
    assert.deepEqual(chart.getInfo(), before)
    assert.equal(JSON.stringify(workbook.save()), contents)
    workbook.undo()
    assert.equal(chart.getType(), 'column')
    assert.deepEqual(chart.getOptions().series[1].dataPoints[0].color, color)
    workbook.redo()
    assert.equal(chart.getType(), 'line')
    await chart.update({ type: 'column' })
    assert.deepEqual(chart.getOptions().series[1].dataPoints[0].color, color)
    assert.deepEqual(chart.getData().series[0].values, [10, 20, 30])
  } finally { studio.dispose() }
})
