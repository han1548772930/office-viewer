import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// P17 integration drafts. The shared Facade owner wires the six Pareto operations.
// No server or test runner has started; these fixtures await the validation phase.
let server, api
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})
afterAll(async () => { await server?.close() })

function gradient() {
  return api.chartLinearGradient({ start: { x: -0.2, y: 0 }, end: { x: 1.2, y: 1 },
    stops: [{ offset: 1, color: 'blue', opacity: 0.4 }, { offset: 0, color: 'rgba(255, 0, 0, 0.8)' }] })
}
function setup() {
  const studio = new api.OfficeStudio(), workbook = api.FOffice.newAPI(studio).getActiveWorkbook(), sheet = workbook.getActiveSheet()
  sheet.getRange('A1:B6').setValues([['Cause', 'Count'], ['A', -5], ['B', 9], ['A', 8], ['zero', 0], ['negative', -1]])
  return { studio, workbook, sheet }
}

test('the six Pareto Builder operations chain, deep-merge detached styles and restore independent defaults', () => {
  const { studio, sheet } = setup()
  try {
    const color = gradient(), point = { size: 11, shape: 'diamond' }, builder = sheet.newChart('pareto').setSource('A1:B6').setCategoryField(0).setValueFields([1])
    assert.equal(builder.setBarStyle({ color, opacity: 0.7, border: { color: 'black', width: 3 }, label: { visible: true } }), builder)
    assert.equal(builder.setCumulativeLineStyle({ color, width: 4, opacity: 0.5, point }), builder)
    assert.equal(builder.setIncludeZeroValues(true), builder)
    const initial = builder.build()
    color.start.x = 0
    point.size = 1
    assert.deepEqual(builder.build(), initial)
    builder.setBarStyle({ border: { opacity: 0.4 } }).setCumulativeLineStyle({ width: 0, point: { color: 'green' } })
    assert.deepEqual(builder.build().options.pareto.barStyle.border, { color: 'black', width: 3, opacity: 0.4 })
    assert.equal(builder.build().options.pareto.cumulativeLineStyle.width, 4)
    assert.deepEqual(builder.build().options.pareto.cumulativeLineStyle.point, { size: 11, shape: 'diamond', color: 'green' })
    builder.setCumulativeLineStyle({ width: -1 }).setCumulativeLineStyle({ width: null, point: { size: null } })
    assert.equal(Object.hasOwn(builder.build().options.pareto.cumulativeLineStyle, 'width'), true)
    assert.equal(builder.build().options.pareto.cumulativeLineStyle.width, undefined)
    const description = builder.describe({ fields: { 0: 'Cause', 1: 'Count' } })
    assert.deepEqual(description.pareto.barStyle.color, gradient())
    assert.deepEqual(description.series[1], { name: 'Count', type: 'column' })
    assert.equal(Object.keys(description.series).length, 1)
    description.pareto.barStyle.color.stops[0].color = 'pink'
    assert.deepEqual(builder.build().options.pareto.barStyle.color, gradient())
    assert.equal(builder.clearBarStyle(), builder)
    assert.equal(builder.clearCumulativeLineStyle(), builder)
    assert.equal(builder.resetIncludeZeroValues(), builder)
    assert.equal(builder.build().options.pareto.barStyle, undefined)
    assert.equal(builder.build().options.pareto.cumulativeLineStyle, undefined)
    assert.equal(builder.build().options.pareto.includeZeroValues, undefined)
    const snapshot = builder.build()
    assert.throws(() => builder.setCumulativeLineStyle({ label: { position: 'outside' } }), TypeError)
    assert.throws(() => builder.setIncludeZeroValues('true'), TypeError)
    assert.deepEqual(builder.build(), snapshot)
    assert.throws(() => sheet.newChart('column').setBarStyle({ color: 'red' }), TypeError)
  } finally { studio.dispose() }
})

test('live Pareto style/zero resets participate in history and preserve signed source data through internal reopen', async () => {
  const { studio, workbook, sheet } = setup()
  let reopened
  try {
    const palette = ['#112233', '#223344', '#334455', '#445566'], chart = await sheet.insertChart(sheet.newChart('pareto').setSource('A1:B6')
      .setCategoryField(0).setValueFields([1]).setPalette(palette).setIncludeZeroValues(true).build())
    assert.deepEqual(chart.getLayout().categories, ['B', 'A', 'zero'])
    const bar = { color: gradient(), opacity: 0.6, label: { visible: true, contentType: 'percentage' } }
    const line = { color: gradient(), width: 6, opacity: 0.4, point: { shape: 'diamond', size: 11 }, label: { visible: true, contentType: 'percentage' } }
    await chart.setBarStyle(bar)
    await chart.setCumulativeLineStyle(line)
    await chart.setCumulativeLineStyle({ width: 0, point: { size: 12 } })
    assert.equal(chart.getOptions().pareto.cumulativeLineStyle.width, 6)
    assert.equal(chart.getOptions().pareto.cumulativeLineStyle.point.size, 12)
    await chart.setIncludeZeroValues(false)
    assert.deepEqual(chart.getLayout().categories, ['B', 'A'])
    workbook.undo()
    assert.deepEqual(chart.getLayout().categories, ['B', 'A', 'zero'])
    workbook.redo()
    assert.deepEqual(chart.getLayout().categories, ['B', 'A'])
    const options = chart.getOptions(), description = chart.getDescription(), paint = chart.getRenderCommands(), raw = chart.getData()
    assert.deepEqual(raw.series[0].values, [-5, 9, 8, 0, -1])
    bar.color.stops[0].color = 'lime'
    line.point.size = 1
    options.pareto.barStyle.color.end.x = 0
    description.pareto.cumulativeLineStyle.color.stops[0].color = 'yellow'
    assert.deepEqual(chart.getOptions().pareto.barStyle.color, gradient())
    assert.equal(chart.getOptions().pareto.cumulativeLineStyle.point.size, 12)
    const snapshot = JSON.parse(JSON.stringify(workbook.save())), serialized = JSON.stringify(snapshot.charts)
    assert.ok(serialized.includes('cumulativeLineStyle'))
    for (const field of ['authored', 'bounds', 'colorStops']) assert.equal(serialized.includes(`"${field}"`), false)
    reopened = new api.OfficeStudio({ workbook: snapshot })
    const restored = api.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet().getChartById(chart.getChartId())
    assert.deepEqual(restored.getOptions(), chart.getOptions())
    assert.deepEqual(restored.getDescription(), chart.getDescription())
    assert.deepEqual(restored.getData(), raw)
    assert.deepEqual(restored.getRenderCommands(), paint)
    await chart.clearCumulativeLineStyle()
    assert.equal(chart.getOptions().pareto.cumulativeLineStyle, undefined)
    assert.equal(chart.getLayout().series.at(-1).color, palette[3])
    assert.equal(chart.getLayout().series.at(-1).points.length, 2)
    workbook.undo()
    assert.equal(chart.getOptions().pareto.cumulativeLineStyle.width, 6)
    workbook.redo()
    assert.equal(chart.getOptions().pareto.cumulativeLineStyle, undefined)
    workbook.undo()
    await chart.clearBarStyle()
    assert.equal(chart.getOptions().pareto.barStyle, undefined)
    assert.equal(chart.getLayout().series[0].color, palette[0])
    workbook.undo()
    assert.deepEqual(chart.getOptions().pareto.barStyle.color, gradient())
    await chart.resetIncludeZeroValues()
    assert.equal(chart.getOptions().pareto.includeZeroValues, undefined)
    assert.deepEqual(chart.getLayout().categories, ['B', 'A'])
    workbook.undo()
    assert.equal(chart.getOptions().pareto.includeZeroValues, false)
    assert.deepEqual(chart.getData(), raw)
  } finally { reopened?.dispose(); studio.dispose() }
})

test('invalid Pareto live operations reject atomically and wrong-type set/clear/reset operations publish no history', async () => {
  const { studio, workbook, sheet } = setup()
  try {
    const chart = await sheet.insertChart(sheet.newChart('pareto').setSource('A1:B6').setCategoryField(0).setValueFields([1])
      .setBarStyle({ color: 'purple' }).setCumulativeLineStyle({ width: 4 }).build())
    const before = chart.getInfo(), contents = JSON.stringify(workbook.save())
    for (const operation of [() => chart.setBarStyle({ opacity: 2 }), () => chart.setCumulativeLineStyle({ label: { position: 'outside' } }),
      () => chart.setCumulativeLineStyle({ point: { color: gradient() } }), () => chart.setCumulativeLineStyle({ width: NaN }),
      () => chart.setIncludeZeroValues('true')]) await assert.rejects(async () => { await operation() }, error => error instanceof TypeError || error instanceof RangeError)
    assert.deepEqual(chart.getInfo(), before)
    assert.equal(JSON.stringify(workbook.save()), contents)
    await chart.update({ type: 'column' })
    const inactive = JSON.stringify(workbook.save())
    for (const operation of [() => chart.setBarStyle({ color: 'red' }), () => chart.clearBarStyle(), () => chart.setCumulativeLineStyle({ width: 3 }),
      () => chart.clearCumulativeLineStyle(), () => chart.setIncludeZeroValues(true), () => chart.resetIncludeZeroValues()]) {
      await assert.rejects(async () => { await operation() }, TypeError)
    }
    assert.equal(JSON.stringify(workbook.save()), inactive)
    workbook.undo()
    assert.equal(chart.getType(), 'pareto')
    assert.deepEqual(chart.getInfo(), before)
    workbook.undo()
    assert.equal(sheet.getCharts().some(candidate => candidate.getChartId() === chart.getChartId()), false)
  } finally { studio.dispose() }
})
