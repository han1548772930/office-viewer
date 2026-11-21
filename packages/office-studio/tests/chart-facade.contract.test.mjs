import assert from 'node:assert/strict'
import { after, before, test } from 'node:test'
import { resolve } from 'node:path'
import { createServer } from 'vite'

let server
let apiModule

before(async () => {
  server = await createServer({
    configFile: resolve(process.cwd(), 'vite.config.ts'),
    root: process.cwd(),
    server: { middlewareMode: true, hmr: false },
    appType: 'custom',
    logLevel: 'error',
  })
  apiModule = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})

after(async () => {
  await server?.close()
})

function setup() {
  const studio = new apiModule.OfficeStudio()
  const officeApi = apiModule.FOffice.newAPI(studio)
  const workbook = officeApi.getActiveWorkbook()
  const sheet = workbook.getActiveSheet()
  sheet.getRange('A1:C4').setValues([
    ['Month', 'North', 'South'],
    ['Jan', 10, 15],
    ['Feb', 20, 18],
    ['Mar', 30, 24],
  ])
  return { studio, workbook, sheet }
}

test('creates editable chart model, reads bound data and emits canvas commands', async () => {
  const { studio, sheet } = setup()
  const chart = await sheet.insertChart(sheet.newChart('bar')
    .setSource('A1:C4')
    .setPosition('E2')
    .setSize(420, 260)
    .setTitle('Revenue')
    .setValueFields([1, 2])
    .build())

  assert.deepEqual(chart.getData(), {
    categories: ['Jan', 'Feb', 'Mar'],
    series: [
      { field: 1, name: 'North', values: [10, 20, 30] },
      { field: 2, name: 'South', values: [15, 18, 24] },
    ],
  })
  const layout = chart.getLayout()
  assert.equal(layout.frame.x, 4 * 96)
  assert.equal(layout.frame.y, 1 * 24)
  assert.equal(layout.series[0].points.length, 3)
  assert.ok(chart.getRenderCommands().some((command) => command.kind === 'rect' && command.fill === '#2563eb'))
  studio.dispose()
})

test('chart data follows source edits, updates are undoable, and invalid updates are atomic', async () => {
  const { studio, workbook, sheet } = setup()
  const chart = await sheet.insertChart(sheet.newChart('line').setSource('A1:C4').setValueFields([1, 2]).build())
  sheet.getRange('B2:C2').setValues([[100, 200]])
  assert.deepEqual(chart.getData().series.map((series) => series.values[0]), [100, 200])

  const before = chart.getInfo()
  await assert.rejects(() => chart.update({ size: { width: 0, height: 100 } }), /Invalid chart size/)
  assert.deepEqual(chart.getInfo(), before)

  await chart.update({ type: 'bar', title: 'Updated' })
  assert.equal(chart.getInfo().type, 'bar')
  workbook.undo()
  assert.equal(sheet.getChartById(chart.getChartId())?.getInfo().type, 'line')
  workbook.redo()
  assert.equal(sheet.getChartById(chart.getChartId())?.getInfo().title, 'Updated')
  await chart.remove()
  assert.equal(sheet.getCharts().length, 0)
  studio.dispose()
})

test('logarithmic axis omits non-positive points without changing source values', async () => {
  const { studio, sheet } = setup()
  sheet.getRange('B2:C4').setValues([[-1, 2], [0, 4], [8, 16]])
  const chart = await sheet.insertChart(sheet.newChart('line').setSource('A1:C4').setYAxis({ scaleType: 'log', logBase: 2 }).setValueFields([1, 2]).build())
  assert.deepEqual(chart.getData().series[0].values, [-1, 0, 8])
  assert.equal(chart.getLayout().series[0].points.length, 1)
  assert.equal(sheet.getRange('B2').getValue(), -1)
  studio.dispose()
})

test('stacked columns accumulate positive and negative values independently', async () => {
  const { studio, sheet } = setup()
  sheet.getRange('B2:C4').setValues([[10, -5], [-4, 8], [2, 3]])
  const chart = await sheet.insertChart(sheet.newChart('stacked-column').setSource('A1:C4').setValueFields([1, 2]).build())
  const layout = chart.getLayout()
  assert.equal(layout.min, -5)
  assert.equal(layout.max, 10)
  assert.equal(layout.series[0].points[0].baseY, layout.series[1].points[0].baseY)
  assert.ok(chart.getRenderCommands().some(command => command.kind === 'rect' && command.fill === '#dc2626'))
  await assert.rejects(() => chart.update({ yAxis: { scaleType: 'log' } }), /linear value axis/)
  studio.dispose()
})

test('area charts split fills at the zero crossing and pie charts draw category slices', async () => {
  const { studio, sheet } = setup()
  sheet.getRange('B2:B4').setValues([[-10], [20], [30]])
  const area = await sheet.insertChart(sheet.newChart('area').setSource('A1:B4').setValueFields([1]).build())
  const areaPolygons = area.getRenderCommands().filter(command => command.kind === 'polygon')
  assert.equal(areaPolygons.length, 3)
  assert.ok(areaPolygons[0].points.length === 3 && areaPolygons[1].points.length === 3)

  sheet.getRange('B2:B4').setValues([[10], [20], [30]])
  const pie = await sheet.insertChart(sheet.newChart('pie').setSource('A1:C4').setValueFields([1]).build())
  assert.ok(pie.getRenderCommands().some(command => command.kind === 'polygon' && command.fill === '#2563eb'))
  const doughnut = await sheet.insertChart(sheet.newChart('doughnut').setSource('A1:C4').setValueFields([1]).build())
  assert.ok(doughnut.getRenderCommands().some(command => command.kind === 'polygon' && command.fill === '#ffffff'))
  await assert.rejects(() => sheet.insertChart(sheet.newChart('pie').setSource('A1:C4').setValueFields([1, 2]).build()), /exactly one value series/)
  studio.dispose()
})
