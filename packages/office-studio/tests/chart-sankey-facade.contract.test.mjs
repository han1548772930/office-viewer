import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// P20 integration drafts. No server, runner, build or browser has started.
let server, api
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})
afterAll(async () => { await server?.close() })

function setup() {
  const studio = new api.OfficeStudio(), workbook = api.FOffice.newAPI(studio).getActiveWorkbook(), sheet = workbook.getActiveSheet()
  sheet.getRange('A1:C6').setValues([['Source', 'Target', 'Weight'], ['R', 'A', 10], ['A', 'B', 4], ['A', 'C', 6], ['B', 'D', 4], ['C', 'E', 6]])
  const builder = sheet.newChart('sankey').setSource('A1:C6').setCategoryField(0).setValueFields([1, 2])
    .setMapping({ edge: { sourceIndex: 0, targetIndex: 1, valueIndex: 2 } }).setSize(520, 340)
  return { studio, workbook, sheet, builder }
}

test('Sankey uses the base builder for detached mappings, inherited node styles, labels and description', () => {
  const { studio, builder } = setup()
  try {
    const mapping = { edge: { sourceIndex: 0, targetIndex: 1, valueIndex: 2 } }
    const style = { fillOpacity: 0.6, color: 'red', label: { visible: true, contentType: 14, position: 'right', format: '0.00' } }
    const point = { color: 'green', border: { color: 'black', width: 2 } }
    assert.equal(builder.setMapping(mapping), builder)
    assert.equal(builder.setAllSeriesStyle(style), builder)
    assert.equal(builder.setSeries(2, { color: 'blue' }), builder)
    assert.equal(builder.setSeries({ name: 'Weight' }, { dataPoints: { 2: point } }), builder)
    const snapshot = builder.build(), description = builder.describe({ fields: { 0: 'Source', 1: 'Target', 2: 'Weight' } })
    mapping.edge.targetIndex = 0; style.label.position = 'inside'; point.border.width = 7
    assert.deepEqual(builder.build(), snapshot)
    assert.equal(description.type, 'sankey')
    assert.deepEqual(description.mapping.edge, { sourceIndex: 0, targetIndex: 1, valueIndex: 2 })
    description.allSeriesStyle.label.visible = false
    assert.equal(builder.build().options.allSeriesStyle.label.visible, true)
    builder.setSeries({ name: 'Weight' }, { dataPoints: { 2: { border: { opacity: 0.4 } } } })
    assert.deepEqual(builder.build().options.namedSeries[0].style.dataPoints[2].border, { color: 'black', width: 2, opacity: 0.4 })
    const beforeInvalid = builder.build()
    assert.throws(() => builder.setAllSeriesStyle({ label: { contentType: 64 } }), TypeError)
    assert.throws(() => builder.setSeries(2, { color: api.chartLinearGradient({ start: { x: 0, y: 0 }, end: { x: 1, y: 0 },
      stops: [{ offset: 0, color: 'red' }, { offset: 1, color: 'blue' }] }) }), TypeError)
    assert.deepEqual(builder.build(), beforeInvalid)
    assert.equal(builder.clearSeries({ name: 'Weight' }), builder)
    assert.equal(builder.build().options.namedSeries.length, 0)
  } finally { studio.dispose() }
})

test('live Sankey node styles and source changes share chart history and preserve author-only JSON through reopen', async () => {
  const { studio, workbook, sheet, builder } = setup()
  let reopened
  try {
    const chart = await sheet.insertChart(builder.setLegend({ visible: false }).build())
    const baseline = chart.getInfo(), baselineData = chart.getData(), baselineLayout = chart.getLayout()
    assert.deepEqual(baselineLayout.sankey.nodes.map(node => node.value), [10, 10, 4, 6, 4, 6])
    const style = { color: 'rgba(255, 0, 0, 0.8)', fillOpacity: 0.5,
      label: { visible: true, contentType: 14, position: 'right', format: '0.00', family: 'Georgia' } }
    assert.equal(await chart.setAllSeriesStyle(style), chart)
    assert.equal(await chart.setSeries({ name: 'Weight' }, { dataPoints: { 2: { color: '#00ff00' } } }), chart)
    const styled = chart.getInfo(), commands = chart.getRenderCommands(), options = chart.getOptions()
    assert.equal(commands.some(command => command.kind === 'text' && command.text === 'Weight,B,4.00'), true)
    assert.equal(commands.some(command => command.kind === 'rect' && command.fill === 'rgba(0, 255, 0, 0.5)'), true)
    style.label.family = 'Arial'; options.namedSeries[0].style.dataPoints[2].color = 'yellow'
    assert.deepEqual(chart.getInfo(), styled)
    workbook.undo()
    assert.equal(chart.getOptions().namedSeries?.length ?? 0, 0)
    workbook.redo()
    assert.deepEqual(chart.getInfo(), styled)
    sheet.getRange('C3').setValue(8)
    const changedData = chart.getData(), changedLayout = chart.getLayout()
    assert.equal(changedData.graph.nodes.find(node => node.name === 'A').value, 14)
    assert.equal(changedLayout.sankey.nodes.find(node => node.name === 'B').value, 8)
    assert.deepEqual(chart.getInfo(), styled)
    workbook.undo()
    assert.deepEqual(chart.getData(), baselineData)
    workbook.redo()
    assert.deepEqual(chart.getData(), changedData)
    const contents = JSON.parse(JSON.stringify(workbook.save())), authored = JSON.stringify(contents.charts)
    for (const field of ['nodeArea', 'band', 'feedbackY', 'hitIndex', 'topology', 'contentDrag', 'sourceVersion']) assert.equal(authored.includes(`"${field}"`), false)
    reopened = new api.OfficeStudio({ workbook: contents })
    const restored = api.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet().getChartById(chart.getChartId())
    assert.deepEqual(restored.getInfo(), styled)
    assert.deepEqual(restored.getData(), changedData)
    assert.deepEqual(restored.getLayout(), changedLayout)
    assert.deepEqual(restored.getRenderCommands(), chart.getRenderCommands())
    await chart.clearAllSeriesStyle()
    assert.equal(chart.getOptions().allSeriesStyle, undefined)
    workbook.undo()
    assert.deepEqual(chart.getInfo(), styled)
    assert.deepEqual(baseline.options, builder.build().options)
  } finally { reopened?.dispose(); studio.dispose() }
})

test('invalid live graph mappings and negative-weight insertions reject without publishing chart history', async () => {
  const { studio, workbook, sheet, builder } = setup()
  try {
    const chart = await sheet.insertChart(builder.build()), baseline = chart.getInfo(), contents = JSON.stringify(workbook.save())
    for (const mapping of [{ edge: { sourceIndex: 0, targetIndex: 0, valueIndex: 2 } },
      { edge: { sourceIndex: 0, targetIndex: 1, valueIndex: 10 } }]) {
      await assert.rejects(async () => { await chart.setMapping(mapping) }, RangeError)
    }
    await assert.rejects(async () => { await chart.setSeries({ name: 'Weight' }, { fillOpacity: 2 }) }, RangeError)
    assert.deepEqual(chart.getInfo(), baseline)
    assert.equal(JSON.stringify(workbook.save()), contents)
    workbook.undo()
    assert.equal(sheet.getCharts().some(candidate => candidate.getChartId() === chart.getChartId()), false)
    sheet.getRange('C3').setValue(-1)
    const invalidSource = JSON.stringify(workbook.save())
    await assert.rejects(async () => { await sheet.insertChart(builder.build()) }, RangeError)
    assert.equal(JSON.stringify(workbook.save()), invalidSource)
  } finally { studio.dispose() }
})
