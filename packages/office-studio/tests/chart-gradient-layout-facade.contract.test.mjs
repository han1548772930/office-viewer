import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// P15 integration drafts. The SSR server starts only when future validation runs.
// This implementation phase has not run servers, tests, browsers or file samples.
let server, api, plotInteraction
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
  plotInteraction = await server.ssrLoadModule('/packages/office-studio/src/features/charts/chart-plot-area-interaction.ts')
})
afterAll(async () => { await server?.close() })

const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-7, `${actual} ≈ ${expected}`)
const manual = { left: 0.2, top: 0.25, width: 0.6, height: 0.5 }
function setup() {
  const studio = new api.OfficeStudio(), workbook = api.FOffice.newAPI(studio).getActiveWorkbook(), sheet = workbook.getActiveSheet()
  sheet.getRange('A1:C4').setValues([['Month', 'North', 'South'], ['Jan', 10, 15], ['Feb', 20, 18], ['Mar', 30, 24]])
  return { studio, workbook, sheet }
}

test('Builder/FChart gradient set/reset and effective descriptions share history and chart-type defaults', async () => {
  const { studio, workbook, sheet } = setup()
  try {
    const builder = sheet.newChart('area').setSource('A1:C4').setValueFields([1]).setAutoGradientFill(false)
    assert.equal(builder.build().options.autoGradientFill, false)
    const before = builder.build().options
    assert.throws(() => builder.setAutoGradientFill('false'), TypeError)
    assert.deepEqual(builder.build().options, before)
    builder.resetAutoGradientFill()
    assert.equal(builder.build().options.autoGradientFill, undefined)
    const chart = await sheet.insertChart(builder.build())
    assert.equal(chart.getDescription().autoGradientFill, true)
    await chart.setAutoGradientFill(false)
    assert.equal(chart.getDescription().autoGradientFill, false)
    workbook.undo()
    assert.equal(chart.getDescription().autoGradientFill, true)
    assert.equal(chart.getOptions().autoGradientFill, undefined)
    workbook.redo()
    assert.equal(chart.getDescription().autoGradientFill, false)
    await chart.resetAutoGradientFill()
    assert.equal(chart.getDescription().autoGradientFill, true)
    workbook.undo()
    assert.equal(chart.getOptions().autoGradientFill, false)
    workbook.redo()
    assert.equal(chart.getOptions().autoGradientFill, undefined)
    await chart.update({ type: 'column' })
    assert.equal(chart.getDescription().autoGradientFill, false)
    const valid = chart.getInfo()
    for (const value of [0, 1, null, {}, 'true']) await assert.rejects(async () => { await chart.setAutoGradientFill(value) }, TypeError)
    assert.deepEqual(chart.getInfo(), valid)
    workbook.undo()
    assert.equal(chart.getInfo().type, 'area')
    assert.equal(chart.getDescription().autoGradientFill, true)
  } finally { studio.dispose() }
})

test('normalized plot setOptions/reset detach inputs, resize content and reject incomplete commits atomically', async () => {
  const { studio, workbook, sheet } = setup()
  try {
    const source = { ...manual }, builder = sheet.newChart('column').setSource('A1:C4').setValueFields([1, 2]).setSize(480, 300).setOptions({ plotArea: source })
    source.left = 0
    assert.deepEqual(builder.build().options.plotArea, manual)
    assert.throws(() => builder.setOptions({ plotArea: { left: 0.3 } }), RangeError)
    assert.deepEqual(builder.build().options.plotArea, manual)
    const chart = await sheet.insertChart(builder.build()), initial = chart.getLayout().plot, valid = chart.getInfo()
    for (const plotArea of [{ left: 0.3 }, { ...manual, width: 1 }, { ...manual, height: 0 }, { ...manual, top: Infinity }]) {
      await assert.rejects(async () => { await chart.setOptions({ plotArea }) }, RangeError)
      assert.deepEqual(chart.getInfo(), valid)
    }
    const description = chart.getDescription()
    description.plotArea.left = 0
    assert.equal(chart.getOptions().plotArea.left, 0.2)
    await chart.update({ size: { width: 960, height: 600 } })
    near(chart.getLayout().plot.width, initial.width * 2)
    near(chart.getLayout().plot.height, initial.height * 2)
    workbook.undo()
    assert.deepEqual(chart.getLayout().plot, initial)
    workbook.redo()
    const grown = chart.getLayout().plot
    await chart.setOptions({ plotArea: undefined })
    assert.equal(chart.getOptions().plotArea, undefined)
    assert.notDeepEqual(chart.getLayout().plot, grown)
    workbook.undo()
    assert.deepEqual(chart.getOptions().plotArea, manual)
    assert.deepEqual(chart.getLayout().plot, grown)
    workbook.redo()
    assert.equal(chart.getOptions().plotArea, undefined)
  } finally { studio.dispose() }
})

test('gradient and plot authorship survive internal JSON save/reopen with data, layout and paints intact', async () => {
  const { studio, workbook, sheet } = setup()
  let reopened
  try {
    const chart = await sheet.insertChart(sheet.newChart('area').setSource('A1:C4').setValueFields([1, 2]).setSize(500, 320)
      .setAutoGradientFill(false).setOptions({ plotArea: manual, allSeriesStyle: { fillOpacity: 0.55, border: { color: 'navy', opacity: 0.4, width: 2 } } }).build())
    await chart.setAutoGradientFill(true)
    const data = chart.getData(), layout = chart.getLayout(), commands = chart.getRenderCommands()
    const snapshot = JSON.parse(JSON.stringify(workbook.save())), serialized = JSON.stringify(snapshot.charts)
    assert.ok(serialized.includes('autoGradientFill') && serialized.includes('plotArea'))
    assert.equal(serialized.includes('colorStops'), false)
    assert.equal(serialized.includes('linearGradient'), false)
    reopened = new api.OfficeStudio({ workbook: snapshot })
    const restored = api.FOffice.newAPI(reopened).getActiveWorkbook(), saved = restored.getActiveSheet().getChartById(chart.getChartId())
    assert.equal(saved.getDescription().autoGradientFill, true)
    assert.deepEqual(saved.getDescription().plotArea, manual)
    assert.deepEqual(saved.getData(), data)
    assert.deepEqual(saved.getLayout(), layout)
    assert.deepEqual(saved.getRenderCommands(), commands)
    await saved.resetAutoGradientFill()
    assert.equal(saved.getOptions().autoGradientFill, undefined)
    assert.equal(saved.getDescription().autoGradientFill, true)
    assert.deepEqual(saved.getOptions().plotArea, manual)
    restored.undo()
    assert.equal(saved.getOptions().autoGradientFill, true)
    restored.redo()
    assert.equal(saved.getOptions().autoGradientFill, undefined)
  } finally { reopened?.dispose(); studio.dispose() }
})

test('one plot drag result uses one existing author command and history entry without changing chart position, size, source or titles', async () => {
  const { studio, workbook, sheet } = setup()
  let subscription
  try {
    const chart = await sheet.insertChart(sheet.newChart('column').setSource('A1:C4').setValueFields([1, 2]).setSize(500, 320)
      .setTitle({ text: 'Main', rotate: 20 }).setSubtitle({ text: 'Sub', color: '#123456' }).setOptions({ plotArea: manual }).build())
    const before = chart.getInfo(), beforeLayout = chart.getLayout(), revision = studio.getWorkbookRevision(workbook.getId()), commands = []
    subscription = studio.events.addEvent('CommandExecuted', event => { if (event.id === 'sheet.command.update-chart') commands.push(event.params) })
    const target = plotInteraction.createChartPlotAreaTarget('column', beforeLayout.frame, beforeLayout.plot, beforeLayout.plot, true)
    const resized = plotInteraction.chartPlotAreaDragLayout(target, before.size, { x: 40, y: 20 }, 'se')
    assert.equal(await chart.setOptions({ plotArea: resized }), chart)
    const after = chart.getInfo()
    assert.equal(commands.length, 1); assert.equal(studio.getWorkbookRevision(workbook.getId()), revision + 1)
    assert.deepEqual(after.options.plotArea, resized); assert.deepEqual(after.position, before.position); assert.deepEqual(after.size, before.size)
    assert.deepEqual(after.source, before.source); assert.deepEqual(after.options.titleStyle, before.options.titleStyle)
    assert.deepEqual(after.options.subtitleStyle, before.options.subtitleStyle)
    workbook.undo(); assert.deepEqual(chart.getInfo(), before); workbook.redo(); assert.deepEqual(chart.getInfo(), after)
  } finally { subscription?.dispose(); studio.dispose() }
})
