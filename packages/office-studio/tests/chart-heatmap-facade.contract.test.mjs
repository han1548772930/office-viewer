import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// P21 Facade/history drafts. Written for the later unified validation phase;
// no runner, server, build or browser has been started for this batch.
let server, api
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})
afterAll(async () => { await server?.close() })

function setup() {
  const studio = new api.OfficeStudio(), facade = api.FOffice.newAPI(studio)
  const workbook = facade.getActiveWorkbook(), sheet = workbook.getActiveSheet()
  sheet.getRange('A1:C4').setValues([['Region', 'Amount', 'Ratio'], ['North', 10, 0.25], ['South', 20, 0.5], ['West', -5, 0]])
  const builder = sheet.newChart('heatmap').setSource('A1:C4').setCategoryField(0).setValueFields([1, 2])
  return { studio, facade, workbook, sheet, builder }
}

test('Heatmap exposes the fixed enum and eight detached fluent operations with independent reset semantics', () => {
  const { studio, facade, builder } = setup()
  try {
    assert.deepEqual(facade.Enum.ChartVisualMapType, { Continuous: 'continuous', Piecewise: 'piecewise' })
    assert.equal(facade.Enum.ChartVisualMapType, api.ChartVisualMapTypes)
    const label = { visible: true, format: '0.00', family: 'Georgia', color: '#123456' }
    assert.equal(builder.setVisualMapType(facade.Enum.ChartVisualMapType.Piecewise), builder)
    assert.equal(builder.setValueRange(-5, 20), builder)
    assert.equal(builder.setCellLabel(label), builder)
    assert.equal(builder.setValueUnit('万元'), builder)
    const baseline = builder.build()
    label.family = 'Arial'
    assert.deepEqual(builder.build(), baseline)
    builder.setCellLabel({ bold: true })
    assert.deepEqual(builder.build().options.heatmap.label, { ...baseline.options.heatmap.label, bold: true })
    builder.setValueRange(0, 0)
    assert.equal(builder.build().options.heatmap.min, 0)
    assert.equal(builder.build().options.heatmap.max, 0)
    assert.equal(builder.resetVisualMapType(), builder)
    assert.equal(builder.build().options.heatmap.visualMapType, undefined)
    assert.equal(builder.build().options.heatmap.unit, '万元')
    assert.equal(builder.clearValueRange(), builder)
    assert.equal(builder.build().options.heatmap.min, undefined)
    assert.equal(builder.build().options.heatmap.max, undefined)
    assert.equal(builder.clearCellLabel(), builder)
    assert.equal(builder.build().options.heatmap.label, undefined)
    assert.equal(builder.clearValueUnit(), builder)
    assert.equal(builder.build().options.heatmap.unit, undefined)
    const reset = builder.build()
    for (const operation of [() => builder.setVisualMapType('gradient'), () => builder.setValueRange(2, 1),
      () => builder.setValueRange(NaN, 1), () => builder.setValueRange(0, Infinity),
      () => builder.setCellLabel({ contentType: 2 }), () => builder.setCellLabel({ visible: 'yes' }),
      () => builder.setValueUnit(10)]) assert.throws(operation, error => error instanceof TypeError || error instanceof RangeError)
    assert.deepEqual(builder.build(), reset)
  } finally { studio.dispose() }
})

test('live Heatmap configuration shares content history, data refresh and authored snapshot reopen', async () => {
  const { studio, workbook, sheet, builder } = setup()
  let reopened
  try {
    const chart = await sheet.insertChart(builder.setValueUnit('元').build()), baseline = chart.getInfo()
    assert.equal(await chart.setVisualMapType('piecewise'), chart)
    assert.equal(await chart.setValueRange(-10, 30), chart)
    assert.equal(await chart.setCellLabel({ visible: true, format: '0.00', bold: true }), chart)
    assert.equal(await chart.setValueUnit('万元'), chart)
    const authored = chart.getInfo(), data = chart.getData(), paint = chart.getRenderCommands()
    workbook.undo()
    assert.equal(chart.getOptions().heatmap.unit, '元')
    workbook.redo()
    assert.deepEqual(chart.getInfo(), authored)
    sheet.getRange('B2').setValue(12)
    const changed = chart.getData()
    assert.equal(changed.series[0].values[0], 12)
    assert.deepEqual(chart.getInfo(), authored)
    workbook.undo()
    assert.deepEqual(chart.getData(), data)
    workbook.redo()
    assert.deepEqual(chart.getData(), changed)
    const saved = JSON.parse(JSON.stringify(workbook.save())), serialized = JSON.stringify(saved.charts)
    for (const field of ['formattedRecords', 'formattedValues', 'heatmapScale', 'pointerGeometry']) assert.equal(serialized.includes(`"${field}"`), false)
    reopened = new api.OfficeStudio({ workbook: saved })
    const restored = api.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet().getChartById(chart.getChartId())
    assert.deepEqual(restored.getInfo(), authored)
    assert.deepEqual(restored.getData(), changed)
    assert.deepEqual(restored.getRenderCommands(), chart.getRenderCommands())
    for (const [operation, field] of [[() => chart.resetVisualMapType(), 'visualMapType'], [() => chart.clearValueUnit(), 'unit'],
      [() => chart.clearCellLabel(), 'label'], [() => chart.clearValueRange(), 'min']]) {
      const before = chart.getInfo()
      assert.equal(await operation(), chart)
      assert.equal(chart.getOptions().heatmap[field], undefined)
      workbook.undo()
      assert.deepEqual(chart.getInfo(), before)
    }
    assert.notDeepEqual(authored, baseline)
    assert.ok(paint.length > 0)
  } finally { reopened?.dispose(); studio.dispose() }
})

test('invalid live and wrong-type Heatmap operations reject without changing the snapshot or undo chain', async () => {
  const { studio, workbook, sheet, builder } = setup()
  try {
    const chart = await sheet.insertChart(builder.build()), before = chart.getInfo(), snapshot = JSON.stringify(workbook.save())
    for (const operation of [() => chart.setVisualMapType('gradient'), () => chart.setValueRange(4, -1),
      () => chart.setValueRange(0, NaN), () => chart.setCellLabel({ contentType: 2 }),
      () => chart.setCellLabel({ color: 'invalid-color' }), () => chart.setValueUnit(null)]) {
      await assert.rejects(async () => { await operation() }, error => error instanceof TypeError || error instanceof RangeError)
    }
    assert.deepEqual(chart.getInfo(), before)
    assert.equal(JSON.stringify(workbook.save()), snapshot)
    await chart.update({ type: 'column' })
    const inactive = JSON.stringify(workbook.save())
    for (const operation of [() => chart.setVisualMapType('continuous'), () => chart.resetVisualMapType(),
      () => chart.setValueRange(0, 1), () => chart.clearValueRange(), () => chart.setCellLabel({ visible: true }),
      () => chart.clearCellLabel(), () => chart.setValueUnit('元'), () => chart.clearValueUnit()]) await assert.rejects(async () => { await operation() }, TypeError)
    assert.equal(JSON.stringify(workbook.save()), inactive)
    workbook.undo()
    assert.deepEqual(chart.getInfo(), before)
    workbook.undo()
    assert.equal(sheet.getCharts().some(item => item.getChartId() === chart.getChartId()), false)
  } finally { studio.dispose() }
})
