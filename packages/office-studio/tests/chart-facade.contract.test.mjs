import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// New Gauge, Radar, Cartesian, named Bubble, P12 axis and P14 pattern cases are drafts for unified validation and have not run in this phase.
let server
let apiModule

beforeAll(async () => {
  server = await createServer({
    configFile: resolve(process.cwd(), 'vite.config.ts'),
    root: process.cwd(),
    server: { middlewareMode: true, hmr: false },
    appType: 'custom',
    logLevel: 'error',
  })
  apiModule = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})

afterAll(async () => {
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
    fields: ['Month', 'North', 'South'],
    records: [['Jan', 10, 15], ['Feb', 20, 18], ['Mar', 30, 24]],
    visibleFields: [0, 1, 2],
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
  assert.ok(chart.getRenderCommands().some((command) => command.kind === 'rect' && command.fill === 'rgba(37, 99, 235, 1)'))
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
  assert.ok(chart.getRenderCommands().some(command => command.kind === 'rect' && command.fill === 'rgba(220, 38, 38, 1)'))
  await assert.rejects(() => chart.update({ yAxis: { scaleType: 'log' } }), /线性数值轴/)
  studio.dispose()
})

test('scatter chart API persists independent X-field bindings and applies numeric log axes', async () => {
  const { studio, workbook, sheet } = setup()
  sheet.getRange('A1:D4').setValues([
    ['Shared X', 'North', 'Other X', 'South'],
    [1, 10, 5, 30], [2, 20, 9, 40], [3, 25, 12, 50],
  ])
  const chart = await sheet.insertChart(sheet.newChart('scatter').setSource('A1:D4')
    .setCategoryField(0).setValueFields([1, 3]).setSeriesToCategory({ 3: [2] })
    .setXAxis({ scaleType: 'log', logBase: 2 }).build())
  assert.deepEqual(chart.getData().series.map((series) => series.xValues), [[1, 2, 3], [5, 9, 12]])
  const layout = chart.getLayout()
  assert.ok(layout.xMin > 0 && layout.xMax > layout.xMin)
  await chart.update({ seriesToCategory: { 3: [0] } })
  assert.deepEqual(chart.getData().series[1].xValues, [1, 2, 3])
  workbook.undo()
  assert.deepEqual(chart.getData().series[1].xValues, [5, 9, 12])
  workbook.redo()
  assert.deepEqual(chart.getInfo().seriesToCategory, { 3: [0] })
  studio.dispose()
})

test('area charts split fills at the zero crossing and pie charts draw category slices', async () => {
  const { studio, sheet } = setup()
  sheet.getRange('B2:B4').setValues([[-10], [20], [30]])
  const area = await sheet.insertChart(sheet.newChart('area').setSource('A1:B4').setValueFields([1]).build())
  const areaPolygons = area.getRenderCommands().filter(command => command.kind === 'polygon')
  const zeroY = area.getLayout().series[0].points[0].baseY
  assert.ok(areaPolygons.length >= 4 && areaPolygons.every(command => command.points.length === 3))
  assert.ok(areaPolygons.some(command => command.points.some(point => point.y < zeroY - 1e-8)))
  assert.ok(areaPolygons.some(command => command.points.some(point => point.y > zeroY + 1e-8)))
  assert.ok(areaPolygons.every(command => command.points.every(point => point.y >= zeroY - 1e-8) || command.points.every(point => point.y <= zeroY + 1e-8)))

  sheet.getRange('B2:B4').setValues([[10], [20], [30]])
  const pie = await sheet.insertChart(sheet.newChart('pie').setSource('A1:C4').setValueFields([1]).build())
  assert.ok(pie.getRenderCommands().some(command => command.kind === 'polygon' && command.fill === 'rgba(37, 99, 235, 1)'))
  const doughnut = await sheet.insertChart(sheet.newChart('doughnut').setSource('A1:C4').setValueFields([1]).build())
  const ring = doughnut.getRenderCommands().filter(command => command.kind === 'polygon'), plot = doughnut.getLayout().plot
  const center = { x: plot.x + plot.width / 2, y: plot.y + plot.height / 2 }
  assert.ok(ring.length > 0)
  for (const command of ring) {
    const sides = command.points.map((point, index) => {
      const next = command.points[(index + 1) % command.points.length]
      return (next.x - point.x) * (center.y - point.y) - (next.y - point.y) * (center.x - point.x)
    })
    assert.ok(sides.some(value => value < -1e-8) && sides.some(value => value > 1e-8), 'Annular triangles leave the center unpainted')
  }
  await assert.rejects(() => sheet.insertChart(sheet.newChart('pie').setSource('A1:C4').setValueFields([1, 2]).build()), /exactly one value series/)
  studio.dispose()
})

test('builder cell placement resolves current hidden and resized axes and persists absolute coordinates', async () => {
  const { studio, workbook, sheet } = setup()
  const builder = sheet.newChart('column').setSource('A1:C4').setValueFields([2]).setPosition('E4')
  sheet.setColumnWidth(0, 140).hideColumns(1)
  sheet.setRowHeightsForced(0, 1, 36).hideRows(1)
  assert.equal(builder.build().positionCell, 'E4')
  const created = await sheet.insertChart(builder.build())
  assert.deepEqual(created.getInfo().position, { x: 332, y: 60 })
  assert.equal(Object.hasOwn(created.getInfo(), 'positionCell'), false)
  const absolute = builder.setAbsolutePosition(17.5, 21.25).build()
  assert.equal(absolute.positionCell, undefined)
  const placed = await sheet.insertChart(absolute)
  assert.deepEqual(placed.getInfo().position, absolute.position)
  const snapshot = workbook.save(), reopened = new apiModule.OfficeStudio({ workbook: snapshot })
  try {
    const saved = apiModule.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet().getChartById(created.getChartId())
    assert.deepEqual(saved.getInfo().position, { x: 332, y: 60 })
  } finally { reopened.dispose(); studio.dispose() }
})

test('builder depth-axis and 3D-view clearing are independent and depth-axis updates survive history and save', async () => {
  const { studio, workbook, sheet } = setup()
  const depth = { title: 'Region', reverse: true, label: { visible: true, color: 'navy' } }
  const builder = sheet.newChart('column3D').setSource('A1:C4').setValueFields([1, 2]).setThreeD({ view: { rotY: 44 } }).setDepthAxis(depth)
  builder.clearThreeD()
  assert.deepEqual(builder.build().options.depthAxis, depth)
  assert.equal(builder.build().options.threeD, undefined)
  builder.setThreeD({ view: { rotY: 44 } }).clearDepthAxis()
  assert.equal(builder.build().options.depthAxis, undefined)
  assert.equal(builder.build().options.threeD.view.rotY, 44)
  const chart = await sheet.insertChart(builder.setDepthAxis(depth).build())
  await chart.update({ options: { depthAxis: { title: 'Updated' } } })
  assert.equal(chart.getOptions().depthAxis.title, 'Updated')
  assert.equal(chart.getOptions().depthAxis.reverse, true)
  workbook.undo()
  assert.equal(chart.getOptions().depthAxis.title, 'Region')
  workbook.redo()
  const reopened = new apiModule.OfficeStudio({ workbook: workbook.save() })
  try {
    const saved = apiModule.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet().getChartById(chart.getChartId())
    assert.deepEqual(saved.getOptions().depthAxis, { ...depth, title: 'Updated' })
    assert.equal(saved.getOptions().threeD.view.rotY, 44)
  } finally { reopened.dispose(); studio.dispose() }
})

test('public Gauge mode, pointer-shape and preset exports are available through the same Facade enum contract', () => {
  const { studio } = setup()
  try {
    const api = apiModule.FOffice.newAPI(studio)
    assert.deepEqual(api.Enum.ChartGaugeMode, { Modern: 'modern', Classic: 'classic' })
    assert.deepEqual(api.Enum.ChartGaugePointerShape, { Needle: 'needle', Slim: 'slim', Arrow: 'arrow' })
    assert.deepEqual(apiModule.ChartGaugeModes, api.Enum.ChartGaugeMode)
    assert.deepEqual(apiModule.ChartGaugePointerShapes, api.Enum.ChartGaugePointerShape)
    assert.deepEqual(apiModule.CHART_GAUGE_PRESETS.map(preset => preset.value), ['modern', 'classic', 'statusDial', 'progress', 'minimal'])
    assert.equal(apiModule.chartGaugePresetOptions('progress').pointer.visible, false)
  } finally { studio.dispose() }
})

test('Gauge builder value fields synchronize existing mapping and FChart changes really switch the bound data', async () => {
  const { studio, workbook, sheet } = setup()
  try {
    const builder = sheet.newChart('gauge').setSource('A1:C4').setGaugeValueField(1).setValueFields([2])
    assert.deepEqual(builder.build().valueFields, [2])
    assert.deepEqual(builder.build().options.mapping.gauge, { valueIndex: 2, labelIndex: 0 })
    const before = builder.build()
    for (const invalid of [[], [1, 2], [-1], [0.5], [NaN], [Infinity], Array(1)]) {
      assert.throws(() => builder.setValueFields(invalid), RangeError)
      assert.deepEqual(builder.build(), before)
    }
    const chart = await sheet.insertChart(builder.build())
    assert.deepEqual(chart.getData().series[0].values, [15, 18, 24])
    assert.equal(chart.getData().series[0].field, 2)
    await chart.setGaugeValueField(1)
    assert.deepEqual(chart.getData().series[0].values, [10, 20, 30])
    assert.deepEqual(chart.getInfo().valueFields, [1])
    workbook.undo()
    assert.deepEqual(chart.getData().series[0].values, [15, 18, 24])
    workbook.redo()
    assert.deepEqual(chart.getOptions().mapping.gauge, { valueIndex: 1, labelIndex: 0 })
    sheet.getRange('E1:E4').setValues([['Metric'], [40], [50], [60]])
    const implicit = sheet.newChart('gauge').setSource('E1:E4').setValueFields([0])
    assert.equal(implicit.build().options?.mapping?.gauge, undefined)
    const single = await sheet.insertChart(implicit.build())
    assert.deepEqual(single.getData().categories, ['Metric', 'Metric', 'Metric'])
    assert.deepEqual(single.getData().series[0].values, [40, 50, 60])
  } finally { studio.dispose() }
})

const gaugeStyleGroups = [
  ['setGaugeAxisLine', 'resetGaugeAxisLine', 'axisLine', { visible: false, width: 24, color: 'gray' }],
  ['setGaugeAxisTick', 'resetGaugeAxisTick', 'axisTick', { visible: true, splitNumber: 3, length: 6, width: 2, color: 'orange' }],
  ['setGaugeSplitLine', 'resetGaugeSplitLine', 'splitLine', { visible: true, length: 10, width: 3, color: 'navy' }],
  ['setGaugeAxisLabel', 'resetGaugeAxisLabel', 'axisLabel', { visible: true, distance: -10, format: '0.0', family: 'Georgia', fontSize: 15, bold: true, italic: true, color: 'purple' }],
  ['setGaugePointer', 'resetGaugePointer', 'pointer', { visible: true, length: 0.8, width: 6, color: 'blue', shape: 'arrow' }],
  ['setGaugeAnchor', 'resetGaugeAnchor', 'anchor', { visible: true, size: 12, color: 'white', borderColor: 'blue', borderWidth: 2 }],
  ['setGaugeProgress', 'resetGaugeProgress', 'progress', { visible: true, width: 14, color: 'lime', rounded: true }],
  ['setGaugeTitle', 'resetGaugeTitle', 'title', { visible: true, offsetY: 0.9, family: 'Georgia', fontSize: 13, color: 'green', bold: true, italic: false }],
  ['setGaugeDetail', 'resetGaugeDetail', 'detail', { visible: true, offsetY: 0.2, format: '0.00%', family: 'Georgia', fontSize: 18, color: 'navy', bold: false, italic: true }],
]

test('every Gauge style setter and group reset shares builder and FChart semantics, history and internal reopen', async () => {
  const { studio, workbook, sheet } = setup()
  try {
    const builder = sheet.newChart('gauge').setSource('A1:C4').setGaugeValueField(1).setGaugeMode('classic').setLegend({ visible: false })
    for (const [setter, resetter, group, value] of gaugeStyleGroups) {
      builder[setter](value)
      assert.deepEqual(builder.build().options.gauge[group], value, setter)
      const field = Object.keys(value).find(key => key !== 'visible')
      builder[setter]({ [field]: null })
      assert.equal(builder.build().options.gauge[group][field], undefined, setter)
      assert.equal(builder.build().options.gauge[group].visible, value.visible)
      builder[resetter]()
      assert.equal(builder.build().options.gauge[group], undefined, resetter)
      assert.equal(builder.build().options.gauge.mode, 'classic')
    }
    const chart = await sheet.insertChart(builder.build())
    for (const [setter, resetter, group, value] of gaugeStyleGroups) {
      await chart[setter](value)
      assert.deepEqual(chart.getOptions().gauge[group], value, setter)
      const detached = chart.getOptions()
      detached.gauge[group].visible = !value.visible
      assert.equal(chart.getOptions().gauge[group].visible, value.visible)
      const field = Object.keys(value).find(key => key !== 'visible')
      await chart[setter]({ [field]: null })
      assert.equal(chart.getOptions().gauge[group][field], undefined, setter)
      assert.equal(chart.getOptions().gauge[group].visible, value.visible)
      await chart[setter](value)
      await chart[resetter]()
      assert.equal(chart.getOptions().gauge[group], undefined, resetter)
      workbook.undo()
      assert.deepEqual(chart.getOptions().gauge[group], value, resetter)
      workbook.redo()
      assert.equal(chart.getOptions().gauge[group], undefined, resetter)
    }
    for (const [setter, , , value] of gaugeStyleGroups) await chart[setter](value)
    const expected = chart.getOptions(), expectedData = chart.getData(), expectedCommands = chart.getRenderCommands()
    const reopened = new apiModule.OfficeStudio({ workbook: workbook.save() })
    try {
      const saved = apiModule.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet().getChartById(chart.getChartId())
      assert.deepEqual(saved.getOptions(), expected)
      assert.deepEqual(saved.getData(), expectedData)
      assert.deepEqual(saved.getRenderCommands(), expectedCommands)
    } finally { reopened.dispose() }
  } finally { studio.dispose() }
})

test('Gauge scale, ticks, angles and ranges support partial negative scales, null resets and atomic range rejection', async () => {
  const { studio, workbook, sheet } = setup()
  try {
    const builder = sheet.newChart('gauge').setSource('A1:C4').setGaugeValueField(1)
      .setGaugeScale({ min: -100, max: -10, splitNumber: 4 }).setGaugeScale({ max: -5 })
      .setGaugeTicks({ majorDivisions: 6, minorDivisionsPerMajor: 3 }).setGaugeAngles({ startAngle: -360, endAngle: 90 })
      .setGaugeRanges([{ from: -100, to: -70, color: 'red' }, { from: -30, to: -5, color: 'blue' }])
    assert.deepEqual([builder.build().options.gauge.min, builder.build().options.gauge.max, builder.build().options.gauge.splitNumber], [-100, -5, 6])
    assert.equal(builder.build().options.gauge.axisTick.splitNumber, 3)
    const before = builder.build()
    assert.throws(() => builder.setGaugeScale({ max: -20 }), RangeError)
    assert.deepEqual(builder.build(), before)
    const chart = await sheet.insertChart(builder.build())
    await chart.setGaugeScale({ max: -1 })
    assert.equal(chart.getOptions().gauge.min, -100)
    await chart.setGaugeTicks({ majorDivisions: 4, minorDivisionsPerMajor: 2 })
    assert.deepEqual([chart.getOptions().gauge.splitNumber, chart.getOptions().gauge.axisTick.splitNumber], [4, 2])
    await chart.setGaugeAngles({ startAngle: 360, endAngle: 0 })
    assert.deepEqual([chart.getOptions().gauge.startAngle, chart.getOptions().gauge.endAngle], [360, 0])
    const stable = chart.getInfo()
    await assert.rejects(() => chart.setGaugeScale({ max: -200 }), RangeError)
    assert.deepEqual(chart.getInfo(), stable)
    await chart.resetGaugeRanges()
    await chart.setGaugeScale({ min: null, max: null, splitNumber: null })
    await chart.setGaugeTicks({ majorDivisions: null, minorDivisionsPerMajor: null })
    await chart.setGaugeAngles({ startAngle: null, endAngle: null })
    assert.deepEqual([chart.getOptions().gauge.min, chart.getOptions().gauge.max, chart.getOptions().gauge.splitNumber, chart.getOptions().gauge.axisTick.splitNumber], [undefined, undefined, undefined, undefined])
    assert.deepEqual([chart.getOptions().gauge.startAngle, chart.getOptions().gauge.endAngle, chart.getOptions().gauge.ranges], [undefined, undefined, undefined])
    await chart.setGaugeMode('progress')
    await chart.resetGaugeMode()
    assert.equal(chart.getOptions().gauge.mode, undefined)
    await chart.setGaugeScale({ min: -100, max: -1, splitNumber: 7 })
    await chart.resetGaugeScale()
    assert.deepEqual([chart.getOptions().gauge.min, chart.getOptions().gauge.max, chart.getOptions().gauge.splitNumber], [undefined, undefined, undefined])
    await chart.setGaugeAngles({ startAngle: 180, endAngle: 0 })
    await chart.resetGaugeAngles()
    assert.deepEqual([chart.getOptions().gauge.startAngle, chart.getOptions().gauge.endAngle], [undefined, undefined])
    workbook.undo()
    assert.deepEqual([chart.getOptions().gauge.startAngle, chart.getOptions().gauge.endAngle], [180, 0])
    workbook.redo()
    builder.resetGaugeRanges().resetGaugeScale().resetGaugeAngles().resetGaugeMode()
    assert.deepEqual([builder.build().options.gauge.min, builder.build().options.gauge.max, builder.build().options.gauge.startAngle, builder.build().options.gauge.mode], [undefined, undefined, undefined, undefined])
  } finally { studio.dispose() }
})

test('Gauge presets and whole-option clearing retain mapping, source data and other chart settings through history and reopen', async () => {
  const { studio, workbook, sheet } = setup()
  try {
    const builder = sheet.newChart('gauge').setSource('A1:C4').setGaugeValueField(2)
      .setOptions({ appearance: { background: 'navy' }, allSeriesStyle: { fillOpacity: 0.4 } })
      .setGaugePointer({ shape: 'arrow', color: 'red' }).setGaugeDetail({ format: '0.000' })
    for (const preset of apiModule.CHART_GAUGE_PRESETS) {
      builder.setGaugePreset(preset.value)
      assert.deepEqual(builder.build().options.mapping.gauge, { valueIndex: 2, labelIndex: 0 })
      assert.equal(builder.build().options.gauge.pointer.color, undefined)
      assert.equal(builder.build().options.gauge.detail.format, undefined)
      assert.equal(builder.build().options.appearance.background, 'navy')
    }
    const chart = await sheet.insertChart(builder.build())
    for (const preset of apiModule.CHART_GAUGE_PRESETS) {
      await chart.setGaugePreset(preset.value)
      assert.equal(chart.getOptions().gauge.mode, preset.value === 'classic' ? 'classic' : 'modern')
      assert.equal(chart.getOptions().gauge.pointer.visible, preset.value !== 'progress')
      assert.equal(chart.getOptions().gauge.progress.visible, preset.value !== 'minimal')
      assert.deepEqual(chart.getData().series[0].values, [15, 18, 24])
    }
    const styled = chart.getOptions()
    await chart.setGauge(undefined)
    assert.equal(chart.getOptions().gauge, undefined)
    assert.deepEqual(chart.getOptions().mapping.gauge, { valueIndex: 2, labelIndex: 0 })
    assert.equal(chart.getOptions().appearance.background, 'navy')
    assert.equal(chart.getOptions().allSeriesStyle.fillOpacity, 0.4)
    workbook.undo()
    assert.deepEqual(chart.getOptions().gauge, styled.gauge)
    workbook.redo()
    const reopened = new apiModule.OfficeStudio({ workbook: workbook.save() })
    try {
      const saved = apiModule.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet().getChartById(chart.getChartId())
      assert.equal(saved.getOptions().gauge, undefined)
      assert.deepEqual(saved.getOptions().mapping.gauge, { valueIndex: 2, labelIndex: 0 })
      assert.deepEqual(saved.getData().series[0].values, [15, 18, 24])
    } finally { reopened.dispose() }
    builder.setGauge(undefined)
    assert.equal(builder.build().options.gauge, undefined)
    assert.deepEqual(builder.build().options.mapping.gauge, { valueIndex: 2, labelIndex: 0 })
  } finally { studio.dispose() }
})

test('Gauge sparse ranges, unsafe prototypes and invalid nested values fail without changing model or content history', async () => {
  const { studio, workbook, sheet } = setup()
  try {
    const chart = await sheet.insertChart(sheet.newChart('gauge').setSource('A1:C4').setGaugeValueField(1).setGaugeMode('classic').build())
    await chart.setGaugePointer({ color: 'blue' })
    const before = chart.getInfo()
    class Unsafe { width = 7 }
    for (const invalid of [new Unsafe(), Object.create(null), JSON.parse('{"__proto__":{}}'), { constructor: 'unsafe' }, { prototype: 'unsafe' }]) {
      await assert.rejects(() => chart.setGauge(invalid), TypeError)
      await assert.rejects(() => chart.setGauge({ pointer: invalid }), TypeError)
      await assert.rejects(() => chart.setGaugePointer(invalid), TypeError)
      assert.deepEqual(chart.getInfo(), before)
    }
    await assert.rejects(() => chart.setGaugeRanges(Array(1)), TypeError)
    await assert.rejects(() => chart.setGauge({ ranges: Array(1) }), TypeError)
    await assert.rejects(() => chart.setGaugePointer({ width: -1 }), RangeError)
    await assert.rejects(() => chart.setGaugeRanges([{ from: 20, to: 60, color: 'red' }, { from: 50, to: 80, color: 'blue' }]), RangeError)
    assert.deepEqual(chart.getInfo(), before)
    workbook.undo()
    assert.equal(chart.getOptions().gauge.pointer, undefined)
    workbook.redo()
    assert.equal(chart.getOptions().gauge.pointer.color, 'blue')
  } finally { studio.dispose() }
})

test('Radar public setters and resets share builder, history and internal reopen behavior', async () => {
  const { studio, workbook, sheet } = setup()
  try {
    const api = apiModule.FOffice.newAPI(studio)
    assert.deepEqual(api.Enum.ChartRadarShape, { Polygon: 'polygon', Circle: 'circle' })
    assert.deepEqual(apiModule.ChartRadarShapes, api.Enum.ChartRadarShape)
    const label = { visible: true, format: '0.0', family: 'Georgia', color: 'navy', fontSize: 14, bold: true, italic: true }
    const builder = sheet.newChart('radar').setSource('A1:C4').setValueFields([1, 2])
      .setShape('circle').setFill(true).setRadarAxisLabel(label)
      .setOptions({ appearance: { font: { fontSize: 16, color: 'green' } } })
    const chart = await sheet.insertChart(builder.build())
    const actions = [['setShape', 'resetShape', 'shape', 'circle'], ['setFill', 'resetFill', 'fill', true], ['setRadarAxisLabel', 'resetRadarAxisLabel', 'axisLabel', label]]
    for (const [set, reset, key, value] of actions) {
      builder[set](value)
      builder[reset]()
      assert.equal(builder.build().options.radar[key], undefined)
      await chart[set](value)
      await chart[reset]()
      assert.equal(chart.getOptions().radar[key], undefined)
      workbook.undo()
      assert.deepEqual(chart.getOptions().radar[key], value)
      workbook.redo()
      assert.equal(chart.getOptions().radar[key], undefined)
    }
    await chart.setRadarAxisLabel(label)
    await chart.setRadarAxisLabel({ fontSize: null, color: null, italic: false })
    assert.equal(chart.getOptions().radar.axisLabel.fontSize, undefined)
    assert.equal(chart.getOptions().radar.axisLabel.color, undefined)
    assert.equal(chart.getOptions().radar.axisLabel.italic, false)
    assert.equal(chart.getOptions().radar.axisLabel.family, 'Georgia')
    assert.ok(chart.getRenderCommands().some(command => command.kind === 'text' && command.color === 'rgba(0, 128, 0, 1)' && /16px Georgia/.test(command.font)))
    const patch = { visible: true, fontSize: 18, color: 'purple' }
    const pending = chart.setRadarAxisLabel(patch)
    patch.fontSize = 222
    await pending
    assert.equal(chart.getOptions().radar.axisLabel.fontSize, 18)
    await chart.setShape('circle')
    await chart.setFill(true)
    const expected = chart.getOptions(), expectedScene = chart.getRenderCommands(), reopened = new apiModule.OfficeStudio({ workbook: workbook.save() })
    try {
      const saved = apiModule.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet().getChartById(chart.getChartId())
      assert.deepEqual(saved.getOptions(), expected)
      assert.deepEqual(saved.getRenderCommands(), expectedScene)
      await saved.setRadar(undefined)
      assert.equal(saved.getOptions().radar, undefined)
    } finally { reopened.dispose() }
  } finally { studio.dispose() }
})

test('Radar malformed style, logarithmic scale and non-Radar calls reject atomically without adding history', async () => {
  const { studio, workbook, sheet } = setup()
  try {
    const builder = sheet.newChart('radar').setSource('A1:C4').setValueFields([1])
    builder.setShape('circle')
    const built = builder.build()
    assert.throws(() => builder.setShape('unsupported'), TypeError)
    assert.deepEqual(builder.build(), built)
    assert.throws(() => builder.setRadarAxisLabel({ constructor: 'unsafe' }), TypeError)
    const chart = await sheet.insertChart(built)
    await chart.setFill(true)
    const before = chart.getInfo()
    for (const invalid of [{ axisLabel: { fontSize: -1 } }, { axisLabel: null, shape: 'unsupported' }, { fill: 1 }]) await assert.rejects(() => chart.setRadar(invalid))
    await assert.rejects(() => chart.setRadarAxisLabel(Object.create({ visible: true })), TypeError)
    await assert.rejects(() => chart.setYAxis({ scaleType: 'log' }), /线性|linear/)
    assert.deepEqual(chart.getInfo(), before)
    workbook.undo()
    assert.equal(chart.getOptions().radar.fill, undefined)
    workbook.redo()
    assert.equal(chart.getOptions().radar.fill, true)
    const line = await sheet.insertChart(sheet.newChart('line').setSource('A1:C4').build()), lineBefore = line.getInfo()
    assert.throws(() => sheet.newChart('line').setShape('circle'), TypeError)
    await assert.rejects(() => line.setRadarAxisLabel({ visible: true }), TypeError)
    assert.deepEqual(line.getInfo(), lineBefore)
  } finally { studio.dispose() }
})

test('Cartesian series restoration clears authored targets and resumes shared defaults through history and reopen', async () => {
  const { studio, workbook, sheet } = setup()
  try {
    const api = apiModule.FOffice.newAPI(studio), shared = { color: 'red', fillOpacity: 0.6,
      border: { color: 'blue', width: 4, opacity: 0.5 }, label: { visible: true, format: '0.0' }, point: { shape: 'rect', size: 8, color: 'green' }, rightAxis: true }
    assert.deepEqual(Object.values(api.Enum.ChartAllSeriesStyleTarget), ['border', 'color', 'fillOpacity', 'label', 'point', 'rightAxis'])
    assert.deepEqual(apiModule.ChartSeriesClearTargets, api.Enum.ChartSeriesClearTarget)
    const specific = { name: 'Authored', color: 'purple', border: { color: 'navy', width: 3 }, point: { size: 16 }, dataPoints: { 0: { color: 'orange' } } }
    const builder = sheet.newChart('line').setSource('A1:C4').setValueFields([1, 2]).setAllSeriesStyle(shared).setSeries(1, specific)
    builder.clearSeries(1, 'point')
    assert.equal(builder.build().options.series[1].point, undefined)
    assert.deepEqual(builder.build().options.allSeriesStyle.point, shared.point)
    builder.setSeries(1, specific)
    const chart = await sheet.insertChart(builder.build())
    await chart.setSeries(1, { color: null, border: { color: null }, point: { size: null } })
    const cleared = chart.getOptions().series[1]
    assert.equal(cleared.color, undefined); assert.equal(cleared.border.color, undefined); assert.equal(cleared.point.size, undefined)
    assert.equal(cleared.border.width, 3); assert.equal(cleared.name, 'Authored')
    assert.ok(chart.getRenderCommands().some(command => command.kind === 'rect' && command.width === 8 && command.fill === 'rgba(0, 128, 0, 0.6)'))
    workbook.undo()
    assert.equal(chart.getOptions().series[1].point.size, 16)
    workbook.redo()
    assert.equal(chart.getOptions().series[1].point.size, undefined)
    const reopened = new apiModule.OfficeStudio({ workbook: workbook.save() })
    try {
      const saved = apiModule.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet().getChartById(chart.getChartId())
      assert.deepEqual(saved.getRenderCommands(), chart.getRenderCommands())
      assert.equal(saved.getOptions().series[1].border.color, undefined)
    } finally { reopened.dispose() }
    for (const target of Object.values(api.Enum.ChartSeriesClearTarget)) {
      await chart.setSeries(1, { ...specific, ...shared })
      await chart.clearSeries(1, target)
      assert.equal(chart.getOptions().series[1][target], undefined)
      assert.deepEqual(chart.getOptions().allSeriesStyle, shared)
    }
    await chart.clearSeries(1)
    assert.equal(chart.getOptions().series[1], undefined)
    await chart.setSeries(1, { color: 'purple' })
    await chart.setSeries(2, { color: 'orange' })
    await chart.clearAllSeriesOverrides()
    assert.deepEqual(chart.getOptions().series, {})
    assert.deepEqual(chart.getOptions().allSeriesStyle, shared)
    for (const target of Object.values(api.Enum.ChartAllSeriesStyleTarget)) {
      await chart.setAllSeriesStyle(shared)
      await chart.clearAllSeriesStyle(target)
      assert.equal(chart.getOptions().allSeriesStyle[target], undefined)
    }
    await chart.setSeries(2, { color: 'purple' })
    await chart.clearAllSeriesStyle()
    assert.equal(chart.getOptions().allSeriesStyle, undefined)
    assert.equal(chart.getOptions().series[2].color, 'purple')
    const before = chart.getInfo()
    await assert.rejects(() => chart.clearSeries(2, 'unsupported'), TypeError)
    await assert.rejects(() => chart.setAllSeriesStyle(Object.create({ color: 'red' })), TypeError)
    assert.deepEqual(chart.getInfo(), before)
  } finally { studio.dispose() }
})

test('Cartesian axis targets and titles restore independently and explicit linear intent clears logarithm bases', async () => {
  const { studio, workbook, sheet } = setup()
  try {
    const api = apiModule.FOffice.newAPI(studio), axis = { scaleType: 'log', logBase: 2, min: 1, max: 128, title: 'Units', reverse: true,
      line: { color: 'red', width: 3 }, tick: { visible: true, length: 6 }, format: '0.0', label: { fontSize: 14, bold: true }, gridLine: { color: 'navy' } }
    assert.deepEqual(Object.values(api.Enum.ChartAxisTarget), ['min', 'max', 'reverse', 'line', 'tick', 'format', 'label', 'gridLine'])
    assert.deepEqual(apiModule.ChartAxisTargets, api.Enum.ChartAxisTarget)
    const builder = sheet.newChart('line').setSource('A1:C4').setValueFields([1, 2]).setYAxis(axis).setRightYAxis(axis)
    builder.setRightYAxis({ max: 256 })
    assert.equal(builder.build().options.rightYAxis.min, 1); assert.equal(builder.build().options.rightYAxis.max, 256)
    builder.setRightYAxis({ scaleType: 'linear' })
    assert.equal(builder.build().options.rightYAxis.logBase, undefined)
    builder.clearYAxis('label')
    assert.equal(builder.build().yAxis.label, undefined); assert.equal(builder.build().yAxis.min, 1)
    builder.setYAxis(axis)
    const chart = await sheet.insertChart(builder.build())
    for (const target of Object.values(api.Enum.ChartAxisTarget)) {
      await chart.setYAxis(axis)
      await chart.clearYAxis(target)
      assert.equal(chart.getInfo().yAxis[target], undefined)
      assert.equal(chart.getInfo().yAxis.title, 'Units')
      assert.equal(chart.getInfo().yAxis.scaleType, 'log')
    }
    await chart.setYAxis({ scaleType: 'linear' })
    assert.equal(chart.getInfo().yAxis.logBase, undefined)
    workbook.undo()
    assert.equal(chart.getInfo().yAxis.logBase, 2)
    workbook.redo()
    assert.equal(chart.getInfo().yAxis.logBase, undefined)
    await chart.setXAxisTitle('Period'); await chart.setYAxisTitle('Measure'); await chart.setRightYAxisTitle('Secondary')
    await chart.clearXAxisTitle(); await chart.clearYAxisTitle(); await chart.clearRightYAxisTitle()
    assert.equal(chart.getInfo().xAxis.title, undefined); assert.equal(chart.getInfo().yAxis.title, undefined); assert.equal(chart.getOptions().rightYAxis.title, undefined)
    await chart.clearXAxis(); await chart.clearYAxis(); await chart.clearRightYAxis()
    assert.deepEqual(chart.getInfo().xAxis, { scaleType: 'linear' })
    assert.deepEqual(chart.getInfo().yAxis, { scaleType: 'linear' })
    assert.deepEqual(chart.getInfo().options.rightYAxis, { scaleType: 'linear' })
    const before = chart.getInfo()
    await assert.rejects(() => chart.clearYAxis('unsupported'), TypeError)
    await assert.rejects(() => chart.setRightYAxis({ min: 5, max: 4 }), RangeError)
    assert.deepEqual(chart.getInfo(), before)
    const reopened = new apiModule.OfficeStudio({ workbook: workbook.save() })
    try {
      const saved = apiModule.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet().getChartById(chart.getChartId())
      assert.deepEqual(saved.getInfo().yAxis, { scaleType: 'linear' })
      assert.deepEqual(saved.getOptions().rightYAxis, { scaleType: 'linear' })
    } finally { reopened.dispose() }
  } finally { studio.dispose() }
})

test('Boxplot uses inherited Cartesian styles and axes across updates, undo, save and internal reopen', async () => {
  const { studio, workbook, sheet } = setup()
  try {
    const samples = [['Sample', 'First distribution', 'Second distribution'],
      ['A', 1, 10], ['B', 2, 20], ['C', 3, 30], ['D', 4, 40], ['E', 5, 50], ['F', 100, 1000]]
    sheet.getRange('A1:C7').setValues(samples)
    const builder = sheet.newChart('boxplot').setSource('A1:C7').setValueFields([1, 2])
      .setYAxis({ min: 0, max: 200 }).setRightYAxis({ min: 0, max: 2000 })
      .setAllSeriesStyle({ border: { color: 'navy', width: 3, opacity: 0.5 }, fillOpacity: 0.4, point: { shape: 'rect', size: 8 } })
      .setSeries(1, { color: 'red' }).setSeries(2, { color: 'green', rightAxis: true, label: { visible: true, position: 'right', format: '0.0' } })
    const chart = await sheet.insertChart(builder.build())
    assert.deepEqual(chart.getData().series.map(series => series.values), [[1, 2, 3, 4, 5, 100], [10, 20, 30, 40, 50, 1000]])
    assert.deepEqual(chart.getLayout().series.map(series => [series.field, series.points[0].value]), [[1, 3.5], [2, 35]])
    assert.equal(chart.getLayout().series[1].rightAxis, true)
    assert.ok(chart.getRenderCommands().some(command => command.kind === 'text' && command.text === '35.0'))
    assert.ok(chart.getRenderCommands().some(command => command.kind === 'line' && command.color === 'rgba(0, 0, 128, 0.5)'))
    const rightY = chart.getLayout().series[1].points[0].y
    await chart.setSeries(2, { rightAxis: null, label: null })
    assert.equal(chart.getLayout().series[1].rightAxis, undefined)
    assert.notEqual(chart.getLayout().series[1].points[0].y, rightY)
    assert.ok(!chart.getRenderCommands().some(command => command.kind === 'text' && command.text === '35.0'))
    workbook.undo()
    assert.equal(chart.getLayout().series[1].rightAxis, true)
    assert.equal(chart.getLayout().series[1].points[0].y, rightY)
    workbook.redo()
    assert.equal(chart.getLayout().series[1].rightAxis, undefined)
    const expected = chart.getRenderCommands(), reopened = new apiModule.OfficeStudio({ workbook: workbook.save() })
    try {
      const saved = apiModule.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet().getChartById(chart.getChartId())
      assert.deepEqual(saved.getRenderCommands(), expected)
      assert.equal(saved.getInfo().type, 'boxplot')
      assert.deepEqual(saved.getData().series.map(series => series.values), chart.getData().series.map(series => series.values))
      await saved.clearSeries(1, 'color')
      assert.equal(saved.getOptions().series[1].color, undefined)
      assert.deepEqual(saved.getOptions().allSeriesStyle.border, { color: 'navy', width: 3, opacity: 0.5 })
    } finally { reopened.dispose() }
  } finally { studio.dispose() }
})

test('Bubble field and name styles stay independent through nullable clears, source regrouping, undo and internal reopen', async () => {
  const { studio, workbook, sheet } = setup()
  try {
    sheet.getRange('A1:E5').setValues([['X', 'Y', 'Size', 'Label', 'Group'], [1, 10, 4, 'N1', 'North'], [2, 20, 9, 'S1', 'South'],
      [3, 30, 16, 'N2', 'North'], [4, 40, 25, 'S2', 'South']])
    const mapping = { xIndex: 0, yIndex: 1, sizeIndex: 2, labelIndex: 3, seriesIndex: 4 }, selector = { name: 'North' }
    const shared = { color: 'gray', point: { size: 8 }, label: { visible: true, contentType: 'series' } }, field = { color: 'blue' }
    const named = { name: 'North display', color: 'red', fillOpacity: 0.6, border: { color: 'navy', width: 3 }, point: { size: 12 }, rightAxis: true,
      label: { visible: true, contentType: 'series' }, dataPoints: { 0: { color: 'purple', shape: 'rect', size: 18 } } }
    const builder = sheet.newChart('bubble').setSource('A1:E5').setValueFields([1]).setBubbleMapping(mapping)
      .setAllSeriesStyle(shared).setSeries(1, field).setSeries(selector, named).setRightYAxis({ min: 0, max: 100 })
    mapping.yIndex = 2; selector.name = 'South'; named.point.size = 99
    assert.equal(builder.build().options.mapping.bubble.yIndex, 1)
    assert.equal(builder.build().options.namedSeries[0].name, 'North'); assert.equal(builder.build().options.namedSeries[0].style.point.size, 12)
    const chart = await sheet.insertChart(builder.build())
    const authored = { ...named, point: { size: 12 } }
    assert.deepEqual(chart.getData().series.map(series => series.name), ['North', 'South'])
    assert.deepEqual(chart.getLayout().series.map(series => [series.name, series.color, Boolean(series.rightAxis)]), [['North display', 'red', true], ['South', 'blue', false]])
    assert.ok(chart.getRenderCommands().some(command => command.kind === 'rect' && command.width === 18 && command.fill === 'rgba(128, 0, 128, 0.6)'))
    sheet.getRange('E2').setValues([['Moved']]); sheet.getRange('E4').setValues([['Moved']])
    assert.deepEqual(chart.getData().series.map(series => series.name), ['Moved', 'South'])
    assert.equal(chart.getLayout().series[0].color, 'blue'); assert.equal(chart.getOptions().namedSeries[0].name, 'North')
    sheet.getRange('E2').setValues([['North']]); sheet.getRange('E4').setValues([['North']])
    assert.equal(chart.getLayout().series[0].name, 'North display'); assert.equal(chart.getLayout().series[0].color, 'red')
    const api = apiModule.FOffice.newAPI(studio)
    for (const target of Object.values(api.Enum.ChartSeriesClearTarget)) {
      await chart.setSeries({ name: 'North' }, authored)
      await chart.clearSeries({ name: 'North' }, target)
      assert.equal(chart.getOptions().namedSeries.find(entry => entry.name === 'North').style[target], undefined)
      assert.deepEqual(chart.getOptions().series[1], field)
    }
    await chart.setSeries({ name: 'North' }, authored)
    await chart.setSeries({ name: 'North' }, { name: null, color: null, point: { size: null }, label: null })
    assert.equal(chart.getLayout().series[0].name, 'North'); assert.equal(chart.getLayout().series[0].color, 'blue')
    assert.equal(chart.getOptions().namedSeries[0].style.point.size, undefined)
    workbook.undo()
    assert.equal(chart.getLayout().series[0].name, 'North display'); assert.equal(chart.getLayout().series[0].color, 'red')
    workbook.redo()
    assert.equal(chart.getLayout().series[0].name, 'North'); assert.equal(chart.getLayout().series[0].color, 'blue')
    await chart.setSeries({ name: 'North display' }, { color: 'green' })
    assert.equal(chart.getLayout().series[0].color, 'blue')
    assert.deepEqual(chart.getOptions().namedSeries.map(entry => entry.name), ['North', 'North display'])
    const expected = chart.getRenderCommands(), reopened = new apiModule.OfficeStudio({ workbook: workbook.save() })
    try {
      const saved = apiModule.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet().getChartById(chart.getChartId())
      assert.deepEqual(saved.getRenderCommands(), expected)
      assert.deepEqual(JSON.parse(JSON.stringify(saved.getOptions())), JSON.parse(JSON.stringify(chart.getOptions())))
      assert.deepEqual(saved.getData().series.map(series => series.name), ['North', 'South'])
    } finally { reopened.dispose() }
    await chart.clearSeries({ name: 'North' })
    assert.deepEqual(chart.getOptions().namedSeries.map(entry => entry.name), ['North display'])
    await chart.clearAllSeriesOverrides()
    assert.deepEqual(chart.getOptions().namedSeries, []); assert.deepEqual(chart.getOptions().series, {})
    assert.deepEqual(chart.getOptions().allSeriesStyle, shared)
    builder.clearAllSeriesOverrides()
    assert.deepEqual(builder.build().options.namedSeries, []); assert.deepEqual(builder.build().options.series, {})
  } finally { studio.dispose() }
})

test('name selectors reject ambiguous source headers and malformed author entries atomically while field styles remain available', async () => {
  const { studio, sheet } = setup()
  try {
    const chart = await sheet.insertChart(sheet.newChart('line').setSource('A1:C4').setValueFields([1, 2]).build())
    sheet.getRange('B1:C1').setValues([['Duplicate', 'Duplicate']])
    const before = chart.getInfo()
    await assert.rejects(() => chart.setSeries({ name: 'Duplicate' }, { color: 'red' }), /匹配多个/)
    for (const selector of [{}, { name: 1 }, { name: 'Duplicate', field: 1 }, '1', -1, 1_048_576]) await assert.rejects(() => chart.setSeries(selector, { color: 'red' }))
    await assert.rejects(() => chart.setOptions({ namedSeries: [{ name: 'Absent', style: {} }, { name: 'Absent', style: {} }] }), RangeError)
    await assert.rejects(() => chart.setOptions({ namedSeries: Array(2) }), TypeError)
    assert.deepEqual(chart.getInfo(), before)
    await chart.setSeries(1, { color: 'red' }); await chart.setSeries(2, { color: 'blue' })
    assert.deepEqual(chart.getLayout().series.map(series => [series.name, series.color]), [['Duplicate', 'red'], ['Duplicate', 'blue']])
    await chart.setSeries({ name: 'Absent' }, { color: 'orange' })
    assert.equal(chart.getOptions().namedSeries[0].name, 'Absent')
    assert.deepEqual(chart.getLayout().series.map(series => series.color), ['red', 'blue'])
  } finally { studio.dispose() }
})

test('Bubble mapping setter and clear preserve other mappings and fallback fields through history and reject invalid roles', async () => {
  const { studio, workbook, sheet } = setup()
  try {
    sheet.getRange('A1:C4').setValues([['X', 'Y', 'Size'], [1, 10, 4], [2, 20, 9], [3, 30, 16]])
    const builder = sheet.newChart('bubble').setSource('A1:C4').setValueFields([1, 2]).setMapping({ categoryIndexes: [0] })
      .setBubbleMapping({ xIndex: 0, yIndex: 1, sizeIndex: 2 }).clearBubbleMapping()
    assert.equal(builder.build().options.mapping.bubble, undefined); assert.deepEqual(builder.build().options.mapping.categoryIndexes, [0])
    const chart = await sheet.insertChart(builder.build())
    assert.equal(chart.getData().series[0].field, 1)
    await chart.setBubbleMapping({ xIndex: 0, yIndex: 2 })
    assert.equal(chart.getData().series[0].field, 2); assert.deepEqual(chart.getData().series[0].values, [4, 9, 16])
    await chart.clearBubbleMapping()
    assert.equal(chart.getOptions().mapping.bubble, undefined); assert.deepEqual(chart.getOptions().mapping.categoryIndexes, [0])
    assert.equal(chart.getData().series[0].field, 1)
    workbook.undo(); assert.equal(chart.getData().series[0].field, 2)
    workbook.redo(); assert.equal(chart.getData().series[0].field, 1)
    const before = chart.getInfo()
    await assert.rejects(() => chart.setBubbleMapping({ xIndex: 0, yIndex: 99 }), RangeError)
    await assert.rejects(() => chart.setBubbleMapping({ xIndex: 0 }), RangeError)
    assert.deepEqual(chart.getInfo(), before)
    assert.throws(() => sheet.newChart('line').setBubbleMapping({ xIndex: 0, yIndex: 1 }), TypeError)
    const line = await sheet.insertChart(sheet.newChart('line').setSource('A1:C4').setValueFields([1]).build()), lineBefore = line.getInfo()
    await assert.rejects(() => line.setBubbleMapping({ xIndex: 0, yIndex: 1 }), TypeError)
    await assert.rejects(() => line.clearBubbleMapping(), TypeError)
    assert.deepEqual(line.getInfo(), lineBefore)
  } finally { studio.dispose() }
})

test('P12 axis title styles, alignment and elapsed dates retain author intent through history and internal reopen', async () => {
  const { studio, workbook, sheet } = setup()
  try {
    sheet.getRange('A1:C4').setValues([['Date', 'North', 'South'], ['1900-02-28', 10, 15], ['1900-02-29', 20, 18], ['1900-03-03', 30, 24]])
    const title = { family: 'Georgia', fontSize: 22, bold: true }, pointer = { indicatorLineColor: 'red', indicatorLineType: 'dashed', indicatorLabelColor: 'navy', indicatorLabelTextColor: 'white' }
    const builder = sheet.newChart('line').setSource('A1:C4').setValueFields([1, 2]).setUseDateAxis(true)
      .setXAxis({ title: 'Calendar', titleStyle: title, titleAlignment: 'start', min: 58, max: 64 })
      .setYAxis({ title: 'Units', titleStyle: { fontSize: 20 }, titleAlignment: 'end' }).setAxisPointer(pointer)
    title.fontSize = 99; pointer.indicatorLineColor = 'blue'
    const chart = await sheet.insertChart(builder.build())
    assert.equal(chart.getInfo().xAxis.titleStyle.fontSize, 22)
    assert.equal(chart.getOptions().axisPointer.indicatorLineColor, 'red')
    const xs = chart.getLayout().series[0].points.map(point => point.x)
    assert.ok(Math.abs((xs[2] - xs[1]) / (xs[1] - xs[0]) - 3) < 1e-9)
    await chart.setXAxis({ titleStyle: { bold: null, italic: true }, titleAlignment: 'end' })
    assert.equal(chart.getInfo().xAxis.titleStyle.bold, undefined)
    assert.equal(chart.getInfo().xAxis.titleStyle.family, 'Georgia')
    assert.equal(chart.getInfo().xAxis.titleAlignment, 'end')
    workbook.undo()
    assert.equal(chart.getInfo().xAxis.titleStyle.bold, true)
    assert.equal(chart.getInfo().xAxis.titleAlignment, 'start')
    workbook.redo()
    const expected = chart.getRenderCommands(), reopened = new apiModule.OfficeStudio({ workbook: workbook.save() })
    try {
      const saved = apiModule.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet().getChartById(chart.getChartId())
      assert.deepEqual(saved.getRenderCommands(), expected)
      assert.equal(saved.getInfo().xAxis.titleStyle.fontSize, 22)
      assert.equal(saved.getInfo().xAxis.titleAlignment, 'end')
      assert.equal(saved.getOptions().mapping.useDateAxis, true)
      assert.equal(saved.getOptions().axisPointer.indicatorLabelColor, 'navy')
    } finally { reopened.dispose() }
  } finally { studio.dispose() }
})

test('P12 pointer clear targets reset one field, preserve other authorship and stay independent of date-axis resets', async () => {
  const { studio, workbook, sheet } = setup()
  try {
    sheet.getRange('A2:A4').setValues([[59], [60], [61]])
    const api = apiModule.FOffice.newAPI(studio), style = { indicatorLineColor: 'red', indicatorLineType: 'dotted', indicatorLabelColor: 'navy', indicatorLabelTextColor: 'white' }
    assert.deepEqual(Object.values(api.Enum.AxisAlignEnum), ['start', 'center', 'end'])
    const chart = await sheet.insertChart(sheet.newChart('line').setSource('A1:C4').setValueFields([1]).setUseDateAxis(true).setAxisPointer(style).build())
    for (const target of Object.values(api.Enum.ChartAxisPointerTarget)) {
      await chart.setAxisPointer(style)
      await chart.clearAxisPointer(target)
      assert.equal(chart.getOptions().axisPointer[target], undefined)
      for (const field of Object.keys(style).filter(field => field !== target)) assert.equal(chart.getOptions().axisPointer[field], style[field])
    }
    await chart.setAxisPointer(style)
    await chart.setAxisPointer({ indicatorLineColor: null })
    assert.equal(chart.getOptions().axisPointer.indicatorLineColor, undefined)
    assert.equal(chart.getOptions().axisPointer.indicatorLabelColor, 'navy')
    await chart.resetUseDateAxis()
    assert.equal(chart.getOptions().mapping.useDateAxis, false)
    assert.equal(chart.getOptions().axisPointer.indicatorLabelColor, 'navy')
    workbook.undo(); assert.equal(chart.getOptions().mapping.useDateAxis, true)
    workbook.redo(); assert.equal(chart.getOptions().mapping.useDateAxis, false)
    await chart.clearAxisPointer()
    assert.equal(chart.getOptions().axisPointer, undefined)
    workbook.undo(); assert.equal(chart.getOptions().axisPointer.indicatorLabelColor, 'navy')
  } finally { studio.dispose() }
})

test('P12 invalid date, title and pointer candidates reject before publishing a content history record', async () => {
  const { studio, workbook, sheet } = setup()
  try {
    const chart = await sheet.insertChart(sheet.newChart('line').setSource('A1:C4').setValueFields([1]).setTitle('Accepted').build())
    const before = chart.getInfo()
    await assert.rejects(() => chart.setUseDateAxis(true), /无法作为日期轴值/)
    await assert.rejects(() => chart.setXAxis({ titleStyle: { fontSize: 0 } }), RangeError)
    await assert.rejects(() => chart.setXAxis({ titleAlignment: 'invalid' }), TypeError)
    await assert.rejects(() => chart.setAxisPointer({ indicatorLineType: 'invalid' }), TypeError)
    await assert.rejects(() => chart.clearAxisPointer('invalid'), TypeError)
    assert.deepEqual(chart.getInfo(), before)
    workbook.undo()
    assert.equal(sheet.getChartById(chart.getChartId()), null)
  } finally { studio.dispose() }
})

test('P14 series patterns detach author inputs, preserve null overrides and resume inheritance through history and internal reopen', async () => {
  const { studio, workbook, sheet } = setup()
  try {
    const allPattern = { preset: 'horz', foregroundColor: '#008000', backgroundColor: '#ffffff' }
    const namePattern = { preset: 'upDiag', foregroundColor: '#ff0000', backgroundColor: '#ffff00' }
    const builder = sheet.newChart('column').setSource('A1:C4').setValueFields([1, 2])
      .setAllSeriesStyle({ fillPattern: allPattern, fillOpacity: 0.8 }).setSeries(1, { fillPattern: null })
      .setSeries({ name: 'North' }, { fillPattern: namePattern, dataPoints: { 0: { fillPattern: null }, 1: { fillPattern: undefined } } })
    allPattern.foregroundColor = '#000000'; namePattern.preset = 'vert'
    const chart = await sheet.insertChart(builder.build())
    assert.equal(chart.getOptions().allSeriesStyle.fillPattern.foregroundColor, '#008000')
    assert.equal(chart.getOptions().series[1].fillPattern, null)
    assert.equal(chart.getOptions().namedSeries[0].style.fillPattern.preset, 'upDiag')
    assert.equal(chart.getOptions().namedSeries[0].style.dataPoints[0].fillPattern, null)
    assert.ok(chart.getRenderCommands().some(command => command.kind === 'polygon' && command.fill === 'rgba(255, 0, 0, 0.8)'))
    await chart.setSeries({ name: 'North' }, { fillPattern: null })
    assert.equal(chart.getOptions().namedSeries[0].style.fillPattern, null)
    workbook.undo(); assert.equal(chart.getOptions().namedSeries[0].style.fillPattern.preset, 'upDiag')
    workbook.redo(); assert.equal(chart.getOptions().namedSeries[0].style.fillPattern, null)
    await chart.setSeries({ name: 'North' }, { fillPattern: undefined })
    assert.equal(chart.getOptions().namedSeries[0].style.fillPattern, undefined)
    const snapshot = JSON.parse(JSON.stringify(workbook.save())), reopened = new apiModule.OfficeStudio({ workbook: snapshot })
    try {
      const saved = apiModule.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet().getChartById(chart.getChartId())
      assert.equal(saved.getOptions().series[1].fillPattern, null)
      assert.equal(saved.getOptions().namedSeries[0].style.dataPoints[0].fillPattern, null)
      assert.deepEqual(saved.getOptions().allSeriesStyle.fillPattern, chart.getOptions().allSeriesStyle.fillPattern)
      assert.deepEqual(saved.getRenderCommands(), chart.getRenderCommands())
    } finally { reopened.dispose() }
  } finally { studio.dispose() }
})

test('P14 malformed patterns fail atomically and keep the published clear-target contracts intact', async () => {
  const { studio, workbook, sheet } = setup()
  try {
    sheet.getRange('A2:A4').setValues([[1], [2], [3]])
    const chart = await sheet.insertChart(sheet.newChart('scatter').setSource('A1:C4').setValueFields([1]).build())
    const before = chart.getInfo(), api = apiModule.FOffice.newAPI(studio)
    assert.equal(Object.values(api.Enum.ChartAllSeriesStyleTarget).length, 6)
    assert.equal(Object.values(api.Enum.ChartSeriesClearTarget).length, 7)
    await assert.rejects(() => chart.setSeries(1, { fillPattern: { preset: 'missing', foregroundColor: 'red', backgroundColor: 'white' } }), TypeError)
    await assert.rejects(() => chart.setAllSeriesStyle({ fillPattern: { preset: 'horz', foregroundColor: 'red' } }), TypeError)
    await assert.rejects(() => chart.setSeries({ name: 'North' }, { dataPoints: { 1: { fillPattern: { preset: 'horz', foregroundColor: 'red', backgroundColor: null } } } }), TypeError)
    assert.deepEqual(chart.getInfo(), before)
    workbook.undo(); assert.equal(sheet.getChartById(chart.getChartId()), null)
  } finally { studio.dispose() }
})
