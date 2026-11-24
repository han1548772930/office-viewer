import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { createChart, readChartData } from '../src/features/charts/chart-engine.ts'
import { CHART_TYPE_CATALOG } from '../src/features/charts/chart-catalog.ts'
import { chartManualPlotFrame, chartPlotAreaFromFrame, supportsChartPlotArea, validateChartPlotArea } from '../src/features/charts/chart-plot-area.ts'
import { chartPlotFrame, layoutChartData, resolveChartPlotAreaForEditing } from '../src/features/charts/chart-layout.ts'
import { chartPlotAreaControls } from '../src/features/charts/chart-plot-area-controls.ts'
import { createChartSettingsDraft, chartSettingsPatch } from '../src/features/charts/chart-settings-draft.ts'
import { projectChartSettings } from '../src/features/charts/chart-settings-data.ts'
import { cloneChartOptions, mergeChartOptions } from '../src/features/charts/chart-options.ts'
import { convertChartType } from '../src/features/charts/chart-type-conversion.ts'
import { chartAxisPointerCommands, chartAxisPointerGeometry } from '../src/features/charts/chart-axis-pointer.ts'
import { chartRadarGeometry } from '../src/features/charts/chart-radar.ts'
import { chart3DScene } from '../src/features/charts/chart-3d.ts'
import { chartEChartsOption } from '../src/features/charts/echarts-option.ts'
import { applyExcelWorkbookView, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'

// P15 drafts. UI/browser, Office and performance acceptance remain unexecuted.
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-7, `${actual} ≈ ${expected}`)
const manual = { left: 0.2, top: 0.25, width: 0.6, height: 0.5 }
function model(type = 'column', options = {}) {
  return { id: 'plot-area', sheetId: 'sheet', type, revision: 0, categoryField: 0, valueFields: [1],
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 3, endCol: 1 } }, position: { x: 100, y: 40 },
    size: { width: 500, height: 300 }, legend: { visible: false, position: 'bottom' }, xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, options }
}
function fixture(type, options = {}) {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.cells = Object.fromEntries([['X', 'Units'], [1, 10], [2, 20], [3, 30]].flatMap((row, r) => row.map((value, c) => [`${r}:${c}`, createCell(r, c, value)])))
  const chart = createChart(workbook, sheet.id, model(type, options))
  return { workbook, chart, data: readChartData(workbook, chart) }
}

test('manual rectangles are normalized to the whole frame and survive movement and resize', () => {
  const chart = model('column', { plotArea: { ...manual } }), bounds = chartPlotFrame(chart)
  assert.deepEqual(bounds.plot, { x: 200, y: 115, width: 300, height: 150 })
  assert.deepEqual(chartPlotAreaFromFrame(bounds.frame, bounds.plot), manual)
  const moved = chartPlotFrame({ ...chart, position: { x: 300, y: 200 }, size: { width: 1000, height: 600 } })
  assert.deepEqual(moved.plot, { x: 500, y: 350, width: 600, height: 300 })
  const editing = resolveChartPlotAreaForEditing(chart)
  editing.left = 0
  assert.equal(chart.options.plotArea.left, 0.2)
})

test('automatic editing seeds include title/font/legend and only active secondary source styles', () => {
  const source = [{ field: 1, name: 'Units' }], base = model('column', { namedSeries: [{ name: 'Dormant', style: { rightAxis: true } }], rightYAxis: { title: 'Secondary', titleStyle: { fontSize: 20 } } })
  const active = chartPlotFrame(base, source), dormant = chartPlotFrame(base, [{ field: 1, name: 'Dormant' }])
  assert.ok(active.plot.width > dormant.plot.width)
  const titled = { ...base, title: 'Title', subtitle: 'Sub', legend: { visible: true, position: 'top' }, options: { ...base.options, titleStyle: { fontSize: 22 } } }
  const seed = resolveChartPlotAreaForEditing(titled, source)
  assert.ok(seed.top > resolveChartPlotAreaForEditing(base, source).top)
  const automatic = chartPlotFrame(titled, source)
  const reconstructed = chartManualPlotFrame(titled.type, automatic.frame, seed)
  for (const key of ['x', 'y', 'width', 'height']) near(reconstructed[key], automatic.plot[key])
  const tiny = resolveChartPlotAreaForEditing({ ...base, size: { width: 1, height: 1 } }, source)
  assert.doesNotThrow(() => validateChartPlotArea(tiny))
  assert.ok(tiny.width > 0 && tiny.height > 0)
})

test('manual mode edits complete rectangles, restores automatic coordinates on clear and rejects invalid apply', () => {
  const config = model('column', { autoGradientFill: true }), source = [{ field: 1, name: 'Units' }], controls = chartPlotAreaControls(source)
  const mode = controls.find(control => control.key === 'plotArea.manual')
  assert.equal(mode.read(config), false)
  mode.write(config, true)
  const initial = { ...config.options.plotArea }, left = controls.find(control => control.key === 'plotArea.left'), width = controls.find(control => control.key === 'plotArea.width')
  left.write(config, 12.5)
  assert.equal(config.options.plotArea.left, 0.125)
  width.write(config, 100)
  assert.throws(() => cloneChartOptions(config.options), RangeError)
  width.write(config, 60)
  assert.doesNotThrow(() => cloneChartOptions(config.options))
  width.write(config, undefined)
  near(config.options.plotArea.width, initial.width)
  assert.equal(config.options.plotArea.left, 0.125)
  mode.write(config, false)
  assert.equal(config.options.plotArea, undefined)
  assert.equal(config.options.autoGradientFill, true)
})

test('invalid plot drafts retain source series and default coordinates until corrected', () => {
  const { chart, data } = fixture('column', { namedSeries: [{ name: 'Units', style: { rightAxis: true } }], rightYAxis: { title: 'Secondary' } })
  const original = structuredClone(chart), draft = createChartSettingsDraft(chart), initial = projectChartSettings(draft, undefined, data)
  const controls = chartPlotAreaControls(initial.data.series)
  controls.find(control => control.key === 'plotArea.manual').write(draft.config, true)
  const automatic = { ...draft.config.options.plotArea }
  controls.find(control => control.key === 'plotArea.left').write(draft.config, 95)
  controls.find(control => control.key === 'plotArea.width').write(draft.config, 90)
  assert.throws(() => chartSettingsPatch(draft), /绘图区/)
  const invalid = projectChartSettings(draft, undefined, data)
  assert.equal(invalid.error, undefined)
  assert.deepEqual(invalid.data, initial.data)
  const corrected = chartPlotAreaControls(invalid.data.series)
  corrected.find(control => control.key === 'plotArea.width').write(draft.config, undefined)
  corrected.find(control => control.key === 'plotArea.left').write(draft.config, undefined)
  for (const key of ['left', 'top', 'width', 'height']) near(draft.config.options.plotArea[key], automatic[key])
  assert.doesNotThrow(() => chartSettingsPatch(draft))
  assert.deepEqual(chart, original)
})

test('plot patches replace atomically and invalid/incomplete rectangles never change author inputs', () => {
  const original = { plotArea: { ...manual, sourceTag: 'previous' }, autoGradientFill: false, appearance: { background: 'white' } }, before = structuredClone(original)
  const replacement = { left: 0.1, top: 0.1, width: 0.8, height: 0.7 }, result = mergeChartOptions(original, { plotArea: replacement })
  assert.deepEqual(result.plotArea, replacement)
  replacement.left = 0.8
  assert.equal(result.plotArea.left, 0.1)
  assert.throws(() => mergeChartOptions(original, { plotArea: { left: 0.3 } }), RangeError)
  assert.deepEqual(original, before)
  const reset = mergeChartOptions(original, { plotArea: undefined })
  assert.equal(reset.plotArea, undefined)
  assert.equal(reset.autoGradientFill, false)
  assert.equal(reset.appearance.background, 'white')
  for (const value of [null, [], new Date(), {}, { ...manual, left: -0.01 }, { ...manual, width: 0 }, { ...manual, height: Infinity }, { ...manual, top: NaN }, { ...manual, width: 0.9 }, { ...manual, top: 0.8 }]) {
    assert.throws(() => cloneChartOptions({ plotArea: value }))
  }
})

test('capability follows all 31 directory variants and inactive author settings survive type conversion', () => {
  const families = ['column', 'bar', 'line', 'area', 'scatter', 'bubble', 'combination', 'pie', 'radar']
  for (const entry of CHART_TYPE_CATALOG) assert.equal(supportsChartPlotArea(entry.value), families.includes(entry.family), entry.id)
  assert.equal(CHART_TYPE_CATALOG.filter(entry => supportsChartPlotArea(entry.value)).length, 31)
  const original = model('column', { plotArea: { ...manual }, autoGradientFill: true }), hidden = convertChartType(original, 'gauge')
  assert.deepEqual(hidden.options.plotArea, manual)
  assert.equal(resolveChartPlotAreaForEditing(hidden), undefined)
  assert.deepEqual(chartPlotFrame(hidden).plot, chartPlotFrame({ ...hidden, options: { ...hidden.options, plotArea: undefined } }).plot)
  const restored = convertChartType(hidden, 'column')
  assert.deepEqual(restored.options.plotArea, manual)
  assert.deepEqual(chartPlotFrame(restored).plot, chartPlotFrame(original).plot)
})

test('Cartesian input/hit guides use the same authored rectangle as content and the native grid', () => {
  const { workbook, chart, data } = fixture('column', { plotArea: { ...manual } }), layout = layoutChartData(data, chart)
  const point = layout.series[0].points[1], geometry = chartAxisPointerGeometry(data, chart, layout), commands = chartAxisPointerCommands(geometry, point)
  assert.ok(commands.some(command => command.kind === 'line' && command.y1 === layout.plot.y && command.y2 === layout.plot.y + layout.plot.height))
  assert.ok(commands.some(command => command.kind === 'line' && command.x1 === layout.plot.x && command.x2 === layout.plot.x + layout.plot.width))
  assert.deepEqual(chartAxisPointerCommands(geometry, { x: layout.plot.x - 1, y: point.y }), [])
  for (const series of layout.series) for (const point of series.points) assert.ok(point.x >= layout.plot.x && point.x <= layout.plot.x + layout.plot.width && point.y >= layout.plot.y && point.y <= layout.plot.y + layout.plot.height)
  const native = chartEChartsOption(workbook, chart)
  assert.equal(native.grid.containLabel, false)
  near(native.grid.left, chart.size.width * manual.left)
  near(native.grid.top, chart.size.height * manual.top)
  near(chart.size.width - native.grid.left - native.grid.right, layout.plot.width)
  near(chart.size.height - native.grid.top - native.grid.bottom, layout.plot.height)
})

test('Pie and Radar centers/radii consume the authored plot instead of the chart center', () => {
  for (const type of ['pie', 'doughnut', 'radar']) {
    const { workbook, chart, data } = fixture(type, { plotArea: { ...manual } }), layout = layoutChartData(data, chart), native = chartEChartsOption(workbook, chart)
    if (type === 'radar') {
      const radar = chartRadarGeometry(data, chart, layout.plot), pointer = chartAxisPointerGeometry(data, chart, layout)
      assert.deepEqual(pointer.radar.center, radar.center)
      assert.deepEqual(native.radar.center, [radar.center.x - chart.position.x, radar.center.y - chart.position.y])
      near(native.radar.radius, radar.radius)
    } else {
      near(native.series[0].left, chart.size.width * manual.left)
      near(native.series[0].top, chart.size.height * manual.top)
    }
    for (const point of layout.series[0].points) assert.ok(point.x >= layout.plot.x && point.x <= layout.plot.x + layout.plot.width && point.y >= layout.plot.y && point.y <= layout.plot.y + layout.plot.height)
  }
})

test('3D geometry-only layout emits no paint and shares projected points with the rendered scene', () => {
  for (const type of ['column3D', 'bar3D', 'line3D', 'area3D', 'pie3D']) {
    const { workbook, chart, data } = fixture(type, { plotArea: { ...manual }, autoGradientFill: true }), layout = layoutChartData(data, chart)
    const geometry = chart3DScene(data, chart, layout, undefined, undefined, false), scene = chart3DScene(data, chart, layout)
    assert.deepEqual(geometry.commands, [])
    assert.deepEqual(geometry.points, scene.points)
    assert.deepEqual(layout.series.map(series => series.points), scene.points)
    assert.ok(scene.commands.length > 0)
    const native = chartEChartsOption(workbook, chart)
    assert.ok(native.graphic.elements.length > 0)
    assert.ok(scene.commands.filter(command => command.clip).every(command => command.clip.x === layout.plot.x && command.clip.y === layout.plot.y && command.clip.width === layout.plot.width && command.clip.height === layout.plot.height || command.clip.x === layout.frame.x && command.clip.y === layout.frame.y))
  }
})

test('editable Excel-view snapshots retain normalized plot/gradient settings independently of original models', () => {
  const { workbook, chart } = fixture('column', { plotArea: { ...manual }, autoGradientFill: true })
  workbook.charts = [chart]
  const view = toExcelWorkbookView(workbook)
  assert.deepEqual(view.charts[0].options.plotArea, manual)
  view.charts[0].options.plotArea.left = 0.1
  view.charts[0].size.width = 1000
  const reopened = applyExcelWorkbookView(workbook, view)
  assert.equal(chart.options.plotArea.left, 0.2)
  assert.equal(reopened.charts[0].options.plotArea.left, 0.1)
  assert.equal(reopened.charts[0].options.autoGradientFill, true)
  near(chartPlotFrame(reopened.charts[0]).plot.width, 600)
})
