import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { createChart, readChartData } from '../src/features/charts/chart-engine.ts'
import { CHART_TYPE_CATALOG } from '../src/features/charts/chart-catalog.ts'
import { chartPlotFrame } from '../src/features/charts/chart-layout.ts'
import { chartPlotAreaControls } from '../src/features/charts/chart-plot-area-controls.ts'
import { validateChartPlotArea } from '../src/features/charts/chart-plot-area.ts'
import { cloneChartOptions, mergeChartOptions } from '../src/features/charts/chart-options.ts'
import { chartEChartsOption } from '../src/features/charts/echarts-option.ts'
import { CHART_PLOT_AREA_HANDLES, createChartPlotAreaTarget, hitChartPlotArea, chartPlotAreaHandlePoints,
  chartPlotAreaOutlineCommands, chartPlotAreaDragLayout, chartPlotAreasEqual } from '../src/features/charts/chart-plot-area-interaction.ts'

// Drafts only. This batch has not run tests, a browser, types or a benchmark.
const frame = { x: 0, y: 0, width: 500, height: 300 }, plot = { x: 80, y: 50, width: 300, height: 180 }
const automatic = { x: 54, y: 12, width: 434, height: 248 }
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-9, `${actual} ≈ ${expected}`)
const target = () => createChartPlotAreaTarget('column', frame, plot, automatic, false)
function fixture(type = 'line', options = {}) {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  const rows = [['Region', 'Main'], ['North', 7], ['South', 4]]
  sheet.cells = Object.fromEntries(rows.flatMap((row, r) => row.map((value, c) => [`${r}:${c}`, createCell(r, c, value)])))
  const chart = createChart(workbook, sheet.id, { type,
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 2, endCol: 1 } }, categoryField: 0, valueFields: [1],
    position: { x: 100, y: 40 }, size: { width: 500, height: 300 }, title: 'Main', subtitle: 'Sub',
    legend: { visible: true, position: 'bottom' }, options })
  workbook.charts = [chart]
  return { workbook, chart, data: readChartData(workbook, chart) }
}

test('surface plot targets are detached and follow the 31 capable variants without persisting view geometry', () => {
  const original = structuredClone({ frame, plot, automatic }), selected = target()
  assert.deepEqual(selected.bounds, plot); assert.equal(selected.manual, false)
  near(selected.area.left, 0.16); near(selected.area.top, 1 / 6); near(selected.area.width, 0.6)
  selected.frame.x = 10; selected.plot.width = 12; selected.automatic.left = 0
  assert.deepEqual({ frame, plot, automatic }, original)
  const families = ['column', 'bar', 'line', 'area', 'scatter', 'bubble', 'combination', 'pie', 'radar']
  let count = 0
  for (const entry of CHART_TYPE_CATALOG) {
    const supported = createChartPlotAreaTarget(entry.value, frame, plot, automatic, false) !== undefined
    assert.equal(supported, families.includes(entry.family), entry.id)
    if (supported) count++
  }
  assert.equal(count, 31)
  assert.throws(() => createChartPlotAreaTarget('column', { ...frame, width: 0 }, plot, automatic, false), RangeError)
})

test('only visible clipped plot edges hit before selection; the selected body excludes neighbouring charts', () => {
  const selected = target(), center = { x: 230, y: 140 }
  assert.equal(hitChartPlotArea(selected, center), false)
  assert.equal(hitChartPlotArea(selected, center, 5, true), true)
  assert.equal(hitChartPlotArea(selected, { x: plot.x - 4, y: center.y }), true)
  assert.equal(hitChartPlotArea(selected, { x: plot.x - 6, y: center.y }), false)
  assert.equal(hitChartPlotArea(selected, { x: -1, y: center.y }, 100, true), false)
  assert.equal(hitChartPlotArea(selected, { x: NaN, y: 100 }), false)
  const clipped = createChartPlotAreaTarget('pie', { x: 0, y: 0, width: 100, height: 100 },
    { x: 90, y: 20, width: 30, height: 50 }, { x: 5, y: 5, width: 90, height: 90 }, false)
  assert.deepEqual(clipped.bounds, { x: 90, y: 20, width: 10, height: 50 })
  assert.equal(hitChartPlotArea(clipped, { x: 101, y: 40 }, 5, true), false)
  const invisible = createChartPlotAreaTarget('column', { x: 0, y: 0, width: 10, height: 10 },
    { x: 54, y: 12, width: 1, height: 1 }, { x: 54, y: 12, width: 1, height: 1 }, false)
  assert.equal(invisible.bounds, undefined); assert.equal(hitChartPlotArea(invisible, { x: 5, y: 5 }, 100, true), false)
  assert.deepEqual(chartPlotAreaHandlePoints(invisible), []); assert.deepEqual(chartPlotAreaOutlineCommands(invisible, { x: 0, y: 0 }, 'blue'), [])
  assert.doesNotThrow(() => validateChartPlotArea(invisible.area))
})

test('eight handle locations and the Canvas outline describe one clipped rectangle in chart coordinates', () => {
  const selected = target(), handles = chartPlotAreaHandlePoints(selected)
  assert.equal(CHART_PLOT_AREA_HANDLES.length, 8); assert.equal(new Set(handles.map(point => point.handle)).size, 8)
  assert.deepEqual(handles.find(point => point.handle === 'nw'), { handle: 'nw', x: 80, y: 50 })
  assert.deepEqual(handles.find(point => point.handle === 'se'), { handle: 'se', x: 380, y: 230 })
  assert.deepEqual(handles.find(point => point.handle === 'n'), { handle: 'n', x: 230, y: 50 })
  const commands = chartPlotAreaOutlineCommands(selected, { x: 100, y: 40 }, '#123456', 1.5)
  assert.equal(commands.length, 4)
  assert.deepEqual(commands[0], { kind: 'line', x1: 180, y1: 90, x2: 480, y2: 90, color: '#123456', width: 1.5,
    clip: { x: 100, y: 40, width: 500, height: 300 } })
  assert.equal(commands[3].x2, commands[0].x1); assert.equal(commands[3].y2, commands[0].y1)
})

test('moving uses the original plot and retains its dimensions while clamping the whole rectangle', () => {
  const selected = target(), original = structuredClone(selected)
  const moved = chartPlotAreaDragLayout(selected, frame, { x: 50, y: 30 })
  near(moved.left, selected.area.left + 0.1); near(moved.top, selected.area.top + 0.1)
  assert.equal(moved.width, selected.area.width); assert.equal(moved.height, selected.area.height)
  const negative = chartPlotAreaDragLayout(selected, frame, { x: -1000, y: -1000 }), positive = chartPlotAreaDragLayout(selected, frame, { x: 1000, y: 1000 })
  assert.equal(negative.left, 0); assert.equal(negative.top, 0)
  near(positive.left + positive.width, 1); near(positive.top + positive.height, 1)
  assert.deepEqual(selected, original)
  assert.throws(() => chartPlotAreaDragLayout(selected, { width: 0, height: 100 }, { x: 0, y: 0 }), RangeError)
  assert.throws(() => chartPlotAreaDragLayout(selected, frame, { x: Infinity, y: 0 }), RangeError)
})

test('each resize direction fixes its opposite edges and remains positive inside the frame', () => {
  const selected = target(), before = selected.area
  for (const { handle } of CHART_PLOT_AREA_HANDLES) {
    const resized = chartPlotAreaDragLayout(selected, frame, { x: 40, y: 20 }, handle)
    assert.doesNotThrow(() => validateChartPlotArea(resized), handle)
    if (handle.includes('w')) near(resized.left + resized.width, before.left + before.width)
    else near(resized.left, before.left)
    if (handle.includes('n')) near(resized.top + resized.height, before.top + before.height)
    else near(resized.top, before.top)
    if (!handle.includes('w') && !handle.includes('e')) near(resized.width, before.width)
    if (!handle.includes('n') && !handle.includes('s')) near(resized.height, before.height)
  }
  for (const { handle } of CHART_PLOT_AREA_HANDLES) {
    const resized = chartPlotAreaDragLayout(selected, frame, { x: handle.includes('w') ? 2000 : -2000, y: handle.includes('n') ? 2000 : -2000 }, handle)
    assert.doesNotThrow(() => validateChartPlotArea(resized), handle)
    assert.ok(resized.width >= 8 / frame.width - 1e-12); assert.ok(resized.height >= 8 / frame.height - 1e-12)
  }
  const tiny = createChartPlotAreaTarget('column', frame, { x: 80, y: 50, width: 2, height: 3 }, automatic, true)
  const unchanged = chartPlotAreaDragLayout(tiny, frame, { x: 0, y: 0 }, 'se')
  assert.equal(chartPlotAreasEqual(tiny.area, unchanged), true)
  const smallerThanAutomaticSeed = createChartPlotAreaTarget('column', frame, { x: 80, y: 50, width: 0.00001, height: 0.00002 }, automatic, true)
  near(smallerThanAutomaticSeed.area.width, 0.00001 / frame.width)
  assert.equal(chartPlotAreasEqual(smallerThanAutomaticSeed.area, chartPlotAreaDragLayout(smallerThanAutomaticSeed, frame, { x: 0, y: 0 }, 'se')), true)
})

test('coordinate restoration keeps other authored values; an incompatible partial restore stays an invalid draft until corrected', () => {
  const { chart, data } = fixture('column', { plotArea: { left: 0.2, top: 0.25, width: 0.6, height: 0.5 }, autoGradientFill: true,
    titleStyle: { color: '#123456' } }), config = structuredClone(chart), controls = chartPlotAreaControls(data.series)
  const left = controls.find(control => control.key === 'plotArea.left'), width = controls.find(control => control.key === 'plotArea.width')
  const automaticFrame = chartPlotFrame({ ...chart, options: { ...chart.options, plotArea: undefined } }, data.series)
  width.reset(config)
  near(config.options.plotArea.width, automaticFrame.plot.width / chart.size.width)
  assert.equal(config.options.plotArea.left, 0.2); assert.equal(config.options.plotArea.top, 0.25); assert.equal(config.options.plotArea.height, 0.5)
  assert.throws(() => cloneChartOptions(config.options), RangeError)
  left.reset(config)
  assert.doesNotThrow(() => cloneChartOptions(config.options))
  controls.find(control => control.key === 'plotArea.manual').reset(config)
  assert.equal(config.options.plotArea, undefined); assert.equal(config.options.autoGradientFill, true)
  assert.deepEqual(config.options.titleStyle, chart.options.titleStyle)
  assert.deepEqual(chart.options.plotArea, { left: 0.2, top: 0.25, width: 0.6, height: 0.5 })
})

test('native and shared options expose their exact painting plot plus independent automatic restoration geometry', () => {
  for (const type of ['line', 'column', 'bar', 'area', 'pie', 'doughnut', 'radar', 'column3D', 'pie3D']) {
    for (const shared of [false, true]) {
      const { workbook, chart, data } = fixture(type, { plotArea: { left: 0.15, top: 0.2, width: 0.65, height: 0.6 },
        ...(shared ? { allSeriesStyle: { label: { visible: true, position: 'insideEnd' } } } : {}) })
      const local = { ...chart, position: { x: 0, y: 0 } }, expected = chartPlotFrame(local, data.series), option = chartEChartsOption(workbook, chart)
      assert.deepEqual(option.officeChartPlotArea.frame, expected.frame, type)
      assert.deepEqual(option.officeChartPlotArea.plot, expected.plot, type)
      assert.equal(option.officeChartPlotArea.manual, true)
      const reset = chartPlotFrame({ ...local, options: { ...local.options, plotArea: undefined } }, data.series)
      near(option.officeChartPlotArea.automatic.left, reset.plot.x / chart.size.width)
      const rebuilt = option.officeChartLegend?.rebuild?.({ Main: false })
      if (rebuilt) assert.deepEqual(rebuilt.plotArea.plot, rebuilt.option.officeChartPlotArea?.plot ?? expected.plot)
      assert.equal(JSON.stringify(workbook).includes('officeChartPlotArea'), false)
    }
  }
})

test('plot edits and automatic reset preserve both title authors and geometry-independent options', () => {
  const { workbook, chart } = fixture('column', { autoGradientFill: false,
    titleStyle: { manualLayout: { left: 0.2, top: 0.1 }, rotate: 25 }, subtitleStyle: { color: '#123456' } })
  const before = structuredClone(chart), selected = chartEChartsOption(workbook, chart).officeChartPlotArea
  const options = mergeChartOptions(chart.options, { plotArea: chartPlotAreaDragLayout(selected, chart.size, { x: 20, y: 15 }) })
  assert.deepEqual(options.titleStyle, before.options.titleStyle); assert.deepEqual(options.subtitleStyle, before.options.subtitleStyle)
  const reset = mergeChartOptions(options, { plotArea: undefined })
  assert.equal(reset.plotArea, undefined); assert.equal(reset.autoGradientFill, false)
  assert.deepEqual(reset.titleStyle, before.options.titleStyle); assert.deepEqual(chart, before)
})
