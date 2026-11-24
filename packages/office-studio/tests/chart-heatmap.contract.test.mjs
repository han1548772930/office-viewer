import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { createChart, readChartData } from '../src/features/charts/chart-engine.ts'
import { projectChartRecords } from '../src/features/charts/chart-record-projection.ts'
import { deriveChartData, chartHeatmapPointStyle } from '../src/features/charts/chart-derived.ts'
import { drawDerivedChartData } from '../src/features/charts/chart-draw.ts'
import { layoutDerivedChartData } from '../src/features/charts/chart-layout.ts'
import { chartEChartsOption } from '../src/features/charts/echarts-option.ts'
import { chartAxisPointerCommands, chartAxisPointerGeometry } from '../src/features/charts/chart-axis-pointer.ts'
import { createChartHeatmapLegend, drawChartHeatmapCells } from '../src/features/charts/chart-heatmap-draw.ts'
import { chartHeatmapFocusedCells, createChartHeatmapPointerGeometry, hitChartHeatmap } from '../src/features/charts/chart-heatmap-interaction.ts'
import { chartHeatmapBoundaryText, chartHeatmapCellRect, chartHeatmapExtent, chartHeatmapLabelAnchor, chartHeatmapLabelText,
  chartHeatmapPieceIndex, chartHeatmapScaleColor, chartHeatmapValueRatio, chartHeatmapValueVisible, createChartHeatmapScale,
  hitChartHeatmapCell, resolveChartHeatmapLabel } from '../src/features/charts/chart-heatmap.ts'
import { chartCommonControls, chartSeriesControls } from '../src/features/charts/chart-settings-controls.ts'
import { chartDataMappingControls, chartTypeSettingsControls } from '../src/features/charts/chart-type-settings-controls.ts'
import { cloneChartDraftConfig } from '../src/features/charts/chart-settings-draft.ts'

// P21 implementation drafts only. No test, server, browser, build or benchmark
// has run in this phase. Browser/Office samples remain separate acceptance work.
const palette = ['#0000ff', '#ff0000']
function model(patch = {}) {
  return { id: 'heatmap-draft', sheetId: 'sheet', revision: 0, type: 'heatmap', categoryField: 0, valueFields: [1, 2],
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 3, endCol: 2 } },
    position: { x: 0, y: 0 }, size: { width: 520, height: 360 }, palette,
    legend: { visible: true, position: 'bottom' }, xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, ...patch }
}
const source = () => ({ categories: ['North', 'South', 'West'], series: [{ field: 1, name: 'Units', values: [1, 3, 5] }, { field: 2, name: 'Revenue', values: [2, 4, 6] }] })
const center = rect => ({ x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 })
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-7, `${actual} ≈ ${expected}`)
const texts = commands => commands.filter(command => command.kind === 'text').map(command => command.text)
function prepare(data = source(), chart = model(), theme = {}) {
  const derived = deriveChartData(data, chart, { preparedData: data }), layout = layoutDerivedChartData(derived, chart)
  const scale = createChartHeatmapScale(chartHeatmapExtent(data, chart), chart.palette, chart.options?.heatmap?.visualMapType)
  return { data, chart, derived, layout, scale, commands: drawDerivedChartData(derived, chart, theme),
    pointer: chartAxisPointerGeometry(data, chart, layout, theme), option: chartEChartsOption({ sheets: [] }, chart, theme, { preparedData: data }) }
}

test('automatic extents use actual signed finite data, an empty 0/0 extent and detached author bounds', () => {
  const values = source(), chart = model(), before = structuredClone([values, chart])
  assert.deepEqual(chartHeatmapExtent(values, chart), { min: 1, max: 6 })
  assert.deepEqual(chartHeatmapExtent({ categories: [], series: [] }, chart), { min: 0, max: 0 })
  assert.deepEqual(chartHeatmapExtent({ categories: ['A', 'B', 'C', 'D'], series: [{ field: 1, name: 'Negative', values: [-8, -2, NaN, Infinity] }] }, chart), { min: -8, max: -2 })
  assert.deepEqual(chartHeatmapExtent(values, model({ options: { heatmap: { min: 10 } } })), { min: 6, max: 10 })
  assert.deepEqual(chartHeatmapExtent(values, model({ options: { heatmap: { min: 4, max: 4 } } })), { min: 4, max: 4 })
  assert.throws(() => chartHeatmapExtent(values, model({ options: { heatmap: { min: 7, max: 1 } } })), RangeError)
  assert.throws(() => createChartHeatmapScale({ min: 0, max: Infinity }, palette), RangeError)
  assert.deepEqual([values, chart], before)
})

test('constant ranges resolve the midpoint color and a single visible piece without dividing by zero', () => {
  for (const type of ['continuous', 'piecewise']) {
    const scale = createChartHeatmapScale({ min: 7, max: 7 }, ['black', 'white'], type)
    assert.equal(chartHeatmapValueRatio(7, scale.extent), 0.5)
    assert.equal(chartHeatmapScaleColor(scale, 7), 'rgba(128, 128, 128, 1)')
    assert.equal(scale.pieces.length, 1)
    assert.equal(chartHeatmapValueVisible(scale, 8), type === 'continuous')
  }
})

test('extreme and subnormal finite ranges keep eight stable bounds and finite endpoint/midpoint colors', () => {
  for (const extent of [{ min: -Number.MAX_VALUE, max: Number.MAX_VALUE }, { min: Number.MIN_VALUE, max: Number.MIN_VALUE * 9 }]) {
    const scale = createChartHeatmapScale(extent, palette, 'piecewise')
    assert.equal(scale.pieces.length, 8)
    assert.equal(scale.pieces[0].lower, extent.min); assert.equal(scale.pieces.at(-1).upper, extent.max)
    assert.equal(scale.pieces.every(piece => Number.isFinite(piece.lower) && Number.isFinite(piece.upper) && piece.lower < piece.upper), true)
    assert.equal(chartHeatmapScaleColor(scale, extent.min), 'rgba(0, 0, 255, 1)')
    assert.equal(chartHeatmapScaleColor(scale, extent.max), 'rgba(255, 0, 0, 1)')
  }
  const continuous = createChartHeatmapScale({ min: -Number.MAX_VALUE, max: Number.MAX_VALUE }, palette)
  assert.equal(chartHeatmapScaleColor(continuous, 0), 'rgba(128, 0, 128, 1)')
})

test('eight pieces are independent of palette length and boundary values belong to the preceding piece', () => {
  const scale = createChartHeatmapScale({ min: 0, max: 8 }, palette, 'piecewise')
  assert.equal(scale.pieces.length, 8)
  assert.deepEqual([0, 1, 1.01, 7, 8, -1, 9].map(value => chartHeatmapPieceIndex(scale, value)), [0, 0, 1, 6, 7, -1, -1])
  assert.equal(chartHeatmapValueVisible(scale, -1), false)
  assert.equal(chartHeatmapValueVisible(createChartHeatmapScale(scale.extent, palette), -1), true)
  assert.notEqual(scale.pieces[1].color, scale.pieces[0].color)
})

test('fields are X columns, categories are Y rows and the first category is at the bottom by default', () => {
  const result = prepare(source(), model({ legend: { visible: false, position: 'bottom' } }))
  const first = result.layout.series[0].points[0], lastField = result.layout.series[1].points[0]
  near(first.x, result.layout.plot.x + result.layout.plot.width / 4)
  near(lastField.x, result.layout.plot.x + result.layout.plot.width * 3 / 4)
  near(first.y, result.layout.plot.y + result.layout.plot.height * 5 / 6)
  const header = result.commands.find(command => command.kind === 'text' && command.text === 'Units')
  const category = result.commands.find(command => command.kind === 'text' && command.text === 'North')
  near(header.x, first.x); assert.ok(header.y > result.layout.plot.y + result.layout.plot.height)
  near(category.y, first.y); assert.ok(category.x < result.layout.plot.x)
  assert.equal(result.layout.rightMin, undefined)
})

test('both category-axis reversals and geometric hits preserve original field/category identities', () => {
  for (const reverseX of [false, true]) for (const reverseY of [false, true]) {
    const chart = model({ xAxis: { scaleType: 'linear', reverse: reverseX }, yAxis: { scaleType: 'linear', reverse: reverseY } }), data = source()
    const plot = { x: 10, y: 20, width: 240, height: 180 }
    for (let seriesPosition = 0; seriesPosition < 2; seriesPosition++) for (let categoryIndex = 0; categoryIndex < 3; categoryIndex++) {
      const rect = chartHeatmapCellRect(data, chart, plot, seriesPosition, categoryIndex)
      assert.deepEqual(hitChartHeatmapCell(data, chart, plot, center(rect)), { seriesPosition, categoryIndex })
    }
    const first = chartHeatmapCellRect(data, chart, plot, 0, 0)
    near(first.x, reverseX ? 130 : 10); near(first.y, reverseY ? 20 : 140)
    assert.equal(hitChartHeatmapCell(data, chart, plot, { x: 9, y: 20 }), undefined)
    assert.equal(chartHeatmapCellRect(data, chart, plot, 2, 0), undefined)
  }
})

test('source formatted text follows field selection and visible-field projection without becoming numeric input', () => {
  const records = { fields: ['Region', 'Units', 'Rate'], records: [['North', 1234.5, 0.125], ['South', 2, 0.5]],
    formattedRecords: [['North', '$1,234.50', '12.5%'], ['South', '2.00', '50.0%']], visibleFields: [0, 2] }
  const before = structuredClone(records), projected = projectChartRecords(records, model({ valueFields: [2, 1] }))
  assert.deepEqual(projected.series.map(series => series.field), [2])
  assert.deepEqual(projected.series[0].values, [0.125, 0.5])
  assert.deepEqual(projected.series[0].formattedValues, ['12.5%', '50.0%'])
  assert.deepEqual(records, before)
})

test('ordinary Sheets source formatting uses numberFormat while retaining raw values and source identity', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.cells = { '0:0': createCell(0, 0, 'Region'), '0:1': createCell(0, 1, 'Rate'), '1:0': createCell(1, 0, 'North'), '1:1': createCell(1, 1, 0.125) }
  sheet.cells['1:1'].style = { numberFormat: '0.0%' }; sheet.cells['1:1'].display = 'stale display'
  const chart = createChart(workbook, sheet.id, { ...model({ valueFields: [1] }), source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 } } })
  const data = readChartData(workbook, chart)
  assert.deepEqual(data.series[0].values, [0.125])
  assert.deepEqual(data.series[0].formattedValues, ['12.5%'])
  assert.equal(data.series[0].field, 1)
})

test('cell labels use source text, explicit format, 12px/family defaults and one value-only resolver', () => {
  const chart = model({ options: { appearance: { font: { family: 'Aptos', fontSize: 24 } }, allSeriesStyle: { label: { visible: true, contentType: 14 } }, heatmap: { unit: ' kg ' } } })
  const data = { categories: ['A', 'B'], series: [{ field: 1, name: 'Rate', values: [0.125, 1234.5], formattedValues: ['12.5%', '$1,234.50'] }] }
  const result = prepare(data, chart), label = resolveChartHeatmapLabel(chart, chart.options.allSeriesStyle.label, '#123456')
  assert.equal(label.fontSize, 12); assert.equal(label.family, 'Aptos'); assert.equal(label.color, '#123456')
  assert.equal(chartHeatmapLabelText(0.125, '12.5%', label), '12.5%')
  assert.equal(chartHeatmapLabelText(0.125, '12.5%', { format: '0.00' }), '0.13')
  assert.equal(chartHeatmapLabelText(0.125, '', {}), '')
  for (const text of ['12.5%', '$1,234.50']) assert.equal(texts(result.commands).filter(value => value === text).length, 1)
  assert.equal(result.commands.filter(command => command.kind === 'text' && ['12.5%', '$1,234.50'].includes(command.text)).every(command => command.font.includes('12px Aptos')), true)
  assert.equal(texts(result.commands).some(text => text.includes('Rate,A') || text.endsWith('% kg')), false)
})

test('inside, top and bottom label positions consume the cell rectangle in both adapters', () => {
  const rect = { x: 10, y: 20, width: 80, height: 60 }
  near(chartHeatmapLabelAnchor(rect, { position: 'inside', fontSize: 12 }).y, 50)
  near(chartHeatmapLabelAnchor(rect, { position: 'top', fontSize: 12 }).y, 9)
  near(chartHeatmapLabelAnchor(rect, { position: 'bottom', fontSize: 12 }).y, 91)
  for (const position of ['inside', 'top', 'bottom']) {
    const result = prepare(source(), model({ options: { heatmap: { label: { visible: true, position } } } }))
    assert.equal(result.option.series[0].data[0].label.position, position)
    const cellCommands = []; drawChartHeatmapCells(cellCommands, result.data, result.chart, result.layout, '#334155', palette, result.scale)
    const text = cellCommands.find(command => command.kind === 'text' && command.text === '1'), point = result.layout.series[0].points[0]
    assert.equal(position === 'inside' ? text.y === point.y : position === 'top' ? text.y < point.y : text.y > point.y, true)
  }
})

test('visible continuous color legends follow all four positions and append trimmed units only at boundaries', () => {
  for (const position of ['top', 'right', 'bottom', 'left']) {
    const result = prepare(source(), model({ legend: { visible: true, position, font: { color: '#abcdef', family: 'Georgia' } }, options: { heatmap: { unit: ' kg ' } } }))
    const legend = result.pointer.heatmap.legend
    assert.ok(legend.bar); assert.equal(legend.horizontal, position === 'top' || position === 'bottom')
    assert.deepEqual(texts(legend.commands), legend.horizontal ? ['1 kg', '6 kg'] : ['6 kg', '1 kg'])
    assert.equal(legend.commands.filter(command => command.kind === 'text').every(command => command.color === '#abcdef' && command.font.includes('Georgia')), true)
    assert.equal(result.option.legend.show, false)
    assert.ok(result.option.graphic.elements.length > 0)
  }
  assert.equal(chartHeatmapBoundaryText(7, '   '), '7')
  const hidden = prepare(source(), model({ legend: { visible: false, position: 'bottom' } })).pointer.heatmap.legend
  assert.deepEqual(hidden.commands, []); assert.equal(hidden.bar, undefined)
})

test('piecewise legends expose eight actual intervals without series keys or continuous-only units', () => {
  const result = prepare(source(), model({ options: { heatmap: { visualMapType: 'piecewise', min: 0, max: 8, unit: 'kg' } } }))
  const legend = result.pointer.heatmap.legend
  assert.equal(legend.pieces.length, 8)
  assert.equal(texts(legend.commands).every(text => !text.includes('kg') && !['Units', 'Revenue'].includes(text)), true)
  assert.equal(result.option.visualMap.pieces.length, 8)
  assert.equal(result.option.visualMap.selectedMode, false); assert.equal(result.option.visualMap.calculable, false)
})

test('native heatmap points share field/category order, resolved author colors, labels and the visible Canvas legend', () => {
  const data = source(), chart = model({ options: { heatmap: { label: { visible: true, format: '0.00' } },
    namedSeries: [{ name: 'Revenue', style: { color: 'purple', dataPoints: { 1: { color: 'green', fillOpacity: 0.4, border: { color: 'black', width: 2 } } } } }] } })
  const result = prepare(data, chart), native = result.option
  assert.deepEqual(native.xAxis.data, ['Units', 'Revenue']); assert.deepEqual(native.yAxis.data, data.categories)
  assert.equal(native.yAxis.inverse, false)
  assert.deepEqual(native.series[0].data.map(point => point.value), [[0, 0, 1], [0, 1, 3], [0, 2, 5], [1, 0, 2], [1, 1, 4], [1, 2, 6]])
  assert.equal(native.series[0].data.every(point => point.visualMap === false), true)
  assert.equal(native.series[0].data[4].itemStyle.color, 'rgba(0, 128, 0, 0.4)')
  assert.equal(native.series[0].data[4].label.formatter, '4.00')
  assert.ok(result.commands.some(command => command.kind === 'rect' && command.fill === native.series[0].data[4].itemStyle.color && command.strokeWidth === 2))
  assert.equal(native.officeChartPointer.mode, 'graphic')
  assert.equal(chartHeatmapPointStyle(chart, 2, 1, 0, 2, result.scale.extent, palette, 'Revenue', result.scale).color, 'purple')
})

test('piecewise values outside authored bounds stay in native data and public layout while their marks and labels are inactive', () => {
  const data = { categories: ['Below', 'Inside', 'Above'], series: [{ field: 1, name: 'Units', values: [-1, 4, 9] }] }
  const result = prepare(data, model({ options: { heatmap: { visualMapType: 'piecewise', min: 0, max: 8, label: { visible: true } } } }))
  assert.deepEqual(result.layout.series[0].points.map(point => point.value), [-1, 4, 9])
  assert.deepEqual(result.option.series[0].data.map(point => point.itemStyle.opacity), [0, 1, 0])
  assert.deepEqual(result.option.series[0].data.map(point => point.label.show), [false, true, false])
  const marks = []; drawChartHeatmapCells(marks, data, result.chart, result.layout, '#334155', palette, result.scale)
  assert.equal(marks.filter(command => command.kind === 'rect').length, 1)
  assert.deepEqual(texts(marks), ['4'])
})

test('continuous legend hover links a bounded value band and unbounded endpoint data without changing authors', () => {
  const data = { categories: ['Below', 'Low', 'High', 'Above'], series: [{ field: 1, name: 'Units', values: [-100, 0, 10, 100] }] }
  const chart = model({ options: { heatmap: { min: 0, max: 10 } } }), result = prepare(data, chart), geometry = result.pointer.heatmap
  const before = structuredClone([data, chart]), bar = geometry.legend.bar, bands = geometry.bands
  const low = hitChartHeatmap(geometry, { x: bar.x, y: bar.y + 6 }), high = hitChartHeatmap(geometry, { x: bar.x + bar.width, y: bar.y + 6 })
  assert.deepEqual(chartHeatmapFocusedCells(geometry, low), [0, 1]); assert.deepEqual(chartHeatmapFocusedCells(geometry, high), [2, 3])
  assert.ok(chartAxisPointerCommands(result.pointer, { x: bar.x, y: bar.y + 6 }).some(command => command.kind === 'polygon'))
  assert.deepEqual(chartAxisPointerCommands(result.pointer), [])
  assert.equal(geometry.bands, bands); assert.deepEqual([data, chart], before)
})

test('piecewise legend hits use the drawn swatch/label rectangle and the original data-point indices', () => {
  const data = { categories: ['Min', 'Boundary', 'Next', 'Max'], series: [{ field: 1, name: 'Units', values: [0, 1, 1.01, 8] }] }
  const chart = model({ legend: { visible: true, position: 'right' }, options: { heatmap: { visualMapType: 'piecewise', min: 0, max: 8 } } })
  const result = prepare(data, chart), geometry = createChartHeatmapPointerGeometry(data, chart, result.layout, palette, '#334155')
  const swatch = geometry.legend.pieces.find(piece => piece.pieceIndex === 0), hit = hitChartHeatmap(geometry, center(swatch.hit))
  assert.deepEqual(hit, { kind: 'piece', pieceIndex: 0 }); assert.deepEqual(chartHeatmapFocusedCells(geometry, hit), [0, 1])
  assert.equal(hitChartHeatmap(geometry, { x: -1, y: -1 }), undefined)
})

test('cell pointers show X field names and Y categories for both axis reversals and suppress invalid-cell emphasis', () => {
  const data = source(); data.series[0].values[1] = null
  for (const reverse of [false, true]) {
    const chart = model({ legend: { visible: false, position: 'bottom' }, xAxis: { scaleType: 'linear', reverse }, yAxis: { scaleType: 'linear', reverse } })
    const result = prepare(data, chart), point = result.layout.series[1].points.find(point => point.categoryIndex === 0)
    assert.deepEqual(texts(chartAxisPointerCommands(result.pointer, point)), ['Revenue', 'North'])
    const empty = chartHeatmapCellRect(data, chart, result.layout.plot, 0, 1)
    assert.equal(hitChartHeatmap(result.pointer.heatmap, center(empty)), undefined)
  }
})

test('Heatmap controls use fixed grouping/defaults, hide label details and reset detached author fields', () => {
  const chart = model(), original = structuredClone(chart), common = chartCommonControls(chart), specific = chartTypeSettingsControls(chart)
  const legendIndex = common.findIndex(control => control.key === 'legend.position')
  assert.equal(common[legendIndex + 1].key, 'heatmap.visualMapType')
  assert.deepEqual(specific.filter(control => !control.visible || control.visible(chart)).map(control => control.key), ['heatmap.min', 'heatmap.max', 'heatmap.unit', 'heatmap.label.visible'])
  assert.equal(specific.some(control => control.key.endsWith('.contentType')), false)
  assert.deepEqual(chartDataMappingControls(chart, 3).slice(0, 2).map(control => control.label), ['纵轴分类字段', '横轴数值字段'])
  const draft = cloneChartDraftConfig(chart), toggle = specific.find(control => control.key === 'heatmap.label.visible')
  toggle.write(draft, true)
  const position = specific.find(control => control.key === 'heatmap.label.position'), format = specific.find(control => control.key === 'heatmap.label.format')
  assert.deepEqual(position.choices.map(choice => choice.value), ['inside', 'top', 'bottom'])
  position.write(draft, 'bottom'); format.write(draft, '0.00')
  assert.equal(format.visible(draft), true)
  position.reset(draft); assert.equal(draft.options.heatmap.label.position, undefined); assert.equal(draft.options.heatmap.label.format, '0.00')
  toggle.reset(draft); assert.equal(draft.options.heatmap.label, undefined)
  assert.deepEqual(chart, original)
  assert.equal(chartSeriesControls(1).find(control => control.key.endsWith('.label.contentType')).visible(chart), false)
  assert.equal(chartSeriesControls(1).find(control => control.key.endsWith('.point.size')).visible(chart), false)
})
