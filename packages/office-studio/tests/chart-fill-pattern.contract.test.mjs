import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { createChart, readChartData, updateChart } from '../src/features/charts/chart-engine.ts'
import { chartPointStyle } from '../src/features/charts/chart-derived.ts'
import { drawChartData, drawChartSector } from '../src/features/charts/chart-draw.ts'
import { cloneChartOptions, mergeChartOptions, chartSeriesStyle } from '../src/features/charts/chart-options.ts'
import { chartSeriesStyleValues, patchChartSeriesOverride } from '../src/features/charts/chart-cartesian-api.ts'
import { CHART_FILL_PATTERN_PRESETS, chartFillPatternPaint } from '../src/features/charts/chart-fill-pattern.ts'
import { chartFillPatternTile } from '../src/features/charts/chart-fill-pattern-geometry.ts'
import { chartPolygon, chartRect, chartSymbol } from '../src/features/charts/chart-geometry.ts'
import { chartSeriesControls } from '../src/features/charts/chart-settings-controls.ts'
import { chartEChartsOption } from '../src/features/charts/echarts-option.ts'

// P14 implementation drafts. No test in this file has run in this implementation phase.
const pattern = (preset = 'pct50', foregroundColor = '#ff0000', backgroundColor = '#ffff00') => ({ preset, foregroundColor, backgroundColor })
const area = points => Math.abs(points.reduce((sum, point, index) => { const next = points[(index + 1) % points.length]; return sum + point.x * next.y - next.x * point.y }, 0)) / 2
const fillArea = (commands, color) => commands.filter(command => command.kind === 'polygon' && command.fill === color).reduce((sum, command) => sum + area(command.points), 0)
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-7, `${actual} ≈ ${expected}`)
function model(type = 'column', options = {}) {
  return { id: 'pattern', sheetId: 'sheet', type, revision: 0, categoryField: 0, valueFields: [1],
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 3, endCol: 1 } }, position: { x: 0, y: 0 },
    size: { width: 360, height: 240 }, legend: { visible: false, position: 'bottom' }, xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, options }
}
function workbookModel(type, options = {}) {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.cells = Object.fromEntries([['X', 'Units'], [1, 10], [2, 20], [3, 30]].flatMap((row, r) => row.map((value, c) => [`${r}:${c}`, createCell(value)])))
  return { workbook, chart: createChart(workbook, sheet.id, model(type, options)) }
}

test('pattern omission inherits, null disables, and point overrides remain independent across every style layer', () => {
  const global = pattern('horz'), named = pattern('upDiag'), point = pattern('smCheck')
  const options = { allSeriesStyle: { fillPattern: global, fillOpacity: 0.8, dataPoints: { 0: { fillPattern: point } } },
    series: { 1: { fillPattern: null, border: { color: 'black' } } }, namedSeries: [{ name: 'Units', style: { fillPattern: named, dataPoints: { 1: { fillPattern: null } } } }] }
  const chart = model('column', options), style = chartSeriesStyle(chart, 1, 0, undefined, 'Units')
  assert.deepEqual(style.fillPattern, named)
  assert.deepEqual(chartPointStyle(style, 0).fillPattern, point)
  assert.equal(chartPointStyle(style, 1).fillPattern, null)
  assert.equal(style.fillOpacity, 0.8)
  assert.equal(style.border.color, 'black')
  const resetName = patchChartSeriesOverride(options, { name: 'Units' }, chartSeriesStyleValues({ fillPattern: undefined }))
  const afterName = chartSeriesStyle(model('column', { ...options, ...resetName }), 1, 0, undefined, 'Units')
  assert.equal(afterName.fillPattern, null)
  const resetField = patchChartSeriesOverride({ ...options, ...resetName }, 1, chartSeriesStyleValues({ fillPattern: undefined }))
  assert.deepEqual(chartSeriesStyle(model('column', { ...options, ...resetName, ...resetField }), 1, 0, undefined, 'Units').fillPattern, global)
})

test('nullable API preserves a solid override, detaches full patterns and replaces pattern objects atomically', () => {
  const source = { fillPattern: pattern('horz'), dataPoints: { 2: { fillPattern: null, color: null } }, label: { bold: null } }
  const values = chartSeriesStyleValues(source)
  source.fillPattern.foregroundColor = '#000000'
  assert.equal(values.fillPattern.foregroundColor, '#ff0000')
  assert.equal(values.dataPoints[2].fillPattern, null)
  assert.equal(values.dataPoints[2].color, undefined)
  assert.equal(values.label.bold, undefined)
  assert.equal(chartSeriesStyleValues({ fillPattern: null }).fillPattern, null)
  const merged = mergeChartOptions({ allSeriesStyle: { fillPattern: { ...pattern('horz'), sourceTag: 'old' }, border: { width: 2 } } },
    { allSeriesStyle: { fillPattern: pattern('vert') } })
  assert.deepEqual(merged.allSeriesStyle.fillPattern, pattern('vert'))
  assert.equal(merged.allSeriesStyle.border.width, 2)
  assert.equal(JSON.parse(JSON.stringify(chartSeriesStyleValues({ fillPattern: null }))).fillPattern, null)
})

test('all DrawingML presets paint distinct tiles, and foreground plus background form a disjoint full tile', () => {
  const masks = new Set()
  for (const [preset] of CHART_FILL_PATTERN_PRESETS) {
    const tile = chartFillPatternTile(preset), pixels = Array(tile.width * tile.height).fill(0)
    for (const [rectangles, value] of [[tile.rectangles, 1], [tile.backgroundRectangles, 2]]) for (const rect of rectangles) {
      assert.ok(Object.isFrozen(rect))
      for (let y = rect.y; y < rect.y + rect.height; y++) for (let x = rect.x; x < rect.x + rect.width; x++) {
        assert.equal(pixels[y * tile.width + x], 0, `${preset} overlaps a translucent pixel`)
        pixels[y * tile.width + x] = value
      }
    }
    assert.ok(pixels.every(value => value !== 0), `${preset} leaves a hole`)
    assert.ok(pixels.includes(1) && pixels.includes(2), `${preset} has no visible two-color pattern`)
    masks.add(pixels.join(''))
    assert.equal(chartFillPatternTile(preset), tile)
  }
  assert.equal(masks.size, 54)
})

test('percentage patterns retain their coverage and line variants retain their direction and density', () => {
  for (const percent of [5, 10, 20, 25, 30, 40, 50, 60, 70, 75, 80, 90]) {
    const tile = chartFillPatternTile(`pct${percent}`), coverage = tile.rectangles.reduce((sum, rect) => sum + rect.width * rect.height, 0) / 256
    assert.ok(Math.abs(coverage - percent / 100) <= 1 / 64, `${percent}% coverage`)
  }
  const horizontal = chartFillPatternTile('horz').rectangles, vertical = chartFillPatternTile('vert').rectangles
  assert.ok(horizontal.every(rect => rect.width === 16 && rect.height === 1))
  assert.ok(vertical.every(rect => rect.width === 1 && rect.height === 16))
  const coverage = preset => chartFillPatternTile(preset).rectangles.reduce((sum, rect) => sum + rect.width * rect.height, 0)
  assert.ok(coverage('ltHorz') < coverage('horz'))
  assert.ok(coverage('horz') < coverage('dkHorz'))
})

test('patterned polygons conserve area without a full background under the translucent foreground', () => {
  const commands = [], paint = chartFillPatternPaint(pattern(), 0.25), rectangle = [{ x: 0, y: 0 }, { x: 16, y: 0 }, { x: 16, y: 16 }, { x: 0, y: 16 }]
  chartPolygon(commands, rectangle, 'blue', { color: 'black', width: 2, opacity: 0.8 }, { x: 0, y: 0, width: 16, height: 16 }, paint)
  near(fillArea(commands, 'rgba(255, 0, 0, 0.25)'), 128)
  near(fillArea(commands, 'rgba(255, 255, 0, 0.25)'), 128)
  const borders = commands.filter(command => command.kind === 'line')
  assert.equal(borders.length, 4)
  assert.ok(borders.every(command => command.color === 'rgba(0, 0, 0, 0.8)' && command.width === 2))
})

test('foreground and background are clipped to concave marks and annular sectors, including the inner hole', () => {
  const paint = pattern('smCheck'), commands = []
  chartPolygon(commands, [{ x: 0, y: 0 }, { x: 16, y: 0 }, { x: 16, y: 4 }, { x: 4, y: 4 }, { x: 4, y: 16 }, { x: 0, y: 16 }], 'blue', undefined, undefined, paint)
  near(fillArea(commands, paint.foregroundColor) + fillArea(commands, paint.backgroundColor), 112)
  assert.ok(commands.every(command => command.points.every(point => point.x <= 4 + 1e-8 || point.y <= 4 + 1e-8)))
  const annulus = [], center = { x: 32, y: 32 }
  drawChartSector(annulus, center, 12, 24, 0, Math.PI, 'blue', undefined, undefined, paint)
  assert.ok(annulus.length > 0)
  assert.ok(annulus.every(command => command.points.every(point => Math.hypot(point.x - center.x, point.y - center.y) >= 12 * Math.cos(Math.PI / 96) - 1e-7)))
})

test('patterned markers and rectangles use fill opacity independently from a retained border', () => {
  const source = pattern('diagCross'), before = structuredClone(source), commands = [], paint = chartFillPatternPaint(source, 0.6)
  chartRect(commands, { x: 0, y: 0, width: 16, height: 16 }, 'blue', { color: 'black', opacity: 0.5, width: 2 }, undefined, paint)
  chartSymbol(commands, { x: 32, y: 32 }, 20, 'blue', 'diamond', { color: 'black', opacity: 0.5, width: 2 }, undefined, paint)
  assert.ok(commands.some(command => command.kind === 'polygon' && command.fill === 'rgba(255, 0, 0, 0.6)'))
  assert.ok(commands.filter(command => command.kind === 'line').every(command => command.color === 'rgba(0, 0, 0, 0.5)'))
  assert.deepEqual(source, before)
})

test('Scatter, Heatmap, Boxplot, Area and Pie consume patterns through the shared ECharts geometry path', () => {
  for (const type of ['scatter', 'heatmap', 'boxplot', 'area', 'pie']) {
    const { workbook, chart } = workbookModel(type, { allSeriesStyle: { fillPattern: pattern(), fillOpacity: 1 } })
    const option = chartEChartsOption(workbook, chart), scene = drawChartData(readChartData(workbook, chart), chart)
    assert.ok(option.graphic.elements.length > 0, type)
    assert.equal(option.series, undefined, type)
    assert.ok(scene.some(command => command.kind === 'polygon' && command.fill === 'rgba(255, 0, 0, 1)'), `${type} foreground`)
    assert.ok(scene.some(command => command.kind === 'polygon' && command.fill === 'rgba(255, 255, 0, 1)'), `${type} background`)
    assert.equal(option.officeChartPointer?.mode, ['scatter', 'heatmap', 'boxplot', 'area'].includes(type) ? 'graphic' : undefined)
  }
})

test('invalid patterns reject candidate updates without changing the chart or its source', () => {
  const { workbook, chart } = workbookModel('column'), before = structuredClone(chart), cells = structuredClone(workbook.sheets[0].cells)
  for (const invalid of [{ ...pattern(), preset: 'missing' }, { preset: 'horz', foregroundColor: 'red' }, { ...pattern(), backgroundColor: 'not-a-color' },
    { ...pattern(), foregroundColor: null }, [], 'horz']) {
    assert.throws(() => updateChart(workbook, chart, { options: { series: { 1: { fillPattern: invalid } } } }), TypeError)
    assert.deepEqual(chart, before)
    assert.deepEqual(workbook.sheets[0].cells, cells)
  }
  assert.equal(cloneChartOptions({ allSeriesStyle: { fillPattern: null } }).allSeriesStyle.fillPattern, null)
})

test('pattern controls distinguish inheritance from solid fills and series color selection clears inherited patterns', () => {
  const config = model('column', { allSeriesStyle: { fillPattern: pattern('vert') } }), controls = chartSeriesControls(1)
  const control = suffix => controls.find(entry => entry.key === `series.1.${suffix}`)
  const preset = control('fillPattern.preset'), fg = control('fillPattern.foregroundColor'), bg = control('fillPattern.backgroundColor')
  assert.equal(preset.read(config), 'inherit')
  assert.equal(fg.visible(config), false)
  preset.write(config, 'none'); assert.equal(config.options.series[1].fillPattern, null)
  preset.write(config, 'inherit'); assert.equal(config.options.series[1].fillPattern, undefined)
  preset.write(config, 'upDiag'); assert.deepEqual(config.options.series[1].fillPattern, pattern('upDiag'))
  assert.equal(fg.visible(config), true)
  fg.write(config, '#008000'); bg.write(config, '#ffffff')
  assert.equal(chartSeriesStyle(config, 1, 0).fillPattern.foregroundColor, '#008000')
  control('color').write(config, '#0000ff')
  assert.equal(config.options.series[1].fillPattern, null)
  assert.equal(chartSeriesStyle(config, 1, 0).fillPattern, null)
  assert.equal(preset.visible(model('wordCloud')), false)
})
