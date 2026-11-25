import assert from 'node:assert/strict'
import { test } from 'vitest'
import * as echarts from 'echarts'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { chartEChartsOption } from '../src/features/charts/echarts-option.ts'
import { ChartNativeElementIndex } from '../src/features/charts/echarts-element-interaction.ts'
import { chartCanvasCommandElement, chartCanvasElementKey, chartCanvasElementRoute, chartCanvasBoxplotBindings,
  chartCanvasBoxplotRuntimeId, chartCanvasPieSliceMark, chartCanvasSeriesBindings, chartCanvasSeriesMark } from '../src/features/charts/chart-element-interaction.ts'
import { drawChartData } from '../src/features/charts/chart-draw.ts'
import { deriveChartData } from '../src/features/charts/chart-derived.ts'
import { layoutDerivedChartData } from '../src/features/charts/chart-layout.ts'
import { chartBoxplotGeometry, chartBoxplotStyle } from '../src/features/charts/chart-boxplot.ts'
import { chartPieLayout } from '../src/features/charts/chart-pie-layout.ts'
import { chartPieMaximumExplosion } from '../src/features/charts/chart-pie.ts'
import { cloneChartDraftConfig } from '../src/features/charts/chart-settings-draft.ts'
import { chartLinearGradient } from '../src/features/charts/chart-color.ts'

// P21 drafts only. Real SVG SSR/ZRender geometry is specified here; nothing was run.
// Every option has a real source-sheet seed even when prepared data is supplied.
const size = { width: 600, height: 420 }, palette = ['#112233', '#223344', '#334455', '#445566']
const source = () => ({ categories: ['A', 'B', 'C', 'D'], series: [
  { field: 1, name: 'Display Alpha', sourceName: 'Original', values: [8, 5, 3, 2], xValues: [1, 2, 3, 4] },
  { field: 2, name: 'Display Beta', sourceName: 'Comparison', values: [2, 3, 5, 7], xValues: [1, 2, 3, 4] },
] })
const model = (type, options = {}) => ({ id: 'mark-style-geometry', sheetId: 'sheet', revision: 0, type,
  categoryField: 0, valueFields: [1, 2], source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 4, endCol: 2 } },
  position: { x: 0, y: 0 }, size, palette, legend: { visible: false, position: 'bottom' },
  xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, options })
const json = value => JSON.parse(JSON.stringify(value))
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-7, `${actual} ≈ ${expected}`)
const markCommands = (commands, kind) => commands.filter(command => chartCanvasCommandElement(command)?.directEdit?.kind === kind)
function option(chart, prepared, computed = { preparedData: prepared }) {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.id = chart.sheetId; workbook.activeSheetId = sheet.id; sheet.cells = {}
  sheet.cells['0:0'] = createCell(0, 0, 'Category')
  for (const series of prepared.series) sheet.cells[`0:${series.field}`] = createCell(0, series.field, series.sourceName ?? series.name)
  for (const [index, category] of prepared.categories.entries()) {
    sheet.cells[`${index + 1}:0`] = createCell(index + 1, 0, category)
    for (const series of prepared.series) sheet.cells[`${index + 1}:${series.field}`] = createCell(index + 1, series.field, series.values[index] ?? null)
  }
  workbook.charts = [chart]
  return chartEChartsOption(workbook, chart, {}, computed)
}
function render(input) {
  const runtime = input.officeChartElements ?? { series: [], axes: [] }, value = { ...input }
  for (const key of Object.keys(value)) if (key.startsWith('officeChart')) delete value[key]
  const instance = echarts.init(null, null, { renderer: 'svg', ssr: true, ...size })
  instance.setOption(value, { notMerge: true, lazyUpdate: false }); instance.renderToSVGString()
  return { option: value, runtime, instance, index: new ChartNativeElementIndex(instance, value, runtime, size.width, size.height), dispose: () => instance.dispose() }
}
function measuredHit(index, element) {
  const target = index.target(element); assert.ok(target, `Missing ${chartCanvasElementKey(element)}`)
  const { bounds } = target, accept = point => { const hit = index.hit(point); return hit?.key === target.key ? { point, hit } : undefined }
  const center = accept({ x: bounds.x + bounds.width / 2, y: bounds.y + bounds.height / 2 }); if (center) return center
  for (let y = bounds.y; y <= bounds.y + bounds.height; y += 0.5) for (let x = bounds.x; x <= bounds.x + bounds.width; x += 0.5) {
    const result = accept({ x, y }); if (result) return result
  }
  assert.fail(`Target is shielded or clipped: ${target.key}`)
}
const samples = [100, null, -100, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 100, -100]
const statistics = { min: 0, q1: 1.25, median: 4.5, q3: 7.75, max: 9,
  outliers: [-100, -100, 100, 100], outlierSourceIndices: [2, 14, 0, 13] }
function boxplot(options = {}) {
  const prepared = { categories: samples.map((_value, index) => `Sample ${index + 1}`), series: [{ field: 9, name: 'Display samples', sourceName: 'Samples', values: [...samples] }] },
    chart = model('boxplot', { boxplot: { quartileMethod: 'inclusive', whiskerMethod: 'iqr' }, ...options })
  chart.valueFields = [9]; chart.source.range.endCol = 9; chart.source.range.endRow = samples.length
  chart.yAxis = { scaleType: 'linear', min: -120, max: 120 }
  return { prepared, chart, computed: { preparedData: prepared, boxplots: [{ field: 9, seriesPosition: 0, statistics }] } }
}

test('native Cartesian points keep exact source indices while measured aliases and line paths expose dedicated seriesStyle owners', () => {
  const prepared = source()
  for (const type of ['line', 'column', 'bar', 'scatter', 'combination']) {
    const chart = model(type, { series: { 1: { label: { visible: true, position: 'top' } } } }), value = render(option(chart, prepared)),
      binding = chartCanvasSeriesBindings(prepared, chart.options)[0], point = { kind: 'dataPoint', binding, dataIndex: 1 }
    try {
      const measured = measuredHit(value.index, point)
      assert.equal(measured.hit.element.binding.field, 1); assert.equal(measured.hit.element.binding.sourceName, 'Original')
      assert.equal(value.index.hit(measured.point, true).key, 'series:field:1')
      assert.ok(value.index.target(chartCanvasSeriesMark(binding)))
      assert.ok(value.runtime.series.some(entry => entry.seriesMark?.binding.seriesId === 'field:1'))
      assert.ok(value.index.list().some(target => target.element.kind === 'dataLabel' && target.element.binding.seriesId === 'field:1'))
    } finally { value.dispose() }
  }
})

test('shared Cartesian line and area paths carry the same dedicated mark as native source aliases while labels remain source-bound', () => {
  const prepared = source(), color = chartLinearGradient({ start: { x: 0, y: 0 }, end: { x: 1, y: 1 },
    stops: [{ offset: 0, color: 'red' }, { offset: 1, color: 'blue' }] })
  for (const type of ['line', 'area']) {
    const chart = model(type, { series: { 1: { color, point: { size: 8 }, label: { visible: true, position: 'top' } } } }),
      commands = drawChartData(prepared, chart), marks = markCommands(commands, 'seriesStyle'), value = render(option(chart, prepared)),
      binding = chartCanvasSeriesBindings(prepared, chart.options)[0]
    assert.ok(marks.some(command => command.kind === (type === 'area' ? 'polygon' : 'line')))
    assert.ok(marks.every(command => command.chartCanvasElement.directEdit.binding.seriesId === 'field:1' || command.chartCanvasElement.directEdit.binding.seriesId === 'field:2'))
    assert.ok(commands.some(command => command.kind === 'text' && chartCanvasCommandElement(command)?.kind === 'dataLabel'))
    try {
      assert.equal(value.option.series, undefined); assert.ok(value.index.target(chartCanvasSeriesMark(binding)))
      measuredHit(value.index, { kind: 'dataPoint', binding, dataIndex: 1 })
    } finally { value.dispose() }
  }
})

test('native and shared Radar polygons bind seriesStyle to the original field or exact-name author instead of the native series position', () => {
  const prepared = source()
  for (const shared of [false, true]) {
    const chart = model('radar', { series: { 1: { color: '#112233' } },
      namedSeries: [{ name: 'Original', style: { color: '#abcdef', ...(shared ? { point: { shape: 'emptyPin', size: 9 } } : {}) } }] }),
      binding = chartCanvasSeriesBindings(prepared, chart.options)[0], value = render(option(chart, prepared))
    try {
      const hit = measuredHit(value.index, chartCanvasSeriesMark(binding)).hit
      assert.equal(hit.element.directEdit.binding.seriesId, 'name:"Original"'); assert.equal(hit.element.directEdit.binding.field, 1)
      assert.ok(value.index.target({ kind: 'series', binding }))
      if (shared) assert.ok(markCommands(drawChartData(prepared, chart), 'seriesStyle').some(command => command.kind === 'polygon'))
      else assert.equal(value.runtime.series[0].seriesOnly, true)
    } finally { value.dispose() }
  }
})

test('native Boxplot summary and duplicate outliers have stable runtime IDs and separately measured original sample ownership', () => {
  const { prepared, chart, computed } = boxplot({ series: { 9: { color: '#112233', dataPoints: {
    2: { size: 18, shape: 'rect' }, 14: { size: 8, shape: 'diamond' }, 0: { size: 20, shape: 'rect' }, 13: { size: 6, shape: 'circle' },
  } } } }), derived = deriveChartData(prepared, chart, computed), before = json(derived.boxes),
    binding = chartCanvasBoxplotBindings(derived.boxes, chart.options)[0], value = render(option(chart, prepared, computed))
  try {
    assert.deepEqual(value.option.series.map(series => series.id), [chartCanvasBoxplotRuntimeId(chart.id, 'summary'), chartCanvasBoxplotRuntimeId(chart.id, 'outliers')])
    assert.equal(value.runtime.series[0].items[0].directEdit.kind, 'seriesStyle')
    assert.deepEqual(value.runtime.series[1].items.map(item => item.dataIndex), [2, 14, 0, 13])
    measuredHit(value.index, chartCanvasSeriesMark(binding))
    for (const dataIndex of [2, 14, 0, 13]) {
      const hit = measuredHit(value.index, { kind: 'dataPoint', binding, dataIndex }).hit
      assert.equal(hit.element.binding.field, 9); assert.equal(hit.element.dataIndex, dataIndex)
    }
    assert.equal(value.index.target({ kind: 'dataPoint', binding, dataIndex: 1 }), undefined)
    assert.deepEqual(json(derived.boxes), before)
  } finally { value.dispose() }
})

test('shared Boxplot summary geometry and median labels use seriesStyle while each real outlier retains its field/name sample style', () => {
  const { prepared, chart, computed } = boxplot({ namedSeries: [{ name: 'Samples', style: { color: '#112233',
    fillPattern: { preset: 'ltUpDiag', foregroundColor: 'black', backgroundColor: 'white' }, label: { visible: true, position: 'inside' },
    dataPoints: { 13: { color: 'red', size: 14, shape: 'rect', fillPattern: null } },
  } }] }), derived = deriveChartData(prepared, chart, computed), binding = chartCanvasBoxplotBindings(derived.boxes, chart.options)[0],
    layout = layoutDerivedChartData(derived, chart, palette), geometry = chartBoxplotGeometry(chart, layout, chartBoxplotStyle(chart, derived.boxes[0], 0, palette), 0, 1),
    commands = drawChartData(prepared, chart, {}, computed), summary = markCommands(commands, 'seriesStyle'), value = render(option(chart, prepared, computed))
  assert.ok(summary.some(command => command.kind === 'text')); assert.ok(summary.some(command => command.kind === 'line'))
  assert.ok(summary.every(command => command.chartCanvasElement.directEdit.binding.seriesId === 'name:"Samples"'))
  const point = commands.find(command => command.kind === 'rect' && chartCanvasCommandElement(command)?.dataIndex === 13)
  assert.ok(point); assert.equal(point.width, 14); near(point.x + point.width / 2, geometry.x); near(point.y + point.height / 2, geometry.y(100))
  try {
    assert.equal(value.option.series, undefined); measuredHit(value.index, { kind: 'dataPoint', binding, dataIndex: 13 })
    assert.ok(value.index.target(chartCanvasSeriesMark(binding)))
  } finally { value.dispose() }
})

test('native Pie labels and their guide lines expose the global pieStyle author while unstyled source sectors retain ordinary point targets', () => {
  const prepared = source(), chart = model('pie', { pie: { label: { visible: true, position: 'outside', fontSize: 12 } } }),
    value = render(option(chart, prepared)), binding = chartCanvasSeriesBindings(prepared, chart.options)[0]
  try {
    measuredHit(value.index, { kind: 'dataPoint', binding, dataIndex: 0 })
    const hit = measuredHit(value.index, { kind: 'mark', directEdit: { kind: 'pieStyle' } }).hit
    assert.deepEqual(chartCanvasElementRoute(hit.element), { tab: 'style', section: 'pieStyle', focus: { kind: 'pieLabel' } })
    assert.equal(value.index.list().some(target => target.element.kind === 'dataLabel'), false)
    assert.equal(value.runtime.series[0].pieLabels, true)
    const commands = markCommands(drawChartData(prepared, chart), 'pieStyle')
    assert.ok(commands.some(command => command.kind === 'text')); assert.ok(commands.some(command => command.kind === 'line'))
  } finally { value.dispose() }
})

test('ordinary Pie independent slice authors match native and shared sector hits and reject a forged native source-index override', () => {
  const prepared = source(), chart = model('pie', { pie: { label: { visible: false }, sliceStyles: { 2: { color: '#abcdef', fillOpacity: 0.4, border: { width: 2 } } } } }),
    binding = chartCanvasSeriesBindings(prepared, chart.options)[0], target = chartCanvasPieSliceMark(prepared.series[0], 2), native = render(option(chart, prepared))
  try {
    measuredHit(native.index, target); measuredHit(native.index, { kind: 'dataPoint', binding, dataIndex: 0 })
    assert.equal(native.index.target({ kind: 'dataPoint', binding, dataIndex: 2 }), undefined)
    assert.equal(native.runtime.series[0].pieSliceMarks[2].sourcePointIndex, 2)
    assert.match(native.option.series[0].data[2].itemStyle.color, /0\.4/); assert.equal(native.option.series[0].data[2].itemStyle.borderWidth, 2)
    const runtime = { ...native.runtime, series: native.runtime.series.map(entry => ({ ...entry,
      ...(entry.pieSliceMarks ? { pieSliceMarks: { 2: { ...entry.pieSliceMarks[2], sourcePointIndex: 3, sourceName: 'Other source' } } } : {}),
    })) }, guarded = new ChartNativeElementIndex(native.instance, native.option, runtime, size.width, size.height)
    assert.equal(guarded.target(target), undefined); assert.ok(guarded.target({ kind: 'dataPoint', binding, dataIndex: 2 }))
  } finally { native.dispose() }
  const shared = cloneChartDraftConfig(chart); shared.options.pie.sliceStyles[2].explosion = 25
  const value = render(option(shared, prepared))
  try {
    assert.equal(value.option.series, undefined); measuredHit(value.index, target)
    assert.equal(value.index.target({ kind: 'dataPoint', binding, dataIndex: 2 }), undefined)
    assert.ok(markCommands(drawChartData(prepared, shared), 'pieSliceStyle').every(command => command.chartCanvasElement.directEdit.sourcePointIndex === 2))
  } finally { value.dispose() }
})

test('both composite Pie plots retain original slice indices, global labels and a generated aggregate with separate author metadata', () => {
  const prepared = source(), before = json(prepared)
  for (const type of ['pieOfPie', 'barOfPie']) {
    const chart = model(type, { pie: { label: { visible: true, position: 'outside' }, aggregateLabel: 'Others', aggregateStyle: { color: '#445566' },
      sliceStyles: { 2: { color: '#abcdef', fillOpacity: 0.5 } }, secondaryPlot: { split: { type: 'position', count: 2 } } } }),
      commands = drawChartData(prepared, chart), value = render(option(chart, prepared))
    try {
      const primary = measuredHit(value.index, chartCanvasPieSliceMark(prepared.series[0], 0, 'primary')).hit,
        secondary = measuredHit(value.index, chartCanvasPieSliceMark(prepared.series[0], 2, 'secondary')).hit
      assert.equal(primary.element.directEdit.plot, 'primary'); assert.equal(secondary.element.directEdit.plot, 'secondary')
      assert.equal(secondary.element.directEdit.sourceField, 1); assert.equal(secondary.element.directEdit.sourcePointIndex, 2)
      measuredHit(value.index, { kind: 'mark', directEdit: { kind: 'pieAggregateStyle', sourceField: 1, sourceName: 'Original', aggregateIndex: 4 } })
      measuredHit(value.index, { kind: 'mark', directEdit: { kind: 'pieStyle' } })
      assert.equal(value.index.target(chartCanvasPieSliceMark(prepared.series[0], 4)), undefined)
      assert.ok(commands.every(command => chartCanvasCommandElement(command)?.kind !== 'dataPoint'))
      assert.ok(markCommands(commands, 'pieAggregateStyle').some(command => command.kind === 'text' && command.text === 'Others'))
      assert.ok(markCommands(commands, 'pieSliceStyle').some(command => command.chartCanvasElement.directEdit.sourcePointIndex === 2
        && command.kind === (type === 'barOfPie' ? 'rect' : 'polygon')))
    } finally { value.dispose() }
  }
  assert.deepEqual(json(prepared), before)
})

test('independent explosion moves real sector centers and inactive secondary bar explosion cannot shrink or move either plot', () => {
  const prepared = source(), plot = { x: 10, y: 10, width: 560, height: 350 }, chart = model('pie', { pie: { sliceStyles: { 0: { explosion: 20 } } } }),
    layout = chartPieLayout(prepared, chart, plot), exploded = layout.sectors.find(sector => sector.part.index === 0), still = layout.sectors.find(sector => sector.part.index === 3)
  near(Math.hypot(exploded.center.x - still.center.x, exploded.center.y - still.center.y), exploded.outer * 0.2)
  const bar = model('barOfPie', { pie: { sliceStyles: { 0: { explosion: 20 }, 2: { explosion: 90 } }, secondaryPlot: { split: { type: 'position', count: 2 } } } }),
    cleared = cloneChartDraftConfig(bar)
  delete cleared.options.pie.sliceStyles[2].explosion
  assert.equal(chartPieMaximumExplosion(prepared, bar), 20)
  assert.deepEqual(chartPieLayout(prepared, bar, plot), chartPieLayout(prepared, cleared, plot))
  assert.equal(bar.options.pie.sliceStyles[2].explosion, 90)
  const composite = model('pieOfPie', { pie: { sliceStyles: { 2: { explosion: 30 } }, secondaryPlot: { split: { type: 'position', count: 2 } } } }),
    compound = chartPieLayout(prepared, composite, plot), moved = compound.sectors.find(sector => sector.secondary && sector.part.index === 2),
    sibling = compound.sectors.find(sector => sector.secondary && sector.part.index === 3)
  near(Math.hypot(moved.center.x - sibling.center.x, moved.center.y - sibling.center.y), moved.outer * 0.3)
})

test('native Boxplot filtering preserves raw sample identities after nulls, duplicate values, hidden points and invalid logarithmic summaries', () => {
  const { prepared, chart, computed } = boxplot({ series: { 9: { dataPoints: { 0: { size: 0 }, 13: { size: 12 } } } } })
  chart.yAxis = { scaleType: 'log', logBase: 10, min: 1, max: 200 }
  const value = render(option(chart, prepared, computed)), binding = value.runtime.series[0].items[0].directEdit.binding
  try {
    assert.deepEqual(value.runtime.series[1].items.map(item => item.dataIndex), [13])
    assert.deepEqual(value.option.series[0].data[0].value, [])
    assert.ok(value.index.target(chartCanvasSeriesMark(binding)))
    measuredHit(value.index, { kind: 'dataPoint', binding, dataIndex: 13 })
    for (const dataIndex of [0, 1, 2, 14]) assert.equal(value.index.target({ kind: 'dataPoint', binding, dataIndex }), undefined)
    assert.deepEqual(statistics.outlierSourceIndices, [2, 14, 0, 13]); assert.equal(statistics.min, 0)
  } finally { value.dispose() }
  chart.options.boxplot.showOutliers = false
  const hidden = render(option(chart, prepared, computed))
  try { assert.deepEqual(hidden.runtime.series[1].items, []); assert.equal(hidden.index.list().some(target => target.element.kind === 'dataPoint'), false) }
  finally { hidden.dispose() }
})
