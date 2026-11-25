import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// P18 rendering drafts only. Real chart projection/drawing is used when run later.
let server, model, engine, native, paths, geometry
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[model, engine, native, paths, geometry] = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/domain/model.ts'),
    ...['chart-engine', 'echarts-option', 'chart-line-path', 'chart-geometry']
      .map(name => server.ssrLoadModule(`/packages/office-studio/src/features/charts/${name}.ts`)),
  ])
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
const rows = [['Label', 'Value'], ['A', 9], ['B', 2], ['C', 8]]
const near = (a, b) => assert.ok(Math.abs(a - b) < 1e-7, `${a} ≈ ${b}`)
const lines = commands => commands.filter(item => item.kind === 'line' && item.width === 7)
const datum = value => value && typeof value === 'object' && !Array.isArray(value) ? value.value : value
const nativeSeries = value => { const series = native.chartEChartsOption(value.workbook, value.chart).series; return Array.isArray(series) ? series : [series] }
function make(type, input = rows, authored = {}, fields = [1]) {
  const workbook = model.createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.cells = Object.fromEntries(input.flatMap((row, r) => row.map((value, c) => [`${r}:${c}`, model.createCell(r, c, value)])))
  const chart = engine.createChart(workbook, sheet.id, { type, position: { x: 0, y: 0 }, size: { width: 480, height: 320 },
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: input.length - 1, endCol: input[0].length - 1 } },
    categoryField: 0, valueFields: fields, legend: { visible: false, position: 'bottom' },
    options: { autoGradientFill: false, allSeriesStyle: { color: '#123456', border: { width: 7 }, point: { shape: 'none' } }, ...authored },
  }, 'shape-rendering')
  return { workbook, sheet, chart }
}
const commands = value => engine.chartRenderCommands(value.workbook, value.chart)
const layout = value => engine.chartLayout(value.workbook, value.chart)

test('native Line and Area map the three authored shapes without changing numeric data or axes', () => {
  for (const type of ['line', 'area']) {
    const value = make(type), before = json(value.sheet.cells), data = engine.readChartData(value.workbook, value.chart)
    for (const shape of ['line', 'smooth', 'step']) {
      value.chart.options.area = { lineStyle: shape }
      const option = native.chartEChartsOption(value.workbook, value.chart), series = option.series[0]
      assert.equal(series.type, 'line'); assert.equal(series.smooth, shape === 'smooth')
      assert.equal(series.step, shape === 'step' ? 'middle' : false)
      assert.equal(series.smoothMonotone, shape === 'smooth' ? 'x' : undefined)
      assert.deepEqual(series.data.map(datum), [9, 2, 8]); assert.equal(series.connectNulls, false)
      assert.equal(!!series.areaStyle, type === 'area')
      assert.equal(option.yAxis.type, 'value')
      assert.deepEqual(engine.readChartData(value.workbook, value.chart), data)
    }
    assert.deepEqual(value.sheet.cells, before)
  }
})

test('shared Canvas commands draw genuinely different smooth/step geometry and keep original point anchors', () => {
  const value = make('line'), original = layout(value).series[0].points, data = engine.readChartData(value.workbook, value.chart)
  const results = new Map()
  for (const [shape, count] of [['line', 2], ['smooth', 24], ['step', 6]]) {
    value.chart.options.area = { lineStyle: shape }
    const paint = lines(commands(value)); assert.equal(paint.length, count); results.set(shape, paint)
    assert.deepEqual(layout(value).series[0].points, original)
    near(paint[0].x1, original[0].x); near(paint[0].y1, original[0].y)
    near(paint.at(-1).x2, original.at(-1).x); near(paint.at(-1).y2, original.at(-1).y)
    assert.deepEqual(engine.readChartData(value.workbook, value.chart), data)
  }
  for (const line of results.get('step')) assert.ok(line.x1 === line.x2 || line.y1 === line.y2)
  assert.ok(results.get('smooth').some(line => line.y1 !== line.y2 && line.x1 !== line.x2))
  assert.notDeepEqual(results.get('smooth'), results.get('line'))
})

test('stacked and percent-stacked area boundaries share identical curve/step samples at every layer', () => {
  const input = [['Label', 'Lower', 'Upper'], ['A', 4, 2], ['B', 8, 3], ['C', 6, 5]]
  for (const type of ['areaStacked', 'areaPercentStacked']) {
    const value = make(type, input, {}, [1, 2]), cells = json(value.sheet.cells), original = layout(value)
    for (const shape of ['line', 'smooth', 'step']) {
      value.chart.options.area = { lineStyle: shape }
      const current = layout(value), first = current.series[0].points, second = current.series[1].points
      assert.deepEqual(current.series, original.series)
      for (let segment = 1; segment < first.length; segment++) {
        const lower = paths.chartLineSegment(first, segment, shape), upper = paths.chartLineSegment(second, segment, shape)
        assert.equal(lower.length, upper.length)
        for (let index = 0; index < lower.length; index++) {
          near(lower[index].x, upper[index].x); near(lower[index].y, upper[index].baseY)
        }
      }
      const series = nativeSeries(value)
      assert.ok(series.every(item => item.stack && item.areaStyle && item.step === (shape === 'step' ? 'middle' : false)))
      if (type === 'areaPercentStacked') for (let index = 0; index < 3; index++) near(datum(series[0].data[index]) + datum(series[1].data[index]), 100)
      assert.ok(commands(value).some(item => item.kind === 'polygon'))
    }
    assert.deepEqual(value.sheet.cells, cells)
  }
})

test('Combination changes only line/area shapes while column data and secondary axis identities remain intact', () => {
  const input = [['Label', 'Bars', 'Line', 'Area'], ['A', 4, 9, 3], ['B', 7, 2, 8], ['C', 2, 8, 5]]
  const value = make('combination', input, { area: { smooth: true }, series: {
    1: { type: 'column' }, 2: { type: 'line', rightAxis: true }, 3: { type: 'area' },
  }, rightYAxis: { scaleType: 'linear' } }, [1, 2, 3])
  const legacy = nativeSeries(value), bars = json(legacy[0]), original = layout(value).series
  assert.equal(legacy[1].smooth, false); assert.equal(legacy[2].smooth, true)
  for (const shape of ['line', 'smooth', 'step']) {
    value.chart.options.area = { lineStyle: shape }
    const series = nativeSeries(value)
    assert.deepEqual(json(series[0]), bars); assert.equal(series[0].type, 'bar')
    for (const item of series.slice(1)) { assert.equal(item.smooth, shape === 'smooth'); assert.equal(item.step, shape === 'step' ? 'middle' : false) }
    assert.equal(series[1].yAxisIndex, 1); assert.equal(series[2].yAxisIndex, 0)
    assert.deepEqual(layout(value).series, original)
  }
})

test('Break, Link and Zero keep their actual data/connectivity semantics for all three shapes', () => {
  const input = [['Label', 'Value'], ['A', 9], ['B', null], ['C', 2], ['D', 8]]
  for (const shape of ['line', 'smooth', 'step']) for (const strategy of ['break', 'link', 'zero']) {
    const value = make('line', input, { area: { lineStyle: shape }, invalidValueStrategy: strategy }), before = json(value.sheet.cells)
    const series = nativeSeries(value)[0], factor = shape === 'smooth' ? 12 : shape === 'step' ? 3 : 1
    assert.equal(series.connectNulls, strategy === 'link')
    assert.deepEqual(series.data.map(datum), strategy === 'zero' ? [9, 0, 2, 8] : [9, null, 2, 8])
    assert.equal(lines(commands(value)).length, factor * (strategy === 'break' ? 1 : strategy === 'link' ? 2 : 3))
    assert.deepEqual(value.sheet.cells, before)
  }
})

test('date geometry remains finite for reversed irregular dates and repeated dates, with missing X excluded', () => {
  const input = [['Date', 'Value'], [45005, 7], [45001, null], [45003, 2], [45003, 8], [null, 12]]
  const value = make('line', input, { mapping: { useDateAxis: true }, invalidValueStrategy: 'link' })
  value.chart.xAxis.reverse = true
  const initial = layout(value), cells = json(value.sheet.cells)
  assert.ok(initial.dateAxis); assert.equal(initial.series[0].points.length, 3)
  const repeated = initial.series[0].points.filter(item => item.xValue === 45003)
  assert.equal(repeated.length, 2); near(repeated[0].x, repeated[1].x)
  for (const shape of ['smooth', 'step']) {
    value.chart.options.area = { lineStyle: shape }
    assert.deepEqual(layout(value).series[0].points, initial.series[0].points)
    assert.ok(lines(commands(value)).length > 0)
    for (const line of lines(commands(value))) for (const key of ['x1', 'y1', 'x2', 'y2']) assert.ok(Number.isFinite(line[key]))
    assert.equal(nativeSeries(value)[0].connectNulls, true)
  }
  assert.deepEqual(value.sheet.cells, cells)
})

test('projected 3D shapes change real faces/edges while preserving stack, camera anchors and single opacity application', () => {
  const input = [['Label', 'Lower', 'Upper'], ['A', 4, 2], ['B', 8, 3], ['C', 6, 5]]
  for (const type of ['line3D', 'area3D', 'areaStacked3D', 'areaPercentStacked3D']) {
    const value = make(type, input, { threeD: { view: { rotX: 35, rotY: 43, perspective: 25 } } }, [1, 2])
    const initial = layout(value).series, before = json(value.sheet.cells), scene = new Map()
    for (const [shape, edgeCount] of [['line', 8], ['smooth', 96], ['step', 24]]) {
      value.chart.options.area = { lineStyle: shape }
      const paint = commands(value); scene.set(shape, paint)
      assert.equal(lines(paint).length, edgeCount); assert.deepEqual(layout(value).series, initial)
      for (const line of lines(paint)) for (const key of ['x1', 'y1', 'x2', 'y2']) assert.ok(Number.isFinite(line[key]))
      const option = native.chartEChartsOption(value.workbook, value.chart)
      assert.ok(option.graphic.elements.length > 0)
    }
    assert.notDeepEqual(scene.get('smooth'), scene.get('line')); assert.notDeepEqual(scene.get('step'), scene.get('line'))
    assert.deepEqual(value.sheet.cells, before)
  }
  const value = make('line3D')
  value.chart.options.allSeriesStyle.border.opacity = 0.4
  assert.ok(lines(commands(value)).every(line => line.color === geometry.chartColorOpacity('#123456', 0.4)))
  const visiblePolygons = commands(value).filter(item => item.kind === 'polygon').length
  value.chart.options.allSeriesStyle.border.visible = false
  assert.equal(lines(commands(value)).length, 0)
  assert.ok(commands(value).filter(item => item.kind === 'polygon').length < visiblePolygons)
})

test('Scatter never gains line geometry: Zero fills missing Y but does not invent coordinates for missing X', () => {
  const input = [['X', 'Y'], [1, 6], [2, null], [null, 8], [4, 3]]
  for (const strategy of ['break', 'link', 'zero']) {
    const value = make('scatter', input, { area: { lineStyle: 'step' }, invalidValueStrategy: strategy,
      allSeriesStyle: { point: { shape: 'circle', size: 8 } } }), before = json(value.sheet.cells)
    const series = nativeSeries(value)[0], points = series.data.map(datum), computed = layout(value).series[0].points
    assert.equal(series.type, 'scatter'); assert.equal(series.smooth, undefined); assert.equal(series.step, undefined)
    assert.equal(points.length, strategy === 'zero' ? 3 : 2); assert.equal(computed.length, points.length)
    assert.equal(points.some(value => value[0] === null), false)
    assert.equal(points.some(value => value[0] === 2 && value[1] === 0), strategy === 'zero')
    assert.equal(lines(commands(value)).length, 0); assert.deepEqual(value.sheet.cells, before)
  }
})
