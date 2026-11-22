import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

let server
let createBlankWorkbook
let createChart
let chartEChartsOption

beforeAll(async () => {
  server = await createServer({
    configFile: resolve(process.cwd(), 'vite.config.ts'),
    root: process.cwd(),
    server: { middlewareMode: true, hmr: false },
    appType: 'custom',
    logLevel: 'error',
  })
  const model = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
  const engine = await server.ssrLoadModule('/packages/office-studio/src/features/charts/chart-engine.ts')
  const adapter = await server.ssrLoadModule('/packages/office-studio/src/features/charts/echarts-option.ts')
  createBlankWorkbook = model.createBlankWorkbook
  createChart = engine.createChart
  chartEChartsOption = adapter.chartEChartsOption
})

afterAll(async () => {
  await server?.close()
})

function workbookWithColumnData() {
  const workbook = createBlankWorkbook()
  const sheet = workbook.sheets[0]
  sheet.cells = {
    '0:0': { value: 'Month' }, '0:1': { value: 'North' }, '0:2': { value: 'South' }, '0:3': { value: 'X2' },
    '1:0': { value: 'Jan' }, '1:1': { value: 10 }, '1:2': { value: 15 }, '1:3': { value: 5 },
    '2:0': { value: 'Feb' }, '2:1': { value: null }, '2:2': { value: 18 }, '2:3': { value: 9 },
    '3:0': { value: 'Mar' }, '3:1': { value: 30 }, '3:2': { value: 24 }, '3:3': { value: 12 },
  }
  return { workbook, sheet }
}

function chart(workbook, sheet, type, patch = {}) {
  return createChart(workbook, sheet.id, {
    type,
    source: { range: { startRow: 0, startCol: 0, endRow: 3, endCol: 3 }, orientation: 'columns' },
    position: { x: 0, y: 0 },
    size: { width: 420, height: 280 },
    valueFields: type === 'pie' || type === 'doughnut' ? [1] : [1, 2],
    ...patch,
  })
}

test('maps each stored category chart type and preserves axis, legend and theme settings', () => {
  const { workbook, sheet } = workbookWithColumnData()
  const cases = [
    ['bar', 'bar'], ['stacked-column', 'bar'], ['line', 'line'], ['area', 'line'],
  ]

  for (const [type, seriesType] of cases) {
    const model = chart(workbook, sheet, type, {
      title: 'Sales', legend: { visible: true, position: 'right' },
      xAxis: { title: 'Period', scaleType: 'linear' },
      yAxis: { title: 'Units', scaleType: type === 'stacked-column' ? 'linear' : 'log', logBase: 2 },
    })
    const option = chartEChartsOption(workbook, model, {
      textColor: '#202020', axisColor: '#505050', splitLineColor: '#eeeeee', palette: ['#123456'],
    })

    assert.equal(option.title.text, 'Sales')
    assert.equal(option.animation, false)
    assert.deepEqual(option.color, ['#123456'])
    assert.equal(option.legend.show, true)
    assert.equal(option.legend.orient, 'vertical')
    assert.deepEqual(option.legend.data, ['North', 'South'])
    assert.equal(option.xAxis.type, 'category')
    assert.deepEqual(option.xAxis.data, ['Jan', 'Feb', 'Mar'])
    assert.equal(option.yAxis.type, type === 'stacked-column' ? 'value' : 'log')
    if (type !== 'stacked-column') assert.equal(option.yAxis.logBase, 2)
    assert.equal(option.series[0].type, seriesType)
    assert.equal(option.series[0].data[1], null)
    if (type === 'stacked-column') assert.equal(option.series[0].stack, 'total')
    if (type === 'area') assert.equal(option.series[0].areaStyle.opacity, 0.28)
  }
})

test('maps pie and doughnut categories to separate named slices', () => {
  const { workbook, sheet } = workbookWithColumnData()
  for (const [type, radius] of [['pie', '74%'], ['doughnut', ['52%', '74%']]]) {
    const option = chartEChartsOption(workbook, chart(workbook, sheet, type))
    assert.equal(option.series[0].type, 'pie')
    assert.deepEqual(option.series[0].radius, radius)
    assert.deepEqual(option.legend.data, ['Jan', 'Feb', 'Mar'])
    assert.deepEqual(option.series[0].data.map(item => item.name), ['Jan', 'Feb', 'Mar'])
    assert.deepEqual(option.series[0].data.map(item => item.value), [10, 0, 30])
    assert.equal(option.series[0].stillShowZeroSum, false)
    assert.equal(option.series[0].showEmptyCircle, false)
  }
})

test('scatter uses each series xValues, skips invalid pairs and omits nonpositive log-axis values', () => {
  const { workbook, sheet } = workbookWithColumnData()
  sheet.cells['1:0'].value = 1
  sheet.cells['2:0'].value = -2
  sheet.cells['3:0'].value = 3
  sheet.cells['1:3'].value = 5
  sheet.cells['2:3'].value = 8
  sheet.cells['3:3'].value = 12
  const model = chart(workbook, sheet, 'scatter', {
    categoryField: 0,
    valueFields: [1, 2],
    seriesToCategory: { 2: [3] },
    xAxis: { title: 'Independent variable', scaleType: 'log', logBase: 10 },
    yAxis: { title: 'Response', scaleType: 'linear' },
    legend: { visible: false, position: 'bottom' },
  })
  const option = chartEChartsOption(workbook, model)

  assert.equal(option.xAxis.type, 'log')
  assert.equal(option.xAxis.logBase, 10)
  assert.equal(option.yAxis.type, 'value')
  assert.equal(option.legend.show, false)
  assert.deepEqual(option.series.map(series => series.data), [[[1, 10], [3, 30]], [[5, 15], [8, 18], [12, 24]]])
})

test('missing independent scatter X values remain missing instead of using the shared X column', () => {
  const { workbook, sheet } = workbookWithColumnData()
  sheet.cells['1:0'].value = 1
  sheet.cells['2:0'].value = 2
  sheet.cells['3:0'].value = 3
  sheet.cells['2:3'].value = null
  const model = chart(workbook, sheet, 'scatter', { seriesToCategory: { 2: [3] } })
  const before = JSON.stringify(workbook)
  const option = chartEChartsOption(workbook, model)
  assert.deepEqual(option.series[1].data, [[5, 15], [12, 24]])
  assert.equal(JSON.stringify(workbook), before)
})

test('radar uses shared category axes and preserves signed and missing values by series', () => {
  const { workbook, sheet } = workbookWithColumnData()
  sheet.cells['1:1'].value = -10
  const model = chart(workbook, sheet, 'radar', {
    title: 'Comparison', xAxis: { title: 'Month' }, yAxis: { title: 'Units' },
  })
  const option = chartEChartsOption(workbook, model, { axisColor: '#595959', textColor: '#262626' })
  assert.equal(option.xAxis, undefined)
  assert.equal(option.yAxis, undefined)
  assert.deepEqual(option.title.map(title => title.text), ['Comparison', 'Month', 'Units'])
  assert.deepEqual(option.legend.data, ['North', 'South'])
  assert.deepEqual(option.radar.indicator, ['Jan', 'Feb', 'Mar'].map(name => ({ name, min: -10, max: 30 })))
  assert.equal(option.radar.axisName.color, '#262626')
  assert.equal(option.radar.axisLine.lineStyle.color, '#595959')
  assert.deepEqual(option.series.map(series => [series.type, series.name, series.data[0].value]), [
    ['radar', 'North', [-10, null, 30]], ['radar', 'South', [15, 18, 24]],
  ])
  assert.throws(() => chart(workbook, sheet, 'radar', { yAxis: { scaleType: 'log' } }), /linear value axis/)
})

test('top and side legends reserve space while following explicit theme colors', () => {
  const { workbook, sheet } = workbookWithColumnData()
  const theme = { backgroundColor: '#202020', textColor: '#eeeeee', palette: ['#ccddff'] }
  const top = chartEChartsOption(workbook, chart(workbook, sheet, 'line', { title: 'Title', legend: { position: 'top' } }), theme)
  assert.ok(top.legend.top > top.title.top)
  assert.ok(top.grid.top > top.legend.top)
  assert.equal(top.title.textStyle.color, '#eeeeee')
  assert.equal(top.legend.textStyle.color, '#eeeeee')
  assert.equal(top.backgroundColor, '#202020')
  const right = chartEChartsOption(workbook, chart(workbook, sheet, 'line', { legend: { position: 'right' } }), theme)
  const left = chartEChartsOption(workbook, chart(workbook, sheet, 'line', { legend: { position: 'left' } }), theme)
  assert.ok(right.grid.right > top.grid.right)
  assert.ok(left.grid.left > top.grid.left)
  right.color[0] = '#000000'
  assert.deepEqual(theme.palette, ['#ccddff'])
  const pie = chartEChartsOption(workbook, chart(workbook, sheet, 'pie', { title: 'Title', legend: { position: 'top' } }), theme)
  assert.equal(pie.series[0].label.color, '#eeeeee')
  assert.ok(pie.series[0].top > pie.legend.top)
})
