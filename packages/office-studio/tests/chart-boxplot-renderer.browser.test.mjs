import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { resolve } from 'node:path'
import { chromium } from '@playwright/test'
import { createServer } from 'vite'

// P23 actual Canvas drafts only. No runner, server, browser, pixels or benchmark has run.
// Interpolated quartiles/IQR/source-marker overrides are explicit local Office extensions.
const fixturePath = '/__chart-boxplot-renderer-test__'
const fixture = `<!doctype html><html lang="en"><head><meta charset="utf-8"></head><body><script type="module">
import { getInstanceByDom } from 'echarts/core'
import { createBlankWorkbook, createCell } from '/packages/office-studio/src/domain/model.ts'
import { createChart, readChartData } from '/packages/office-studio/src/features/charts/chart-engine.ts'
import { deriveChartData } from '/packages/office-studio/src/features/charts/chart-derived.ts'
import { chartBoxplotGeometry, chartBoxplotStyle } from '/packages/office-studio/src/features/charts/chart-boxplot.ts'
import { layoutDerivedChartData } from '/packages/office-studio/src/features/charts/chart-layout.ts'
import { chartPalette } from '/packages/office-studio/src/features/charts/chart-options.ts'
import { chartEChartsOption } from '/packages/office-studio/src/features/charts/echarts-option.ts'
import { EChartsCanvasRenderer } from '/packages/office-studio/src/features/charts/echarts-renderer.ts'
const renderer = new EChartsCanvasRenderer(), size = { width: 560, height: 400, devicePixelRatio: 1 }
const samples = [100, null, -100, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 100, -100]
const points = {
  0: { color: 'red', size: 10, shape: 'rect', fillOpacity: 0.25, border: { color: 'navy', width: 2, opacity: 0.8 } },
  2: { color: 'orange', size: 8, shape: 'diamond' },
  3: { color: 'gold', size: 18, shape: 'circle' },
  13: { color: 'blue', size: 14, shape: 'square', fillOpacity: 0.75 },
  14: { color: 'purple', size: 12, shape: 'triangle' },
}
function model(boxplot = {}, dataPoints = points, values = samples) {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  const table = [['Sample', 'Measurements'], ...values.map((value, index) => ['Sample ' + (index + 1), value])]
  sheet.cells = Object.fromEntries(table.flatMap((row, r) => row.map((value, c) => [r + ':' + c, createCell(r, c, value)])))
  const chart = createChart(workbook, sheet.id, { type: 'boxplot', categoryField: 0, valueFields: [1],
    source: { range: { startRow: 0, startCol: 0, endRow: table.length - 1, endCol: 1 }, orientation: 'columns' },
    position: { x: 0, y: 0 }, size: { width: size.width, height: size.height }, title: 'Boxplot renderer draft',
    legend: { visible: true, position: 'bottom' }, yAxis: { min: -120, max: 240 },
    options: { boxplot: { quartileMethod: 'inclusive', whiskerMethod: 'iqr', ...boxplot },
      series: { 1: { color: 'green', point: { shape: 'circle', size: 4 }, dataPoints } } } })
  return { workbook, chart }
}
function derived(workbook, chart) { return deriveChartData(readChartData(workbook, chart), chart) }
function geometry(workbook, chart) {
  const result = derived(workbook, chart), palette = chartPalette(chart), layout = layoutDerivedChartData(result, chart, palette)
  return { result, layout, box: chartBoxplotGeometry(chart, layout, chartBoxplotStyle(chart, result.boxes[0], 0, palette), 0, 1) }
}
window.boxplotRendererFixture = { renderer, size, model, points, derived, geometry, chartEChartsOption, getInstanceByDom }
</script></body></html>`
let server, browser, page, baseURL
const errors = []
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(),
    resolve: { alias: { echarts: resolve(process.cwd(), 'packages/office-studio/node_modules/echarts') } },
    plugins: [{ name: 'chart-boxplot-renderer-fixture', configureServer(vite) {
      vite.middlewares.use(fixturePath, async (_request, response, next) => {
        try { response.setHeader('Content-Type', 'text/html; charset=utf-8'); response.end(await vite.transformIndexHtml(fixturePath, fixture)) }
        catch (error) { next(error) }
      })
    } }], server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen()
  const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate an HTTP port')
  baseURL = `http://127.0.0.1:${address.port}${fixturePath}`
  browser = await chromium.launch({ headless: true })
})
beforeEach(async () => {
  errors.length = 0; page = await browser.newPage()
  page.on('pageerror', error => errors.push(error.message))
  await page.goto(baseURL); await page.waitForFunction(() => Boolean(window.boxplotRendererFixture))
})
afterEach(async () => { await page?.close(); assert.deepEqual(errors, []) })
afterAll(async () => { await browser?.close(); await server?.close() })

test('actual native Canvas retains duplicate source IDs, independent marker paths and summary paint at shared data coordinates', async () => {
  const result = await page.evaluate(() => {
    const f = window.boxplotRendererFixture, { workbook, chart } = f.model(), author = JSON.stringify(chart), source = JSON.stringify(workbook.sheets[0].cells)
    const expected = f.geometry(workbook, chart), option = f.chartEChartsOption(workbook, chart), frame = f.renderer.render(chart.id, option, f.size)
    const instance = f.getInstanceByDom(frame.source), data = instance.getModel().getSeriesByIndex(1).getData(), summary = instance.getModel().getSeriesByIndex(0).getData()
    const paths = [], items = Array.from({ length: data.count() }, (_value, index) => {
      const group = data.getItemGraphicEl(index), path = group.getSymbolPath(), datum = option.series[1].data[index]
      paths.push(path)
      const sourceIndex = Number(datum.id.slice(datum.id.lastIndexOf(':') + 1)), value = datum.value[1], position = data.getItemLayout(index)
      return { id: data.getId(index), sourceIndex, value, name: data.getName(index), symbol: data.getItemVisual(index, 'symbol'), size: data.getItemVisual(index, 'symbolSize'),
        position, expected: [expected.box.x, expected.box.y(value)], fill: path.style.fill, stroke: path.style.stroke, lineWidth: path.style.lineWidth,
        opacity: path.style.opacity, scale: [path.scaleX, path.scaleY] }
    })
    const boxPath = summary.getItemGraphicEl(0), separatePaths = new Set(paths).size === paths.length
    const cached = f.renderer.render(chart.id, f.chartEChartsOption(workbook, chart), f.size)
    const answer = { chartId: chart.id, items, separatePaths, summary: { fill: boxPath.style.fill, stroke: boxPath.style.stroke, opacity: boxPath.style.opacity,
      five: option.series[0].data[0].value }, identities: expected.result.boxes[0].statistics.outlierSourceIndices,
      cached: cached.source === frame.source && cached.sourceVersion === frame.sourceVersion,
      author: JSON.stringify(chart) === author, source: JSON.stringify(workbook.sheets[0].cells) === source }
    f.renderer.retain(new Set())
    return answer
  })
  assert.deepEqual(result.identities, [2, 14, 0, 13])
  assert.deepEqual(result.items.map(item => item.id), [2, 14, 0, 13].map(index => `${result.chartId}:box:0:outlier:${index}`))
  assert.deepEqual(result.items.map(item => [item.symbol, item.size]), [['diamond', 8], ['triangle', 12], ['rect', 10], ['rect', 14]])
  assert.deepEqual(result.items.map(item => item.fill), ['rgba(255, 165, 0, 1)', 'rgba(128, 0, 128, 1)', 'rgba(255, 0, 0, 0.25)', 'rgba(0, 0, 255, 0.75)'])
  for (const item of result.items) {
    for (const coordinate of [0, 1]) assert.ok(Math.abs(item.position[coordinate] - item.expected[coordinate]) < 1e-5)
    assert.deepEqual(item.scale, [item.size / 2, item.size / 2]); assert.equal(item.opacity, 1); assert.equal(item.name, 'Measurements')
  }
  assert.deepEqual(result.items[2].position, result.items[3].position)
  assert.equal(result.items[2].stroke, 'rgba(0, 0, 128, 0.8)'); assert.equal(result.items[2].lineWidth, 2)
  assert.deepEqual(result.summary, { fill: 'rgba(0, 128, 0, 0.45)', stroke: 'rgba(0, 128, 0, 1)', opacity: 1, five: [0, 1.25, 4.5, 7.75, 9] })
  for (const flag of ['separatePaths', 'cached', 'author', 'source']) assert.equal(result[flag], true, flag)
})

test('actual graphic outlier paint and independent border remain retained through guide-only hover, leave and cache reuse', async () => {
  const result = await page.evaluate(() => {
    const f = window.boxplotRendererFixture, points = { ...f.points, 13: { ...f.points[13], shape: 'emptySquare' } }, { workbook, chart } = f.model({}, points)
    const author = JSON.stringify(chart), expected = f.geometry(workbook, chart), option = f.chartEChartsOption(workbook, chart), frame = f.renderer.render(chart.id, option, f.size), instance = f.getInstanceByDom(frame.source)
    const display = () => instance.getZr().storage.getDisplayList(), red = display().find(element => element.type === 'rect' && element.style.fill === 'rgba(255, 0, 0, 0.25)' && element.shape.width === 10)
    const title = display().find(element => element.type === 'tspan' && element.style.text === chart.title), legend = display().find(element => element.type === 'tspan' && element.style.text === '样本分布')
    const borders = display().filter(element => element.type === 'line' && element.style.stroke === 'rgba(0, 0, 128, 0.8)' && element.style.lineWidth === 2)
    const calls = [], original = instance.setOption.bind(instance)
    instance.setOption = (patch, flags) => { calls.push(patch); return original(patch, flags) }
    const point = { x: expected.box.x, y: expected.layout.plot.y + expected.layout.plot.height / 2 }, hover = f.renderer.updatePointer(chart.id, point)
    const repeated = f.renderer.updatePointer(chart.id, point), cached = f.renderer.render(chart.id, f.chartEChartsOption(workbook, chart), f.size), leave = f.renderer.updatePointer(chart.id)
    const answer = { graphic: option.series === undefined, center: [red.shape.x + 5, red.shape.y + 5], expected: [expected.box.x, expected.box.y(100)], borderCount: borders.length,
      retained: [red, title, legend, ...borders].every(element => display().includes(element)), onlyGuides: calls.every(call => Object.keys(call).join() === 'graphic'),
      replace: calls[0]?.graphic.elements[0].$action === 'replace', remove: calls[1]?.graphic.elements[0].$action === 'remove',
      repeatEmpty: repeated === undefined, cached: cached.source === frame.source && cached.sourceVersion === hover.sourceVersion,
      versions: hover.sourceVersion > frame.sourceVersion && leave.sourceVersion > hover.sourceVersion, author: JSON.stringify(chart) === author }
    f.renderer.retain(new Set())
    return answer
  })
  for (const coordinate of [0, 1]) assert.ok(Math.abs(result.center[coordinate] - result.expected[coordinate]) < 1e-5)
  assert.equal(result.borderCount, 4)
  for (const flag of ['graphic', 'retained', 'onlyGuides', 'replace', 'remove', 'repeatEmpty', 'cached', 'versions', 'author']) assert.equal(result[flag], true, flag)
})

test('actual fixed-default boxes keep full-sample whiskers, omit outliers and preserve singleton slots without activating unsupported sample markers', async () => {
  const result = await page.evaluate(() => {
    const f = window.boxplotRendererFixture, points = { ...f.points, 0: { ...f.points[0], shape: 'emptySquare' } }, answers = []
    for (const values of [undefined, [100]]) {
      const { workbook, chart } = f.model({ quartileMethod: undefined, whiskerMethod: undefined }, points, values), author = JSON.stringify(chart)
      const expected = f.derived(workbook, chart), option = f.chartEChartsOption(workbook, chart), frame = f.renderer.render(chart.id, option, f.size), instance = f.getInstanceByDom(frame.source)
      const boxes = instance.getModel().getSeriesByIndex(0).getData(), outliers = instance.getModel().getSeriesByIndex(1).getData()
      answers.push({ native: option.series !== undefined, statistics: expected.boxes[0].statistics ?? null, five: option.series[0].data[0].value,
        slotCount: boxes.count(), boxVisible: Boolean(boxes.getItemGraphicEl(0)), outlierCount: outliers.count(), author: JSON.stringify(chart) === author })
      f.renderer.retain(new Set())
    }
    return answers
  })
  assert.deepEqual(result[0].statistics, { min: -100, q1: 1, median: 4.5, q3: 8, max: 100, outliers: [], outlierSourceIndices: [] })
  assert.deepEqual(result[0].five, [-100, 1, 4.5, 8, 100]); assert.equal(result[0].boxVisible, true)
  assert.equal(result[1].statistics, null); assert.deepEqual(result[1].five, []); assert.equal(result[1].boxVisible, false)
  for (const item of result) { assert.equal(item.native, true); assert.equal(item.slotCount, 1); assert.equal(item.outlierCount, 0); assert.equal(item.author, true) }
})

test('actual source refresh activates an inlier preset at its original row, retains surviving duplicate IDs and hides outliers without deleting author styles', async () => {
  const result = await page.evaluate(() => {
    const f = window.boxplotRendererFixture, { workbook, chart } = f.model(), authoredPoints = JSON.stringify(chart.options.series[1].dataPoints)
    const first = f.renderer.render(chart.id, f.chartEChartsOption(workbook, chart), f.size)
    const sheet = workbook.sheets[0]
    sheet.cells['1:1'].value = 1; sheet.cells['1:1'].display = '1'
    sheet.cells['4:1'].value = 200; sheet.cells['4:1'].display = '200'
    const option = f.chartEChartsOption(workbook, chart), refresh = f.renderer.render(chart.id, option, f.size), instance = f.getInstanceByDom(refresh.source)
    const data = instance.getModel().getSeriesByIndex(1).getData(), items = Array.from({ length: data.count() }, (_value, index) => {
      const path = data.getItemGraphicEl(index).getSymbolPath()
      return { id: data.getId(index), fill: path.style.fill, size: data.getItemVisual(index, 'symbolSize') }
    })
    const statistics = f.derived(workbook, chart).boxes[0].statistics, cached = f.renderer.render(chart.id, f.chartEChartsOption(workbook, chart), f.size)
    chart.options.boxplot.showOutliers = false; chart.revision += 1
    const hidden = f.renderer.render(chart.id, f.chartEChartsOption(workbook, chart), f.size), hiddenCount = instance.getModel().getSeriesByIndex(1).getData().count()
    const retainedStatistics = JSON.stringify(f.derived(workbook, chart).boxes[0].statistics) === JSON.stringify(statistics)
    chart.options.boxplot.showOutliers = true; chart.revision += 1
    const shown = f.renderer.render(chart.id, f.chartEChartsOption(workbook, chart), f.size), restored = instance.getModel().getSeriesByIndex(1).getData()
    const answer = { chartId: chart.id, items, sourceIndices: statistics.outlierSourceIndices, hiddenCount, retainedStatistics,
      restoredIds: Array.from({ length: restored.count() }, (_value, index) => restored.getId(index)),
      points: JSON.stringify(chart.options.series[1].dataPoints) === authoredPoints,
      cached: cached.source === refresh.source && cached.sourceVersion === refresh.sourceVersion,
      versions: refresh.sourceVersion > first.sourceVersion && hidden.sourceVersion > refresh.sourceVersion && shown.sourceVersion > hidden.sourceVersion }
    f.renderer.retain(new Set())
    return answer
  })
  const ids = [2, 14, 13, 3].map(index => `${result.chartId}:box:0:outlier:${index}`)
  assert.deepEqual(result.sourceIndices, [2, 14, 13, 3]); assert.deepEqual(result.items.map(item => item.id), ids); assert.deepEqual(result.restoredIds, ids)
  assert.equal(result.items[2].fill, 'rgba(0, 0, 255, 0.75)'); assert.equal(result.items[2].size, 14)
  assert.equal(result.items[3].fill, 'rgba(255, 215, 0, 1)'); assert.equal(result.items[3].size, 18)
  assert.equal(result.hiddenCount, 0)
  for (const flag of ['retainedStatistics', 'points', 'cached', 'versions']) assert.equal(result[flag], true, flag)
})
