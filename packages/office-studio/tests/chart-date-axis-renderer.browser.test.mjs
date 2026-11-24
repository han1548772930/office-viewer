import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { resolve } from 'node:path'
import { chromium } from '@playwright/test'
import { createServer } from 'vite'

// P22 actual Canvas drafts only. No runner, server, browser, pixels or benchmark has run.
const fixturePath = '/__chart-date-axis-renderer-test__'
const fixture = `<!doctype html><html lang="en"><head><meta charset="utf-8"></head><body><script type="module">
import { getInstanceByDom } from 'echarts/core'
import { createBlankWorkbook, createCell } from '/packages/office-studio/src/domain/model.ts'
import { createChart, readChartData } from '/packages/office-studio/src/features/charts/chart-engine.ts'
import { deriveChartData } from '/packages/office-studio/src/features/charts/chart-derived.ts'
import { layoutDerivedChartData } from '/packages/office-studio/src/features/charts/chart-layout.ts'
import { chartEChartsOption } from '/packages/office-studio/src/features/charts/echarts-option.ts'
import { EChartsCanvasRenderer } from '/packages/office-studio/src/features/charts/echarts-renderer.ts'
const renderer = new EChartsCanvasRenderer(), size = { width: 560, height: 400, devicePixelRatio: 1 }
function model(type = 'line', date1904 = false) {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  workbook.date1904 = date1904
  const dates = date1904 ? [90, 30, 59, 0] : ['2024-03-31', '2024-01-31', '2024-02-29', '2024-01-01']
  const boxDates = date1904 ? ['1904-03-31', '1904-01-31', '1904-02-29'] : ['2024-03-31', '2024-01-31', '2024-02-29']
  const table = type === 'boxplot' ? [['Sample', ...boxDates], ['One', 10, 3, 6], ['Two', 20, 4, 7], ['Three', 30, 5, 8]]
    : [['Date', 'Units', 'Size'], ...dates.map((date, index) => [date, [30, 10, 20, 11][index], [3, 1, 2, 4][index]])]
  sheet.cells = Object.fromEntries(table.flatMap((row, r) => row.map((value, c) => [r + ':' + c, createCell(r, c, value)])))
  const chart = createChart(workbook, sheet.id, { type, categoryField: 0, valueFields: type === 'boxplot' ? [1, 2, 3] : type === 'bubble' ? [1, 2] : [1],
    source: { range: { startRow: 0, startCol: 0, endRow: table.length - 1, endCol: table[0].length - 1 }, orientation: 'columns' },
    position: { x: 0, y: 0 }, size: { width: size.width, height: size.height }, legend: { visible: false, position: 'bottom' },
    xAxis: { baseTimeUnit: 'months', majorTimeUnit: 'months', majorUnit: 1, label: { family: 'Georgia', fontSize: 14 } },
    yAxis: { min: 0, max: 100 }, options: { mapping: { useDateAxis: true, ...(type === 'bubble' ? { bubble: { xIndex: 0, yIndex: 1, sizeIndex: 2 } } : {}) } } })
  return { workbook, chart }
}
function layout(workbook, chart) { return layoutDerivedChartData(deriveChartData(readChartData(workbook, chart), chart), chart) }
function dateGlyphs(instance, prefix = '2024') {
  return instance.getZr().storage.getDisplayList().filter(element => element.type === 'tspan' && new RegExp('^' + prefix + '-\\\\d{2}-01$').test(element.style.text))
}
window.dateAxisRendererFixture = { renderer, size, model, layout, dateGlyphs, chartEChartsOption, getInstanceByDom }
</script></body></html>`
let server, browser, page, baseURL
const errors = []
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(),
    resolve: { alias: { echarts: resolve(process.cwd(), 'packages/office-studio/node_modules/echarts') } },
    plugins: [{ name: 'chart-date-axis-renderer-fixture', configureServer(vite) {
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
  await page.goto(baseURL); await page.waitForFunction(() => Boolean(window.dateAxisRendererFixture))
})
afterEach(async () => { await page?.close(); assert.deepEqual(errors, []) })
afterAll(async () => { await browser?.close(); await server?.close() })

test('actual native monthly line keeps every sorted source row and places shared first-day glyphs once on its own data extent', async () => {
  const result = await page.evaluate(() => {
    const f = window.dateAxisRendererFixture, { workbook, chart } = f.model(), author = JSON.stringify(chart)
    const expected = f.layout(workbook, chart), option = f.chartEChartsOption(workbook, chart), frame = f.renderer.render(chart.id, option, f.size), instance = f.getInstanceByDom(frame.source)
    const data = instance.getModel().getSeriesByIndex(0).getData(), extent = instance.getModel().getComponent('xAxis').axis.scale.getExtent()
    const positions = Array.from({ length: data.count() }, (_, index) => ({ id: data.getId(index), point: data.getItemLayout(index) }))
    const labels = instance.getOption().graphic[0].elements.filter(element => element.type === 'text' && /^2024-\d{2}-01$/.test(element.style?.text))
    const glyphs = f.dateGlyphs(instance), nativeDateLabelsHidden = instance.getModel().getComponent('xAxis').option.axisLabel.show === false
    const commonPoints = expected.series[0].points.map(point => ({ id: chart.id + ':series:0:point:' + point.categoryIndex, point: [point.x, point.y] }))
    const authorUnchanged = JSON.stringify(chart) === author, expectedExtent = [expected.xMin, expected.xMax], pointerMode = option.officeChartPointer.mode
    f.renderer.retain(new Set())
    return { positions, commonPoints, extent, expectedExtent, labels: labels.map(element => ({ text: element.style.text, x: element.x, y: element.y })),
      glyphs: glyphs.map(element => ({ text: element.style.text, font: element.style.font })), nativeDateLabelsHidden, authorUnchanged, pointerMode }
  })
  assert.deepEqual(result.extent, result.expectedExtent)
  assert.equal(result.positions.length, 4)
  for (let index = 0; index < result.positions.length; index++) {
    assert.equal(result.positions[index].id, result.commonPoints[index].id)
    for (const coordinate of [0, 1]) assert.ok(Math.abs(result.positions[index].point[coordinate] - result.commonPoints[index].point[coordinate]) < 1e-5)
  }
  assert.deepEqual(result.labels.map(label => label.text), ['2024-01-01', '2024-02-01', '2024-03-01'])
  assert.deepEqual(result.glyphs.map(glyph => glyph.text), ['2024-01-01', '2024-02-01', '2024-03-01'])
  assert.equal(result.glyphs.every(glyph => /14px.*Georgia/.test(glyph.font)), true)
  assert.equal(result.nativeDateLabelsHidden, true); assert.equal(result.authorUnchanged, true); assert.equal(result.pointerMode, 'graphic')
})

test('graphic date hover selects the first original row within a shared month and retains native content and axis glyphs through leave', async () => {
  const result = await page.evaluate(() => {
    const f = window.dateAxisRendererFixture, { workbook, chart } = f.model(), author = JSON.stringify(chart), option = f.chartEChartsOption(workbook, chart)
    const frame = f.renderer.render(chart.id, option, f.size), instance = f.getInstanceByDom(frame.source), content = instance.getModel().getSeriesByIndex(0).getData().getItemGraphicEl(0)
    const glyphs = f.dateGlyphs(instance), layout = option.officeChartPointer.geometry.layout, x = layout.plot.x + layout.categoryPositions[1] * layout.plot.width
    const hover = f.renderer.updatePointer(chart.id, { x: x + 1, y: layout.plot.y + layout.plot.height / 2 })
    const graphics = instance.getOption().graphic[0].elements, pointerTexts = graphics.filter(element => element.type === 'text' && String(element.id).startsWith(chart.id + ':axis-pointer')).map(element => element.style.text)
    const hoveredGlyphs = f.dateGlyphs(instance), contentRetained = content === instance.getModel().getSeriesByIndex(0).getData().getItemGraphicEl(0)
    const leave = f.renderer.updatePointer(chart.id), after = f.dateGlyphs(instance), pointerRemoved = !instance.getOption().graphic[0].elements.some(element => String(element.id).startsWith(chart.id + ':axis-pointer'))
    const retained = glyphs.every((element, index) => element === hoveredGlyphs[index] && element === after[index]), authorUnchanged = JSON.stringify(chart) === author
    f.renderer.retain(new Set())
    return { pointerTexts, count: after.length, retained, contentRetained, pointerRemoved, authorUnchanged, newer: hover.sourceVersion > frame.sourceVersion && leave.sourceVersion > hover.sourceVersion }
  })
  assert.equal(result.pointerTexts[0], '2024-01-31')
  assert.equal(result.count, 3)
  for (const flag of ['retained', 'contentRetained', 'pointerRemoved', 'authorUnchanged', 'newer']) assert.equal(result[flag], true, flag)
})

test('calendar major-unit edits change glyph density without point positions and changing the basic unit restores elapsed-day distances', async () => {
  const result = await page.evaluate(() => {
    const f = window.dateAxisRendererFixture, { workbook, chart } = f.model(), option = f.chartEChartsOption(workbook, chart)
    const initial = f.renderer.render(chart.id, option, f.size), instance = f.getInstanceByDom(initial.source)
    const pointPositions = () => { const data = instance.getModel().getSeriesByIndex(0).getData(); return Array.from({ length: data.count() }, (_, index) => data.getItemLayout(index)[0]) }
    const before = pointPositions()
    chart.xAxis.majorUnit = 2; chart.revision += 1
    const interval = f.renderer.render(chart.id, f.chartEChartsOption(workbook, chart), f.size), afterInterval = pointPositions(), intervalTexts = f.dateGlyphs(instance).map(element => element.style.text)
    delete chart.xAxis.baseTimeUnit; chart.xAxis.majorUnit = 1; chart.revision += 1
    const elapsed = f.renderer.render(chart.id, f.chartEChartsOption(workbook, chart), f.size), afterElapsed = pointPositions(), actualExtent = instance.getModel().getComponent('xAxis').axis.scale.getExtent()
    const expected = f.layout(workbook, chart), authorMajorUnit = chart.xAxis.majorUnit, authorBase = Object.hasOwn(chart.xAxis, 'baseTimeUnit')
    f.renderer.retain(new Set())
    return { before, afterInterval, afterElapsed, intervalTexts, actualExtent, expectedExtent: [expected.xMin, expected.xMax], authorMajorUnit, authorBase,
      newer: interval.sourceVersion > initial.sourceVersion && elapsed.sourceVersion > interval.sourceVersion }
  })
  assert.deepEqual(result.afterInterval, result.before)
  assert.deepEqual(result.intervalTexts, ['2024-01-01', '2024-03-01'])
  assert.notEqual(result.afterElapsed[0], result.afterElapsed[1])
  assert.ok(Math.abs((result.afterElapsed[2] - result.afterElapsed[1]) / (result.afterElapsed[3] - result.afterElapsed[2]) - 29 / 31) < 1e-8)
  assert.deepEqual(result.actualExtent, result.expectedExtent)
  assert.equal(result.authorMajorUnit, 1); assert.equal(result.authorBase, false); assert.equal(result.newer, true)
})

test('1904 month labels survive native Bubble and graphic Bar, 3D and Boxplot renderers without shifting the epoch', async () => {
  const result = await page.evaluate(() => {
    const f = window.dateAxisRendererFixture, results = []
    for (const type of ['bubble', 'bar-horizontal', 'column3D', 'boxplot']) {
      const { workbook, chart } = f.model(type, true), author = JSON.stringify(chart), option = f.chartEChartsOption(workbook, chart), frame = f.renderer.render(chart.id, option, f.size), instance = f.getInstanceByDom(frame.source)
      const glyphs = f.dateGlyphs(instance, '1904'), expected = f.layout(workbook, chart)
      results.push({ type, labels: glyphs.map(element => element.style.text), sourceMin: expected.dateSourceMin, date1904: expected.date1904,
        native: instance.getModel().getSeries().length > 0, author: JSON.stringify(chart) === author })
      f.renderer.retain(new Set())
    }
    return results
  })
  for (const item of result) {
    assert.deepEqual(item.labels, ['1904-01-01', '1904-02-01', '1904-03-01'], item.type)
    assert.equal(item.date1904, true, item.type); assert.equal(item.author, true, item.type)
    assert.equal(item.native, item.type === 'bubble', item.type)
    assert.equal(item.sourceMin, item.type === 'boxplot' ? 30 : 0, item.type)
  }
})
