import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { resolve } from 'node:path'
import { chromium } from '@playwright/test'
import { createServer } from 'vite'

// P12 actual Canvas drafts only. No runner, server, browser or pixel check has run.
const fixturePath = '/__chart-axis-title-renderer-test__'
const fixture = `<!doctype html><html lang="en"><head><meta charset="utf-8"></head><body><script type="module">
import { getInstanceByDom } from 'echarts/core'
import { createBlankWorkbook, createCell } from '/packages/office-studio/src/domain/model.ts'
import { createChart } from '/packages/office-studio/src/features/charts/chart-engine.ts'
import { chartEChartsOption } from '/packages/office-studio/src/features/charts/echarts-option.ts'
import { EChartsCanvasRenderer } from '/packages/office-studio/src/features/charts/echarts-renderer.ts'
const renderer = new EChartsCanvasRenderer(), size = { width: 520, height: 360, devicePixelRatio: 1 }
function model(type = 'line') {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  const table = [['Region', 'Units', 'Revenue'], ['North', 1, 2], ['South', 3, 4], ['West', 5, 6]]
  sheet.cells = Object.fromEntries(table.flatMap((row, r) => row.map((value, c) => [r + ':' + c, createCell(r, c, value)])))
  const chart = createChart(workbook, sheet.id, { type, categoryField: 0, valueFields: [1, 2],
    source: { range: { startRow: 0, startCol: 0, endRow: 3, endCol: 2 }, orientation: 'columns' },
    position: { x: 0, y: 0 }, size: { width: size.width, height: size.height }, title: 'Axis title renderer draft',
    xAxis: { title: 'Physical horizontal title', titleAlignment: 'start', titleStyle: { family: 'Georgia', fontSize: 12 } },
    yAxis: { title: 'Physical vertical title', titleAlignment: 'start', titleStyle: { family: 'Georgia', fontSize: 12 } },
    legend: { visible: type === 'heatmap', position: 'bottom' } })
  return { workbook, chart }
}
function titleElements(instance) {
  return instance.getZr().storage.getDisplayList().filter(element => element.type === 'tspan'
    && (element.style.text === 'Physical horizontal title' || element.style.text === 'Physical vertical title'))
}
function bounds(element) {
  const rect = element.getBoundingRect().clone(), transform = element.getComputedTransform()
  if (transform) rect.applyTransform(transform)
  return { x: rect.x, y: rect.y, width: rect.width, height: rect.height }
}
window.axisTitleRendererFixture = { renderer, size, model, bounds, titleElements, chartEChartsOption, getInstanceByDom }
</script></body></html>`
let server, browser, page, baseURL
const errors = []
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(),
    resolve: { alias: { echarts: resolve(process.cwd(), 'packages/office-studio/node_modules/echarts') } },
    plugins: [{ name: 'chart-axis-title-renderer-fixture', configureServer(vite) {
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
  await page.goto(baseURL); await page.waitForFunction(() => Boolean(window.axisTitleRendererFixture))
})
afterEach(async () => { await page?.close(); assert.deepEqual(errors, []) })
afterAll(async () => { await browser?.close(); await server?.close() })

test('actual native-content titles are single Canvas glyph runs against physical edges and do not move when axes reverse', async () => {
  const result = await page.evaluate(() => {
    const f = window.axisTitleRendererFixture, { workbook, chart } = f.model(), author = JSON.stringify(chart)
    const option = f.chartEChartsOption(workbook, chart), frame = f.renderer.render(chart.id, option, f.size), instance = f.getInstanceByDom(frame.source)
    const titles = f.titleElements(instance), horizontal = titles.find(element => element.style.text === 'Physical horizontal title'), vertical = titles.find(element => element.style.text === 'Physical vertical title')
    // GraphicModel flattens group children into elements with parentId.
    const graphics = instance.getOption().graphic[0].elements, textGraphics = graphics.filter(element => element.type === 'text')
    const horizontalOption = textGraphics.find(element => element.style?.text === 'Physical horizontal title'), verticalOption = textGraphics.find(element => element.style?.text === 'Physical vertical title')
    const plot = { x: option.grid.left, y: option.grid.top, width: f.size.width - option.grid.left - option.grid.right, height: f.size.height - option.grid.top - option.grid.bottom }
    const before = titles.map(element => ({ text: element.style.text, font: element.style.font, bounds: f.bounds(element) }))
    const horizontalBounds = f.bounds(horizontal), verticalBounds = f.bounds(vertical)
    const horizontalInside = horizontalBounds.x >= plot.x - 1 && horizontalBounds.x + horizontalBounds.width <= plot.x + plot.width + 1
    const verticalInside = verticalBounds.y >= plot.y - 1 && verticalBounds.y + verticalBounds.height <= plot.y + plot.height + 1
    const authorBeforeReverse = JSON.stringify(chart) === author
    chart.xAxis.reverse = true; chart.yAxis.reverse = true; chart.revision += 1
    const reverse = f.renderer.render(chart.id, f.chartEChartsOption(workbook, chart), f.size)
    const after = f.titleElements(f.getInstanceByDom(reverse.source)).map(element => ({ text: element.style.text, font: element.style.font, bounds: f.bounds(element) }))
    const emptyNativeNames = [instance.getModel().getComponent('xAxis').option.name, instance.getModel().getComponent('yAxis').option.name].every(name => !name)
    f.renderer.retain(new Set())
    return { before, after, count: titles.length, xStart: horizontalOption.x === plot.x, yStart: verticalOption.y === plot.y,
      horizontalInside, verticalInside,
      emptyNativeNames, authorBeforeReverse, newer: reverse.sourceVersion > frame.sourceVersion }
  })
  assert.equal(result.count, 2)
  assert.deepEqual(result.after, result.before)
  assert.equal(result.before.every(title => /12px.*Georgia/.test(title.font)), true)
  for (const flag of ['xStart', 'yStart', 'horizontalInside', 'verticalInside', 'emptyNativeNames', 'authorBeforeReverse', 'newer']) assert.equal(result[flag], true, flag)
})

test('native and graphic hover retain title glyphs, native content and Heatmap legend without changing author history', async () => {
  const result = await page.evaluate(() => {
    const f = window.axisTitleRendererFixture, results = []
    for (const type of ['line', 'heatmap']) {
      const { workbook, chart } = f.model(type), author = JSON.stringify(chart), option = f.chartEChartsOption(workbook, chart), frame = f.renderer.render(chart.id, option, f.size)
      const instance = f.getInstanceByDom(frame.source), titles = f.titleElements(instance), data = instance.getModel().getSeriesByIndex(0).getData(), content = data.getItemGraphicEl(0)
      const point = type === 'heatmap' ? { x: option.officeChartPointer.geometry.heatmap.legend.bar.x + option.officeChartPointer.geometry.heatmap.legend.bar.width / 2,
        y: option.officeChartPointer.geometry.heatmap.legend.bar.y + 6 } : { x: option.grid.left + 60, y: option.grid.top + 60 }
      const hover = f.renderer.updatePointer(chart.id, point), after = f.titleElements(instance), leave = f.renderer.updatePointer(chart.id)
      results.push({ type, count: after.length, retained: titles.every((element, index) => element === after[index]),
        content: Boolean(content) && content === instance.getModel().getSeriesByIndex(0).getData().getItemGraphicEl(0),
        newer: hover.sourceVersion > frame.sourceVersion && leave.sourceVersion > hover.sourceVersion, author: JSON.stringify(chart) === author,
        legend: type !== 'heatmap' || instance.getOption().graphic[0].elements.some(element => String(element.id).startsWith(chart.id + ':heatmap-legend')) })
      f.renderer.retain(new Set())
    }
    return results
  })
  for (const resultItem of result) {
    assert.equal(resultItem.count, 2, resultItem.type)
    for (const flag of ['retained', 'content', 'newer', 'author', 'legend']) assert.equal(resultItem[flag], true, `${resultItem.type}: ${flag}`)
  }
})
