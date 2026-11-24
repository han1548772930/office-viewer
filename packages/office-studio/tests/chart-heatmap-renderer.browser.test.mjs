import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { resolve } from 'node:path'
import { chromium } from '@playwright/test'
import { createServer } from 'vite'

// P21 drafts for actual ECharts Canvas geometry, native visual precedence and
// forwarded hover. No runner, dev server, browser or pixel check has started.
const fixturePath = '/__chart-heatmap-renderer-test__'
const fixture = `<!doctype html><html lang="en"><head><meta charset="utf-8"></head><body><script type="module">
import { getInstanceByDom } from 'echarts/core'
import { createBlankWorkbook, createCell } from '/packages/office-studio/src/domain/model.ts'
import { createChart } from '/packages/office-studio/src/features/charts/chart-engine.ts'
import { chartEChartsOption } from '/packages/office-studio/src/features/charts/echarts-option.ts'
import { EChartsCanvasRenderer } from '/packages/office-studio/src/features/charts/echarts-renderer.ts'
const renderer = new EChartsCanvasRenderer(), size = { width: 520, height: 360, devicePixelRatio: 1 }
function model(options = {}) {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  const table = [['Region', 'Units', 'Revenue'], ['North', 1, 2], ['South', 3, 4], ['West', 5, 6]]
  sheet.cells = Object.fromEntries(table.flatMap((row, r) => row.map((value, c) => [r + ':' + c, createCell(r, c, value)])))
  for (let row = 1; row <= 3; row++) sheet.cells[row + ':2'].style = { numberFormat: '0.00' }
  const chart = createChart(workbook, sheet.id, { type: 'heatmap', categoryField: 0, valueFields: [1, 2],
    source: { range: { startRow: 0, startCol: 0, endRow: 3, endCol: 2 }, orientation: 'columns' },
    position: { x: 0, y: 0 }, size: { width: size.width, height: size.height }, palette: ['#0000ff', '#ff0000'],
    title: 'Heatmap renderer draft', legend: { visible: true, position: 'bottom' }, options })
  return { workbook, chart }
}
const center = rect => ({ x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 })
window.heatmapRendererFixture = { renderer, size, model, center, chartEChartsOption, getInstanceByDom }
</script></body></html>`
let server, browser, page, baseURL
const errors = []
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(),
    resolve: { alias: { echarts: resolve(process.cwd(), 'packages/office-studio/node_modules/echarts') } },
    plugins: [{ name: 'chart-heatmap-renderer-fixture', configureServer(vite) {
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
  await page.goto(baseURL); await page.waitForFunction(() => Boolean(window.heatmapRendererFixture))
})
afterEach(async () => { await page?.close(); assert.deepEqual(errors, []) })
afterAll(async () => { await browser?.close(); await server?.close() })

test('actual native Canvas rectangles and text preserve field X/category Y, source formatting and resolved author colors', async () => {
  const result = await page.evaluate(() => {
    const f = window.heatmapRendererFixture, { workbook, chart } = f.model({ heatmap: { label: { visible: true }, unit: ' kg ' },
      namedSeries: [{ name: 'Revenue', style: { color: 'purple', dataPoints: { 1: { color: 'green', fillOpacity: 0.4 } } } }] })
    const author = JSON.stringify(chart), option = f.chartEChartsOption(workbook, chart), frame = f.renderer.render(chart.id, option, f.size)
    const instance = f.getInstanceByDom(frame.source), native = instance.getModel().getSeriesByIndex(0).getData(), geometry = option.officeChartPointer.geometry
    const points = Array.from({ length: native.count() }, (_value, index) => {
      const element = native.getItemGraphicEl(index), shape = element.shape, identity = option.series[0].data[index].value
      const projected = geometry.layout.series[identity[0]].points.find(point => point.categoryIndex === identity[1])
      return { x: shape.x + shape.width / 2, y: shape.y + shape.height / 2, expectedX: projected.x, expectedY: projected.y,
        fill: element.style.fill, text: element.getTextContent()?.style.text, fontSize: element.getTextContent()?.style.fontSize }
    })
    const legendId = chart.id + ':heatmap-legend', graphic = instance.getOption().graphic[0].elements
    return { points, legend: graphic.some(element => String(element.id).startsWith(legendId)), nativeLegendHidden: instance.getOption().legend[0].show === false,
      formatted: geometry.data.series[1].formattedValues, authorUnchanged: JSON.stringify(chart) === author }
  })
  for (const point of result.points) {
    assert.ok(Math.abs(point.x - point.expectedX) < 1, 'native X center')
    assert.ok(Math.abs(point.y - point.expectedY) < 1, 'native Y center')
    assert.equal(point.fontSize, 12)
  }
  assert.equal(result.points[3].fill, 'rgba(128, 0, 128, 1)')
  assert.equal(result.points[4].fill, 'rgba(0, 128, 0, 0.4)')
  assert.equal(result.points[4].text, '4.00')
  assert.deepEqual(result.formatted, ['2.00', '4.00', '6.00'])
  assert.equal(result.legend, true); assert.equal(result.nativeLegendHidden, true)
  assert.equal(result.authorUnchanged, true)
})

test('colorbar hover patches only transient graphic guides, reuses the cache and leaves native content/range and authors intact', async () => {
  const result = await page.evaluate(() => {
    const f = window.heatmapRendererFixture, { workbook, chart } = f.model({ heatmap: { min: 0, max: 8 } })
    const option = f.chartEChartsOption(workbook, chart), baseline = f.renderer.render(chart.id, option, f.size), instance = f.getInstanceByDom(baseline.source)
    const author = JSON.stringify(chart), model = instance.getModel(), native = model.getSeriesByIndex(0).getData(), element = native.getItemGraphicEl(0), fill = element.style.fill
    const range = JSON.stringify(model.getComponent('visualMap').option.range), bar = option.officeChartPointer.geometry.heatmap.legend.bar
    const calls = [], original = instance.setOption.bind(instance)
    instance.setOption = (patch, flags) => { calls.push(patch); return original(patch, flags) }
    const point = { x: bar.x + bar.width * 3 / 8, y: bar.y + 6 }, hover = f.renderer.updatePointer(chart.id, point)
    const repeated = f.renderer.updatePointer(chart.id, point), cached = f.renderer.render(chart.id, f.chartEChartsOption(workbook, chart), f.size)
    const focused = calls[0]?.graphic.elements[0], leave = f.renderer.updatePointer(chart.id)
    return { newer: hover.sourceVersion > baseline.sourceVersion, repeatEmpty: repeated === undefined, cached: cached.sourceVersion === hover.sourceVersion && cached.source === baseline.source,
      leaveNewer: leave.sourceVersion > hover.sourceVersion, remove: calls[1]?.graphic.elements[0].$action === 'remove',
      onlyGuides: calls.every(call => Object.keys(call).join() === 'graphic'), group: focused?.id === chart.id + ':axis-pointer' && focused.$action === 'replace',
      focusedCell: focused?.children.some(child => (child.type === 'rect' ? child : child.children?.[0])?.style.stroke !== undefined),
      retained: instance.getModel().getSeriesByIndex(0).getData().getItemGraphicEl(0) === element && element.style.fill === fill,
      range: JSON.stringify(instance.getModel().getComponent('visualMap').option.range) === range, author: JSON.stringify(chart) === author }
  })
  for (const flag of ['newer', 'repeatEmpty', 'cached', 'leaveNewer', 'remove', 'onlyGuides', 'group', 'focusedCell', 'retained', 'range', 'author']) assert.equal(result[flag], true, flag)
})

test('native eight-piece colors match the drawn intervals, including exact boundaries and inactive out-of-range values', async () => {
  const result = await page.evaluate(() => {
    const f = window.heatmapRendererFixture, { workbook, chart } = f.model({ heatmap: { visualMapType: 'piecewise', min: 0, max: 8, label: { visible: true } } })
    const sheet = workbook.sheets[0]; sheet.cells['1:1'].value = -1; sheet.cells['2:1'].value = 1; sheet.cells['3:1'].value = 1.01
    const option = f.chartEChartsOption(workbook, chart), frame = f.renderer.render(chart.id, option, f.size), instance = f.getInstanceByDom(frame.source)
    const data = instance.getModel().getSeriesByIndex(0).getData(), runtime = option.officeChartPointer.geometry.heatmap
    const cell0 = data.getItemGraphicEl(0), cell1 = data.getItemGraphicEl(1), cell2 = data.getItemGraphicEl(2), swatch = runtime.legend.pieces.find(piece => piece.pieceIndex === 0)
    const hover = f.renderer.updatePointer(chart.id, f.center(swatch.hit))
    return { pieceCount: runtime.legend.pieces.length, boundary: cell1.style.fill === runtime.scale.pieces[0].color,
      next: cell2.style.fill === runtime.scale.pieces[1].color, inactive: cell0.style.opacity === 0,
      hiddenLabel: !cell0.getTextContent() || cell0.getTextContent().ignore === true, hover: hover.sourceVersion > frame.sourceVersion,
      unchanged: chart.options.heatmap.min === 0 && chart.options.heatmap.max === 8 }
  })
  assert.equal(result.pieceCount, 8)
  for (const flag of ['boundary', 'next', 'inactive', 'hiddenLabel', 'hover', 'unchanged']) assert.equal(result[flag], true, flag)
})

test('patterned graphic charts share the visible legend and hover while theme/source/DPR refresh and disposal replace its runtime owner', async () => {
  const result = await page.evaluate(() => {
    const f = window.heatmapRendererFixture, { workbook, chart } = f.model({ heatmap: { visualMapType: 'piecewise' }, allSeriesStyle: {
      fillPattern: { preset: 'pct50', foregroundColor: '#ff0000', backgroundColor: '#ffff00' } } })
    const author = JSON.stringify(chart), option = f.chartEChartsOption(workbook, chart), baseline = f.renderer.render(chart.id, option, f.size)
    const runtime = option.officeChartPointer.geometry.heatmap, point = f.center(runtime.legend.pieces[0].hit)
    const hover = f.renderer.updatePointer(chart.id, point), instance = f.getInstanceByDom(baseline.source)
    workbook.sheets[0].cells['1:1'].value = 20
    const changedOption = f.chartEChartsOption(workbook, chart, { textColor: '#ffffff', backgroundColor: '#111111', palette: ['#00ff00', '#ff00ff'] })
    const changed = f.renderer.render(chart.id, changedOption, { ...f.size, devicePixelRatio: 2 }), newInstance = f.getInstanceByDom(changed.source)
    const changedHit = changedOption.officeChartPointer.geometry.heatmap.legend.pieces[0]
    const currentHover = f.renderer.updatePointer(chart.id, f.center(changedHit.hit))
    f.renderer.retain(new Set())
    return { graphic: option.series === undefined, hover: hover.sourceVersion > baseline.sourceVersion, current: currentHover.sourceVersion > changed.sourceVersion,
      resetRange: changedOption.officeChartPointer.geometry.heatmap.scale.extent.max === 20, themed: changedOption.officeChartPointer.geometry.heatmap.color === '#ffffff',
      oldReleased: instance.isDisposed() && baseline.source.width === 1 && baseline.source.height === 1,
      released: newInstance.isDisposed() && changed.source.width === 1 && changed.source.height === 1 && f.renderer.updatePointer(chart.id, point) === undefined,
      author: JSON.stringify(chart) === author }
  })
  for (const flag of ['graphic', 'hover', 'current', 'resetRange', 'themed', 'oldReleased', 'released', 'author']) assert.equal(result[flag], true, flag)
})
