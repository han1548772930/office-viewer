import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { resolve } from 'node:path'
import { chromium } from '@playwright/test'
import { createServer } from 'vite'

// P20 renderer drafts: geometry, retained runtime and events; no pixel snapshots.
// No server, runner or browser has started during this implementation phase.
const fixturePath = '/__chart-sankey-renderer-test__'
const fixture = `<!doctype html><html lang="en"><head><meta charset="utf-8"></head><body><script type="module">
import { getInstanceByDom } from 'echarts/core'
import { createBlankWorkbook, createCell } from '/packages/office-studio/src/domain/model.ts'
import { createChart } from '/packages/office-studio/src/features/charts/chart-engine.ts'
import { chartEChartsOption } from '/packages/office-studio/src/features/charts/echarts-option.ts'
import { EChartsCanvasRenderer } from '/packages/office-studio/src/features/charts/echarts-renderer.ts'
const renderer = new EChartsCanvasRenderer(), size = { width: 520, height: 340, devicePixelRatio: 1 }
function model() {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  const table = [['Source', 'Target', 'Weight'], ['R', 'A', 10], ['A', 'B', 4], ['A', 'C', 6], ['B', 'D', 4], ['C', 'E', 6]]
  sheet.cells = Object.fromEntries(table.flatMap((row, r) => row.map((value, c) => [r + ':' + c, createCell(r, c, value)])))
  const chart = createChart(workbook, sheet.id, { type: 'sankey', categoryField: 0, valueFields: [1, 2],
    source: { range: { startRow: 0, startCol: 0, endRow: 5, endCol: 2 }, orientation: 'columns' },
    position: { x: 0, y: 0 }, size: { width: size.width, height: size.height }, title: 'Sankey renderer draft',
    legend: { visible: false, position: 'bottom' }, options: { mapping: { edge: { sourceIndex: 0, targetIndex: 1, valueIndex: 2 } } } })
  return { workbook, chart }
}
const center = node => ({ x: node.x + node.width / 2, y: node.y + node.height / 2 })
window.sankeyRendererFixture = { renderer, size, model, center, chartEChartsOption, getInstanceByDom }
</script></body></html>`

let server, browser, page, baseURL
const errors = []
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(),
    resolve: { alias: { echarts: resolve(process.cwd(), 'packages/office-studio/node_modules/echarts') } },
    plugins: [{ name: 'chart-sankey-renderer-fixture', configureServer(vite) {
      vite.middlewares.use(fixturePath, async (_request, response, next) => {
        try { response.setHeader('Content-Type', 'text/html; charset=utf-8'); response.end(await vite.transformIndexHtml(fixturePath, fixture)) }
        catch (error) { next(error) }
      })
    } }], server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen()
  const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate an HTTP port')
  baseURL = `http://127.0.0.1:${address.port}${fixturePath}`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)
beforeEach(async () => {
  errors.length = 0; page = await browser.newPage()
  page.on('pageerror', error => errors.push(error.message))
  await page.goto(baseURL)
  await page.waitForFunction(() => Boolean(window.sankeyRendererFixture))
})
afterEach(async () => {
  await page?.evaluate(() => window.sankeyRendererFixture?.renderer.dispose())
  await page?.close(); assert.deepEqual(errors, [])
})
afterAll(async () => { await browser?.close(); await server?.close() })

test('trajectory hover replaces only the Sankey content group, caches repeated hits and leaves author data unchanged', async () => {
  const result = await page.evaluate(() => {
    const f = window.sankeyRendererFixture, { workbook, chart } = f.model(), option = f.chartEChartsOption(workbook, chart)
    const runtime = option.officeChartPointer, first = f.renderer.render(chart.id, option, f.size), instance = f.getInstanceByDom(first.source)
    const before = JSON.stringify(chart), original = instance.setOption.bind(instance), patches = []
    const chrome = instance.getOption().graphic[0].elements.filter(element => element.id !== runtime.groupId).map(element => element.id)
    instance.setOption = (value, ...args) => { patches.push(value); return original(value, ...args) }
    const link = runtime.scene.layout.links.find(link => link.source === 'A' && link.target === 'B'), point = link.points[8]
    const hit = f.renderer.getContentHit(chart.id, point), shown = f.renderer.updatePointer(chart.id, point)
    const repeated = f.renderer.updatePointer(chart.id, point), cached = f.renderer.render(chart.id, f.chartEChartsOption(workbook, chart), f.size)
    const hidden = f.renderer.updatePointer(chart.id)
    return { hit, expected: link.index, shown: shown.sourceVersion, cached: cached.sourceVersion, hidden: hidden.sourceVersion,
      repeated: repeated === undefined, sameSource: cached.source === first.source, sameChart: before === JSON.stringify(chart),
      onlyContent: patches.length === 2 && patches.every(patch => Object.keys(patch).join() === 'graphic'
        && patch.graphic.elements.length === 1 && patch.graphic.elements[0].id === runtime.groupId && patch.graphic.elements[0].$action === 'replace'),
      chrome: JSON.stringify(chrome) === JSON.stringify(instance.getOption().graphic[0].elements.filter(element => element.id !== runtime.groupId).map(element => element.id)),
      stripped: !Object.hasOwn(instance.getOption(), 'officeChartPointer') }
  })
  assert.deepEqual(result.hit, { kind: 'link', index: result.expected })
  assert.equal(result.shown, 2); assert.equal(result.cached, 2); assert.equal(result.hidden, 3)
  for (const flag of ['repeated', 'sameSource', 'sameChart', 'onlyContent', 'chrome', 'stripped']) assert.equal(result[flag], true, flag)
})

test('node drags retain view positions across identical authored renders, cancel exactly and publish no model mutation', async () => {
  const result = await page.evaluate(() => {
    const f = window.sankeyRendererFixture, { workbook, chart } = f.model(), option = f.chartEChartsOption(workbook, chart)
    const runtime = option.officeChartPointer, first = f.renderer.render(chart.id, option, f.size), node = runtime.scene.layout.nodes.find(node => node.name === 'B')
    const initial = JSON.stringify(runtime.scene.layout), authored = JSON.stringify(chart), origin = f.center(node)
    const started = f.renderer.beginContentDrag(chart.id, origin), target = { x: origin.x + 35, y: origin.y + 10 }
    const updated = f.renderer.updateContentDrag(chart.id, target), moved = JSON.stringify(runtime.scene.layout)
    const invalidMove = f.renderer.updateContentDrag(chart.id, { x: NaN, y: origin.y })
    const committed = f.renderer.endContentDrag(chart.id), cached = f.renderer.render(chart.id, f.chartEChartsOption(workbook, chart), f.size)
    const hitAfterRender = f.renderer.getContentHit(chart.id, target), retained = JSON.stringify(runtime.scene.layout) === moved
    f.renderer.beginContentDrag(chart.id, target)
    f.renderer.updateContentDrag(chart.id, { x: origin.x - 35, y: origin.y + 25 })
    const cancelled = f.renderer.endContentDrag(chart.id, true), restored = JSON.stringify(runtime.scene.layout) === moved
    return { start: started.kind, nodeIndex: started.nodeIndex, expected: node.index, changed: initial !== moved,
      versionIncreases: updated.sourceVersion > first.sourceVersion && committed.sourceVersion > updated.sourceVersion
        && cancelled.sourceVersion > cached.sourceVersion, cachedVersion: cached.sourceVersion === committed.sourceVersion,
      invalidMove: invalidMove === undefined, retained, restored, hitAfterRender, sameChart: authored === JSON.stringify(chart),
      secondEnd: f.renderer.endContentDrag(chart.id) === undefined }
  })
  assert.equal(result.start, 'sankey-node'); assert.equal(result.nodeIndex, result.expected)
  assert.deepEqual(result.hitAfterRender, { kind: 'node', index: result.expected })
  for (const flag of ['changed', 'versionIncreases', 'cachedVersion', 'invalidMove', 'retained', 'restored', 'sameChart', 'secondEnd']) assert.equal(result[flag], true, flag)
})

test('source, size, theme and DPR changes reset view-only positions and end stale content drags', async () => {
  const result = await page.evaluate(() => {
    const f = window.sankeyRendererFixture, { workbook, chart } = f.model(), resets = []
    let option = f.chartEChartsOption(workbook, chart), frame = f.renderer.render(chart.id, option, f.size), size = { ...f.size }, theme = {}
    for (const change of ['source', 'size', 'theme', 'dpr']) {
      const node = option.officeChartPointer.scene.layout.nodes.find(node => node.name === 'B'), point = f.center(node)
      f.renderer.beginContentDrag(chart.id, point)
      f.renderer.updateContentDrag(chart.id, { x: point.x + 35, y: point.y + 10 })
      const previous = frame, instance = f.getInstanceByDom(frame.source)
      if (change === 'source') workbook.sheets[0].cells['2:2'].value = 8
      if (change === 'size') { chart.size = { width: 560, height: 360 }; size = { ...size, ...chart.size } }
      if (change === 'theme') theme = { textColor: 'white', backgroundColor: '#111111' }
      if (change === 'dpr') size = { ...size, devicePixelRatio: 2 }
      option = f.chartEChartsOption(workbook, chart, theme)
      frame = f.renderer.render(chart.id, option, size, theme)
      const automatic = option.officeChartPointer.scene.layout.nodes.find(node => node.name === 'B')
      resets.push({ change, stale: f.renderer.endContentDrag(chart.id) === undefined,
        automatic: f.renderer.getContentHit(chart.id, f.center(automatic))?.index === automatic.index,
        newer: frame.sourceVersion > previous.sourceVersion,
        released: change !== 'dpr' || (previous.source !== frame.source && instance.isDisposed() && f.getInstanceByDom(previous.source) === undefined) })
    }
    return resets
  })
  for (const reset of result) for (const flag of ['stale', 'automatic', 'newer', 'released']) assert.equal(reset[flag], true, `${reset.change}: ${flag}`)
})

test('retain and dispose release topology, hit regions and active drags with their Canvas owner', async () => {
  const result = await page.evaluate(() => {
    const f = window.sankeyRendererFixture, { workbook, chart } = f.model(), option = f.chartEChartsOption(workbook, chart)
    const frame = f.renderer.render(chart.id, option, f.size), instance = f.getInstanceByDom(frame.source)
    const point = f.center(option.officeChartPointer.scene.layout.nodes[0])
    f.renderer.beginContentDrag(chart.id, point); f.renderer.retain(new Set())
    const retained = [f.renderer.getContentHit(chart.id, point), f.renderer.updateContentDrag(chart.id, point), f.renderer.endContentDrag(chart.id)]
    const rebuilt = f.renderer.render(chart.id, f.chartEChartsOption(workbook, chart), f.size), rebuiltInstance = f.getInstanceByDom(rebuilt.source)
    f.renderer.beginContentDrag(chart.id, point); f.renderer.dispose(); f.renderer.dispose()
    return { absent: retained.every(value => value === undefined), oldReleased: instance.isDisposed() && f.getInstanceByDom(frame.source) === undefined,
      newReleased: rebuiltInstance.isDisposed() && f.getInstanceByDom(rebuilt.source) === undefined,
      size: [frame.source.width, frame.source.height, rebuilt.source.width, rebuilt.source.height],
      finalHit: f.renderer.getContentHit(chart.id, point) === undefined, finalDrag: f.renderer.endContentDrag(chart.id) === undefined }
  })
  for (const flag of ['absent', 'oldReleased', 'newReleased', 'finalHit', 'finalDrag']) assert.equal(result[flag], true, flag)
  assert.deepEqual(result.size, [1, 1, 1, 1])
})
