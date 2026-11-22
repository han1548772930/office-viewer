import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { resolve } from 'node:path'
import { chromium } from '@playwright/test'
import { createServer } from 'vite'

const fixturePath = '/__echarts-renderer-test__'
const fixture = `<!doctype html><html lang="en"><head><meta charset="utf-8"></head><body><script type="module">
import { getInstanceByDom } from 'echarts/core'
import { createBlankWorkbook } from '/packages/office-studio/src/domain/model.ts'
import { createChart } from '/packages/office-studio/src/features/charts/chart-engine.ts'
import { chartEChartsOption } from '/packages/office-studio/src/features/charts/echarts-option.ts'
import { EChartsCanvasRenderer } from '/packages/office-studio/src/features/charts/echarts-renderer.ts'

const renderer = new EChartsCanvasRenderer()
const size = { width: 420, height: 280, devicePixelRatio: 1 }
function model(type) {
  const workbook = createBlankWorkbook()
  const sheet = workbook.sheets[0]
  sheet.cells = {
    '0:0': { value: 'X' }, '0:1': { value: 'Value' }, '0:2': { value: 'Other' },
    '1:0': { value: 1 }, '1:1': { value: 5 }, '1:2': { value: 7 },
    '2:0': { value: 2 }, '2:1': { value: 12 }, '2:2': { value: 9 },
    '3:0': { value: 3 }, '3:1': { value: 20 }, '3:2': { value: 15 },
  }
  const chart = createChart(workbook, sheet.id, {
    type,
    source: { range: { startRow: 0, startCol: 0, endRow: 3, endCol: 2 }, orientation: 'columns' },
    position: { x: 0, y: 0 }, size: { width: size.width, height: size.height },
    valueFields: type === 'pie' || type === 'doughnut' ? [1] : [1, 2],
    legend: { visible: false, position: 'bottom' },
  })
  return { workbook, chart }
}
function pixels(canvas) {
  const context = canvas.getContext('2d')
  if (!context) throw new Error('Canvas2D is unavailable')
  const bytes = context.getImageData(0, 0, canvas.width, canvas.height).data
  let hash = 2166136261
  let nontransparent = 0
  let blue = 0
  for (let index = 0; index < bytes.length; index += 4) {
    if (bytes[index + 3] > 0) nontransparent++
    if (bytes[index + 3] > 0 && bytes[index + 2] > bytes[index] + 30 && bytes[index + 2] > bytes[index + 1] + 20) blue++
    for (let channel = 0; channel < 4; channel++) hash = Math.imul(hash ^ bytes[index + channel], 16777619)
  }
  return { hash: hash >>> 0, nontransparent, blue }
}
window.rendererFixture = { renderer, size, model, chartEChartsOption, getInstanceByDom, pixels }
</script></body></html>`

let server
let browser
let page
let baseURL
const errors = []

beforeAll(async () => {
  server = await createServer({
    configFile: false,
    root: process.cwd(),
    resolve: { alias: { echarts: resolve(process.cwd(), 'packages/office-studio/node_modules/echarts') } },
    plugins: [{
      name: 'echarts-renderer-fixture',
      configureServer(vite) {
        vite.middlewares.use(fixturePath, async (_request, response, next) => {
          try {
            response.setHeader('Content-Type', 'text/html; charset=utf-8')
            response.end(await vite.transformIndexHtml(fixturePath, fixture))
          } catch (error) {
            next(error)
          }
        })
      },
    }],
    server: { host: '127.0.0.1', port: 0 },
    logLevel: 'error',
  })
  await server.listen()
  const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate an HTTP port')
  baseURL = `http://127.0.0.1:${address.port}${fixturePath}`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)

beforeEach(async () => {
  errors.length = 0
  page = await browser.newPage()
  page.on('pageerror', error => errors.push(error.message))
  await page.goto(baseURL)
  await page.waitForFunction(() => Boolean(window.rendererFixture))
})

afterEach(async () => {
  await page?.evaluate(() => window.rendererFixture?.renderer.dispose())
  await page?.close()
  assert.deepEqual(errors, [])
})

afterAll(async () => {
  await browser?.close()
  await server?.close()
})

test('all supported workbook chart types paint real series pixels into a Canvas', async () => {
  const results = await page.evaluate(() => {
    const fixture = window.rendererFixture
    return ['bar', 'line', 'area', 'pie', 'doughnut', 'stacked-column', 'scatter', 'radar'].map(type => {
      const { workbook, chart } = fixture.model(type)
      const frame = fixture.renderer.render(type, fixture.chartEChartsOption(workbook, chart), fixture.size)
      return { type, version: frame.sourceVersion, width: frame.source.width, height: frame.source.height, ...fixture.pixels(frame.source) }
    })
  })
  for (const result of results) {
    assert.equal(result.version, 1, result.type)
    assert.deepEqual([result.width, result.height], [420, 280], result.type)
    assert.ok(result.nontransparent > 100, `${result.type} has no rendered content`)
    assert.ok(result.blue > 20, `${result.type} has no rendered blue data series`)
  }
})

test('cached options reuse the source and version while data changes repaint and increment the version', async () => {
  const result = await page.evaluate(() => {
    const fixture = window.rendererFixture
    const { workbook, chart } = fixture.model('bar')
    const option = fixture.chartEChartsOption(workbook, chart)
    const first = fixture.renderer.render('chart', option, fixture.size)
    const firstPixels = fixture.pixels(first.source)
    const cached = fixture.renderer.render('chart', structuredClone(option), fixture.size)
    workbook.sheets[0].cells['1:1'].value = 35
    const updated = fixture.renderer.render('chart', fixture.chartEChartsOption(workbook, chart), fixture.size)
    return {
      cachedSameSource: cached.source === first.source,
      cachedVersion: cached.sourceVersion,
      firstVersion: first.sourceVersion,
      updatedSameSource: updated.source === first.source,
      updatedVersion: updated.sourceVersion,
      beforeHash: firstPixels.hash,
      afterHash: fixture.pixels(updated.source).hash,
    }
  })
  assert.equal(result.cachedSameSource, true)
  assert.equal(result.cachedVersion, result.firstVersion)
  assert.equal(result.updatedSameSource, true)
  assert.equal(result.updatedVersion, result.firstVersion + 1)
  assert.notEqual(result.afterHash, result.beforeHash)
})

test('resize reuses the instance and DPR changes replace it with a correctly sized backing store', async () => {
  const result = await page.evaluate(() => {
    const fixture = window.rendererFixture
    const { workbook, chart } = fixture.model('line')
    const option = fixture.chartEChartsOption(workbook, chart)
    const first = fixture.renderer.render('chart', option, fixture.size)
    const originalInstance = fixture.getInstanceByDom(first.source)
    const resized = fixture.renderer.render('chart', option, { ...fixture.size, width: 500, height: 320 })
    const resize = {
      sameSource: resized.source === first.source,
      sameInstance: fixture.getInstanceByDom(resized.source) === originalInstance,
      backingSize: [resized.source.width, resized.source.height],
      version: resized.sourceVersion,
    }
    const retina = fixture.renderer.render('chart', option, { width: 500, height: 320, devicePixelRatio: 2 })
    return {
      resize,
      replacedSource: retina.source !== resized.source,
      oldInstanceReleased: fixture.getInstanceByDom(resized.source) === undefined && originalInstance.isDisposed(),
      backingSize: [retina.source.width, retina.source.height],
      version: retina.sourceVersion,
      blue: fixture.pixels(retina.source).blue,
    }
  })
  assert.deepEqual(result.resize, { sameSource: true, sameInstance: true, backingSize: [500, 320], version: 2 })
  assert.equal(result.replacedSource, true)
  assert.equal(result.oldInstanceReleased, true)
  assert.deepEqual(result.backingSize, [1000, 640])
  assert.equal(result.version, 3)
  assert.ok(result.blue > 20)
})

test('retain and dispose release chart instances through the ECharts public instance registry', async () => {
  const result = await page.evaluate(() => {
    const fixture = window.rendererFixture
    const { workbook, chart } = fixture.model('bar')
    const option = fixture.chartEChartsOption(workbook, chart)
    const kept = fixture.renderer.render('kept', option, fixture.size)
    const removed = fixture.renderer.render('removed', option, fixture.size)
    const keptInstance = fixture.getInstanceByDom(kept.source)
    const removedInstance = fixture.getInstanceByDom(removed.source)
    fixture.renderer.retain(new Set(['kept']))
    const retained = fixture.getInstanceByDom(kept.source) === keptInstance
    const removedReleased = fixture.getInstanceByDom(removed.source) === undefined && removedInstance.isDisposed()
    fixture.renderer.dispose()
    fixture.renderer.dispose()
    return {
      retained, removedReleased,
      keptReleased: fixture.getInstanceByDom(kept.source) === undefined && keptInstance.isDisposed(),
      keptSize: [kept.source.width, kept.source.height],
      removedSize: [removed.source.width, removed.source.height],
    }
  })
  assert.equal(result.retained, true)
  assert.equal(result.removedReleased, true)
  assert.equal(result.keptReleased, true)
  assert.deepEqual(result.keptSize, [1, 1])
  assert.deepEqual(result.removedSize, [1, 1])
})

test('invalid sizes are rejected before creating or changing an existing chart surface', async () => {
  const result = await page.evaluate(() => {
    const fixture = window.rendererFixture
    const { workbook, chart } = fixture.model('bar')
    const option = fixture.chartEChartsOption(workbook, chart)
    const frame = fixture.renderer.render('chart', option, fixture.size)
    const failures = []
    for (const patch of [{ width: 0 }, { height: -1 }, { devicePixelRatio: 0 }, { width: NaN }, { height: Infinity }]) {
      try {
        fixture.renderer.render('chart', option, { ...fixture.size, ...patch })
        failures.push('accepted')
      } catch (error) {
        failures.push(error.name)
      }
    }
    const cached = fixture.renderer.render('chart', option, fixture.size)
    return { failures, sameSource: cached.source === frame.source, version: cached.sourceVersion, size: [cached.source.width, cached.source.height] }
  })
  assert.deepEqual(result.failures, Array(5).fill('RangeError'))
  assert.equal(result.sameSource, true)
  assert.equal(result.version, 1)
  assert.deepEqual(result.size, [420, 280])
})
