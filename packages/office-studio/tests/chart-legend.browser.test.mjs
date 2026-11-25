import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { resolve } from 'node:path'
import { chromium } from '@playwright/test'
import { createServer } from 'vite'

// Draft: real Canvas hit/layout, retained views and production Overlay events.
// No browser, dev server, runner or screenshot/pixel check has started in this batch.
let server, browser, page, baseURL
const errors = []
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(), server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen()
  const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate an HTTP port')
  baseURL = `http://127.0.0.1:${address.port}/packages/office-studio/tests/fixtures/chart-legend.html`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)
beforeEach(async () => {
  errors.length = 0; page = await browser.newPage()
  page.on('pageerror', error => errors.push(error.message))
  await page.goto(baseURL); await page.waitForFunction(() => Boolean(window.chartLegendFixture))
})
afterEach(async () => { await page?.evaluate(() => window.chartLegendFixture?.dispose()); await page?.close(); assert.deepEqual(errors, []) })
afterAll(async () => { await browser?.close(); await server?.close() })

test('native multiple/single/closed modes change actual series visibility, retain cache/resize/DPR view state and leave authors unchanged', async () => {
  const result = await page.evaluate(() => {
    const f = window.chartLegendFixture, value = f.model({ type: 'line' }), before = JSON.stringify(value.chart), first = f.paint(value).frame
    const target = f.hitPoint(value.chart.id, 'Series 1'), changed = f.renderer.activateLegend(value.chart.id, target)
    const surface = f.surfaces.get(value.chart.id), model = surface.instance.getModel(), hidden = model.isSeriesFiltered(model.getSeriesByIndex(1))
    const cached = f.paint(value).frame, resized = f.paint(value, 2, { backgroundColor: '#202020' }).frame
    const retained = f.renderer.getLegendView(value.chart.id).selected['Series 1'] === false
    const single = f.model({ legend: { selectMode: 'single' } }); f.paint(single)
    f.renderer.activateLegend(single.chart.id, f.hitPoint(single.chart.id, 'Series 2'))
    const selected = f.renderer.getLegendView(single.chart.id).selected
    const closed = f.model({ legend: { selectMode: undefined } }), closedFrame = f.paint(closed).frame
    const unchanged = f.renderer.activateLegend(closed.chart.id, f.hitPoint(closed.chart.id, 'Series 1'))
    return { hidden, changed: changed.sourceVersion > first.sourceVersion, cached: cached.sourceVersion === changed.sourceVersion,
      sameCanvas: first.source === cached.source, replacedCanvas: resized.source !== first.source, retained,
      single: Object.values(selected).filter(Boolean).length === 1 && selected['Series 2'],
      closed: unchanged === undefined && f.renderer.getLegendView(closed.chart.id).selected['Series 1'] === true,
      closedMode: f.getInstanceByDom(closedFrame.source).getOption().legend[0].selectedMode === false,
      author: JSON.stringify(value.chart) === before }
  })
  for (const [key, value] of Object.entries(result)) assert.equal(value, true, key)
})

test('shared Canvas fallbacks remove real content marks and restore them while native paging keeps all names reachable', async () => {
  const result = await page.evaluate(() => {
    const f = window.chartLegendFixture, value = f.model({ shared: true }), before = JSON.stringify(value.chart)
    const first = f.paint(value), count = f.graphicRectangles(first.instance).length
    f.renderer.activateLegend(value.chart.id, f.hitPoint(value.chart.id, 'Series 1'))
    const fewer = f.graphicRectangles(first.instance).length < count
    f.renderer.activateLegend(value.chart.id, f.hitPoint(value.chart.id, 'Series 1'))
    const restored = f.graphicRectangles(first.instance).length === count
    const many = f.model({ count: 16, width: 300, legend: { selectMode: 'close', wrap: false } }), initial = f.paint(many).frame
    const moved = f.renderer.activateLegend(many.chart.id, f.hitPoint(many.chart.id, 'pageNext')), view = f.renderer.getLegendView(many.chart.id)
    const cached = f.paint(many).frame
    return { fewer, restored, author: before === JSON.stringify(value.chart), page: view.scrollDataIndex > 0,
      version: moved.sourceVersion > initial.sourceVersion, cached: cached.sourceVersion === moved.sourceVersion,
      retainedPage: f.renderer.getLegendView(many.chart.id).scrollDataIndex === view.scrollDataIndex,
      names: Object.keys(view.selected).length === 16 && Object.values(view.selected).every(Boolean) }
  })
  for (const [key, value] of Object.entries(result)) assert.equal(value, true, key)
})

test('production Overlay forwards zoomed/scrolled/read-only clicks as view operations, cancels drags and retains selection through theme/DPR changes', async () => {
  const baseline = await page.evaluate(() => window.chartLegendFixture.mountOverlay({ scale: 150, readonly: true, shared: true }))
  await page.waitForFunction(() => document.querySelector('.chart-hitbox')?.dataset.rendered === 'true')
  const point = await page.evaluate(() => window.chartLegendFixture.overlayPoint('Series 1'))
  await page.mouse.click(point.x, point.y)
  await page.waitForFunction(() => window.chartLegendFixture.overlayState().selected['Series 1'] === false)
  const selected = await page.evaluate(() => window.chartLegendFixture.overlayState())
  assert.equal(selected.readOnly, true); assert.equal(selected.updates, 0); assert.equal(selected.removals, 0)
  assert.equal(selected.revision, baseline.revision); assert.equal(selected.author, baseline.author)
  const drag = await page.evaluate(() => window.chartLegendFixture.overlayPoint('Series 0'))
  await page.mouse.move(drag.x, drag.y); await page.mouse.down(); await page.mouse.move(drag.x + 16, drag.y); await page.mouse.up()
  const beforeTheme = await page.evaluate(() => window.chartLegendFixture.overlayState())
  assert.equal(beforeTheme.selected['Series 0'], true)
  await page.evaluate(async () => {
    await window.chartLegendFixture.setDark(true); await window.chartLegendFixture.setZoom(125)
    const scroller = document.querySelector('.canvas-scroll'); scroller.scrollLeft = 24; scroller.scrollTop = 12
    scroller.dispatchEvent(new Event('scroll'))
  })
  await page.waitForFunction(previousVersion => {
    const state = window.chartLegendFixture.overlayState()
    return state.sourceVersion > previousVersion && state.zoom === 125 && state.dark === true
      && state.bitmapWidth === state.expectedWidth && state.bitmapHeight === state.expectedHeight
  }, beforeTheme.sourceVersion)
  const final = await page.evaluate(() => window.chartLegendFixture.overlayState())
  assert.equal(final.selected['Series 1'], false); assert.equal(final.updates, 0); assert.equal(final.author, baseline.author)
})

test('a partially dispatched legend failure keeps the old view and forces a complete repaint before cache reuse', async () => {
  const result = await page.evaluate(() => {
    const f = window.chartLegendFixture, value = f.model({ shared: true }), first = f.paint(value), before = JSON.stringify(value.chart)
    const originalCount = f.graphicRectangles(first.instance).length, target = f.hitPoint(value.chart.id, 'Series 1')
    const dispatch = first.instance.dispatchAction
    let failed = false
    first.instance.dispatchAction = function (...args) {
      const result = dispatch.apply(this, args)
      if (args[0]?.type === 'legendToggleSelect') throw new Error('Injected legend action failure')
      return result
    }
    try { f.renderer.activateLegend(value.chart.id, target) } catch (error) { failed = error.message === 'Injected legend action failure' }
    finally { first.instance.dispatchAction = dispatch }
    const retained = f.renderer.getLegendView(value.chart.id).selected['Series 1'] === true, recovered = f.paint(value), cached = f.paint(value)
    return { failed, retained, author: JSON.stringify(value.chart) === before,
      painted: recovered.instance.getOption().legend[0].selected['Series 1'] === true && f.graphicRectangles(recovered.instance).length === originalCount,
      refreshed: recovered.frame.sourceVersion > first.frame.sourceVersion, cached: cached.frame.sourceVersion === recovered.frame.sourceVersion }
  })
  for (const [key, value] of Object.entries(result)) assert.equal(value, true, key)
})
