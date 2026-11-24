import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium } from '@playwright/test'
import { createServer } from 'vite'
import { resolve } from 'node:path'

// P25 real Canvas drafts only. No server, browser, pixels or benchmark ran.
const fixturePath = '/__chart-candlestick-renderer-draft__'
const fixture = `<!doctype html><html lang="en"><head><meta charset="utf-8"></head><body><script type="module">
import { getInstanceByDom } from 'echarts/core'
import { candlestickFixture, CANDLE_HLC_MAPPING } from '/packages/office-studio/tests/fixtures/chart-candlestick-fixtures.mjs'
import { createBlankWorkbook, createCell } from '/packages/office-studio/src/domain/model.ts'
import { createChart } from '/packages/office-studio/src/features/charts/chart-engine.ts'
import { chartEChartsOption } from '/packages/office-studio/src/features/charts/echarts-option.ts'
import { EChartsCanvasRenderer } from '/packages/office-studio/src/features/charts/echarts-renderer.ts'
const renderer = new EChartsCanvasRenderer(), size = { width: 560, height: 400, devicePixelRatio: 1 }
function candle(options = {}) {
  const result = candlestickFixture(options)
  result.chart.size = { width: size.width, height: size.height }
  return result
}
function histogram(binGap) {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0], table = [['Row', 'Value'], ['A', 0], ['B', 1], ['C', 2], ['D', 3], ['E', 4]]
  sheet.cells = Object.fromEntries(table.flatMap((row, r) => row.map((value, c) => [r + ':' + c, createCell(r, c, value)])))
  const chart = createChart(workbook, sheet.id, { type: 'histogram', categoryField: 0, valueFields: [1],
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 5, endCol: 1 } },
    position: { x: 0, y: 0 }, size: { width: size.width, height: size.height },
    options: { histogram: { binning: { mode: 'count', count: 4 }, underflowThreshold: 1, overflowThreshold: 3, binGap } } }, 'histogram-renderer-draft')
  return { workbook, chart }
}
window.statisticalTypesRendererDraft = { renderer, size, candle, histogram, hlc: CANDLE_HLC_MAPPING, chartEChartsOption, getInstanceByDom }
</script></body></html>`
let server, browser, page, address
const errors = []
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(),
    resolve: { alias: { echarts: resolve(process.cwd(), 'packages/office-studio/node_modules/echarts') } },
    plugins: [{ name: 'chart-candlestick-renderer-draft', configureServer(vite) {
      vite.middlewares.use(fixturePath, async (_request, response, next) => {
        try { response.setHeader('Content-Type', 'text/html; charset=utf-8'); response.end(await vite.transformIndexHtml(fixturePath, fixture)) }
        catch (error) { next(error) }
      })
    } }], server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen()
  const bound = server.httpServer.address()
  if (!bound || typeof bound === 'string') throw new Error('Vite did not allocate a port')
  address = `http://127.0.0.1:${bound.port}${fixturePath}`
  browser = await chromium.launch({ headless: true })
})
beforeEach(async () => {
  errors.length = 0; page = await browser.newPage()
  page.on('pageerror', error => errors.push(error.message))
  await page.goto(address); await page.waitForFunction(() => Boolean(window.statisticalTypesRendererDraft))
})
afterEach(async () => { await page?.close(); assert.deepEqual(errors, []) })
afterAll(async () => { await browser?.close(); await server?.close() })

test('actual native Canvas keeps Rising/Falling/Doji paint independent and omits the missing OHLC row', async () => {
  const result = await page.evaluate(() => {
    const f = window.statisticalTypesRendererDraft, { workbook, chart } = f.candle({ candlestick: {
      candleWidth: 40, rising: { color: 'red', fillOpacity: 0.25, border: { color: 'navy', opacity: 0.8, width: 3 } },
      falling: { color: 'blue', hollow: true, border: { color: 'green', opacity: 0.6, width: 2 } },
      doji: { border: { color: 'black', opacity: 0.4, width: 2 } },
    } }), before = JSON.stringify(chart), option = f.chartEChartsOption(workbook, chart)
    const frame = f.renderer.render(chart.id, option, f.size), instance = f.getInstanceByDom(frame.source), data = instance.getModel().getSeriesByIndex(0).getData()
    const items = [0, 1, 2].map(index => {
      const path = data.getItemGraphicEl(index), layout = data.getItemLayout(index)
      return { sign: layout.sign, fill: path.style.fill, stroke: path.style.stroke, width: path.style.lineWidth }
    })
    const cached = f.renderer.render(chart.id, f.chartEChartsOption(workbook, chart), f.size)
    const answer = { native: option.series[0].type, large: option.series[0].large, items,
      empty: data.getItemGraphicEl(3) === undefined, cached: frame.source === cached.source && frame.sourceVersion === cached.sourceVersion,
      author: JSON.stringify(chart) === before }
    f.renderer.retain(new Set()); return answer
  })
  assert.equal(result.native, 'candlestick'); assert.equal(result.large, false)
  assert.deepEqual(result.items, [
    { sign: 1, fill: 'rgba(255, 0, 0, 0.25)', stroke: 'rgba(0, 0, 128, 0.8)', width: 3 },
    { sign: -1, fill: 'rgba(0, 0, 255, 0)', stroke: 'rgba(0, 128, 0, 0.6)', width: 2 },
    { sign: 0, fill: 'rgba(78, 85, 101, 1)', stroke: 'rgba(0, 0, 0, 0.4)', width: 2 },
  ])
  assert.equal(result.empty, true); assert.equal(result.cached, true); assert.equal(result.author, true)
})

test('actual HLC graphic Canvas retains dashed high/low strokes and Dot close lengths through a marker clear', async () => {
  const result = await page.evaluate(() => {
    const f = window.statisticalTypesRendererDraft, { workbook, chart } = f.candle({ mapping: { candlestick: f.hlc }, candlestick: {
      highLowLine: { border: { color: 'navy', width: 2, dashType: 'dashed', opacity: 0.5 } },
      closeMarker: { point: { color: 'green', shape: 'dot', size: 12 } },
    } })
    chart.yAxis.reverse = true
    const option = f.chartEChartsOption(workbook, chart), frame = f.renderer.render(chart.id, option, f.size), instance = f.getInstanceByDom(frame.source)
    const display = instance.getZr().storage.getDisplayList(), marks = display.filter(path => path.type === 'line' && path.style.stroke === 'rgba(0, 128, 0, 1)')
    const ranges = display.filter(path => path.type === 'line' && path.style.stroke === 'rgba(0, 0, 128, 0.5)')
    const lengths = marks.map(path => path.shape.x2 - path.shape.x1), lineWidths = marks.map(path => path.style.lineWidth)
    chart.options.candlestick.closeMarker.point.shape = 'none'; chart.revision++
    const cleared = f.renderer.render(chart.id, f.chartEChartsOption(workbook, chart), f.size)
    const remaining = f.getInstanceByDom(cleared.source).getZr().storage.getDisplayList()
      .filter(path => path.type === 'line' && path.style.stroke === 'rgba(0, 128, 0, 1)')
    const answer = { graphic: option.series === undefined, lengths, lineWidths, rangeSegments: ranges.length,
      rangesVertical: ranges.every(path => path.shape.x1 === path.shape.x2), remaining: remaining.length }
    f.renderer.retain(new Set()); return answer
  })
  assert.equal(result.graphic, true); assert.deepEqual(result.lengths, [6, 6, 6, 6]); assert.deepEqual(result.lineWidths, [1, 1, 1, 1])
  assert.ok(result.rangeSegments > 4); assert.equal(result.rangesVertical, true); assert.equal(result.remaining, 0)
})

test('actual Histogram Canvas uses materialized labels and the same valid/default gap ratio', async () => {
  const result = await page.evaluate(() => {
    const f = window.statisticalTypesRendererDraft, runs = []
    for (const gap of [0.25, 2]) {
      const { workbook, chart } = f.histogram(gap), option = f.chartEChartsOption(workbook, chart)
      const frame = f.renderer.render(chart.id, option, f.size), instance = f.getInstanceByDom(frame.source), data = instance.getModel().getSeriesByIndex(0).getData()
      const axis = instance.getModel().getComponent('xAxis').axis, band = axis.getBandWidth()
      runs.push({ labels: option.xAxis.data, gap: option.series[0].barCategoryGap,
        widths: [0, 1, 2, 3].map(index => Math.abs(data.getItemLayout(index).width) / band),
        counts: option.series[0].data, authorGap: chart.options.histogram.binGap })
      f.renderer.retain(new Set())
    }
    return runs
  })
  for (const run of result) assert.deepEqual(run.labels, ['≤ 1', '[2, 2.5]', '(2.5, 3]', '> 3'])
  assert.equal(result[0].gap, '25%'); assert.equal(result[0].authorGap, 0.25)
  assert.ok(Math.abs(Number.parseFloat(result[1].gap) - 0.7) < 1e-12); assert.equal(result[1].authorGap, 2)
  assert.deepEqual(result[0].counts, [2, 1, 1, 1])
  for (const width of result[0].widths) assert.ok(Math.abs(width - 0.75) < 1e-5)
  for (const width of result[1].widths) assert.ok(Math.abs(width - 0.993) < 1e-5)
})
