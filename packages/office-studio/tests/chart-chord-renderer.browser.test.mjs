import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { resolve } from 'node:path'
import { chromium } from '@playwright/test'
import { createServer } from 'vite'

// P28 browser drafts: actual ECharts Canvas, measured text and retained runtime.
// No server, browser, runner, WASM build or pixel snapshot ran in this phase.
const fixturePath = '/__chart-chord-renderer-test__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"></head><body><script type="module">
import { getInstanceByDom, graphic } from 'echarts/core'
import { createBlankWorkbook, createCell } from '/packages/office-studio/src/domain/model.ts'
import { createChart } from '/packages/office-studio/src/features/charts/chart-engine.ts'
import { chartEChartsOption } from '/packages/office-studio/src/features/charts/echarts-option.ts'
import { EChartsCanvasRenderer } from '/packages/office-studio/src/features/charts/echarts-renderer.ts'
const renderer = new EChartsCanvasRenderer(), size = { width: 560, height: 400, devicePixelRatio: 1 }
function model(names = ['B', 'A', 'C', 'D']) {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0], [b, a, c, d] = names
  const table = [['Source', 'Target', 'Weight'], [b, a, 2], [a, b, 3], [c, d, 4], [d, c, 1], [b, b, 9]]
  sheet.cells = Object.fromEntries(table.flatMap((row, r) => row.map((value, col) => {
    const cell = createCell(r, col, value)
    if (col === 2 && r > 0) { cell.style.numberFormat = '0.00'; cell.display = value.toFixed(2) }
    return [r + ':' + col, cell]
  })))
  const chart = createChart(workbook, sheet.id, { type: 'chord', categoryField: 0, valueFields: [1, 2],
    source: { range: { startRow: 0, startCol: 0, endRow: 5, endCol: 2 }, orientation: 'columns' },
    position: { x: 0, y: 0 }, size: { width: size.width, height: size.height }, title: 'Chord renderer draft',
    subtitle: 'Aggregate weights', legend: { visible: true, position: 'bottom' },
    options: { mapping: { chord: { sourceIndex: 0, targetIndex: 1, valueIndex: 2 } },
      chord: { node: { opacity: 0.45, borderWidth: 2 }, ribbon: { opacity: 0.35 } } } })
  return { workbook, chart }
}
function nodePoint(runtime, index = 0) {
  const layout = runtime.scene.layout, node = layout.nodes[index], angle = (node.start + node.end) / 2
  const radius = (layout.innerRadius + layout.outerRadius) / 2
  return { x: layout.center.x + Math.cos(angle) * radius, y: layout.center.y + Math.sin(angle) * radius }
}
function texts(instance) {
  const found = []
  for (const root of instance.getZr().storage.getRoots()) root.traverse(element => {
    if (element instanceof graphic.Text) found.push(element)
  })
  return found
}
window.chordRendererFixture = { renderer, size, model, nodePoint, texts, chartEChartsOption, getInstanceByDom }
</script></body></html>`

let server, browser, page, baseURL
const errors = []
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(),
    resolve: { alias: { echarts: resolve(process.cwd(), 'packages/office-studio/node_modules/echarts') } },
    plugins: [{ name: 'chart-chord-renderer-fixture', configureServer(vite) {
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
  await page.waitForFunction(() => Boolean(window.chordRendererFixture))
  await page.evaluate(() => document.fonts.ready)
})
afterEach(async () => {
  await page?.evaluate(() => window.chordRendererFixture?.renderer.dispose())
  await page?.close(); assert.deepEqual(errors, [])
})
afterAll(async () => { await browser?.close(); await server?.close() })

test('Chord hover replaces only content, caches repeated hits and restores the exact authored base paint', async () => {
  const result = await page.evaluate(() => {
    const f = window.chordRendererFixture, { workbook, chart } = f.model(), option = f.chartEChartsOption(workbook, chart)
    const runtime = option.officeChartPointer, first = f.renderer.render(chart.id, option, f.size)
    const instance = f.getInstanceByDom(first.source), original = instance.setOption.bind(instance), patches = []
    const authored = JSON.stringify(chart), scene = JSON.stringify(runtime.scene), point = f.nodePoint(runtime)
    const chrome = instance.getOption().graphic[0].elements.filter(element => element.id !== runtime.groupId).map(element => element.id)
    const baseChildren = option.graphic.elements.find(element => element.id === runtime.groupId).children
    instance.setOption = (value, ...args) => { patches.push(value); return original(value, ...args) }
    const tooltip = f.renderer.getContentTooltip(chart.id, point), shown = f.renderer.updatePointer(chart.id, point)
    const repeat = f.renderer.updatePointer(chart.id, { x: point.x + 0.01, y: point.y })
    const cached = f.renderer.render(chart.id, f.chartEChartsOption(workbook, chart), f.size)
    const hidden = f.renderer.updatePointer(chart.id), secondLeave = f.renderer.updatePointer(chart.id)
    return { tooltip, versions: [first.sourceVersion, shown.sourceVersion, cached.sourceVersion, hidden.sourceVersion],
      repeated: repeat === undefined && secondLeave === undefined, sameSurface: first.source === cached.source,
      unchanged: authored === JSON.stringify(chart) && scene === JSON.stringify(runtime.scene),
      contentOnly: patches.length === 2 && patches.every(patch => Object.keys(patch).join() === 'graphic'
        && patch.graphic.elements.length === 1 && patch.graphic.elements[0].id === runtime.groupId
        && patch.graphic.elements[0].$action === 'replace'),
      restored: JSON.stringify(patches.at(-1).graphic.elements[0].children) === JSON.stringify(baseChildren),
      chrome: JSON.stringify(chrome) === JSON.stringify(instance.getOption().graphic[0].elements.filter(element => element.id !== runtime.groupId).map(element => element.id)),
      stripped: !Object.hasOwn(instance.getOption(), 'officeChartPointer'),
      noLegend: instance.getOption().legend === undefined || instance.getOption().legend.length === 0 }
  })
  assert.deepEqual(result.tooltip, { id: 'chord-node:0', kind: 'node', index: 0, title: 'B', lines: ['Weight：5.00'] })
  assert.deepEqual(result.versions, [1, 2, 2, 3])
  for (const key of ['repeated', 'sameSurface', 'unchanged', 'contentOnly', 'restored', 'chrome', 'stripped', 'noLegend']) assert.equal(result[key], true, key)
})

test('rotated label hits use actual ZRender font bounds and transforms; font zero removes label hit regions', async () => {
  const result = await page.evaluate(() => {
    const f = window.chordRendererFixture, { workbook, chart } = f.model(['Wide 中文 WWW', 'iiii', 'C', 'D'])
    chart.options.chord.label = { rotation: 27, fontSize: 18, distance: 20, family: 'Georgia', position: 'outside' }
    const option = f.chartEChartsOption(workbook, chart), runtime = option.officeChartPointer
    const frame = f.renderer.render(chart.id, option, f.size), instance = f.getInstanceByDom(frame.source)
    const text = f.texts(instance).find(element => element.style.text === 'Wide 中文 WWW')
    const bounds = text.getBoundingRect(), transform = text.getComputedTransform()
    const expected = [[bounds.x, bounds.y], [bounds.x + bounds.width, bounds.y],
      [bounds.x + bounds.width, bounds.y + bounds.height], [bounds.x, bounds.y + bounds.height]].map(([x, y]) => ({
        x: transform ? transform[0] * x + transform[2] * y + transform[4] : x,
        y: transform ? transform[1] * x + transform[3] * y + transform[5] : y,
      }))
    const measured = runtime.hitIndex.labels.find(label => label.index === 0)
    const center = { x: expected.reduce((sum, point) => sum + point.x, 0) / 4, y: expected.reduce((sum, point) => sum + point.y, 0) / 4 }
    const tooltip = f.renderer.getContentTooltip(chart.id, center)
    const delta = Math.max(...expected.flatMap((point, index) => [Math.abs(point.x - measured.polygon[index].x), Math.abs(point.y - measured.polygon[index].y)]))
    const count = runtime.hitIndex.labels.length
    chart.options.chord.label.fontSize = 0
    const zero = f.chartEChartsOption(workbook, chart); f.renderer.render(chart.id, zero, f.size)
    return { delta, count, tooltip, rotated: Math.abs(expected[1].x - expected[0].x) > 1 && Math.abs(expected[1].y - expected[0].y) > 1,
      zeroLabels: zero.officeChartPointer.hitIndex.labels.length,
      noNameText: !f.texts(instance).some(element => element.style.text === 'Wide 中文 WWW') }
  })
  assert.ok(result.delta < 1e-8); assert.equal(result.count, 4); assert.equal(result.rotated, true)
  assert.equal(result.tooltip.id, 'chord-node:0'); assert.equal(result.tooltip.title, 'Wide 中文 WWW')
  assert.equal(result.zeroLabels, 0); assert.equal(result.noNameText, true)
})

test('none focus still supplies aggregate tooltips; count mapping reports occurrence totals without content drags', async () => {
  const result = await page.evaluate(() => {
    const f = window.chordRendererFixture, { workbook, chart } = f.model()
    chart.options.chord.emphasis = { focus: 'none' }
    let option = f.chartEChartsOption(workbook, chart), frame = f.renderer.render(chart.id, option, f.size), point = f.nodePoint(option.officeChartPointer)
    const weighted = f.renderer.getContentTooltip(chart.id, point), unchanged = f.renderer.updatePointer(chart.id, point) === undefined
    const version = f.renderer.render(chart.id, f.chartEChartsOption(workbook, chart), f.size).sourceVersion
    chart.options.mapping.chord = { sourceIndex: 0, targetIndex: 1 }; chart.valueFields = [1]
    option = f.chartEChartsOption(workbook, chart); frame = f.renderer.render(chart.id, option, f.size); point = f.nodePoint(option.officeChartPointer)
    return { weighted, unchanged, version, count: f.renderer.getContentTooltip(chart.id, point),
      noDrag: f.renderer.beginContentDrag(chart.id, point) === undefined,
      outside: f.renderer.getContentTooltip(chart.id, { x: -1, y: -1 }) === undefined }
  })
  assert.deepEqual(result.weighted.lines, ['Weight：5.00']); assert.deepEqual(result.count.lines, ['计数：2'])
  assert.equal(result.version, 1)
  for (const key of ['unchanged', 'noDrag', 'outside']) assert.equal(result[key], true, key)
})

test('equal normalized geometry refreshes weighted and formatted tooltip facts while retaining the cached hover bitmap', async () => {
  const result = await page.evaluate(() => {
    const f = window.chordRendererFixture, { workbook, chart } = f.model(), option = f.chartEChartsOption(workbook, chart)
    const initial = f.renderer.render(chart.id, option, f.size), point = f.nodePoint(option.officeChartPointer)
    const shown = f.renderer.updatePointer(chart.id, point), labels = option.officeChartPointer.hitIndex.labels
    for (let row = 1; row <= 4; row++) { const cell = workbook.sheets[0].cells[row + ':2']; cell.value *= 2; cell.display = cell.value.toFixed(2) }
    const refreshed = f.chartEChartsOption(workbook, chart), next = f.renderer.render(chart.id, refreshed, f.size)
    const weighted = f.renderer.getContentTooltip(chart.id, point), repeated = f.renderer.updatePointer(chart.id, point)
    for (let row = 1; row <= 4; row++) { const cell = workbook.sheets[0].cells[row + ':2']; cell.style.numberFormat = '0.0'; cell.display = cell.value.toFixed(1) }
    const formatted = f.chartEChartsOption(workbook, chart), formattedFrame = f.renderer.render(chart.id, formatted, f.size)
    const formattedTooltip = f.renderer.getContentTooltip(chart.id, point), hidden = f.renderer.updatePointer(chart.id)
    return { versions: [initial.sourceVersion, shown.sourceVersion, next.sourceVersion, formattedFrame.sourceVersion, hidden.sourceVersion],
      weighted, formattedTooltip, cachedHit: repeated === undefined, sameSurface: initial.source === next.source,
      retainedLabels: refreshed.officeChartPointer.hitIndex.labels === labels && formatted.officeChartPointer.hitIndex.labels === labels,
      sums: refreshed.officeChartPointer.scene.layout.nodes.map(node => node.value) }
  })
  assert.deepEqual(result.versions, [1, 2, 2, 2, 3]); assert.deepEqual(result.sums, [10, 10, 10, 10])
  assert.deepEqual(result.weighted.lines, ['Weight：10.00']); assert.deepEqual(result.formattedTooltip.lines, ['Weight：10.0'])
  for (const key of ['cachedHit', 'sameSurface', 'retainedLabels']) assert.equal(result[key], true, key)
})

test('theme, size and DPR changes rebuild hit regions and release replaced Canvas owners', async () => {
  const result = await page.evaluate(() => {
    const f = window.chordRendererFixture, { workbook, chart } = f.model(), changes = []
    let size = { ...f.size }, theme = {}, option = f.chartEChartsOption(workbook, chart), frame = f.renderer.render(chart.id, option, size)
    for (const change of ['theme', 'size', 'dpr']) {
      f.renderer.updatePointer(chart.id, f.nodePoint(option.officeChartPointer))
      const previous = frame, oldInstance = f.getInstanceByDom(frame.source)
      if (change === 'theme') theme = { textColor: '#ffffff', backgroundColor: '#111111' }
      if (change === 'size') { chart.size = { width: 600, height: 440 }; size = { ...size, ...chart.size } }
      if (change === 'dpr') size = { ...size, devicePixelRatio: 2 }
      option = f.chartEChartsOption(workbook, chart, theme); frame = f.renderer.render(chart.id, option, size, theme)
      changes.push({ change, newer: frame.sourceVersion > previous.sourceVersion, labels: option.officeChartPointer.hitIndex.labels.length,
        hit: f.renderer.getContentTooltip(chart.id, f.nodePoint(option.officeChartPointer)).index === 0,
        backing: frame.source.width === Math.round(size.width * size.devicePixelRatio) && frame.source.height === Math.round(size.height * size.devicePixelRatio),
        released: change !== 'dpr' || (frame.source !== previous.source && oldInstance.isDisposed() && f.getInstanceByDom(previous.source) === undefined) })
    }
    return changes
  })
  for (const change of result) {
    assert.equal(change.labels, 4)
    for (const key of ['newer', 'hit', 'backing', 'released']) assert.equal(change[key], true, `${change.change}: ${key}`)
  }
})

test('retain and dispose release tooltip facts, label regions and every Canvas owner', async () => {
  const result = await page.evaluate(() => {
    const f = window.chordRendererFixture, { workbook, chart } = f.model(), option = f.chartEChartsOption(workbook, chart)
    const frame = f.renderer.render(chart.id, option, f.size), instance = f.getInstanceByDom(frame.source), point = f.nodePoint(option.officeChartPointer)
    f.renderer.updatePointer(chart.id, point); f.renderer.retain(new Set())
    const absent = f.renderer.getContentTooltip(chart.id, point) === undefined && f.renderer.updatePointer(chart.id, point) === undefined
    const rebuilt = f.renderer.render(chart.id, f.chartEChartsOption(workbook, chart), f.size), next = f.getInstanceByDom(rebuilt.source)
    f.renderer.dispose(); f.renderer.dispose()
    return { absent, oldReleased: instance.isDisposed() && f.getInstanceByDom(frame.source) === undefined,
      newReleased: next.isDisposed() && f.getInstanceByDom(rebuilt.source) === undefined,
      sizes: [frame.source.width, frame.source.height, rebuilt.source.width, rebuilt.source.height],
      final: f.renderer.getContentTooltip(chart.id, point) === undefined && f.renderer.updatePointer(chart.id, point) === undefined }
  })
  for (const key of ['absent', 'oldReleased', 'newReleased', 'final']) assert.equal(result[key], true, key)
  assert.deepEqual(result.sizes, [1, 1, 1, 1])
})
