import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium } from '@playwright/test'
import { createServer } from 'vite'
import { resolve } from 'node:path'

// P27 actual native Canvas drafts. No server, browser, pixels or benchmark ran.
const fixturePath = '/__chart-sunburst-renderer-draft__'
const fixture = `<!doctype html><html lang="en"><head><meta charset="utf-8"></head><body><script type="module">
import { getInstanceByDom } from 'echarts/core'
import { sunburstFixture } from '/packages/office-studio/tests/fixtures/chart-sunburst-fixtures.mjs'
import { createCell } from '/packages/office-studio/src/domain/model.ts'
import { chartEChartsOption } from '/packages/office-studio/src/features/charts/echarts-option.ts'
import { EChartsCanvasRenderer } from '/packages/office-studio/src/features/charts/echarts-renderer.ts'
const renderer = new EChartsCanvasRenderer(), size = { width: 560, height: 400, devicePixelRatio: 1 }
function inspect(instance) {
  const data = instance.getModel().getSeriesByIndex(0).getData(), result = []
  data.tree.eachNode('preorder', node => {
    const item = data.getRawDataItem(node.dataIndex)
    if (!item?.semanticPath) return
    const graphic = data.getItemGraphicEl(node.dataIndex)
    result.push({ node, item, graphic, text: graphic?.getTextContent?.() })
  })
  return result
}
function centerPoint(graphic) {
  const shape = graphic.shape, angle = (shape.startAngle + shape.endAngle) / 2, radius = (shape.r0 + shape.r) / 2
  return { x: shape.cx + Math.cos(angle) * radius, y: shape.cy + Math.sin(angle) * radius }
}
window.sunburstRendererDraft = { renderer, size, sunburstFixture, createCell, chartEChartsOption, getInstanceByDom, inspect, centerPoint }
</script></body></html>`
let server, browser, page, address
const errors = []
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(),
    resolve: { alias: { echarts: resolve(process.cwd(), 'packages/office-studio/node_modules/echarts') } },
    plugins: [{ name: 'chart-sunburst-renderer-draft', configureServer(vite) {
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
  errors.length = 0; page = await browser.newPage(); page.on('pageerror', error => errors.push(error.message))
  await page.goto(address); await page.waitForFunction(() => Boolean(window.sunburstRendererDraft))
})
afterEach(async () => { await page?.close(); assert.deepEqual(errors, []) })
afterAll(async () => { await browser?.close(); await server?.close() })

test('actual native layout retains signed parent arcs, positive residuals, flat semantic facts and the true center total', async () => {
  const result = await page.evaluate(() => {
    const f = window.sunburstRendererDraft, { workbook, chart } = f.sunburstFixture({ sunburst: { centerLabel: { visible: true, format: '0.00', fontSize: 21 } } })
    const before = JSON.stringify(chart), option = f.chartEChartsOption(workbook, chart), frame = f.renderer.render(chart.id, option, f.size)
    const instance = f.getInstanceByDom(frame.source), nodes = f.inspect(instance), north = nodes.find(entry => entry.item.name === 'North')
    const parent = north.node.getLayout(), children = north.node.children.map(node => node.getLayout())
    const center = instance.getZr().storage.getDisplayList().find(graphic => graphic.style?.text === '22.00')
    const answer = { legend: option.legend.show, pointer: option.officeChartPointer.mode, own: north.item.semanticSelfValue, value: north.node.getValue(),
      parentSpan: parent.endAngle - parent.startAngle, childrenSpan: children.at(-1).endAngle - children[0].startAngle,
      roots: nodes.filter(entry => entry.item.semanticPath.length === 1 && !entry.item.isSunburstResidual).map(entry => [entry.item.name, entry.item.value]),
      residuals: nodes.filter(entry => entry.item.isSunburstResidual).map(entry => ({ name: entry.item.name, value: entry.item.value, path: entry.item.semanticPath, owner: entry.item.semanticParentId })),
      center: center && { text: center.style.text, fontSize: center.style.fontSize }, native: option.series[0].type,
      author: JSON.stringify(chart) === before, flat: nodes.every(entry => !Object.hasOwn(entry.item, 'semanticNode')) }
    f.renderer.retain(new Set()); return answer
  })
  assert.equal(result.native, 'sunburst'); assert.equal(result.legend, false); assert.equal(result.pointer, 'sunburst')
  assert.equal(result.value, 13); assert.equal(result.own, -5)
  assert.ok(Math.abs(result.parentSpan - 13 / 22 * Math.PI * 2) < 1e-8)
  assert.ok(Math.abs(result.childrenSpan - 18 / 22 * Math.PI * 2) < 1e-8); assert.ok(result.childrenSpan > result.parentSpan)
  assert.deepEqual(result.roots, [['North', 13], ['South', 9]])
  assert.equal(result.residuals.length, 1); assert.deepEqual(result.residuals[0].path, ['South']); assert.equal(result.residuals[0].value, 3)
  assert.equal(result.residuals[0].name, ''); assert.deepEqual(result.center, { text: '22.00', fontSize: 21 })
  assert.equal(result.flat, true); assert.equal(result.author, true)
})

test('actual native label states never expose residual, Empty0 or under-four-percent fallback node names', async () => {
  const result = await page.evaluate(() => {
    const f = window.sunburstRendererDraft, answers = []
    for (const contentType of [0, 10]) {
      const { workbook, sheet, chart } = f.sunburstFixture({ sunburst: { label: { contentType } } })
      sheet.cells['2:2'] = f.createCell(2, 2, 1000)
      const option = f.chartEChartsOption(workbook, chart), frame = f.renderer.render(chart.id, option, f.size), nodes = f.inspect(f.getInstanceByDom(frame.source))
      const hidden = contentType === 0 ? nodes : nodes.filter(entry => entry.item.isSunburstResidual || entry.item.name === 'Leaf A')
      answers.push({ contentType, normal: hidden.every(entry => entry.text.ignore),
        states: hidden.every(entry => ['emphasis', 'select', 'blur'].every(name => entry.text.getState(name).ignore)),
        positiveNames: hidden.filter(entry => entry.item.name).map(entry => entry.item.name) })
      f.renderer.retain(new Set())
    }
    return answers
  })
  assert.equal(result.every(run => run.normal && run.states), true)
  assert.ok(result[0].positiveNames.includes('North')); assert.ok(result[1].positiveNames.includes('Leaf A'))
})

test('native complete-fit retains oversized full text, gradient backgrounds and synchronized state hides', async () => {
  const result = await page.evaluate(() => {
    const f = window.sunburstRendererDraft, background = { type: 'linear-gradient', start: { x: 0, y: 0 }, end: { x: 1, y: 0 }, stops: [{ offset: 0, color: 'red' }, { offset: 1, color: 'blue' }] }
    const { workbook, sheet, chart } = f.sunburstFixture({ sunburst: { label: { fontSize: 24, backgroundColor: background } } })
    const name = '完整旭日图节点'.repeat(50); sheet.cells['2:1'] = f.createCell(2, 1, name); chart.size = { width: 240, height: 160 }
    const option = f.chartEChartsOption(workbook, chart), frame = f.renderer.render(chart.id, option, { ...f.size, ...chart.size })
    const node = f.inspect(f.getInstanceByDom(frame.source)).find(entry => entry.item.name === name), text = node.text
    const answer = { source: text.style.text, hidden: text.ignore, width: text.style.width, height: text.style.height, overflow: text.style.overflow,
      states: ['emphasis', 'select', 'blur'].map(state => ({ hidden: text.getState(state).ignore, background: text.getState(state).style.backgroundColor })),
      background: text.style.backgroundColor }
    f.renderer.retain(new Set()); return answer
  })
  assert.equal(result.source, '完整旭日图节点'.repeat(50)); assert.equal(result.hidden, true)
  assert.equal(result.width, undefined); assert.equal(result.height, undefined); assert.equal(result.overflow, undefined)
  assert.equal(result.background.type, 'linear'); assert.equal(result.background.global, false); assert.equal(result.background.x2, 1)
  assert.equal(result.states.every(state => state.hidden && state.background.type === 'linear'), true)
})

test('actual native sectors use synchronous disjoint pattern tiles with hidden labels and retain paint in view states', async () => {
  const result = await page.evaluate(() => {
    const f = window.sunburstRendererDraft, { workbook, chart } = f.sunburstFixture({ allSeriesStyle: { fillOpacity: 0.5,
      fillPattern: { preset: 'pct50', foregroundColor: 'black', backgroundColor: 'white' } }, sunburst: { label: { visible: false } } })
    const frame = f.renderer.render(chart.id, f.chartEChartsOption(workbook, chart), { ...f.size, devicePixelRatio: 2 })
    const node = f.inspect(f.getInstanceByDom(frame.source)).find(entry => entry.item.name === 'Leaf A'), fill = node.graphic.style.fill
    const tile = fill.image.getContext('2d').getImageData(0, 0, fill.image.width, fill.image.height).data, colors = new Set()
    for (let index = 0; index < tile.length; index += 4) colors.add(Array.from(tile.slice(index, index + 4)).join(','))
    const answer = { hidden: node.text.ignore, opacity: node.graphic.style.opacity, repeat: fill.repeat, size: [fill.image.width, fill.image.height],
      scale: [fill.scaleX, fill.scaleY], colors: Array.from(colors).sort(), states: ['emphasis', 'select', 'blur'].map(name => node.graphic.getState(name).style.fill === fill) }
    f.renderer.retain(new Set()); return answer
  })
  assert.equal(result.hidden, true); assert.equal(result.opacity, 1); assert.equal(result.repeat, 'repeat')
  assert.deepEqual(result.size, [32, 32]); assert.deepEqual(result.scale, [0.5, 0.5])
  assert.deepEqual(result.colors, ['0,0,0,128', '255,255,255,128']); assert.deepEqual(result.states, [true, true, true])
})

test('bitmap hover reports residual own value, ancestor emphasis, leave/reentry and source versions without content mutation', async () => {
  const result = await page.evaluate(() => {
    const f = window.sunburstRendererDraft, { workbook, chart } = f.sunburstFixture(), before = JSON.stringify(chart), option = f.chartEChartsOption(workbook, chart)
    const frame = f.renderer.render(chart.id, option, f.size), instance = f.getInstanceByDom(frame.source)
    const residual = f.inspect(instance).find(entry => entry.item.isSunburstResidual), point = f.centerPoint(residual.graphic)
    const words = () => instance.getZr().storage.getDisplayList().filter(entry => !entry.ignore).map(entry => entry.style?.text ?? '').filter(Boolean)
    const active = f.renderer.updatePointer(chart.id, point), during = words(), tooltip = f.renderer.getContentTooltip(chart.id, point), cached = f.renderer.updatePointer(chart.id, point)
    const south = f.inspect(instance).find(entry => entry.item.name === 'South'), emphasized = south.graphic.currentStates.includes('emphasis')
    const left = f.renderer.updatePointer(chart.id), cleared = words(), absent = f.renderer.getContentTooltip(chart.id, point) === undefined,
      returned = f.renderer.updatePointer(chart.id, point), reentered = f.renderer.getContentTooltip(chart.id, point)
    const answer = { versions: [frame.sourceVersion, active?.sourceVersion, left?.sourceVersion, returned?.sourceVersion], cached: cached === undefined,
      during, cleared, tooltip, absent, reentered, emphasized, author: JSON.stringify(chart) === before }
    f.renderer.retain(new Set()); return answer
  })
  assert.equal(result.cached, true); assert.ok(result.versions.every(Number.isFinite))
  for (let index = 1; index < result.versions.length; index++) assert.ok(result.versions[index] > result.versions[index - 1])
  assert.equal(result.during.some(text => /Weight: 3|自身数值: 3/.test(text)), false)
  assert.deepEqual(result.tooltip.lines, ['Weight: 3 (13.6%)', '自身数值: 3', '子节点合计: 6']); assert.equal(result.tooltip.title, 'South')
  assert.equal(result.absent, true); assert.deepEqual(result.reentered, result.tooltip)
  assert.equal(result.cleared.some(text => /Weight: 3/.test(text)), false)
  assert.equal(result.emphasized, true); assert.equal(result.author, true)
})

test('native theme, resize, DPR replacement and retain/dispose keep source versions and release old surfaces', async () => {
  const result = await page.evaluate(() => {
    const f = window.sunburstRendererDraft, { workbook, chart } = f.sunburstFixture(), before = JSON.stringify(chart)
    const first = f.renderer.render(chart.id, f.chartEChartsOption(workbook, chart), f.size)
    const retained = f.renderer.render(chart.id, f.chartEChartsOption(workbook, chart), f.size)
    const themed = f.renderer.render(chart.id, f.chartEChartsOption(workbook, chart, { backgroundColor: '#141414', textColor: '#eeeeee' }), f.size)
    chart.size = { width: 600, height: 420 }
    const resized = f.renderer.render(chart.id, f.chartEChartsOption(workbook, chart), { ...f.size, ...chart.size })
    const replaced = f.renderer.render(chart.id, f.chartEChartsOption(workbook, chart), { ...f.size, ...chart.size, devicePixelRatio: 2 })
    const same = retained.source === first.source && retained.sourceVersion === first.sourceVersion
    f.renderer.retain(new Set())
    return { same, versions: [first.sourceVersion, themed.sourceVersion, resized.sourceVersion, replaced.sourceVersion], replaced: replaced.source !== resized.source,
      oldSize: [resized.source.width, resized.source.height], released: [replaced.source.width, replaced.source.height], retainedAuthor: JSON.parse(before).options.mapping.hierarchyIndexes.join(',') === chart.options.mapping.hierarchyIndexes.join(',') }
  })
  assert.equal(result.same, true); assert.equal(result.replaced, true); assert.equal(result.retainedAuthor, true)
  for (let index = 1; index < result.versions.length; index++) assert.ok(result.versions[index] > result.versions[index - 1])
  assert.deepEqual(result.oldSize, [1, 1]); assert.deepEqual(result.released, [1, 1])
})
