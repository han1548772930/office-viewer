import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium } from '@playwright/test'
import { createServer } from 'vite'
import { resolve } from 'node:path'

// P26 native Canvas drafts. No server, browser, pixels or benchmark ran.
const fixturePath = '/__chart-treemap-renderer-draft__'
const fixture = `<!doctype html><html lang="en"><head><meta charset="utf-8"></head><body><script type="module">
import { getInstanceByDom } from 'echarts/core'
import { treemapFixture } from '/packages/office-studio/tests/fixtures/chart-treemap-fixtures.mjs'
import { createCell } from '/packages/office-studio/src/domain/model.ts'
import { chartEChartsOption } from '/packages/office-studio/src/features/charts/echarts-option.ts'
import { EChartsCanvasRenderer } from '/packages/office-studio/src/features/charts/echarts-renderer.ts'
const renderer = new EChartsCanvasRenderer(), size = { width: 560, height: 400, devicePixelRatio: 1 }
function inspect(instance) {
  const data = instance.getModel().getSeriesByIndex(0).getData(), result = []
  data.tree.eachNode('preorder', node => {
    const item = data.getRawDataItem(node.dataIndex)
    if (!item?.semanticNode) return
    const graphic = data.getItemGraphicEl(node.dataIndex), texts = []
    if (graphic?.getTextContent?.()) texts.push(graphic.getTextContent())
    graphic?.traverse?.(child => { if (child !== graphic && child.getTextContent?.()) texts.push(child.getTextContent()) })
    result.push({ node, item, graphic, texts })
  })
  return result
}
window.treemapRendererDraft = { renderer, size, treemapFixture, createCell, chartEChartsOption, getInstanceByDom, inspect }
</script></body></html>`
let server, browser, page, address
const errors = []
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(),
    resolve: { alias: { echarts: resolve(process.cwd(), 'packages/office-studio/node_modules/echarts') } },
    plugins: [{ name: 'chart-treemap-renderer-draft', configureServer(vite) {
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
  await page.goto(address); await page.waitForFunction(() => Boolean(window.treemapRendererDraft))
})
afterEach(async () => { await page?.close(); assert.deepEqual(errors, []) })
afterAll(async () => { await browser?.close(); await server?.close() })

test('actual Canvas keeps parent labels when leaves are hidden, native residual identity and depth-tinted fill pixels', async () => {
  const result = await page.evaluate(() => {
    const f = window.treemapRendererDraft, { workbook, chart } = f.treemapFixture({ treemap: { label: { visible: false } } }), before = JSON.stringify(chart)
    const option = f.chartEChartsOption(workbook, chart), frame = f.renderer.render(chart.id, option, f.size), instance = f.getInstanceByDom(frame.source)
    const nodes = f.inspect(instance), parent = nodes.filter(item => item.node.viewChildren?.length), leaf = nodes.filter(item => !item.node.viewChildren?.length)
    const pixels = leaf.filter(item => !item.item.isTreemapResidual).map(item => {
      const box = item.graphic.getBoundingRect().clone(); box.applyTransform(item.graphic.getComputedTransform())
      return { name: item.item.semanticNode.name, rgba: Array.from(frame.source.getContext('2d').getImageData(Math.floor(box.x + box.width / 2), Math.floor(box.y + box.height / 2), 1, 1).data) }
    })
    const answer = { legend: option.legend.show, pointer: option.officeChartPointer.mode,
      parents: parent.map(item => ({ name: item.item.semanticNode.name, visible: item.texts.some(text => !text.ignore) })),
      leaves: leaf.map(item => ({ name: item.item.name, hidden: item.texts.every(text => text.ignore) })),
      residual: nodes.filter(item => item.item.isTreemapResidual).map(item => ({ name: item.item.name, owner: item.item.semanticNode.name, self: item.item.semanticNode.selfValue, value: item.item.value })),
      pixels, author: JSON.stringify(chart) === before }
    f.renderer.retain(new Set()); return answer
  })
  assert.equal(result.legend, false); assert.equal(result.pointer, 'treemap'); assert.equal(result.author, true)
  assert.deepEqual(result.parents, [{ name: 'North', visible: true }, { name: 'South', visible: true }])
  assert.equal(result.leaves.every(item => item.hidden), true)
  assert.deepEqual(result.residual, [{ name: '', owner: 'South', self: 3, value: 3 }])
  for (const pixel of result.pixels) assert.deepEqual(pixel.rgba, pixel.name === 'Leaf A' ? [46, 61, 75, 255] : [225, 240, 255, 255])
})

test('native overlapping layout refills the banner strip and elevates parent text over descendants without truncation', async () => {
  const result = await page.evaluate(() => {
    const f = window.treemapRendererDraft, runs = []
    for (const mode of ['banner', 'overlapping']) {
      const { workbook, chart } = f.treemapFixture({ treemap: { parentLabelLayout: mode, label: { visible: true, fontSize: 12, contentType: 10 } } })
      const option = f.chartEChartsOption(workbook, chart), frame = f.renderer.render(chart.id, option, f.size), nodes = f.inspect(f.getInstanceByDom(frame.source))
      const north = nodes.find(item => item.item.semanticNode.name === 'North' && item.node.viewChildren?.length)
      const children = north.node.viewChildren.map(child => child.getLayout())
      const descendants = nodes.filter(item => item.item.semanticNode.path[0] === 'North' && item !== north)
      runs.push({ mode, firstChildY: Math.min(...children.map(rect => rect.y)), labelHeight: north.node.getLayout().upperLabelHeight,
        parentZ: Math.max(...north.texts.map(text => text.z2)), descendantZ: Math.max(...descendants.flatMap(item => item.texts.map(text => text.z2))),
        unclamped: nodes.every(item => item.texts.every(text => text.beforeUpdate === undefined && text.style.height === undefined && text.style.lineOverflow === undefined)) })
      f.renderer.retain(new Set())
    }
    return runs
  })
  assert.ok(result[0].firstChildY >= result[0].labelHeight)
  assert.ok(result[1].firstChildY < result[1].labelHeight)
  assert.ok(result[1].parentZ > result[1].descendantZ)
  assert.equal(result.every(run => run.unclamped), true)
})

test('native complete-fit hides an oversized full source string and applies real relative label-background gradients', async () => {
  const result = await page.evaluate(() => {
    const f = window.treemapRendererDraft, background = { type: 'linear-gradient', start: { x: 0, y: 0 }, end: { x: 1, y: 0 }, stops: [{ offset: 0, color: 'red' }, { offset: 1, color: 'blue' }] }
    const { workbook, sheet, chart } = f.treemapFixture({ treemap: { parentLabelLayout: 'none', label: { fontSize: 24, backgroundColor: background } } })
    const name = '完整长标签'.repeat(20)
    sheet.cells['1:1'] = f.createCell(1, 1, name); sheet.cells['5:1'] = f.createCell(5, 1, name)
    chart.size = { width: 240, height: 120 }
    const option = f.chartEChartsOption(workbook, chart), frame = f.renderer.render(chart.id, option, { ...f.size, ...chart.size })
    const node = f.inspect(f.getInstanceByDom(frame.source)).find(item => item.item.name === name)
    const text = node.texts[0], answer = { source: text.style.text, ignore: text.ignore, states: ['emphasis', 'select', 'blur'].map(state => text.getState(state).ignore),
      gradient: text.style.backgroundColor, height: text.style.height, lineOverflow: text.style.lineOverflow, hasSetter: typeof text.beforeUpdate === 'function' }
    f.renderer.retain(new Set()); return answer
  })
  assert.equal(result.source, '完整长标签'.repeat(20)); assert.equal(result.ignore, true); assert.deepEqual(result.states, [true, true, true])
  assert.equal(result.gradient.type, 'linear'); assert.equal(result.gradient.global, false); assert.equal(result.gradient.x2, 1)
  assert.equal(result.height, undefined); assert.equal(result.lineOverflow, undefined); assert.equal(result.hasSetter, false)
})

test('native pattern paints the leaf content when all label states omit Text and keeps the outline color independent', async () => {
  const result = await page.evaluate(() => {
    const f = window.treemapRendererDraft, { workbook, chart } = f.treemapFixture({
      allSeriesStyle: { fillPattern: { preset: 'cross', foregroundColor: 'black', backgroundColor: 'white' }, fillOpacity: 0.4,
        border: { visible: true, color: 'red', opacity: 0.75, width: 3 } }, treemap: { parentLabelLayout: 'none', label: { visible: false } },
    }), before = JSON.stringify(chart), option = f.chartEChartsOption(workbook, chart)
    const hide = node => {
      node.label = { ...node.label, show: false }; node.upperLabel = { ...node.upperLabel, show: false }
      for (const name of ['emphasis', 'select', 'blur']) node[name] = { ...node[name], label: { show: false }, upperLabel: { show: false } }
      for (const child of node.children ?? []) hide(child)
    }
    for (const node of option.series[0].data) hide(node)
    const frame = f.renderer.render(chart.id, option, f.size), leaf = f.inspect(f.getInstanceByDom(frame.source)).find(item => item.item.name === 'Leaf A')
    const rectangles = []; leaf.graphic.traverse(child => { if (child.type === 'rect') rectangles.push(child) }); rectangles.sort((a, b) => a.z2 - b.z2)
    const background = rectangles[0], content = rectangles.at(-1), tile = content.style.fill, pixels = tile.image.getContext('2d').getImageData(0, 0, tile.image.width, tile.image.height).data
    const alpha = []
    for (let offset = 3; offset < pixels.length; offset += 4) alpha.push(pixels[offset])
    const answer = { textCount: leaf.texts.length, rectangles: rectangles.length, background: background.style.fill, backgroundOpacity: background.style.opacity,
      contentOpacity: content.style.opacity, repeat: tile.repeat, size: [tile.image.width, tile.image.height], alpha: [...new Set(alpha)],
      states: ['emphasis', 'select', 'blur'].map(name => content.getState(name).style.fill.image === tile.image), author: JSON.stringify(chart) === before }
    f.renderer.retain(new Set()); return answer
  })
  assert.equal(result.textCount, 0); assert.equal(result.rectangles, 2); assert.equal(result.background, 'rgba(255, 0, 0, 0.75)')
  assert.equal(result.backgroundOpacity, 1); assert.equal(result.contentOpacity, 1); assert.equal(result.repeat, 'repeat')
  assert.deepEqual(result.size, [16, 16]); assert.deepEqual(result.alpha, [102]); assert.deepEqual(result.states, [true, true, true]); assert.equal(result.author, true)
})

test('bitmap hover shows semantic residual tooltip and clears/restores ancestor emphasis through leave and reentry', async () => {
  const result = await page.evaluate(() => {
    const f = window.treemapRendererDraft, { workbook, chart } = f.treemapFixture(), before = JSON.stringify(chart), option = f.chartEChartsOption(workbook, chart)
    const frame = f.renderer.render(chart.id, option, f.size), instance = f.getInstanceByDom(frame.source), residual = f.inspect(instance).find(item => item.item.isTreemapResidual)
    const box = residual.graphic.getBoundingRect().clone(); box.applyTransform(residual.graphic.getComputedTransform())
    const point = { x: box.x + box.width / 2, y: box.y + box.height / 2 }, words = () => instance.getZr().storage.getDisplayList().map(item => item.style?.text ?? '').filter(Boolean)
    const active = f.renderer.updatePointer(chart.id, point), during = words(), cached = f.renderer.updatePointer(chart.id, point)
    const emphasized = f.inspect(instance).find(item => item.item.semanticNode.name === 'South' && item.node.viewChildren?.length).graphic.currentStates
    const left = f.renderer.updatePointer(chart.id), cleared = words(), returned = f.renderer.updatePointer(chart.id, point), reentered = words()
    const answer = { active: Boolean(active), cached: cached === undefined, left: Boolean(left), returned: Boolean(returned), during, cleared, reentered,
      emphasis: emphasized.includes('emphasis'), author: JSON.stringify(chart) === before }
    f.renderer.retain(new Set()); return answer
  })
  assert.equal(result.active, true); assert.equal(result.cached, true); assert.equal(result.left, true); assert.equal(result.returned, true)
  assert.ok(result.during.some(text => /Weight: 9/.test(text))); assert.ok(result.during.some(text => /自身数值: 3/.test(text)))
  assert.equal(result.cleared.some(text => /Weight: 9/.test(text)), false); assert.ok(result.reentered.some(text => /Weight: 9/.test(text)))
  assert.equal(result.emphasis, true); assert.equal(result.author, true)
})
