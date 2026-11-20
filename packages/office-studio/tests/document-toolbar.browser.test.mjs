import assert from 'node:assert/strict'
import { after, afterEach, before, beforeEach, test } from 'node:test'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

const fixturePath = '/__document-toolbar-test__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>Docs toolbar interaction fixture</title><style>body{margin:0}#app{height:100vh;min-width:0}</style></head><body><div id="app"></div><script type="module">
import { createApp, h, ref } from 'vue'
import DocumentWorkbench from '/packages/office-studio/src/components/DocumentWorkbench.vue'
import { createBlankDocument } from '/packages/office-studio/src/domain/model'
import { createParagraph } from '/packages/office-studio/src/features/document/model/document-model'
const initial = createBlankDocument()
initial.id = 'toolbar-doc'; initial.name = '工具栏交互文档'
initial.blocks = [createParagraph('甲乙丙丁')]
const model = ref(initial)
const history = ref([])
const redo = ref([])
const copy = (value) => JSON.parse(JSON.stringify(value))
const update = (value) => { history.value.push(copy(model.value)); redo.value = []; model.value = value }
window.documentToolbarFixture = {
  snapshot: () => copy(model.value),
  heading: () => { const next = copy(model.value); next.blocks[0].type = 'heading'; next.blocks[0].level = 1; model.value = next },
}
createApp({ render: () => h(DocumentWorkbench, {
  modelValue: model.value, canUndo: !!history.value.length, canRedo: !!redo.value.length,
  'onUpdate:modelValue': update,
  onUndo: () => { if (history.value.length) { redo.value.push(copy(model.value)); model.value = history.value.pop() } },
  onRedo: () => { if (redo.value.length) { history.value.push(copy(model.value)); model.value = redo.value.pop() } },
}) }).mount('#app')
</script></body></html>`

let server
let browser
let page
let baseURL
let errors
before(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'document-toolbar-interaction-fixture',
    configureServer(vite) {
      vite.middlewares.use(fixturePath, async (_req, res, next) => {
        try { res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.end(await vite.transformIndexHtml(fixturePath, fixture)) } catch (error) { next(error) }
      })
    },
  }], server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen()
  const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate an HTTP port')
  baseURL = `http://127.0.0.1:${address.port}${fixturePath}`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, { timeout: 60_000 })
after(async () => { await browser?.close(); await server?.close() })
beforeEach(async () => {
  errors = []
  page = await browser.newPage({ viewport: { width: 1440, height: 900 } })
  page.on('pageerror', (error) => errors.push(error.message))
  await page.goto(baseURL)
  await expect(page.getByRole('textbox', { name: '文档画布编辑区', exact: true })).toBeVisible()
}, { timeout: 60_000 })
afterEach(async () => { await page?.close() })

test('Docs composes a classic Ribbon and a single footer without duplicating the application header', async () => {
  assert.equal(await page.locator('.document-header').count(), 0)
  assert.equal(await page.locator('.document-statusbar').count(), 1)
  assert.equal(await page.locator('.workbench-ribbon-tabs').evaluate((element) => element.getBoundingClientRect().height), 36)
  assert.equal(await page.locator('.workbench-ribbon-toolbar').evaluate((element) => element.getBoundingClientRect().height), 40)
  assert.equal(await page.locator('.document-statusbar').evaluate((element) => element.getBoundingClientRect().height), 36)
  assert.equal(await page.locator('.document-outline').count(), 0)
  await expect(page.locator('.document-statusbar')).toContainText('无未保存更改')
  assert.deepEqual(errors, [])
})

test('font selection preserves the logical Canvas range and uses the external document history', async () => {
  const bridge = page.getByLabel('文档文字输入', { exact: true })
  await bridge.focus()
  await bridge.press('Control+A')
  const font = page.getByRole('combobox', { name: '字体', exact: true })
  await font.click()
  await page.getByTitle('Georgia', { exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks[0].runs[0].fontFamily)).toBe('Georgia')
  await expect(page.locator('.document-statusbar')).toContainText('未保存更改')
  await page.getByRole('button', { name: '撤销', exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks[0].runs[0].fontFamily)).toBe('Arial')
  await page.getByRole('button', { name: '重做', exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks[0].runs[0].fontFamily)).toBe('Georgia')
  assert.deepEqual(errors, [])
})

test('outline, zoom and focus view operations keep document content and dirty state unchanged', async () => {
  await page.evaluate(() => window.documentToolbarFixture.heading())
  const before = await page.evaluate(() => window.documentToolbarFixture.snapshot())
  await page.getByRole('tab', { name: '视图', exact: true }).click()
  await page.getByRole('button', { name: '文档大纲', exact: true }).click()
  await expect(page.locator('.document-outline')).toBeVisible()
  await page.setViewportSize({ width: 680, height: 900 })
  await expect(page.locator('.document-outline')).toHaveClass(/document-outline--float/)
  await page.locator('.document-statusbar').getByRole('button', { name: '放大', exact: true }).click()
  await page.getByRole('button', { name: '专注模式', exact: true }).click()
  await expect(page.locator('.workbench-ribbon')).toHaveCount(0)
  await page.getByRole('button', { name: '退出专注模式', exact: true }).click()
  await expect(page.locator('.workbench-ribbon')).toBeVisible()
  assert.deepEqual(await page.evaluate(() => window.documentToolbarFixture.snapshot()), before)
  assert.deepEqual(errors, [])
})
