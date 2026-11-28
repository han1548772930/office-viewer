import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

// Deferred actual Ant/Canvas/Facade/view-history integration drafts. Not executed.
const fixturePath = '/__document-zoom-contract__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><style>body{margin:0}#app{height:100vh}</style></head><body><div id="app"></div><script type="module">
import { createApp, h, ref, shallowRef } from 'vue'
import { ConfigProvider, theme } from 'antdv-next'
import { OfficeStudio } from '/packages/office-studio/src/api/office-studio'
import { FDocument } from '/packages/office-studio/src/api/facade'
import DocumentWorkbench from '/packages/office-studio/src/components/DocumentWorkbench.vue'
const source = id => ({ id, title: id, documentStyle: { documentFlavor: 1 }, body: { dataStream: '甲乙丙丁\\r\\n' }, settings: { zoomRatio: id === 'zoom-a' ? 0.8 : 2, kept: 'setting' } })
let studio = new OfficeStudio({ currentUserId: 'reader' })
studio.createDocument(source('zoom-a')); studio.createDocument(source('zoom-b')); studio.activateDocument('zoom-a')
const api = shallowRef(new FDocument(studio, 'zoom-a')), model = shallowRef(api.value.getModel()), variant = ref('classic'), dark = ref(false), notices = []
let subscriptions = [], cancel
function observe() {
  for (const subscription of subscriptions) subscription.dispose()
  subscriptions = [api.value.onModelChanged(() => { model.value = api.value.getModel() }),
    api.value.onZoomRatioChanged(() => { model.value = api.value.getModel() })]
}
observe()
window.documentZoomFixture = {
  snapshot: () => api.value.getSnapshot(),
  history: () => ({ undo: studio.editor.history.units[api.value.id]?.undo.length ?? 0, redo: studio.editor.history.units[api.value.id]?.redo.length ?? 0, dirty: studio.editor.history.isDirty(api.value.id), revision: api.value.getModelRevision() }),
  appearance: (flavor, layout, isDark = false) => { const saved = api.value.getSnapshot(); saved.documentStyle.documentFlavor = flavor; studio.createDocument(saved); api.value = new FDocument(studio, saved.id); model.value = api.value.getModel(); variant.value = layout; dark.value = isDark; observe() },
  switchUnit: id => { studio.activateDocument(id); api.value = new FDocument(studio, id); model.value = api.value.getModel(); observe() },
  saveReopen: () => { const saved = api.value.save(); for (const subscription of subscriptions) subscription.dispose(); studio.dispose(); studio = new OfficeStudio({ document: saved, currentUserId: 'reader' }); studio.activateDocument(saved.id); api.value = new FDocument(studio, saved.id); model.value = api.value.getModel(); observe(); return saved },
  setZoom: ratio => api.value.setZoomRatio(ratio),
  setDark: value => { dark.value = value },
  notices: () => [...notices],
  denyZoom: () => { cancel = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === 'doc.operation.set-zoom-ratio') event.cancel = true }) },
  permitZoom: () => { cancel?.dispose(); cancel = undefined },
  denyEdit: () => { api.value.setPermissionRules([{ id: 'zoom-edit-lock', scope: 'document', action: 'edit', effect: 'deny', userIds: ['reader'] }]) },
}
const app = createApp({ render: () => h(ConfigProvider, { theme: { algorithm: dark.value ? theme.darkAlgorithm : theme.defaultAlgorithm } },
  () => h(DocumentWorkbench, { modelValue: model.value, documentApi: api.value, ribbonVariant: variant.value,
    'onUpdate:modelValue': value => { model.value = value }, onUndo: () => api.value.undo(), onRedo: () => api.value.redo(), onNotify: value => notices.push(value) })) })
app.mount('#app')
window.addEventListener('pagehide', () => { cancel?.dispose(); for (const subscription of subscriptions) subscription.dispose(); app.unmount(); studio.dispose() }, { once: true })
</script></body></html>`

let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'document-zoom-browser-fixture',
    configureServer(vite) { vite.middlewares.use(fixturePath, async (_request, response, next) => { try {
      response.setHeader('Content-Type', 'text/html; charset=utf-8'); response.end(await vite.transformIndexHtml(fixturePath, fixture))
    } catch (error) { next(error) } }) },
  }], server: { host: '127.0.0.1', port: 0 }, appType: 'custom', logLevel: 'error' })
  await server.listen()
  const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate a fixture port')
  baseURL = 'http://127.0.0.1:' + address.port + fixturePath
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)
afterAll(async () => { await browser?.close(); await server?.close() })
beforeEach(async () => {
  errors = []; page = await browser.newPage({ viewport: { width: 1440, height: 900 } })
  page.on('pageerror', error => errors.push(error.message))
  await page.goto(baseURL); await page.locator('.document-statusbar').waitFor()
}, 60_000)
afterEach(async () => { await page?.close() })

for (const flavor of [1, 2]) for (const variant of ['classic', 'grid']) test('real Ribbon and Footer persist source zoom in flavor ' + flavor + ' / ' + variant, async () => {
  await page.evaluate(([flavor, variant]) => window.documentZoomFixture.appearance(flavor, variant), [flavor, variant])
  await page.getByRole('tab', { name: '视图', exact: true }).click()
  const control = page.locator('.ribbon-visible-groups [data-ribbon-item="doc-zoom"] .ant-select')
  await expect(control).toContainText('80%')
  const before = await page.evaluate(() => window.documentZoomFixture.history())
  await control.click(); await page.getByRole('option', { name: '130%', exact: true }).click()
  await expect(page.locator('.document-statusbar .zoom-value')).toContainText('130%')
  assert.equal((await page.evaluate(() => window.documentZoomFixture.snapshot())).settings.zoomRatio, 1.3)
  assert.deepEqual(await page.evaluate(() => window.documentZoomFixture.history()), before)
  await page.getByRole('button', { name: '放大', exact: true }).click()
  assert.equal((await page.evaluate(() => window.documentZoomFixture.snapshot())).settings.zoomRatio, 1.4)
  await page.evaluate(() => window.documentZoomFixture.saveReopen())
  await expect(page.locator('.document-statusbar .zoom-value')).toContainText('140%')
  assert.equal((await page.evaluate(() => window.documentZoomFixture.snapshot())).settings.kept, 'setting')
  assert.deepEqual(errors, [])
})

test('unit switching restores each document zoom and API events update both real controls', async () => {
  await page.evaluate(() => window.documentZoomFixture.setZoom(1.7))
  await expect(page.locator('.document-statusbar .zoom-value')).toContainText('170%')
  await page.evaluate(() => window.documentZoomFixture.switchUnit('zoom-b'))
  await expect(page.locator('.document-statusbar .zoom-value')).toContainText('200%')
  await page.evaluate(() => window.documentZoomFixture.switchUnit('zoom-a'))
  await expect(page.locator('.document-statusbar .zoom-value')).toContainText('170%')
  await page.getByRole('tab', { name: '视图', exact: true }).click()
  await expect(page.locator('.ribbon-visible-groups [data-ribbon-item="doc-zoom"]')).toContainText('170%')
  assert.deepEqual(errors, [])
})

test('canceled zoom retains real source ratio and read-only zoom stays available', async () => {
  const original = await page.evaluate(() => window.documentZoomFixture.snapshot())
  await page.evaluate(() => window.documentZoomFixture.denyZoom())
  await page.getByRole('button', { name: '放大', exact: true }).click()
  assert.deepEqual(await page.evaluate(() => window.documentZoomFixture.snapshot()), original)
  assert.equal((await page.evaluate(() => window.documentZoomFixture.notices())).length, 1)
  await page.evaluate(() => { window.documentZoomFixture.permitZoom(); window.documentZoomFixture.denyEdit() })
  const lockedHistory = await page.evaluate(() => window.documentZoomFixture.history())
  await page.getByRole('button', { name: '放大', exact: true }).click()
  await expect(page.locator('.document-statusbar .zoom-value')).toContainText('90%')
  assert.deepEqual(await page.evaluate(() => window.documentZoomFixture.history()), lockedHistory)
  assert.deepEqual(errors, [])
})

test('Footer and its Ant zoom controls follow the shared theme while view settings stay intact', async () => {
  const footer = page.locator('.document-statusbar')
  const light = await footer.evaluate(element => getComputedStyle(element).backgroundColor)
  const original = await page.evaluate(() => window.documentZoomFixture.snapshot())
  await page.evaluate(() => window.documentZoomFixture.setDark(true))
  await expect.poll(() => footer.evaluate(element => getComputedStyle(element).backgroundColor)).not.toBe(light)
  await expect(page.getByRole('button', { name: '放大', exact: true })).toBeEnabled()
  assert.deepEqual(await page.evaluate(() => window.documentZoomFixture.snapshot()), original)
  assert.deepEqual(errors, [])
})
