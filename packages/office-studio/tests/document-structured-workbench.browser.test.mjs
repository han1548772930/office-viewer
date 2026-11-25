import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

// Deferred browser drafts mount the actual Canvas workbench and antdv-next dialogs.
// Synthetic composition guards are not acceptance for the real Windows IME.
const fixturePath = '/__document-structured-workbench-test__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>Docs structured blocks</title><style>body{margin:0}#app{height:100vh}</style></head><body><div id="app"></div><script type="module">
import { createApp, h, ref, shallowRef } from 'vue'
import { ConfigProvider, theme } from 'antdv-next'
import { OfficeStudio } from '/packages/office-studio/src/api/office-studio'
import { FDocument } from '/packages/office-studio/src/api/facade'
import { createBlankDocument } from '/packages/office-studio/src/domain/model'
import { createParagraph, findDocumentBlock } from '/packages/office-studio/src/features/document/model/document-model'
import { createTextMeasurer, layoutDocument } from '/packages/office-studio/src/features/document/layout/document-layout'
import { sceneOffset } from '/packages/office-studio/src/features/document/rendering/document-scene'
import { createOfficeI18n } from '/packages/office-studio/src/i18n/office-i18n'
import DocumentWorkbench from '/packages/office-studio/src/components/DocumentWorkbench.vue'
const copy = value => JSON.parse(JSON.stringify(value)), standalone = new URLSearchParams(location.search).has('standalone')
const paragraph = (id, text = '') => ({ ...createParagraph(text), id })
const initial = createBlankDocument(); initial.id = 'structured-browser'; initial.name = '结构块交互草稿'
initial.page = { width: 620, height: 540, marginTop: 42, marginBottom: 42, marginLeft: 48, marginRight: 48 }
initial.blocks = [{ ...paragraph('code'), type: 'code', runs: [], children: [paragraph('code-p', 'a😀b'), paragraph('code-last', '  c')],
  code: { language: 'typescript', showLineNumbers: true, tabSize: 4, wrap: true } },
  { ...paragraph('quote'), type: 'quote', runs: [], children: [paragraph('quote-p', 'Quoted text')] }, paragraph('tail', 'Outside')]
if (standalone) delete initial.id
let studio = standalone ? undefined : new OfficeStudio({ document: initial, currentUserId: 'editor' })
studio?.activateDocument(initial.id)
const api = shallowRef(studio ? new FDocument(studio, initial.id) : undefined), model = shallowRef(api.value?.getModel() ?? initial)
const dark = ref(false), variant = ref('classic'), notices = []
function sync() { if (api.value) model.value = api.value.getModel() }
function select(blockId, offset = 0, endId = blockId, endOffset = offset) {
  if (!api.value) return false
  return api.value.setInputSelection({ anchor: { blockId, offset }, focus: { blockId: endId, offset: endOffset } })
}
window.documentStructuredFixture = {
  snapshot: () => copy(api.value?.getModel() ?? model.value), notices,
  source: () => findDocumentBlock(api.value?.getModel() ?? model.value, 'code').children.map(block => block.runs.map(run => run.text).join('')).join('\\n'),
  history: () => studio ? { undo: studio.editor.history.units[api.value.id]?.undo.length ?? 0, redo: studio.editor.history.units[api.value.id]?.redo.length ?? 0 } : undefined,
  select, dark: value => dark.value = value, variant: value => variant.value = value,
  changeSource: value => { api.value.getCode('code').setText(value); sync() },
  replaceSameSnapshot: () => { model.value = copy(model.value) },
  saveAndReopen: () => { const saved = api.value.save(); studio.dispose(); studio = new OfficeStudio({ document: copy(saved), currentUserId: 'editor' }); studio.activateDocument(saved.id); api.value = new FDocument(studio, saved.id); sync(); select('code-p', 0) },
  point: (id, offset = 0) => {
    const viewport = document.querySelector('.document-scroll'), rect = viewport.getBoundingClientRect()
    const layout = layoutDocument(api.value?.getModel() ?? model.value, createTextMeasurer())
    const line = layout.lines.find(line => line.target.blockId === id && offset >= line.start && offset <= line.end)
    if (!line) throw new Error('Paragraph does not have current Canvas geometry')
    const glyph = line.glyphs.find(glyph => glyph.offset >= offset), shift = sceneOffset(layout, { width: rect.width, height: rect.height, scrollLeft: viewport.scrollLeft, scrollTop: viewport.scrollTop, zoom: 100 })
    return { x: rect.x + shift.x + (glyph?.x ?? line.x + line.width), y: rect.y + shift.y + line.y + line.height / 2 }
  },
}
const app = createApp({ render: () => h(ConfigProvider, { theme: { algorithm: dark.value ? theme.darkAlgorithm : theme.defaultAlgorithm } }, {
  default: () => h(DocumentWorkbench, { modelValue: model.value, documentApi: api.value, canUndo: studio?.editor.canUndo, canRedo: studio?.editor.canRedo, ribbonVariant: variant.value,
    'onUpdate:modelValue': next => model.value = next, onNotify: message => notices.push(message) }),
}) })
app.use(createOfficeI18n()); app.mount('#app')
window.addEventListener('beforeunload', () => { app.unmount(); studio?.dispose() }, { once: true })
</script></body></html>`

let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'document-structured-workbench-fixture',
    configureServer(vite) { vite.middlewares.use(fixturePath, async (_req, res, next) => {
      try { res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.end(await vite.transformIndexHtml(fixturePath, fixture)) } catch (error) { next(error) }
    }) },
  }], server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen()
  const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate an HTTP port')
  baseURL = 'http://127.0.0.1:' + address.port + fixturePath
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, { timeout: 60_000 })
afterAll(async () => { await browser?.close(); await server?.close() })
beforeEach(async () => {
  errors = []; page = await browser.newPage({ viewport: { width: 1440, height: 900 }, deviceScaleFactor: 1.25 })
  page.on('pageerror', error => errors.push(error.message)); await page.goto(baseURL)
  await page.waitForFunction(() => !!window.documentStructuredFixture)
  await expect(page.getByRole('textbox', { name: '文档文字输入', exact: true })).toBeAttached()
})
afterEach(async () => { await page?.close(); assert.deepEqual(errors, []) })
const source = () => page.evaluate(() => window.documentStructuredFixture.source())
const history = () => page.evaluate(() => window.documentStructuredFixture.history())
async function contextProperty(blockId, name) {
  const point = await page.evaluate(id => window.documentStructuredFixture.point(id), blockId)
  await page.mouse.click(point.x + 2, point.y, { button: 'right' })
  await page.getByRole('menuitem', { name, exact: true }).click()
  return page.getByRole('dialog')
}

test('code Tab and Shift+Tab use the current range and one undo item for each indentation action', async () => {
  await page.evaluate(() => window.documentStructuredFixture.select('code-last', 3, 'code-p', 1))
  await page.getByRole('textbox', { name: '文档文字输入', exact: true }).focus()
  await page.keyboard.press('Tab')
  await expect.poll(source).toBe('\ta😀b\n\t  c'); assert.equal((await history()).undo, 1)
  await page.keyboard.press('Shift+Tab')
  await expect.poll(source).toBe('a😀b\n  c'); assert.equal((await history()).undo, 2)
})

test('code property dialog rejects a stale revision and a reopened document API', async () => {
  for (const action of ['change', 'reopen']) {
    const dialog = await contextProperty('code-p', '代码块属性')
    await dialog.locator('textarea').fill('local draft')
    await page.evaluate(action => action === 'change' ? window.documentStructuredFixture.changeSource('remote content') : window.documentStructuredFixture.saveAndReopen(), action)
    const before = await source(), previousHistory = await history()
    await dialog.getByRole('button', { name: '应用', exact: true }).click()
    assert.equal(await source(), before); assert.deepEqual(await history(), previousHistory)
    await expect(dialog).toBeVisible()
    await dialog.getByRole('button', { name: '取消', exact: true }).click(); await expect(dialog).toBeHidden()
  }
})

test('standalone code settings accept a missing ID and reject replacement by an identical parent snapshot', async () => {
  await page.goto(baseURL + '?standalone=1'); await page.waitForFunction(() => !!window.documentStructuredFixture)
  let dialog = await contextProperty('code-p', '代码块属性')
  await dialog.locator('textarea').fill('accepted draft')
  await dialog.getByRole('button', { name: '应用', exact: true }).click(); await expect(dialog).toBeHidden()
  assert.equal(await source(), 'accepted draft')
  dialog = await contextProperty('code-p', '代码块属性')
  await dialog.locator('textarea').fill('stale draft')
  await page.evaluate(() => window.documentStructuredFixture.replaceSameSnapshot())
  await dialog.getByRole('button', { name: '应用', exact: true }).click()
  assert.equal(await source(), 'accepted draft'); await expect(dialog).toBeVisible()
})

test('structured properties keep ant controls in both themes and return focus after applying a current edit', async () => {
  for (const dark of [false, true]) {
    await page.evaluate(dark => window.documentStructuredFixture.dark(dark), dark)
    const dialog = await contextProperty('code-p', '代码块属性')
    await expect(dialog.locator('.ant-input-number')).toHaveCount(1)
    await expect(dialog.getByRole('switch')).toHaveCount(2)
    await dialog.locator('textarea').fill(dark ? 'dark edit' : 'light edit')
    await dialog.getByRole('button', { name: '应用', exact: true }).click(); await expect(dialog).toBeHidden()
    await expect(page.getByRole('textbox', { name: '文档文字输入', exact: true })).toBeFocused()
    assert.equal(await source(), dark ? 'dark edit' : 'light edit')
  }
})

test('synthetic composition in the editor blocks a property commit and preserves the current code', async () => {
  const dialog = await contextProperty('code-p', '代码块属性'), before = await source()
  await dialog.locator('textarea').fill('must not commit')
  const input = page.getByRole('textbox', { name: '文档文字输入', exact: true })
  await input.dispatchEvent('compositionstart', { data: '中' })
  await dialog.getByRole('button', { name: '应用', exact: true }).click()
  assert.equal(await source(), before); await expect(dialog).toBeVisible()
  await input.dispatchEvent('compositionend', { data: '' })
})
