import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import { createServer } from 'vite'
import vue from '@vitejs/plugin-vue'

// Unexecuted real Canvas/Ant/Facade/history drafts. Synthetic FontFaceSet
// notifications verify lifecycle wiring; they do not claim real font/IME coverage.
const fixturePath = '/__document-toc-pagination__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><style>body{margin:0}#app{height:100vh}</style></head><body><div id="app"></div><script type="module">
import { createApp, h, KeepAlive, ref, shallowRef, watch } from 'vue'
import { ConfigProvider, theme } from 'antdv-next'
import { OfficeStudio } from '/packages/office-studio/src/api/office-studio'
import { FDocument } from '/packages/office-studio/src/api/facade'
import { waitForDocumentTocPagination } from '/packages/office-studio/src/api/document-toc-pagination'
import { createBlankDocument } from '/packages/office-studio/src/domain/model'
import { createParagraph, runText } from '/packages/office-studio/src/features/document/model/document-model'
import { documentTextFieldRanges } from '/packages/office-studio/src/features/document/model/document-text-fields'
import { documentTocPageNumbers } from '/packages/office-studio/src/features/document/toc/document-toc'
import DocumentWorkbench from '/packages/office-studio/src/components/DocumentWorkbench.vue'
const initial = createBlankDocument(); initial.id = 'toc-canvas'; initial.name = '目录重排'
initial.page = { ...initial.page, width: 420, height: 380, marginTop: 30, marginBottom: 30, marginLeft: 30, marginRight: 30 }
initial.blocks = [{ ...createParagraph('正文输入位置'), id: 'body' }, ...Array.from({ length: 12 }, (_, index) => ({
  ...createParagraph('章节标题 ' + (index + 1)), id: 'heading-' + index, type: 'heading', level: index % 3 + 1,
}))]
const studio = new OfficeStudio({ document: initial, currentUserId: 'editor' }); studio.activateDocument(initial.id)
const api = new FDocument(studio, initial.id), model = shallowRef(api.getModel()), workbench = ref(), dark = ref(false), variant = ref('classic'), visible = ref(true)
const stop = watch(() => studio.editor.document.model, value => { model.value = value }, { flush: 'sync' })
api.setInputSelection({ anchor: { blockId: 'body', offset: 0 }, focus: { blockId: 'body', offset: 0 } })
const state = () => workbench.value?.$.setupState
window.tocFixture = {
  ready: () => state()?.tocLayoutPublication?.complete === true && state().tocLayoutPublication.source === state().model,
  layout: () => { const publication = state()?.tocLayoutPublication; return publication && { complete: publication.complete, fontVersion: publication.fontVersion,
    generation: publication.layoutGeneration, labels: documentTocPageNumbers(publication.layout.headings, publication.layout.pages) } },
  fields: () => documentTextFieldRanges(api.getModel()).filter(range => range.field.field === 'PAGEREF').map(range => ({
    headingId: range.block.metadata.documentToc.headingId, page: range.field.cachedResult })),
  history: () => ({ undo: studio.editor.history.units[api.id]?.undo.length ?? 0, redo: studio.editor.history.units[api.id]?.redo.length ?? 0 }),
  settle: () => waitForDocumentTocPagination(studio, api.id),
  undo: () => api.undo(), redo: () => api.redo(), show: value => { visible.value = value },
  appearance: (isDark, layout) => { dark.value = isDark; variant.value = layout },
  insert: () => studio.executeCommand('doc.command.insert-table-of-contents', { unitId: api.id, title: '目录' }),
  openDialog: () => api.openTableOfContentsDialog(),
  tocText: () => api.getModel().blocks.filter(block => block.metadata.documentToc?.role === 'entry').map(block => runText(block.runs)),
  selectToc: () => { const block = api.getModel().blocks.find(block => block.metadata.documentToc?.role === 'entry');
    return api.setInputSelection({ anchor: { blockId: block.id, offset: 0 }, focus: { blockId: block.id, offset: 0 } }) },
  renameFirstHeading: () => { const selection = { anchor: { blockId: 'heading-0', offset: 0 }, focus: { blockId: 'heading-0', offset: 0 } };
    api.setInputSelection(selection); return api.executeCommand({ id: 'doc.command.insert-text', text: '新标题 ' }, selection) },
}
const app = createApp({ render: () => h(ConfigProvider, { theme: { algorithm: dark.value ? theme.darkAlgorithm : theme.defaultAlgorithm } }, {
  default: () => h(KeepAlive, null, { default: () => visible.value ? h(DocumentWorkbench, {
    ref: workbench, modelValue: model.value, documentApi: api, ribbonVariant: variant.value, canUndo: studio.editor.canUndo, canRedo: studio.editor.canRedo,
    'onUpdate:modelValue': value => { model.value = value }, onUndo: () => api.undo(), onRedo: () => api.redo(),
  }) : null }),
}) })
app.mount('#app')
window.addEventListener('pagehide', () => { stop(); app.unmount(); studio.dispose() }, { once: true })
</script></body></html>`

let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'toc-pagination-fixture', configureServer(vite) {
      vite.middlewares.use(fixturePath, async (_req, res, next) => {
        try { res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.end(await vite.transformIndexHtml(fixturePath, fixture)) }
        catch (error) { next(error) }
      })
    },
  }], server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen()
  const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate a fixture port')
  baseURL = 'http://127.0.0.1:' + address.port + fixturePath
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, { timeout: 60_000 })
afterAll(async () => { await browser?.close(); await server?.close() })
beforeEach(async () => {
  errors = []; page = await browser.newPage({ viewport: { width: 1440, height: 900 } })
  page.on('pageerror', error => errors.push(error.message))
  await page.goto(baseURL)
  await expect.poll(() => page.evaluate(() => window.tocFixture?.ready())).toBe(true)
}, { timeout: 60_000 })
afterEach(async () => { await page?.close() })

test.each(['classic', 'grid'])('the %s automatic TOC menu reflows real page labels in one undo entry', async variant => {
  await page.evaluate(value => window.tocFixture.appearance(value === 'grid', value), variant)
  await page.getByRole('tab', { name: '插入', exact: true }).click()
  await page.getByRole('button', { name: '插入目录', exact: true }).click()
  await page.getByRole('menuitem', { name: '自动目录', exact: true }).click()
  assert.equal((await page.evaluate(() => window.tocFixture.settle())).status, 'settled')
  await expect.poll(() => page.evaluate(() => {
    const { labels } = window.tocFixture.layout(), fields = window.tocFixture.fields()
    return fields.length === 12 && fields.every(field => field.page === String(labels[field.headingId]))
  })).toBe(true)
  assert.deepEqual(await page.evaluate(() => window.tocFixture.history()), { undo: 1, redo: 0 })
  const fields = await page.evaluate(() => window.tocFixture.fields())
  await page.evaluate(() => window.tocFixture.undo())
  await expect.poll(() => page.evaluate(() => window.tocFixture.fields().length)).toBe(0)
  await page.evaluate(() => window.tocFixture.redo())
  await expect.poll(() => page.evaluate(() => window.tocFixture.fields())).toEqual(fields)
  assert.deepEqual(errors, [])
})

test('font invalidation publishes pending state, clears fallback metrics and recomputes without content history', async () => {
  const initial = await page.evaluate(() => window.tocFixture.layout())
  const pending = await page.evaluate(() => {
    document.fonts.dispatchEvent(new Event('loading'))
    return window.tocFixture.layout()
  })
  assert.equal(pending.complete, false)
  assert.ok(pending.fontVersion > initial.fontVersion)
  await page.evaluate(() => document.fonts.dispatchEvent(new Event('loadingdone')))
  await expect.poll(() => page.evaluate(() => window.tocFixture.ready())).toBe(true)
  assert.ok((await page.evaluate(() => window.tocFixture.layout())).generation > initial.generation)
  assert.deepEqual(await page.evaluate(() => window.tocFixture.history()), { undo: 0, redo: 0 })
  assert.deepEqual(errors, [])
})

test.each(['classic', 'grid'])('the %s update dialog and remove command share the actual TOC model and isolated history', async variant => {
  await page.evaluate(value => window.tocFixture.appearance(value === 'grid', value), variant)
  assert.equal(await page.evaluate(() => window.tocFixture.insert()), true)
  await page.evaluate(() => window.tocFixture.renameFirstHeading())
  await expect.poll(() => page.evaluate(() => window.tocFixture.ready())).toBe(true)
  const before = await page.evaluate(() => window.tocFixture.history())
  await page.evaluate(() => window.tocFixture.selectToc())
  await page.getByRole('button', { name: '更新目录', exact: true }).click()
  const dialog = page.getByRole('dialog', { name: '更新目录', exact: true })
  await dialog.getByRole('combobox').click()
  await page.getByRole('option', { name: '更新整个目录', exact: true }).click()
  await dialog.getByRole('button', { name: '应用', exact: true }).click()
  await expect(dialog).toBeHidden()
  assert.equal((await page.evaluate(() => window.tocFixture.settle())).status, 'settled')
  await expect.poll(() => page.evaluate(() => window.tocFixture.tocText()[0]?.startsWith('新标题 '))).toBe(true)
  assert.deepEqual(await page.evaluate(() => window.tocFixture.history()), { undo: before.undo + 1, redo: 0 })
  const text = await page.evaluate(() => window.tocFixture.tocText())
  await page.evaluate(() => window.tocFixture.selectToc())
  await page.getByRole('button', { name: '删除目录', exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.tocFixture.fields().length)).toBe(0)
  assert.deepEqual(await page.evaluate(() => window.tocFixture.history()), { undo: before.undo + 2, redo: 0 })
  await page.evaluate(() => window.tocFixture.undo())
  await expect.poll(() => page.evaluate(() => window.tocFixture.tocText())).toEqual(text)
  await page.evaluate(() => window.tocFixture.redo())
  await expect.poll(() => page.evaluate(() => window.tocFixture.fields().length)).toBe(0)
  assert.deepEqual(errors, [])
})

test('the fixed public dialog operation opens the Ant dialog and releases its owner while the workbench is inactive', async () => {
  assert.equal(await page.evaluate(() => window.tocFixture.insert()), true)
  await page.evaluate(() => window.tocFixture.selectToc())
  const before = await page.evaluate(() => window.tocFixture.history())
  assert.equal(await page.evaluate(() => window.tocFixture.openDialog()), true)
  const dialog = page.getByRole('dialog', { name: '更新目录', exact: true })
  await expect(dialog).toBeVisible()
  await dialog.getByRole('button', { name: '取消', exact: true }).click()
  await expect(dialog).toBeHidden()
  assert.deepEqual(await page.evaluate(() => window.tocFixture.history()), before)
  await page.evaluate(() => window.tocFixture.show(false))
  assert.equal(await page.evaluate(() => window.tocFixture.openDialog()), false)
  await page.evaluate(() => window.tocFixture.show(true))
  await expect.poll(() => page.evaluate(() => window.tocFixture.ready())).toBe(true)
  await page.evaluate(() => window.tocFixture.selectToc())
  assert.equal(await page.evaluate(() => window.tocFixture.openDialog()), true)
  await expect(dialog).toBeVisible()
  assert.deepEqual(await page.evaluate(() => window.tocFixture.history()), before)
  assert.deepEqual(errors, [])
})
