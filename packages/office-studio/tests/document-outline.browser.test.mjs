import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

const fixturePath = '/__document-outline-test__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>Document outline interaction fixture</title></head><body>
<div id="app"></div><script type="module">
import { createApp, h, reactive } from 'vue'
import { Button, ConfigProvider, theme } from 'antdv-next'
import DocumentOutline from '/packages/office-studio/src/components/document/DocumentOutline.vue'
const state = reactive({
  model: { kind: 'document', name: '大纲交互文档', blocks: Array.from({ length: 24 }, (_, index) => ({
    id: 'heading-' + (index + 1), type: 'heading', level: index % 3 + 1,
    runs: [{ id: 'run-' + index, text: '标题 ' + (index + 1) }], style: {}, metadata: {},
  })), page: {}, modified: false, wordCount: 0 },
  activeId: 'heading-1', mode: 'side-bar', width: 256, dark: false, disabled: false,
})
const navigations = []
window.outlineFixture = {
  snapshot: () => ({ model: JSON.stringify(state.model), navigations: [...navigations] }),
  activate: (id) => { state.activeId = id },
  mode: (value) => { state.mode = value },
  rename: (id, text) => { state.model.blocks.find((block) => block.id === id).runs[0].text = text },
  remove: (id) => { state.model.blocks = state.model.blocks.filter((block) => block.id !== id) },
  width: (value) => { state.width = value },
  appearance: (value) => { state.dark = value },
  disabled: (value) => { state.disabled = value },
  title: () => { state.model.blocks.unshift({ id: 'real-title', type: 'paragraph', level: 0,
    runs: [{ id: 'title-run', text: '作者文档标题' }], style: {}, metadata: { namedStyleType: 'title' } }) },
}
createApp({ render: () => h(ConfigProvider, { theme: { algorithm: state.dark ? theme.darkAlgorithm : theme.defaultAlgorithm } }, { default: () =>
 h('main', { style: { position: 'relative', display: 'flex', width: '960px', height: '280px' } }, [
  h(DocumentOutline, { model: state.model, activeId: state.activeId, mode: state.mode, maxHeight: 240, maxWidth: state.width, disabled: state.disabled,
    onNavigate: (item) => { navigations.push(item.id); state.activeId = item.id },
  }),
  h(Button, { id: 'outside', style: { position: 'absolute', left: '600px', top: '30px' } }, { default: () => '外部按钮' }),
]) }) }).mount('#app')
</script></body></html>`

let server
let browser
let page
let baseURL

beforeAll(async () => {
  server = await createServer({
    configFile: false,
    root: process.cwd(),
    publicDir: false,
    logLevel: 'error',
    plugins: [vue(), {
      name: 'document-outline-interaction-fixture',
      configureServer(vite) {
        vite.middlewares.use(fixturePath, async (_request, response, next) => {
          try {
            const html = await vite.transformIndexHtml(fixturePath, fixture)
            response.statusCode = 200
            response.setHeader('Content-Type', 'text/html; charset=utf-8')
            response.end(html)
          } catch (error) {
            next(error)
          }
        })
      },
    }],
    server: { host: '127.0.0.1', port: 0, strictPort: false },
  })
  await server.listen()
  const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('The outline fixture did not bind a TCP port.')
  baseURL = `http://127.0.0.1:${address.port}${fixturePath}`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
})

afterAll(async () => {
  await browser?.close()
  await server?.close()
})

beforeEach(async () => {
  page = await browser.newPage({ viewport: { width: 1100, height: 600 } })
  await page.goto(baseURL)
  await expect(page.getByRole('navigation', { name: '文档大纲' })).toBeVisible()
})

afterEach(async () => { await page?.close() })

test('outline keyboard focus is separate from activation, and Escape collapses and restores focus', async () => {
  const initial = await page.evaluate(() => window.outlineFixture.snapshot())
  const first = page.locator('[data-outline-item="heading-1"]')
  const second = page.locator('[data-outline-item="heading-2"]')
  await first.focus()
  await first.press('ArrowDown')
  await expect(second).toBeFocused()
  assert.deepEqual((await page.evaluate(() => window.outlineFixture.snapshot())).navigations, [])
  await second.press('Enter')
  await expect(second).toHaveAttribute('aria-current', 'location')
  assert.deepEqual((await page.evaluate(() => window.outlineFixture.snapshot())).navigations, ['heading-2'])
  await second.press('End')
  await expect(page.locator('[data-outline-item="heading-24"]')).toBeFocused()
  await page.locator('[data-outline-item="heading-24"]').press('Home')
  await expect(page.locator('[data-outline-item]').first()).toBeFocused()
  await page.locator('[data-outline-item]').first().press('Escape')
  const toggle = page.getByRole('button', { name: '文档大纲', exact: true })
  await expect(toggle).toBeFocused()
  await expect(toggle).toHaveAttribute('aria-expanded', 'false')
  await expect(page.getByRole('navigation', { name: '文档大纲' })).toHaveCount(0)
  await toggle.press('Enter')
  await expect(page.getByRole('navigation', { name: '文档大纲' })).toBeVisible()
  assert.equal((await page.evaluate(() => window.outlineFixture.snapshot())).model, initial.model)
})

test('floating outline starts collapsed and closes on outside interaction or heading activation', async () => {
  await page.evaluate(() => window.outlineFixture.mode('float'))
  const toggle = page.getByRole('button', { name: '文档大纲', exact: true })
  await expect(toggle).toHaveAttribute('aria-expanded', 'false')
  await toggle.click()
  await expect(page.getByRole('navigation', { name: '文档大纲' })).toBeVisible()
  await page.getByRole('button', { name: '外部按钮' }).click()
  await expect(toggle).toHaveAttribute('aria-expanded', 'false')
  await toggle.click()
  await page.locator('[data-outline-item="heading-3"]').click()
  await expect(toggle).toHaveAttribute('aria-expanded', 'false')
  await expect(toggle).toBeFocused()
  assert.deepEqual((await page.evaluate(() => window.outlineFixture.snapshot())).navigations, ['heading-3'])
})

test('active headings scroll into view without stealing focus and update by stable ID after edits', async () => {
  const outside = page.getByRole('button', { name: '外部按钮' })
  await outside.focus()
  await page.evaluate(() => window.outlineFixture.activate('heading-24'))
  const last = page.locator('[data-outline-item="heading-24"]')
  await expect(last).toHaveAttribute('aria-current', 'location')
  await expect.poll(() => page.evaluate(() => {
    const container = document.querySelector('.document-outline__navigation').getBoundingClientRect()
    const item = document.querySelector('[data-outline-item="heading-24"]').getBoundingClientRect()
    return item.top >= container.top && item.bottom <= container.bottom + 1
  })).toBe(true)
  await expect(outside).toBeFocused()
  await page.evaluate(() => window.outlineFixture.rename('heading-24', '重命名的章节'))
  await expect(last).toHaveText('重命名的章节')
  await page.evaluate(() => window.outlineFixture.remove('heading-24'))
  await expect(last).toHaveCount(0)
  await expect(page.locator('[data-outline-item]')).toHaveCount(24)
  await page.evaluate(() => { window.outlineFixture.activate('heading-1'); window.outlineFixture.width(160) })
  await expect.poll(() => page.locator('.document-outline').evaluate((element) => element.getBoundingClientRect().width)).toBe(160)
  await expect(outside).toBeFocused()
})

test('authored Title is navigable and replaces the virtual title after a live model update', async () => {
  await page.evaluate(() => window.outlineFixture.title())
  const title = page.locator('[data-outline-item="real-title"]')
  await expect(title).toHaveText('作者文档标题')
  await expect(page.locator('.document-outline__item.is-title')).toHaveCount(1)
  await title.click()
  await expect(title).toHaveAttribute('aria-current', 'location')
  assert.deepEqual((await page.evaluate(() => window.outlineFixture.snapshot())).navigations, ['real-title'])
})

test('the real Ant controls follow ConfigProvider themes and disabled navigation cannot emit a new target', async () => {
  const panel = page.locator('.document-outline'), first = page.locator('[data-outline-item="heading-1"]')
  const before = await panel.evaluate(element => getComputedStyle(element).backgroundColor)
  await page.evaluate(() => window.outlineFixture.appearance(true))
  await expect.poll(() => panel.evaluate(element => getComputedStyle(element).backgroundColor)).not.toBe(before)
  await expect(first).toHaveClass(/ant-btn/)
  await first.focus()
  await page.evaluate(() => window.outlineFixture.disabled(true))
  await expect(first).toBeDisabled()
  const color = await first.evaluate(element => getComputedStyle(element).color)
  await first.hover({ force: true })
  assert.equal(await first.evaluate(element => getComputedStyle(element).color), color)
  await first.dispatchEvent('click')
  assert.deepEqual((await page.evaluate(() => window.outlineFixture.snapshot())).navigations, [])
  await page.evaluate(() => { window.outlineFixture.disabled(false); window.outlineFixture.appearance(false) })
  await expect.poll(() => panel.evaluate(element => getComputedStyle(element).backgroundColor)).toBe(before)
  await first.focus(); await first.press('Enter')
  assert.deepEqual((await page.evaluate(() => window.outlineFixture.snapshot())).navigations, ['heading-1'])
})
