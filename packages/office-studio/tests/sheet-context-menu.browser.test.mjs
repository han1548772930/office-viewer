import assert from 'node:assert/strict'
import { after, before, beforeEach, test } from 'node:test'
import { createServer } from 'vite'
import vue from '@vitejs/plugin-vue'
import { chromium } from '@playwright/test'

let server
let browser
let page
let baseURL
const pageErrors = []
const fixture = `
import { createApp, h, ref } from 'vue'
import SheetContextMenu from '/packages/office-studio/src/components/spreadsheet/SheetContextMenu.vue'
const open = ref(false)
const hit = ref({ kind: 'cell', row: 2, column: 1, key: 'B3' })
const invocations = ref([])
window.contextFixture = { open, hit, invocations }
createApp({ render() { return h('main', [
  h('button', { id: 'show', onClick: () => { open.value = true } }, 'Open menu'),
  h(SheetContextMenu, { open: open.value, hit: hit.value, x: 48, y: 64, canDeleteRow: true, canDeleteColumn: true, frozen: false, 'onUpdate:open': value => { open.value = value }, onCommand: (id, params) => { invocations.value.push({ id, params }) } }),
  h('output', { id: 'commands' }, JSON.stringify(invocations.value)),
]) } }).mount('#app')
`

before(async () => {
  server = await createServer({
    configFile: false,
    root: process.cwd(),
    cacheDir: 'node_modules/.vite-sheet-context-contract',
    optimizeDeps: { noDiscovery: true, include: ['vue'] },
    plugins: [vue(), {
      name: 'sheet-context-menu-fixture',
      resolveId(id) { if (id === '/__sheet_context_fixture.ts') return '\0sheet-context-fixture' },
      load(id) { if (id === '\0sheet-context-fixture') return fixture },
      configureServer(vite) {
        vite.middlewares.use('/__sheet_context_menu__/', (_req, res) => {
          res.setHeader('Content-Type', 'text/html; charset=utf-8')
          res.end('<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"></head><body><div id="app"></div><script type="module" src="/__sheet_context_fixture.ts"></script></body></html>')
        })
      },
    }],
    appType: 'custom', logLevel: 'error', server: { host: '127.0.0.1', port: 0 },
  })
  await server.listen()
  const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate an HTTP port')
  baseURL = `http://127.0.0.1:${address.port}`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, { timeout: 60_000 })

after(async () => { await page?.close(); await browser?.close(); await server?.close() })

beforeEach(async () => {
  page = await browser.newPage({ viewport: { width: 420, height: 280 } })
  pageErrors.length = 0
  page.on('pageerror', error => pageErrors.push(error.message))
  const response = await page.goto(`${baseURL}/__sheet_context_menu__/`)
  await page.getByRole('button', { name: 'Open menu' }).waitFor({ timeout: 5000 }).catch(async error => {
    throw new Error(`HTTP ${response?.status()} ${response?.url()}\n${error.message}\n${pageErrors.join('\n')}\n${(await page.content()).slice(0, 1500)}`)
  })
})

test('menu renders executable actions, supports keyboard navigation and closes after dispatch', async () => {
  await page.getByRole('button', { name: 'Open menu' }).click()
  const menu = page.getByRole('menu', { name: '单元格上下文菜单' })
  await menu.waitFor()
  assert.equal(await menu.getByRole('menuitem').count(), 17)
  assert.equal(await menu.getByRole('menuitem', { name: '取消冻结' }).isDisabled(), true)
  await page.keyboard.press('ArrowDown')
  await page.keyboard.press('Enter')
  assert.equal(await menu.count(), 0)
  assert.match(await page.locator('#commands').textContent(), /"id":"sheet.command.copy"/)
})

test('disabled constraints remain disabled and Escape closes without dispatch', async () => {
  await page.evaluate(() => {
    window.contextFixture.hit.value = { kind: 'row', row: 4 }
    window.contextFixture.open.value = true
  })
  const menu = page.getByRole('menu')
  await menu.waitFor()
  assert.equal(await menu.getByRole('menuitem', { name: '删除行' }).isDisabled(), false)
  await page.keyboard.press('Escape')
  assert.equal(await menu.count(), 0)
  assert.equal(await page.locator('#commands').textContent(), '[]')
})
