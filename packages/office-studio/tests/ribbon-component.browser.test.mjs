import assert from 'node:assert/strict'
import { before, after, beforeEach, test } from 'node:test'
import { createServer } from 'vite'
import vue from '@vitejs/plugin-vue'
import { chromium } from '@playwright/test'

let server
let browser
let page
let baseURL
const errors = []

const fixture = `
import { createApp, computed, h, ref } from 'vue'
import { Bold, Undo2 } from 'lucide-vue-next'
import WorkbenchRibbon from '/packages/office-studio/src/platform/ribbon/WorkbenchRibbon.vue'
const active = ref('home')
const locked = ref(false)
const showContext = ref(true)
const menusVisible = ref(true)
const direction = ref('ltr')
const family = ref('Arial')
const bold = ref(false)
window.ribbonInvocations = []
window.ribbonFixture = { active, locked, showContext, menusVisible, direction, family, bold }
const tabs = computed(() => menusVisible.value ? [
  { id: 'home', title: '开始', order: 0, groups: [
    { id: 'history', title: '历史', items: [
      { id: 'undo', type: 'button', title: '撤销', icon: Undo2, commandId: 'history.command.undo', disabled: locked.value, disabledReason: '没有历史记录' },
    ] },
    { id: 'font', title: '字体', items: [
      { id: 'family', type: 'select', title: '字体', commandId: 'doc.command.font', value: family.value, width: 112, options: [{ label: 'Arial', value: 'Arial' }, { label: '宋体', value: '宋体' }, { label: '不可用字体', value: 'locked', disabled: true }] },
      { id: 'bold', type: 'button', title: '加粗', icon: Bold, activated: bold.value, commandId: 'doc.command.bold' },
      { id: 'colors', type: 'menu', title: '字体颜色', showLabel: true, children: [
        { id: 'red', type: 'button', title: '红色', commandId: 'doc.command.color', params: '#ef4444' },
        { id: 'blue', type: 'button', title: '蓝色', commandId: 'doc.command.color', params: '#3b82f6' },
      ] },
    ] },
    { id: 'paragraph', title: '段落', items: Array.from({ length: 5 }, (_, n) => ({ id: 'tool-' + n, type: 'button', title: '段落工具' + n, commandId: 'doc.command.tool', params: n, showLabel: true })) },
  ] },
  { id: 'hidden', title: '隐藏', hidden: true, groups: [{ id: 'hidden-group', items: [{ id: 'hidden-tool', type: 'button', title: '隐藏工具', commandId: 'hidden' }] }] },
  { id: 'empty', title: '空分类', groups: [] },
  { id: 'insert', title: '插入', order: 1, groups: [{ id: 'insert', items: [{ id: 'insert-table', type: 'button', title: '插入表格', commandId: 'doc.command.insert-table', showLabel: true }] }] },
  { id: 'context', title: '表格', order: 2, contextual: true, hidden: !showContext.value, groups: [{ id: 'table', items: [{ id: 'delete-table', type: 'button', title: '删除表格', commandId: 'doc.command.delete-table', showLabel: true }] }] },
] : [])
createApp({ render() { return h(WorkbenchRibbon, { tabs: tabs.value, modelValue: active.value, direction: direction.value, 'onUpdate:modelValue': (value) => active.value = value, onExecute: (id, params) => { window.ribbonInvocations.push([id, params]); if (id === 'doc.command.font') family.value = params; if (id === 'doc.command.bold') bold.value = !bold.value } }) } }).mount('#app')
`

before(async () => {
  server = await createServer({
    configFile: false,
    root: process.cwd(),
    plugins: [vue(), {
      name: 'ribbon-browser-fixture',
      resolveId(id) { if (id === '/__ribbon_fixture.ts') return '\0ribbon-fixture' },
      load(id) { if (id === '\0ribbon-fixture') return fixture },
      configureServer(vite) {
        vite.middlewares.use('/__ribbon_contract/', (_req, res) => {
          res.setHeader('Content-Type', 'text/html; charset=utf-8')
          res.end('<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><style>body{margin:0}#app{width:100%;min-width:0}</style></head><body><div id="app"></div><script type="module" src="/__ribbon_fixture.ts"></script></body></html>')
        })
      },
    }],
    appType: 'custom',
    logLevel: 'error',
    server: { host: '127.0.0.1', port: 0 },
  })
  await server.listen()
  const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate an HTTP port')
  baseURL = `http://127.0.0.1:${address.port}`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, { timeout: 60_000 })

after(async () => {
  await page?.close()
  await browser?.close()
  await server?.close()
})

beforeEach(async () => {
  await page?.close()
  page = await browser.newPage({ viewport: { width: 1200, height: 700 } })
  errors.length = 0
  page.on('pageerror', (error) => errors.push(error.message))
  await page.goto(`${baseURL}/__ribbon_contract/`)
  await page.getByRole('tab', { name: '开始', exact: true }).waitFor()
}, { timeout: 60_000 })

test('classic geometry, hidden categories and keyboard navigation use a single roving tab stop', async () => {
  assert.equal(await page.getByRole('tab').count(), 3)
  assert.equal(await page.locator('.workbench-ribbon-tabs').evaluate((element) => element.getBoundingClientRect().height), 36)
  assert.equal(await page.locator('.workbench-ribbon-toolbar').evaluate((element) => element.getBoundingClientRect().height), 40)
  assert.equal(await page.locator('[role="tab"][tabindex="0"]').count(), 1)
  await page.getByRole('tab', { name: '开始', exact: true }).focus()
  await page.keyboard.press('End')
  assert.equal(await page.getByRole('tab', { name: '表格', exact: true }).getAttribute('aria-selected'), 'true')
  await page.keyboard.press('ArrowRight')
  assert.equal(await page.getByRole('tab', { name: '开始', exact: true }).getAttribute('aria-selected'), 'true')
  await page.evaluate(() => { window.ribbonFixture.direction.value = 'rtl' })
  await page.keyboard.press('ArrowRight')
  assert.equal(await page.getByRole('tab', { name: '表格', exact: true }).getAttribute('aria-selected'), 'true')
  await page.keyboard.press('Home')
  assert.equal(await page.getByRole('tab', { name: '开始', exact: true }).getAttribute('aria-selected'), 'true')
  assert.deepEqual(errors, [])
})

test('selectors and activated/disabled buttons reflect editor state and dispatch real command parameters', async () => {
  const font = page.getByRole('combobox', { name: '字体', exact: true })
  assert.equal(await font.count(), 1, await page.getByRole('combobox').evaluateAll((elements) => elements.map((element) => element.outerHTML).join('\n')))
  await font.click()
  await page.getByTitle('宋体', { exact: true }).click()
  await page.waitForFunction(() => window.ribbonInvocations.some(([id, value]) => id === 'doc.command.font' && value === '宋体'))
  await page.getByRole('button', { name: '加粗', exact: true }).click()
  assert.equal(await page.getByRole('button', { name: '加粗', exact: true }).getAttribute('aria-pressed'), 'true')
  await page.evaluate(() => { window.ribbonFixture.locked.value = true })
  assert.equal(await page.getByRole('button', { name: '撤销', exact: true }).isDisabled(), true)
  assert.deepEqual(await page.evaluate(() => window.ribbonInvocations), [['doc.command.font', '宋体'], ['doc.command.bold', undefined]])
  assert.deepEqual(errors, [])
})

test('narrow toolbar exposes overflow controls, nested menus and selectors without document overflow', async () => {
  await page.setViewportSize({ width: 210, height: 700 })
  const more = page.getByRole('button', { name: '更多功能', exact: true })
  await more.waitFor()
  await more.focus()
  await page.keyboard.press('ArrowDown')
  const popup = page.getByRole('menu', { name: '更多功能', exact: true })
  await popup.waitFor()
  await popup.getByRole('button', { name: '段落工具4', exact: true }).click()
  await page.waitForFunction(() => window.ribbonInvocations.some(([id, value]) => id === 'doc.command.tool' && value === 4))
  await more.click()
  await popup.getByRole('button', { name: '字体颜色', exact: true }).click()
  const colors = page.getByRole('menu', { name: '字体颜色', exact: true })
  await colors.getByRole('button', { name: '红色', exact: true }).click()
  await page.waitForFunction(() => window.ribbonInvocations.some(([id, value]) => id === 'doc.command.color' && value === '#ef4444'))
  await more.click()
  await popup.getByRole('combobox', { name: '字体', exact: true }).click()
  await page.getByTitle('宋体', { exact: true }).click()
  await page.waitForFunction(() => window.ribbonInvocations.some(([id, value]) => id === 'doc.command.font' && value === '宋体'))
  await more.focus()
  await page.keyboard.press('ArrowDown')
  await page.keyboard.press('Escape')
  assert.equal(await more.evaluate((element) => element === document.activeElement), true)
  assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth), true)
  assert.deepEqual(errors, [])
})

test('context category removal returns to the previous regular tab without dispatching content edits', async () => {
  await page.getByRole('tab', { name: '插入', exact: true }).click()
  await page.getByRole('tab', { name: '表格', exact: true }).click()
  await page.evaluate(() => { window.ribbonFixture.showContext.value = false })
  await page.waitForFunction(() => window.ribbonFixture.active.value === 'insert')
  assert.equal(await page.getByRole('tab', { name: '插入', exact: true }).getAttribute('aria-selected'), 'true')
  assert.deepEqual(await page.evaluate(() => window.ribbonInvocations), [])
  assert.deepEqual(errors, [])
})

test('resize tracking survives temporarily empty menus and does not reopen stale overflow popups', async () => {
  await page.evaluate(() => { window.ribbonFixture.menusVisible.value = false })
  await page.getByTestId('workbench-ribbon').waitFor({ state: 'detached' })
  await page.evaluate(() => { window.ribbonFixture.menusVisible.value = true })
  await page.getByRole('tab', { name: '开始', exact: true }).waitFor()
  await page.setViewportSize({ width: 210, height: 700 })
  const more = page.getByRole('button', { name: '更多功能', exact: true })
  await more.click()
  await page.getByRole('menu', { name: '更多功能', exact: true }).waitFor()
  await page.setViewportSize({ width: 1200, height: 700 })
  await more.waitFor({ state: 'detached' })
  await page.setViewportSize({ width: 210, height: 700 })
  await more.waitFor()
  assert.equal(await more.getAttribute('aria-expanded'), 'false')
  assert.equal(await page.getByRole('menu', { name: '更多功能', exact: true }).count(), 0)
  assert.deepEqual(errors, [])
})
