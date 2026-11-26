import assert from 'node:assert/strict'
import { afterAll, beforeAll, beforeEach, test } from 'vitest'
import { createServer } from 'vite'
import vue from '@vitejs/plugin-vue'
import { chromium } from '@playwright/test'

let server
let browser
let page
let baseURL
const pageErrors = []
const flattenItems = items => items.flatMap(item => [item, ...flattenItems(item.children ?? [])])
const fixture = `
import { createApp, h, nextTick, ref } from 'vue'
import { Button, ConfigProvider } from 'antdv-next'
import SheetContextMenu from '/packages/office-studio/src/components/spreadsheet/SheetContextMenu.vue'
import SpreadsheetCanvas from '/packages/office-studio/src/components/spreadsheet/SpreadsheetCanvas.vue'
import { createBlankWorkbook } from '/packages/office-studio/src/domain/model.ts'
import { toExcelWorkbookView } from '/packages/office-studio/src/services/excelViewAdapter.ts'
const open = ref(false)
const hit = ref({ kind: 'cell', row: 2, column: 1, key: 'B3' })
const invocations = ref([])
window.contextFixture = { open, hit, invocations }
const canvasRef = ref()
const workbook = createBlankWorkbook()
workbook.id = 'canvas-context-book'; workbook.sheets[0].id = 'canvas-data'
workbook.sheets[0].rowCount = 20; workbook.sheets[0].columnCount = 8
const sheet = ref(toExcelWorkbookView(workbook).sheets[0])
sheet.value.cells = { A1: { value: '2', rawValue: 2 }, A2: { value: '4', rawValue: 4 }, B3: {
  value: '42', rawValue: '42', hyperlink: 'https://example.test',
  sparkline: { type: 'line', sourceRange: 'A1:A2', color: '#246fce', emptyCells: 'gap', axis: 'individual' },
} }
const selectedKey = ref('B3')
const selectedRanges = ref([{ startRow: 2, endRow: 2, startColumn: 1, endColumn: 1, rangeType: 0 }])
const activeRangeIndex = ref(0)
const owner = ref({ revision: 7, generation: 2, userId: 'local-user' })
const canvasEvents = ref([])
window.canvasContextFixture = { sheet, owner,
  search: () => canvasRef.value.featureSearchContext(),
  state: () => JSON.parse(JSON.stringify({ sheet: sheet.value, selectedKey: selectedKey.value, selectedRanges: selectedRanges.value, activeRangeIndex: activeRangeIndex.value,
    owner: owner.value, events: canvasEvents.value })),
  select: async (ranges, key, index = ranges.length - 1) => { selectedRanges.value = ranges; selectedKey.value = key; activeRangeIndex.value = index; await nextTick() },
}
createApp({ render() { return h(ConfigProvider, {}, { default: () => h('main', [
  h(Button, { id: 'show', onClick: () => { open.value = true } }, { default: () => 'Open menu' }),
  h(SheetContextMenu, { open: open.value, hit: hit.value, sheetId: 'sheet-a', x: 48, y: 64, canDeleteRow: true, canDeleteColumn: true, frozen: false, hasHiddenRows: false, hasHiddenColumns: false, 'onUpdate:open': value => { open.value = value }, onCommand: (id, params) => { invocations.value.push({ id, params }) } }),
  h('output', { id: 'commands' }, JSON.stringify(invocations.value)),
  h('div', { style: { height: '220px', marginTop: '8px' } }, [h(SpreadsheetCanvas, { ref: canvasRef,
    workbookId: workbook.id, sheet: sheet.value, selectedKey: selectedKey.value, selectedRange: [selectedKey.value],
    selectedRanges: selectedRanges.value, activeRangeIndex: activeRangeIndex.value, editingKey: null, editValue: '', zoom: 100, canFormat: true,
    currentUserId: owner.value.userId, workbookRevision: owner.value.revision, workbookHistoryGeneration: owner.value.generation,
    onContextCommand: (id, params) => canvasEvents.value.push({ id, params }),
  })]),
]) }) } }).mount('#app')
`

beforeAll(async () => {
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

afterAll(async () => { await page?.close(); await browser?.close(); await server?.close() })

beforeEach(async () => {
  page = await browser.newPage({ viewport: { width: 420, height: 280 } })
  pageErrors.length = 0
  page.on('pageerror', error => pageErrors.push(error.message))
  await page.addInitScript(() => Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined }))
  const response = await page.goto(`${baseURL}/__sheet_context_menu__/`)
  await page.getByRole('button', { name: 'Open menu' }).waitFor({ timeout: 5000 }).catch(async error => {
    throw new Error(`HTTP ${response?.status()} ${response?.url()}\n${error.message}\n${pageErrors.join('\n')}\n${(await page.content()).slice(0, 1500)}`)
  })
})

test('menu renders quick actions, supports submenu navigation and dispatches commands', async () => {
  await page.getByRole('button', { name: 'Open menu' }).click()
  const menu = page.getByRole('menu', { name: '单元格上下文菜单' })
  await menu.waitFor()
  assert.equal(await page.getByRole('group', { name: '快捷操作' }).getByRole('button').count(), 3)
  await page.getByRole('button', { name: '复制', exact: true }).click()
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
  await menu.getByRole('menuitem', { name: '冻结' }).hover()
  assert.equal(await page.getByRole('menuitem', { name: '取消冻结' }).isDisabled(), true)
  await page.keyboard.press('Escape')
  assert.equal(await menu.count(), 0)
  assert.equal(await page.locator('#commands').textContent(), '[]')
})

test('submenu leaf dispatches its mapped command and row menu uses row-specific targets', async () => {
  await page.getByRole('button', { name: 'Open menu' }).click()
  await page.getByRole('menuitem', { name: '选择性复制' }).hover()
  await page.getByRole('menuitem', { name: '仅复制公式' }).click()
  assert.match(await page.locator('#commands').textContent(), /"id":"sheet.command.copy-formulas"/)

  await page.evaluate(() => {
    window.contextFixture.hit.value = { kind: 'row', row: 4 }
    window.contextFixture.open.value = true
  })
  await page.getByRole('menuitem', { name: '在上方插入 1 行' }).click()
  const invocations = JSON.parse(await page.locator('#commands').textContent())
  assert.deepEqual(invocations.at(-1).params, { sheetId: 'sheet-a', ranges: [{ startRow: 4, endRow: 4, startCol: 0, endCol: 16_383 }] })
})

// Deferred real Canvas getter drafts. Synthetic data does not accept Windows native IME, WebGPU or performance.
test('Canvas feature-search getter exposes cell context without opening a menu or changing selection, content or focus', async () => {
  const grid = page.getByRole('grid', { name: '工作表画布' })
  await grid.waitFor(); await grid.focus()
  const before = await page.evaluate(() => window.canvasContextFixture.state())
  const groups = await page.evaluate(() => window.canvasContextFixture.search())
  const items = flattenItems(groups[0].items)
  assert.ok(items.some(item => item.commandId === 'sheet.command.text-to-number'))
  assert.deepEqual(items.find(item => item.commandId === 'sheet.operation.open-cell-hyperlink').params, { key: 'B3' })
  assert.equal(items.find(item => item.commandId === 'sheet.operation.open-sparkline-panel').disabled, false)
  assert.equal(await page.getByRole('menu').count(), 0)
  assert.deepEqual(await page.evaluate(() => window.canvasContextFixture.state()), before)
  assert.equal(await grid.evaluate(element => element === document.activeElement), true)
  assert.deepEqual(pageErrors, [])
})

test('Canvas feature-search getter follows axis selection, hidden ranges, permissions and current outline ownership', async () => {
  await page.getByRole('grid', { name: '工作表画布' }).waitFor()
  await page.evaluate(async () => {
    window.canvasContextFixture.sheet.value.hiddenRows = [2]
    await window.canvasContextFixture.select([{ startRow: 1, endRow: 3, startColumn: 0, endColumn: 7, rangeType: 1 }], 'A2')
  })
  const rows = await page.evaluate(() => window.canvasContextFixture.search()[0].items)
  assert.ok(rows.some(item => item.commandId === 'sheet.command.set-row-height'))
  assert.equal(rows.find(item => item.id === 'show-hidden-rows').disabled, false)
  assert.equal(rows.some(item => item.commandId === 'sheet.command.set-worksheet-col-width'), false)
  const groupRows = rows.find(item => item.id === 'sheet.contextMenu.dimension-outline').children.find(item => item.commandId === 'sheet.command.outline-selected-rows')
  assert.equal(groupRows.params.expectedRevision, 7)
  assert.equal(groupRows.params.expectedHistoryGeneration, 2)
  assert.equal(groupRows.params.expectedUserId, 'local-user')
  assert.deepEqual(groupRows.params.ranges, [{ startRow: 1, endRow: 3, startCol: 0, endCol: 7, rangeType: 1 }])
  await page.evaluate(async () => {
    window.canvasContextFixture.sheet.value.hiddenColumns = [2]
    window.canvasContextFixture.sheet.value.rangeProtectionRules = [{ id: 'locked-column', permissionId: 'locked-column', ownerUserId: 'peer',
      allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView', ranges: [{ startRow: 0, endRow: 19, startCol: 2, endCol: 2 }] }]
    window.canvasContextFixture.owner.value = { revision: 8, generation: 3, userId: 'reader' }
    await window.canvasContextFixture.select([{ startRow: 0, endRow: 19, startColumn: 1, endColumn: 3, rangeType: 2 }], 'B1')
  })
  const columns = await page.evaluate(() => window.canvasContextFixture.search()[0].items)
  assert.ok(columns.some(item => item.commandId === 'sheet.command.set-worksheet-col-width'))
  assert.equal(columns.find(item => item.id === 'show-hidden-columns').disabled, true)
  assert.equal(columns.find(item => item.id === 'insert-column-left').disabled, true)
  const groupColumns = columns.find(item => item.id === 'sheet.contextMenu.dimension-outline').children.find(item => item.commandId === 'sheet.command.outline-selected-columns')
  assert.equal(groupColumns.disabled, true)
  assert.equal(groupColumns.params.expectedRevision, 8)
  assert.equal(groupColumns.params.expectedHistoryGeneration, 3)
  assert.equal(groupColumns.params.expectedUserId, 'reader')
  assert.equal(await page.getByRole('menu').count(), 0)
  assert.deepEqual(pageErrors, [])
})

test('Canvas feature-search uses an earlier active range in a mixed row and column selection', async () => {
  await page.getByRole('grid', { name: '工作表画布' }).waitFor()
  await page.evaluate(() => window.canvasContextFixture.select([
    { startRow: 1, endRow: 2, startColumn: 0, endColumn: 7, rangeType: 1 },
    { startRow: 0, endRow: 19, startColumn: 5, endColumn: 5, rangeType: 2 },
  ], 'A2', 0))
  const before = await page.evaluate(() => window.canvasContextFixture.state())
  const items = await page.evaluate(() => window.canvasContextFixture.search()[0].items)
  assert.ok(items.some(item => item.commandId === 'sheet.command.set-row-height'))
  assert.equal(items.some(item => item.commandId === 'sheet.command.set-worksheet-col-width'), false)
  assert.deepEqual(await page.evaluate(() => window.canvasContextFixture.state()), before)
  assert.equal(await page.getByRole('menu').count(), 0)
  assert.deepEqual(pageErrors, [])
})
