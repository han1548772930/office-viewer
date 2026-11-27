import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'
import { SHEET_MERGE_COMMAND_IDS } from '../src/features/sheet-workbench/sheet-merge-command.ts'
import { SHEET_CLEAR_COMMAND_IDS } from '../src/features/sheet-workbench/sheet-clear-command.ts'

const fixturePath = '/__sheet-range-editor-test__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>Sheets range editing</title><style>body{margin:0}#app{height:100vh}</style></head><body><div id="app"></div><script type="module">
import { createApp, h, ref, watch } from 'vue'
import { ConfigProvider, theme } from 'antdv-next'
import { OfficeStudio } from '/packages/office-studio/src/api/office-studio'
import { FOffice } from '/packages/office-studio/src/api/facade'
import ExcelWorkbench from '/packages/office-studio/src/components/ExcelWorkbench.vue'
import { createBlankWorkbook, createCell } from '/packages/office-studio/src/domain/model'
import { applyExcelWorkbookView, parseCellKey, toExcelWorkbookView } from '/packages/office-studio/src/services/excelViewAdapter'
import { SpreadsheetGeometry } from '/packages/office-studio/src/rendering/spreadsheet/layout'
const copy = value => JSON.parse(JSON.stringify(value))
const initial = createBlankWorkbook()
initial.id = 'range-edit-fixture'; initial.name = '范围编辑'; initial.calculationMode = 'manual'
const sheet = initial.sheets[0]
sheet.defaultColumnWidth = 90; sheet.defaultRowHeight = 26
sheet.cells['0:0'] = createCell(0, 0, null)
sheet.cells['0:1'] = { ...createCell(0, 1, '首项'), richText: [{ text: '首项', style: { italic: true } }], metadata: { vendor: 'kept' } }
sheet.cells['0:1'].style.font.bold = true; sheet.cells['0:1'].style.fill = '#abcdef'
sheet.cells['0:2'] = { ...createCell(0, 2, 11), formula: '=5+6', display: '11', dataType: 'formula', cachedDataType: 'number', hyperlink: 'https://example.test', metadata: { formulaError: 'obsolete', extension: 'kept' } }
sheet.cells['0:2'].style.font.italic = true
const studio = new OfficeStudio({ workbook: initial }), api = FOffice.newAPI(studio)
const standalone = new URLSearchParams(location.search).has('standalone')
let base = copy(initial)
const view = ref(toExcelWorkbookView(initial)), dark = ref(false), variant = ref('classic'), notices = [], requests = [], emitted = [], valueChanges = []
watch(() => studio.editor.workbook.model, model => { if (!standalone) view.value = toExcelWorkbookView(model) }, { flush: 'sync' })
studio.events.addEvent('SheetMergeConfirmationRequested', event => requests.push(copy(event)))
studio.events.addEvent('SheetValueChanged', event => valueChanges.push(copy(event)))
const current = () => standalone ? base : studio.editor.workbook.model
window.rangeEditorFixture = {
  snapshot: () => copy(current()), notifications: () => copy(notices), requests: () => copy(requests),
  revision: () => studio.getWorkbookRevision(current().id), valueChanges: () => copy(valueChanges),
  history: () => { const unit = studio.editor.history.units[studio.activeWorkbookId]; return { undo: unit?.undo.length ?? 0, redo: unit?.redo.length ?? 0, emits: emitted.length } },
  appearance: (value, layout) => { dark.value = value; variant.value = layout },
  command: (id, ranges, extra = {}) => studio.executeCommandSync(id, { unitId: current().id, sheetId: current().sheets[0].id, ranges, ...extra }),
  apiClear: (method, address) => api.getActiveWorkbook().getActiveSheet().getRange(address)[method](),
  change: () => api.getActiveWorkbook().getActiveSheet().getRange('D3').setValue('new revision'),
  protect: () => {
    const model = copy(current()); model.sheets[0].rangeProtectionRules = [{ id: 'locked', permissionId: 'locked-p', ownerUserId: 'owner', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView', ranges: [{ startRow: 0, endRow: 0, startCol: 0, endCol: 2 }] }]
    if (standalone) { base = model; view.value = toExcelWorkbookView(model) } else studio.editor.workbook.replace(model)
  },
  seed: (mode) => {
    const model = copy(initial), sheet = model.sheets[0]
    if (mode === 'merged') {
      sheet.cells = { '0:0': createCell(0, 0, 'first'), '4:4': createCell(4, 4, 'second') }
      sheet.mergedRanges = [{ startRow: 0, endRow: 0, startCol: 0, endCol: 1 }, { startRow: 4, endRow: 4, startCol: 4, endCol: 5 }]
      sheet.cells['0:0'].style.font.bold = true
    }
    if (mode === 'filtered') {
      sheet.cells = {}; for (let row = 0; row < 3; row++) sheet.cells[row + ':0'] = createCell(row, 0, 'row ' + row)
      sheet.hiddenRows = [2]; sheet.autoFilter = { range: { startRow: 0, endRow: 2, startCol: 0, endCol: 0 }, columns: { 0: { kind: 'values', selectedItems: ['row 2'] } }, filteredRows: [1] }
    }
    if (mode === 'hyperlink-selection') {
      sheet.cells['0:2'].hyperlink = null
    }
    if (mode === 'find-column') {
      sheet.cells = { '1:0': createCell(1, 0, 'alpha'), '1:1': createCell(1, 1, 'alpha') }
    }
    if (mode === 'find-workbook') {
      sheet.cells = {}; sheet.name = '空表'
      model.sheets = [
        { ...copy(sheet), id: 'find-before', name: '前表', cells: { '0:0': createCell(0, 0, 'alpha') } },
        sheet,
        { ...copy(sheet), id: 'find-after', name: '后表', cells: { '0:0': createCell(0, 0, 'alpha') } },
      ]
    }
    if (standalone) { base = model; view.value = toExcelWorkbookView(model) } else { studio.editor.workbook.replace(model); studio.editor.history.clear(studio.activeWorkbookId) }
    emitted.length = 0
  },
  save: () => JSON.stringify(current()),
  reopen: value => api.createWorkbook({ ...JSON.parse(value), id: 'range-edit-reopened' }),
  switchWorkbook: () => { const model = createBlankWorkbook(); model.id = 'range-other'; model.name = '其他工作簿'; api.createWorkbook(model) },
  rect: address => {
    const position = parseCellKey(address), element = document.querySelector('.spreadsheet-canvas'), scroll = element.querySelector('.canvas-scroll')
    const bounds = element.getBoundingClientRect(), geometry = new SpreadsheetGeometry(view.value.sheets[0])
    const rect = geometry.cellRect(position.row, position.col, { width: bounds.width, height: bounds.height, scrollLeft: scroll.scrollLeft, scrollTop: scroll.scrollTop, zoom: 1 })
    return { x: bounds.left + rect.x + Math.min(12, rect.width / 2), y: bounds.top + rect.y + rect.height / 2 }
  },
}
createApp({ render: () => h(ConfigProvider, { theme: { algorithm: dark.value ? theme.darkAlgorithm : theme.defaultAlgorithm } }, {
  default: () => h(ExcelWorkbench, { modelValue: view.value, studio: standalone ? undefined : studio,
    canUndo: studio.editor.canUndo, canRedo: studio.editor.canRedo, ribbonVariant: variant.value,
    'onUpdate:modelValue': value => { emitted.push(copy(value)); if (standalone) { base = applyExcelWorkbookView(base, value); view.value = value } else studio.editor.replaceWorkbookFromView(value, studio.regionalSettings.region.value) },
    onUndo: () => api.getActiveWorkbook().undo(), onRedo: () => api.getActiveWorkbook().redo(), onNotify: value => notices.push(copy(value)),
  }),
}) }).mount('#app')
window.addEventListener('pagehide', () => studio.dispose(), { once: true })
</script></body></html>`

let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'sheet-range-editor-fixture', configureServer(vite) { vite.middlewares.use(fixturePath, async (_req, res, next) => {
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
  errors = []; page = await browser.newPage({ viewport: { width: 1800, height: 900 } })
  page.on('pageerror', error => errors.push(error.message))
  await page.addInitScript(() => Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined }))
  await page.goto(baseURL); await expect(grid()).toBeVisible()
}, { timeout: 60_000 })
afterEach(async () => { try { assert.deepEqual(errors, []) } finally { await page?.close() } })
const grid = () => page.getByRole('grid', { name: '工作表画布', exact: true })
const modal = () => page.getByRole('dialog', { name: '合并单元格', exact: true })
const snapshot = () => page.evaluate(() => window.rangeEditorFixture.snapshot())
const history = () => page.evaluate(() => window.rangeEditorFixture.history())
async function select(address) {
  const input = page.getByRole('textbox', { name: '名称框', exact: true })
  await input.fill(address); await input.press('Enter'); await expect(grid()).toBeFocused()
}
async function menu(id, label) {
  const item = page.locator('[data-ribbon-item="' + id + '"]:visible').first()
  await expect(item).toBeVisible(); await item.getByRole('button').click()
  await page.locator('.ribbon-command-menu:visible').getByRole('menuitem', { name: label, exact: true }).click()
}
async function command(id) { await page.locator('[data-ribbon-item="' + id + '"]:visible').first().getByRole('button').click() }
async function contextClear(label, address = 'A1') {
  const point = await page.evaluate(value => window.rangeEditorFixture.rect(value), address)
  await page.mouse.click(point.x, point.y, { button: 'right' })
  await page.getByRole('menuitem', { name: '清除', exact: true }).hover()
  await page.getByRole('menuitem', { name: label, exact: true }).click()
}

test('toolbar merge confirmation cancels without history and restores Canvas focus', async () => {
  await select('A1:C1'); await menu('sheet.menu.merge', '全部合并')
  await expect(modal()).toBeVisible(); await expect(modal()).toContainText('其余 1 项内容')
  assert.equal((await history()).undo, 0); assert.equal((await snapshot()).sheets[0].cells['0:2'].formula, '=5+6')
  await modal().getByRole('button', { name: '取消', exact: true }).click()
  await expect(modal()).toBeHidden(); await expect(grid()).toBeFocused()
  assert.equal((await history()).undo, 0); assert.deepEqual((await snapshot()).sheets[0].mergedRanges, [])
})

test('accepted merge preserves the donor and non-master styles; undo/redo and internal reopen use one history', async () => {
  await select('A1:C1'); await menu('sheet.menu.merge', '全部合并')
  await modal().getByRole('button', { name: '继续合并', exact: true }).click()
  await expect(modal()).toBeHidden(); await expect(grid()).toBeFocused()
  const saved = await snapshot(), sheet = saved.sheets[0]
  assert.equal(sheet.cells['0:0'].value, '首项'); assert.deepEqual(sheet.cells['0:0'].richText, [{ text: '首项', style: { italic: true } }])
  assert.deepEqual(sheet.cells['0:0'].metadata, { vendor: 'kept' }); assert.equal(sheet.cells['0:2'].style.font.italic, true)
  assert.equal(sheet.cells['0:2'].formula, null); assert.equal(sheet.cells['0:2'].hyperlink, null); assert.equal(sheet.cells['0:2'].metadata.formulaError, undefined)
  assert.equal((await history()).undo, 1)
  await command('sheet.command.undo'); assert.equal((await snapshot()).sheets[0].cells['0:2'].formula, '=5+6')
  await command('sheet.command.redo'); assert.equal((await snapshot()).sheets[0].cells['0:2'].formula, null)
  await page.evaluate(value => window.rangeEditorFixture.reopen(JSON.stringify(value)), saved)
  await select('A1'); assert.equal((await snapshot()).sheets[0].cells['0:0'].value, '首项')
})

test('Delete, right-click format clear and public clear share content, topology and history behavior', async () => {
  await select('C1'); await grid().press('Delete')
  const cleared = (await snapshot()).sheets[0].cells['0:2']
  assert.equal(cleared.value, null); assert.equal(cleared.formula, null); assert.equal(cleared.display, ''); assert.equal(cleared.style.font.italic, true)
  assert.equal((await history()).undo, 1)
  await page.evaluate(() => window.rangeEditorFixture.seed('merged')); await select('A1')
  await contextClear('清除格式')
  assert.deepEqual((await snapshot()).sheets[0].mergedRanges, [{ startRow: 4, endRow: 4, startCol: 4, endCol: 5 }])
  assert.equal((await snapshot()).sheets[0].cells['0:0'].value, 'first'); assert.equal((await history()).undo, 1)
  await page.evaluate(() => window.rangeEditorFixture.apiClear('clear', 'E5:F5'))
  assert.deepEqual((await snapshot()).sheets[0].mergedRanges, []); assert.equal((await snapshot()).sheets[0].cells['4:4'], undefined)
  assert.equal((await history()).undo, 2)
})

test('Home clear-format preserves editing and composing drafts, then clears through the existing history and internal save paths', async () => {
  for (const standalone of [false, true]) for (const host of ['cell', 'formula']) for (const variant of ['classic', 'grid']) {
    await page.goto(baseURL + (standalone ? '?standalone=1' : '')); await expect(grid()).toBeVisible()
    await page.evaluate(value => window.rangeEditorFixture.appearance(false, value), variant); await select('B1')
    const before = await snapshot(), depth = await history(), revision = await page.evaluate(() => window.rangeEditorFixture.revision())
    const input = host === 'cell' ? page.locator('.canvas-cell-input input, .canvas-cell-input textarea') : page.getByRole('textbox', { name: '公式编辑', exact: true })
    if (host === 'cell') await grid().press('F2')
    if (host === 'formula') await input.focus()
    await input.fill('未提交草稿')
    const clear = page.locator('[data-ribbon-item="sheet.command.clear-format"]:visible').first().getByRole('button')
    await expect(clear).toBeDisabled(); await clear.evaluate(element => element.click())
    for (const id of ['sheet.command.clear-format', 'sheet.command.clear-selection-format']) {
      await page.getByTestId('workbench-ribbon').evaluate((element, commandId) => element.__vueParentComponent.emit('execute', commandId), id)
    }
    await expect(input).toHaveValue('未提交草稿'); await expect(grid()).not.toBeFocused()
    assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), depth)
    assert.equal(await page.evaluate(() => window.rangeEditorFixture.revision()), revision)
    await input.dispatchEvent('compositionstart'); await expect(clear).toBeDisabled()
    await page.evaluate(value => window.rangeEditorFixture.appearance(true, value), variant); await expect(clear).toBeDisabled()
    for (const id of ['sheet.command.clear-format', 'sheet.command.clear-selection-format']) {
      await page.getByTestId('workbench-ribbon').evaluate((element, commandId) => element.__vueParentComponent.emit('execute', commandId), id)
    }
    await expect(input).toHaveValue('未提交草稿'); await expect(grid()).not.toBeFocused()
    assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), depth)
    assert.equal(await page.evaluate(() => window.rangeEditorFixture.revision()), revision)
    assert.deepEqual(await page.evaluate(() => window.rangeEditorFixture.valueChanges()), [])
    await input.dispatchEvent('compositionend', { data: '未提交草稿' }); await expect(clear).toBeDisabled()
    await input.press('Escape'); await expect(grid()).toBeFocused(); await expect(clear).toBeEnabled(); await clear.click()
    await expect(grid()).toBeFocused()
    const painted = await snapshot(), cell = painted.sheets[0].cells['0:1']
    assert.equal(cell.value, '首项'); assert.equal(cell.style.font.bold, false); assert.equal(cell.style.fill, null)
    assert.deepEqual(cell.metadata, before.sheets[0].cells['0:1'].metadata)
    assert.deepEqual(painted.sheets[0].cells['0:2'], before.sheets[0].cells['0:2'])
    assert.equal((await history()).undo, depth.undo + (standalone ? 0 : 1))
    assert.deepEqual(await page.evaluate(() => window.rangeEditorFixture.valueChanges()), [])
    const saved = await page.evaluate(() => window.rangeEditorFixture.save())
    if (standalone) {
      assert.equal((await history()).emits, depth.emits + 1)
      assert.deepEqual(JSON.parse(saved).sheets[0].cells, painted.sheets[0].cells)
      continue
    }
    await command('sheet.command.undo'); assert.deepEqual(await snapshot(), before)
    await command('sheet.command.redo'); assert.deepEqual(await snapshot(), painted)
    await page.evaluate(value => window.rangeEditorFixture.reopen(value), saved); await expect(grid()).toBeVisible()
    assert.deepEqual((await snapshot()).sheets[0].cells, painted.sheets[0].cells)
  }
})

test('Home content clear and every clear intent preserve input drafts, then clear values with one history or standalone publication', async () => {
  for (const standalone of [false, true]) for (const host of ['cell', 'formula']) for (const variant of ['classic', 'grid']) {
    await page.goto(baseURL + (standalone ? '?standalone=1' : '')); await expect(grid()).toBeVisible()
    await page.evaluate(value => window.rangeEditorFixture.appearance(false, value), variant); await select('B1')
    const before = await snapshot(), depth = await history(), revision = await page.evaluate(() => window.rangeEditorFixture.revision())
    const input = host === 'cell' ? page.locator('.canvas-cell-input input, .canvas-cell-input textarea') : page.getByRole('textbox', { name: '公式编辑', exact: true })
    if (host === 'cell') await grid().press('F2')
    if (host === 'formula') await input.focus()
    await input.fill('未提交清值草稿')
    const clear = page.locator('[data-ribbon-item="sheet.command.clear-content"]:visible').first().getByRole('button')
    for (const composing of [false, true]) {
      if (composing) {
        await input.dispatchEvent('compositionstart')
        await page.evaluate(value => window.rangeEditorFixture.appearance(true, value), variant)
      }
      await expect(clear).toBeDisabled(); await clear.evaluate(element => element.click())
      for (const intent of [
        ...SHEET_CLEAR_COMMAND_IDS.map(commandId => ({ commandId, params: undefined })),
        { commandId: 'sheet.command.clear-range', params: { options: { contentsOnly: true } } },
        { commandId: 'sheet.command.clear-range', params: { options: { formatOnly: true } } },
      ]) {
        await page.getByTestId('workbench-ribbon').evaluate((element, value) => element.__vueParentComponent.emit('execute', value.commandId, value.params), intent)
      }
      await expect(input).toHaveValue('未提交清值草稿'); await expect(grid()).not.toBeFocused()
      assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), depth)
      assert.equal(await page.evaluate(() => window.rangeEditorFixture.revision()), revision)
      assert.deepEqual(await page.evaluate(() => window.rangeEditorFixture.valueChanges()), [])
    }
    await input.dispatchEvent('compositionend', { data: '未提交清值草稿' }); await expect(clear).toBeDisabled()
    await input.press('Escape'); await expect(grid()).toBeFocused(); await expect(clear).toBeEnabled(); await clear.click()
    await expect(grid()).toBeFocused()
    const cleared = await snapshot(), cell = cleared.sheets[0].cells['0:1']
    assert.equal(cell.value, null); assert.equal(cell.formula, null); assert.equal(cell.dataType, 'blank')
    assert.equal(cell.display, ''); assert.deepEqual(cell.style, before.sheets[0].cells['0:1'].style)
    assert.deepEqual(cell.metadata, before.sheets[0].cells['0:1'].metadata)
    assert.deepEqual(cleared.sheets[0].cells['0:2'], before.sheets[0].cells['0:2'])
    assert.equal((await history()).undo, depth.undo + (standalone ? 0 : 1))
    const saved = await page.evaluate(() => window.rangeEditorFixture.save())
    if (standalone) {
      assert.equal((await history()).emits, depth.emits + 1)
      assert.deepEqual(JSON.parse(saved).sheets[0].cells, cleared.sheets[0].cells)
      continue
    }
    await command('sheet.command.undo'); assert.deepEqual(await snapshot(), before)
    await command('sheet.command.redo'); assert.deepEqual(await snapshot(), cleared)
    await page.evaluate(value => window.rangeEditorFixture.reopen(value), saved); await expect(grid()).toBeVisible()
    assert.deepEqual((await snapshot()).sheets[0].cells, cleared.sheets[0].cells)
  }
})

test('toolbar protection creates a rule for the selected range and preserves existing rules through history and internal reopen', async () => {
  for (const variant of ['classic', 'grid']) {
    await page.goto(baseURL); await expect(grid()).toBeVisible()
    await page.evaluate(value => window.rangeEditorFixture.appearance(false, value), variant)
    await page.evaluate(() => window.rangeEditorFixture.protect()); await select('D3')
    const before = await snapshot(), depth = await history()
    await command('sheet.operation.open-range-protection')
    const panel = page.getByRole('dialog', { name: '范围保护', exact: true })
    await expect(panel).toBeVisible()
    await expect(panel.getByRole('textbox', { name: '保护应用区域', exact: true })).toHaveValue('D3')
    const description = panel.getByRole('textbox', { name: '保护说明', exact: true })
    await expect(description).toBeEnabled(); await expect(description).toHaveValue('')
    assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), depth)
    await page.evaluate(value => window.rangeEditorFixture.appearance(true, value), variant)
    await description.fill('当前选区保护'); await panel.getByRole('button', { name: '保存规则', exact: true }).click()
    const saved = await snapshot(), rules = saved.sheets[0].rangeProtectionRules
    assert.equal(rules.length, 2); assert.deepEqual(rules[0], before.sheets[0].rangeProtectionRules[0])
    assert.equal(rules[1].description, '当前选区保护')
    assert.deepEqual(rules[1].ranges, [{ startRow: 2, endRow: 2, startCol: 3, endCol: 3 }])
    assert.equal(rules[1].editState, 'onlyMe'); assert.equal(rules[1].viewState, 'othersCanView')
    assert.deepEqual(saved.sheets[0].cells, before.sheets[0].cells)
    assert.equal((await history()).undo, depth.undo + 1)
    const serialized = await page.evaluate(() => window.rangeEditorFixture.save())
    await panel.getByRole('button', { name: '关闭', exact: true }).click(); await expect(panel).toBeHidden()
    await command('sheet.command.undo'); assert.deepEqual(await snapshot(), before)
    await command('sheet.command.redo'); assert.deepEqual(await snapshot(), saved)
    await page.evaluate(value => window.rangeEditorFixture.reopen(value), serialized); await expect(grid()).toBeVisible()
    assert.deepEqual((await snapshot()).sheets[0].rangeProtectionRules, rules)
    await select('D3')
    const point = await page.evaluate(() => window.rangeEditorFixture.rect('D3'))
    await page.mouse.click(point.x, point.y, { button: 'right' })
    await page.getByRole('menuitem', { name: '保护行列', exact: true }).hover()
    await page.getByRole('menuitem', { name: '查看所有保护范围', exact: true }).click()
    await expect(panel).toBeVisible()
    await expect(panel.getByRole('textbox', { name: '保护应用区域', exact: true })).toHaveValue('A1:C1')
    assert.deepEqual((await snapshot()).sheets[0].rangeProtectionRules, rules)
    await panel.getByRole('button', { name: '关闭', exact: true }).click(); await expect(panel).toBeHidden()
  }
})

test('new protection preserves cell and formula drafts and rejects protected selections while existing read-only management remains available', async () => {
  const panel = page.getByRole('dialog', { name: '范围保护', exact: true })
  for (const host of ['cell', 'formula']) for (const variant of ['classic', 'grid']) {
    await page.goto(baseURL); await expect(grid()).toBeVisible()
    await page.evaluate(value => window.rangeEditorFixture.appearance(false, value), variant); await select('D3')
    const before = await snapshot(), depth = await history(), revision = await page.evaluate(() => window.rangeEditorFixture.revision())
    const input = host === 'cell' ? page.locator('.canvas-cell-input input, .canvas-cell-input textarea') : page.getByRole('textbox', { name: '公式编辑', exact: true })
    if (host === 'cell') await grid().press('F2')
    if (host === 'formula') await input.focus()
    await input.fill('未提交保护草稿')
    const protection = page.locator('[data-ribbon-item="sheet.operation.open-range-protection"]:visible').first().getByRole('button')
    for (const composing of [false, true]) {
      if (composing) {
        await input.dispatchEvent('compositionstart')
        await page.evaluate(value => window.rangeEditorFixture.appearance(true, value), variant)
      }
      await expect(protection).toBeDisabled(); await protection.evaluate(element => element.click())
      for (const params of [{ mode: 'add' }, { mode: 'add', sheetId: before.sheets[0].id }]) {
        await page.getByTestId('workbench-ribbon').evaluate((element, value) => element.__vueParentComponent.emit('execute', 'sheet.operation.open-range-protection', value), params)
      }
      await expect(panel).toBeHidden(); await expect(input).toHaveValue('未提交保护草稿'); await expect(grid()).not.toBeFocused()
      assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), depth)
      assert.equal(await page.evaluate(() => window.rangeEditorFixture.revision()), revision)
    }
    await input.dispatchEvent('compositionend', { data: '未提交保护草稿' }); await expect(protection).toBeDisabled()
    await input.press('Escape'); await expect(grid()).toBeFocused(); await expect(protection).toBeEnabled(); await protection.click()
    await expect(panel).toBeVisible()
    await expect(panel.getByRole('textbox', { name: '保护应用区域', exact: true })).toHaveValue('D3')
    assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), depth)
    await panel.getByRole('button', { name: '关闭', exact: true }).click(); await expect(panel).toBeHidden()
  }
  await page.goto(baseURL); await expect(grid()).toBeVisible()
  await page.evaluate(() => window.rangeEditorFixture.protect()); await select('B1')
  const before = await snapshot(), depth = await history()
  const protection = page.locator('[data-ribbon-item="sheet.operation.open-range-protection"]:visible').first().getByRole('button')
  await expect(protection).toBeDisabled()
  await page.getByTestId('workbench-ribbon').evaluate(element => element.__vueParentComponent.emit('execute', 'sheet.operation.open-range-protection', { mode: 'add' }))
  await expect(panel).toBeHidden(); assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), depth)
  await page.getByTestId('workbench-ribbon').evaluate((element, sheetId) => element.__vueParentComponent.emit('execute', 'sheet.operation.open-range-protection', { sheetId }), before.sheets[0].id)
  await expect(panel).toBeVisible()
  await expect(panel.getByRole('textbox', { name: '保护应用区域', exact: true })).toHaveValue('A1:C1')
  await expect(panel.getByRole('textbox', { name: '保护说明', exact: true })).toBeDisabled()
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), depth)
  await panel.getByRole('button', { name: '关闭', exact: true }).click(); await expect(panel).toBeHidden()
})

test('merge commands preserve cell and formula drafts until Escape, then confirm through the existing history and internal save paths', async () => {
  for (const standalone of [false, true]) for (const host of ['cell', 'formula']) for (const variant of ['classic', 'grid']) {
    await page.goto(baseURL + (standalone ? '?standalone=1' : '')); await expect(grid()).toBeVisible()
    await page.evaluate(value => window.rangeEditorFixture.appearance(false, value), variant); await select('A1:C1')
    const before = await snapshot(), depth = await history(), revision = await page.evaluate(() => window.rangeEditorFixture.revision())
    const requests = await page.evaluate(() => window.rangeEditorFixture.requests())
    const input = host === 'cell' ? page.locator('.canvas-cell-input input, .canvas-cell-input textarea') : page.getByRole('textbox', { name: '公式编辑', exact: true })
    if (host === 'cell') await grid().press('F2')
    if (host === 'formula') await input.focus()
    await input.fill('未提交合并草稿')
    const merge = page.locator('[data-ribbon-item="sheet.menu.merge"]:visible').first().getByRole('button')
    for (const composing of [false, true]) {
      if (composing) {
        await input.dispatchEvent('compositionstart')
        await page.evaluate(value => window.rangeEditorFixture.appearance(true, value), variant)
      }
      await expect(merge).toBeDisabled(); await merge.evaluate(element => element.click())
      for (const id of SHEET_MERGE_COMMAND_IDS) {
        await page.getByTestId('workbench-ribbon').evaluate((element, commandId) => element.__vueParentComponent.emit('execute', commandId), id)
      }
      await expect(input).toHaveValue('未提交合并草稿'); await expect(grid()).not.toBeFocused(); await expect(modal()).toBeHidden()
      assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), depth)
      assert.equal(await page.evaluate(() => window.rangeEditorFixture.revision()), revision)
      assert.deepEqual(await page.evaluate(() => window.rangeEditorFixture.requests()), requests)
      assert.deepEqual(await page.evaluate(() => window.rangeEditorFixture.valueChanges()), [])
    }
    await input.dispatchEvent('compositionend', { data: '未提交合并草稿' }); await expect(merge).toBeDisabled()
    await input.press('Escape'); await expect(grid()).toBeFocused(); await expect(merge).toBeEnabled()
    await select('A1:C1'); await menu('sheet.menu.merge', '全部合并'); await expect(modal()).toBeVisible()
    assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), depth)
    await modal().getByRole('button', { name: '继续合并', exact: true }).click()
    await expect(modal()).toBeHidden(); await expect(grid()).toBeFocused()
    const merged = await snapshot(), cell = merged.sheets[0].cells['0:0']
    assert.equal(cell.value, '首项'); assert.deepEqual(cell.metadata, { vendor: 'kept' })
    assert.deepEqual(cell.richText, before.sheets[0].cells['0:1'].richText)
    assert.equal(merged.sheets[0].cells['0:2'].formula, null)
    assert.deepEqual(merged.sheets[0].mergedRanges, [{ startRow: 0, endRow: 0, startCol: 0, endCol: 2 }])
    assert.equal((await history()).undo, depth.undo + (standalone ? 0 : 1))
    assert.equal((await page.evaluate(() => window.rangeEditorFixture.requests())).length, requests.length + (standalone ? 0 : 1))
    const saved = await page.evaluate(() => window.rangeEditorFixture.save())
    if (standalone) {
      assert.equal((await history()).emits, depth.emits + 1)
      assert.deepEqual(JSON.parse(saved).sheets[0], merged.sheets[0])
      continue
    }
    await command('sheet.command.undo'); assert.deepEqual(await snapshot(), before)
    await command('sheet.command.redo'); assert.deepEqual(await snapshot(), merged)
    await page.evaluate(value => window.rangeEditorFixture.reopen(value), saved); await expect(grid()).toBeVisible()
    assert.deepEqual((await snapshot()).sheets[0], merged.sheets[0])
  }
})

test('range clearing skips filter-hidden rows and processes manually hidden rows using the pre-edit filter scope', async () => {
  await page.evaluate(() => window.rangeEditorFixture.seed('filtered')); await select('A1:A3'); await grid().press('Delete')
  const sheet = (await snapshot()).sheets[0]
  assert.equal(sheet.cells['1:0'].value, 'row 1'); assert.equal(sheet.cells['2:0'].value, null)
  assert.deepEqual(sheet.hiddenRows, [2]); assert.equal((await history()).undo, 1)
})

test('a content revision arriving during confirmation rejects the original request with visible feedback', async () => {
  await select('A1:C1'); await menu('sheet.menu.merge', '全部合并')
  await page.evaluate(() => window.rangeEditorFixture.change())
  await modal().getByRole('button', { name: '继续合并', exact: true }).click()
  await expect(modal().getByRole('alert')).toContainText('工作簿已更改')
  assert.deepEqual((await snapshot()).sheets[0].mergedRanges, []); assert.equal((await history()).undo, 1)
  await modal().getByRole('button', { name: '取消', exact: true }).click(); await expect(grid()).toBeFocused()
})

test('a permission change during confirmation never clears the protected contents', async () => {
  await select('A1:C1'); await menu('sheet.menu.merge', '全部合并')
  await page.evaluate(() => window.rangeEditorFixture.protect())
  await modal().getByRole('button', { name: '继续合并', exact: true }).click()
  await expect(modal().getByRole('alert')).toContainText('工作簿已更改')
  assert.equal((await snapshot()).sheets[0].cells['0:1'].value, '首项'); assert.equal((await history()).undo, 0)
  await modal().getByRole('button', { name: '取消', exact: true }).click()
  const item = page.locator('[data-ribbon-item="sheet.menu.merge"]:visible').first()
  await expect(item.getByRole('button')).toBeDisabled()
  const before = await snapshot(), depth = await history(), revision = await page.evaluate(() => window.rangeEditorFixture.revision())
  await expect(page.locator('[data-ribbon-item="sheet.command.clear-format"]:visible').first().getByRole('button')).toBeDisabled()
  await page.getByTestId('workbench-ribbon').evaluate(element => element.__vueParentComponent.emit('execute', 'sheet.command.clear-format'))
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), depth)
  assert.equal(await page.evaluate(() => window.rangeEditorFixture.revision()), revision)
})

test('unmerge handles two disjoint selections instead of only the active master', async () => {
  await page.evaluate(() => window.rangeEditorFixture.seed('merged')); await select('A1')
  const point = await page.evaluate(() => window.rangeEditorFixture.rect('E5'))
  await page.keyboard.down('Control'); await page.mouse.click(point.x, point.y); await page.keyboard.up('Control')
  await menu('sheet.menu.merge', '取消合并')
  assert.deepEqual((await snapshot()).sheets[0].mergedRanges, []); assert.equal((await history()).undo, 1)
  assert.equal((await snapshot()).sheets[0].cells['0:1'].style.font.bold, true)
})

test('switching workbooks closes the old confirmation and does not mutate either workbook', async () => {
  await select('A1:C1'); await menu('sheet.menu.merge', '全部合并')
  await page.evaluate(() => window.rangeEditorFixture.switchWorkbook())
  await expect(modal()).toBeHidden(); await expect(grid()).toBeVisible()
  assert.equal((await snapshot()).id, 'range-other'); assert.deepEqual((await snapshot()).sheets[0].mergedRanges, [])
  assert.equal((await history()).undo, 0)
})

test('the controlled confirmation inherits the dark theme in both Ribbon layouts and returns focus after Escape', async () => {
  for (const variant of ['classic', 'grid']) {
    await page.evaluate(value => window.rangeEditorFixture.appearance(true, value), variant)
    await select('A1:C1'); await menu('sheet.menu.merge', '全部合并'); await expect(modal()).toBeVisible()
    const color = await modal().evaluate(element => getComputedStyle(element.querySelector('.ant-modal-container') ?? element.querySelector('.ant-modal-content') ?? element).backgroundColor)
    assert.notEqual(color, 'rgb(255, 255, 255)')
    await modal().press('Escape'); await expect(modal()).toBeHidden(); await expect(grid()).toBeFocused()
  }
  assert.equal((await history()).undo, 0)
})

for (const variant of ['classic', 'grid']) test('standalone ' + variant + ' merge cancellation keeps the actual primary and successful unmerge returns editing to the range start', async () => {
  await page.goto(baseURL + '?standalone=1')
  await page.evaluate(variant => window.rangeEditorFixture.appearance(false, variant), variant)
  await select('A1:C1')
  await grid().press('Tab')
  await menu('sheet.menu.merge', '全部合并')
  const before = await snapshot()
  await modal().getByRole('button', { name: '取消', exact: true }).click()
  await expect(grid()).toBeFocused()
  assert.deepEqual(await snapshot(), before)
  assert.equal((await history()).emits, 0)
  await grid().press('F2')
  const input = page.locator('.canvas-cell-input input, .canvas-cell-input textarea')
  await expect(input).toHaveValue('首项')
  await input.press('Escape')
  await menu('sheet.menu.merge', '全部合并')
  await modal().getByRole('button', { name: '继续合并', exact: true }).click()
  await expect(grid()).toBeFocused()
  assert.equal((await history()).emits, 1)
  await menu('sheet.menu.merge', '取消合并')
  assert.equal((await history()).emits, 2)
  assert.deepEqual((await snapshot()).sheets[0].mergedRanges, [])
  await grid().press('F2')
  await expect(input).toHaveValue('首项')
  await input.press('Escape')
  const unmerged = await snapshot()
  await page.evaluate(() => window.rangeEditorFixture.protect())
  const denied = await snapshot(), beforeDenied = await history()
  await expect(page.locator('[data-ribbon-item="sheet.menu.merge"]:visible').first().getByRole('button')).toBeDisabled()
  assert.deepEqual(await snapshot(), denied)
  assert.deepEqual(await history(), beforeDenied)
  assert.equal(unmerged.sheets[0].cells['0:0'].value, '首项')
})

test('standalone confirmation and clear keep explicit rich-text and style reset markers through the parent adapter', async () => {
  await page.goto(baseURL + '?standalone=1'); await expect(grid()).toBeVisible()
  await select('A1:C1'); await menu('sheet.menu.merge', '全部合并')
  await modal().getByRole('button', { name: '继续合并', exact: true }).click(); await expect(grid()).toBeFocused()
  assert.equal((await snapshot()).sheets[0].cells['0:0'].value, '首项'); assert.equal((await snapshot()).sheets[0].cells['0:1'].richText, undefined)
  await select('A1'); await grid().press('Delete')
  assert.equal((await snapshot()).sheets[0].cells['0:0'].richText, undefined)
  await contextClear('清除格式')
  assert.equal((await snapshot()).sheets[0].cells['0:0'].style.font.bold, false)
  assert.deepEqual((await snapshot()).sheets[0].mergedRanges, [])
})

test.each(['classic', 'grid'])('%s single replacement from an unmatched cell follows column order through history and internal reopen', async variant => {
  await page.evaluate(value => { window.rangeEditorFixture.seed('find-column'); window.rangeEditorFixture.appearance(false, value) }, variant)
  await select('B1')
  const before = await snapshot()
  await page.getByRole('tab', { name: '数据', exact: true }).click(); await command('sheet.operation.find')
  const panel = page.getByRole('search'), query = panel.getByRole('textbox', { name: '查找', exact: true }), replacement = panel.getByRole('textbox', { name: '替换', exact: true })
  await expect(panel).toHaveClass(/ant-form-inline/); assert.equal(await panel.evaluate(element => element.tagName), 'FORM')
  const lightBackground = await panel.evaluate(element => getComputedStyle(element).backgroundColor)
  await expect(query).toBeFocused(); await query.fill('alpha'); await replacement.fill('beta')
  await panel.getByRole('combobox', { name: '查找方向', exact: true }).click()
  await page.locator('.ant-select-dropdown:visible').getByText('按列查找', { exact: true }).click()
  await expect(panel).toContainText('2 个匹配项'); await replacement.focus()
  await page.evaluate(value => window.rangeEditorFixture.appearance(true, value), variant)
  await expect.poll(() => panel.evaluate(element => getComputedStyle(element).backgroundColor)).not.toBe(lightBackground)
  await expect(replacement).toBeFocused(); await expect(replacement).toHaveValue('beta'); await expect(query).toHaveValue('alpha')
  assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0)
  await replacement.press('Enter')
  await expect.poll(async () => (await snapshot()).sheets[0].cells['1:1'].value).toBe('beta')
  assert.deepEqual((await snapshot()).sheets[0].cells['1:0'], before.sheets[0].cells['1:0'])
  assert.equal((await history()).undo, 1)
  await query.press('Escape'); await expect(panel).toBeHidden(); await expect(grid()).toBeFocused()
  await page.getByRole('tab', { name: '开始', exact: true }).click(); await command('sheet.command.undo')
  assert.deepEqual((await snapshot()).sheets[0].cells, before.sheets[0].cells)
  await command('sheet.command.redo'); assert.equal((await snapshot()).sheets[0].cells['1:1'].value, 'beta')
  const saved = await page.evaluate(() => window.rangeEditorFixture.save())
  await page.evaluate(value => window.rangeEditorFixture.reopen(value), saved); await select('B1')
  assert.equal((await snapshot()).sheets[0].cells['1:1'].value, 'beta')
  assert.deepEqual((await snapshot()).sheets[0].cells['1:0'], before.sheets[0].cells['1:0'])
})

test.each(['classic', 'grid'])('%s single replacement from an unmatched sheet follows workbook order through history and internal reopen', async variant => {
  await page.evaluate(value => { window.rangeEditorFixture.seed('find-workbook'); window.rangeEditorFixture.appearance(false, value) }, variant)
  await select('A1')
  const before = await snapshot()
  assert.equal(before.activeSheetId, before.sheets[1].id)
  await page.getByRole('tab', { name: '数据', exact: true }).click(); await command('sheet.operation.find')
  const panel = page.getByRole('search'), query = panel.getByRole('textbox', { name: '查找', exact: true }), replacement = panel.getByRole('textbox', { name: '替换', exact: true })
  await expect(query).toBeFocused(); await query.fill('alpha'); await replacement.fill('beta')
  await panel.getByRole('combobox', { name: '查找范围', exact: true }).click()
  await page.locator('.ant-select-dropdown:visible').getByText('工作簿', { exact: true }).click()
  await expect(panel).toContainText('2 个匹配项'); await replacement.focus()
  await page.evaluate(value => window.rangeEditorFixture.appearance(true, value), variant)
  await expect(replacement).toBeFocused(); await expect(replacement).toHaveValue('beta'); await expect(query).toHaveValue('alpha')
  assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0)
  await panel.getByRole('button', { name: '替换', exact: true }).click()
  await expect.poll(async () => (await snapshot()).sheets[2].cells['0:0'].value).toBe('beta')
  assert.deepEqual((await snapshot()).sheets[0].cells, before.sheets[0].cells); assert.deepEqual((await snapshot()).sheets[1].cells, {})
  assert.equal((await history()).undo, 1)
  await panel.getByRole('button', { name: '关闭查找和替换', exact: true }).click(); await expect(panel).toBeHidden(); await expect(grid()).toBeFocused()
  await page.getByRole('tab', { name: '开始', exact: true }).click(); await command('sheet.command.undo')
  assert.deepEqual((await snapshot()).sheets.map(sheet => sheet.cells), before.sheets.map(sheet => sheet.cells))
  await command('sheet.command.redo'); assert.equal((await snapshot()).sheets[2].cells['0:0'].value, 'beta')
  const saved = await page.evaluate(() => window.rangeEditorFixture.save())
  await page.evaluate(value => window.rangeEditorFixture.reopen(value), saved); await select('A1')
  assert.equal((await snapshot()).sheets[2].cells['0:0'].value, 'beta')
  assert.deepEqual((await snapshot()).sheets[0].cells, before.sheets[0].cells); assert.deepEqual((await snapshot()).sheets[1].cells, {})
})

test.each(['classic', 'grid'])('%s hyperlinks use the active range for restrictions and labels through history and internal reopen', async variant => {
  await page.evaluate(value => { window.rangeEditorFixture.seed('hyperlink-selection'); window.rangeEditorFixture.appearance(false, value) }, variant)
  const before = await snapshot()
  await page.getByRole('tab', { name: '插入', exact: true }).click(); await select('A1')
  const formulaPoint = await page.evaluate(() => window.rangeEditorFixture.rect('C1'))
  await page.keyboard.down('Control'); await page.mouse.click(formulaPoint.x, formulaPoint.y); await page.keyboard.up('Control')
  const insert = page.locator('[data-ribbon-item="sheet.operation.insert-hyperlink"]:visible').first().getByRole('button')
  await expect(insert).toBeDisabled(); assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0)
  await select('A1')
  const textPoint = await page.evaluate(() => window.rangeEditorFixture.rect('B1'))
  await page.keyboard.down('Control'); await page.mouse.click(textPoint.x, textPoint.y); await page.keyboard.up('Control')
  await expect(insert).toBeEnabled(); await insert.click()
  const dialog = page.getByRole('dialog', { name: '插入超链接', exact: true })
  await expect(dialog).toBeVisible(); await expect(dialog.locator('.ant-form')).toBeVisible()
  await expect(dialog.getByRole('textbox', { name: '显示文字', exact: true })).toHaveValue('首项')
  const target = dialog.getByRole('textbox', { name: '链接地址', exact: true })
  await target.fill('https://example.test/report')
  await page.evaluate(value => window.rangeEditorFixture.appearance(true, value), variant)
  await expect(target).toBeFocused(); await expect(target).toHaveValue('https://example.test/report')
  await dialog.getByRole('button', { name: '确定', exact: true }).click(); await expect(dialog).toBeHidden(); await expect(grid()).toBeFocused()
  await expect.poll(async () => (await snapshot()).sheets[0].cells['0:1'].hyperlink).toBe('https://example.test/report')
  const linked = (await snapshot()).sheets[0].cells['0:1']
  assert.equal(linked.value, '首项'); assert.deepEqual(linked.richText, before.sheets[0].cells['0:1'].richText)
  assert.deepEqual(linked.metadata, before.sheets[0].cells['0:1'].metadata); assert.deepEqual(linked.style, before.sheets[0].cells['0:1'].style)
  assert.deepEqual((await snapshot()).sheets[0].cells['0:0'], before.sheets[0].cells['0:0'])
  assert.deepEqual((await snapshot()).sheets[0].cells['0:2'], before.sheets[0].cells['0:2']); assert.equal((await history()).undo, 1)
  await page.getByRole('tab', { name: '开始', exact: true }).click(); await command('sheet.command.undo')
  assert.deepEqual((await snapshot()).sheets[0].cells, before.sheets[0].cells)
  await command('sheet.command.redo'); assert.deepEqual((await snapshot()).sheets[0].cells['0:1'], linked)
  const saved = await page.evaluate(() => window.rangeEditorFixture.save())
  await page.evaluate(value => window.rangeEditorFixture.reopen(value), saved); await select('B1')
  assert.deepEqual((await snapshot()).sheets[0].cells['0:1'], linked)
  assert.deepEqual((await snapshot()).sheets[0].cells['0:0'], before.sheets[0].cells['0:0'])
  assert.deepEqual((await snapshot()).sheets[0].cells['0:2'], before.sheets[0].cells['0:2'])
})
