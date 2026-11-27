import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

const fixturePath = '/__sheet-format-painter-test__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>Sheets format painter</title><style>body{margin:0}#app{height:100vh;min-width:0}</style></head><body><div id="app"></div><script type="module">
import { computed, createApp, h, ref, watch } from 'vue'
import { ConfigProvider, theme } from 'antdv-next'
import { OfficeStudio } from '/packages/office-studio/src/api/office-studio'
import { FOffice } from '/packages/office-studio/src/api/facade'
import ExcelWorkbench from '/packages/office-studio/src/components/ExcelWorkbench.vue'
import { createBlankWorkbook, createCell, defaultCellStyle } from '/packages/office-studio/src/domain/model'
import { applyExcelWorkbookView, parseCellKey, toExcelWorkbookView } from '/packages/office-studio/src/services/excelViewAdapter'
import { SpreadsheetGeometry } from '/packages/office-studio/src/rendering/spreadsheet/layout'
import { createSpreadsheetScene } from '/packages/office-studio/src/rendering/spreadsheet/scene'
const copy = value => JSON.parse(JSON.stringify(value))
const initial = createBlankWorkbook(), sheet = initial.sheets[0]
initial.id = 'painter-browser'; initial.name = '格式刷'; initial.calculationMode = 'manual'; initial.activeSheetId = sheet.id = 'painter-main'
sheet.rowCount = 200; sheet.columnCount = 26; sheet.defaultColumnWidth = 90; sheet.defaultRowHeight = 26
sheet.cells['0:0'] = createCell(0, 0, 123.456)
sheet.cells['0:0'].style = { ...defaultCellStyle(), font: { family: 'Courier New', size: 18, bold: true, italic: true, underline: true, strike: false, color: '#123456' }, fill: '#abcdef', horizontal: 'right', vertical: 'bottom', wrap: true, wrapStrategy: 'wrap', numberFormat: '0.00', border: { bottom: { style: 'double', color: '#336699' } }, rotation: 45, rotationVertical: false, shrinkToFit: true, indent: 2, metadata: { vendor: { id: 'source-style' } } }
sheet.cells['0:1'] = createCell(0, 1, 'source B1'); sheet.cells['0:1'].style.font.italic = true
sheet.cells['1:0'] = createCell(1, 0, 'source A2'); sheet.cells['1:0'].style.fill = '#00ff00'
sheet.cells['1:1'] = createCell(1, 1, 'default source')
sheet.cells['3:3'] = { ...createCell(3, 3, .25), formula: '=1/4', dataType: 'formula', cachedDataType: 'number', display: '0.25', metadata: { target: 'formula' }, hyperlink: 'https://target.test' }
sheet.cells['3:4'] = { ...createCell(3, 4, '00123'), dataType: 'text', richText: [{ text: '00123', style: { italic: true } }], metadata: { target: 'text' } }
sheet.cells['4:3'] = createCell(4, 3, 7); sheet.cells['4:4'] = createCell(4, 4, false); sheet.cells['4:4'].style.font.bold = true
sheet.cells['10:13'] = createCell(10, 13, 'outside')
const second = copy(createBlankWorkbook().sheets[0]); second.id = 'painter-second'; second.name = '第二工作表'; second.rowCount = 200; second.columnCount = 26; second.defaultColumnWidth = 90; second.defaultRowHeight = 26
second.cells['3:3'] = createCell(3, 3, 'second target'); initial.sheets.push(second)
const standalone = new URLSearchParams(location.search).has('standalone'), studio = new OfficeStudio({ workbook: initial }), api = FOffice.newAPI(studio)
let base = copy(initial), pointerId = 0
const view = ref(toExcelWorkbookView(initial)), dark = ref(false), variant = ref('classic'), notices = [], executed = [], requests = [], emitted = [], valueChanges = [], canceled = new Set()
const direction = computed(() => studio.regionalSettings.locale.value === 'ar-SA' ? 'rtl' : 'ltr')
watch(() => studio.editor.workbook.model, model => { if (!standalone) view.value = toExcelWorkbookView(model) }, { flush: 'sync' })
studio.events.addEvent('CommandExecuted', event => executed.push(copy(event)))
studio.events.addEvent('SheetValueChanged', event => valueChanges.push(copy(event)))
studio.events.addEvent('SheetFormatPainterConfirmationRequested', event => requests.push(copy(event)))
studio.events.addEvent('BeforeCommandExecute', event => { if (canceled.has(event.id)) event.cancel = true })
document.addEventListener('pointerdown', event => { pointerId = event.pointerId }, true)
const current = () => standalone ? base : studio.editor.workbook.model
const active = () => current().sheets.find(item => item.id === current().activeSheetId) ?? current().sheets[0]
const replace = model => { if (standalone) { base = model; view.value = toExcelWorkbookView(model) } else studio.editor.workbook.replace(model) }
window.painterFixture = {
  snapshot: (id) => copy(id ? studio.getWorkbookModel(id) : current()),
  revision: () => studio.getWorkbookRevision(current().id),
  state: () => copy(studio.sheetFormatPainter.state), source: () => studio.sheetFormatPainter.getSource(),
  history: (id = studio.activeWorkbookId) => { const unit = studio.editor.history.units[id]; return { undo: unit?.undo.length ?? 0, redo: unit?.redo.length ?? 0, emits: emitted.length } },
  commands: () => copy(executed), notifications: () => copy(notices), requests: () => copy(requests),
  valueChanges: () => copy(valueChanges),
  cancel: (id, enabled) => { if (enabled) canceled.add(id); else canceled.delete(id) },
  appearance: (isDark, language = 'zh-CN') => { dark.value = isDark; studio.regionalSettings.setLocale(language) },
  variant: value => { variant.value = value },
  command: (id, ranges, extra = {}) => studio.executeCommandSync(id, { unitId: current().id, sheetId: active().id, ...(ranges ? { ranges } : {}), ...extra }),
  setSourceSize: value => api.getActiveWorkbook().getSheetBySheetId('painter-main').getRange('A1').setFontSize(value),
  mutate: () => api.getActiveWorkbook().getSheetBySheetId('painter-main').getRange('N11').setValue('new revision'),
  mutateStandalone: () => { const model = copy(current()); model.sheets[0].cells['10:13'] = createCell(10, 13, 'new standalone revision'); replace(model) },
  mutateSourceConditional: () => { const model = copy(current()); model.sheets[0].conditionalFormats[0].style.fill = '#3344dd'; replace(model) },
  activateSheet: id => { if (standalone) { base.activeSheetId = id; view.value = toExcelWorkbookView(base) } else { api.getActiveWorkbook().setActiveSheet(id); view.value = toExcelWorkbookView(studio.editor.workbook.model) } },
  activateWorkbook: id => studio.activateWorkbook(id),
  newWorkbook: () => api.createWorkbook({ ...copy(initial), id: 'painter-other-book' }),
  reopen: model => api.createWorkbook({ ...model, id: 'painter-reopened' }),
  seed: mode => {
    const model = copy(initial), source = model.sheets[0]
    if (mode === 'merged' || mode === 'conditional-merged') source.mergedRanges = [{ startRow: 0, endRow: 0, startCol: 0, endCol: 1 }]
    if (mode === 'frozen') { source.freeze = { rows: 1, columns: 1 }; source.hiddenRows = [2] }
    if (mode === 'rich-fonts') {
      Object.assign(source.cells['0:0'].style, { rotation: 0, shrinkToFit: false, indent: 0, wrap: false, wrapStrategy: 'overflow' })
      source.cells['3:4'].style = copy(source.cells['0:0'].style)
      source.cells['3:4'].richText = [{ text: '00123', fontSize: 32, color: '#ff0000', id: 'target-run', vendor: { keep: true } }]
    }
    if (mode === 'conditional' || mode === 'conditional-merged') {
      for (const cell of Object.values(source.cells)) cell.style = defaultCellStyle()
      source.conditionalFormats = [
        { id: 'painter-cf-source', type: 'expression', ranges: [{ startRow: 0, endRow: 0, startCol: 0, endCol: 0 }], operator: 'greaterThan', formula1: '=A1>0', style: { fill: '#35a65f' }, priority: 1, stopIfTrue: true, metadata: { vendor: { keep: true } } },
        { id: 'painter-cf-target', type: 'cellIs', ranges: [{ startRow: 3, endRow: 4, startCol: 3, endCol: 4 }], operator: 'greaterThan', formula1: '0', style: { fill: '#dc3344' }, priority: 2, stopIfTrue: false },
      ]
    }
    if (mode === 'conditional-boundary' || mode === 'conditional-value-formula') {
      for (const target of model.sheets) {
        for (const cell of Object.values(target.cells)) cell.style = defaultCellStyle()
        for (const [row, value] of [[0, 7], [1, 2], [2, 9], [3, 6], [4, 10]]) target.cells[row + ':0'] = createCell(row, 0, value)
      }
      source.conditionalFormats = [{ id: 'painter-cf-boundary', type: 'expression', ranges: [{ startRow: 1, endRow: 4, startCol: 0, endCol: 0 }],
        operator: 'greaterThan', formula1: '=A1>5', style: { fill: '#35a65f', vendor: { style: ['keep'] } }, priority: 1, stopIfTrue: false,
        vendor: { root: ['keep'] }, config: { custom: { config: ['keep'] } }, metadata: { vendor: { keep: ['keep'] } } }]
      if (mode === 'conditional-value-formula') Object.assign(source.conditionalFormats[0], { type: 'highlightCell', subType: 'formula', formula1: '', value: '=A1>5' })
    }
    if (mode === 'conditional-imported') {
      source.conditionalFormats = [
        { id: 'cf-imported-expression', type: 'highlightCell', subType: 'formula', ranges: [{ startRow: 0, endRow: 3, startCol: 0, endCol: 0 }],
          operator: 'greaterThan', formula1: '=UNSUPPORTED([VendorColumn])', style: { fill: '#35a65f' }, priority: 1, stopIfTrue: false, metadata: { vendor: ['keep'] } },
        { id: 'cf-imported-number', type: 'cellIs', ranges: [{ startRow: 3, endRow: 3, startCol: 3, endCol: 3 }],
          operator: 'greaterThan', formula1: '=UNSUPPORTED([VendorColumn])', style: { fill: '#35a65f' }, priority: 2, stopIfTrue: false },
      ]
    }
    if (mode === 'conditional-independent') {
      source.conditionalFormats = [
        { id: 'cf-independent-first', type: 'cellIs', ranges: [{ startRow: 0, endRow: 3, startCol: 0, endCol: 0 }],
          operator: 'greaterThan', formula1: '5', style: { fill: '#35a65f' }, priority: 1, stopIfTrue: false, metadata: { vendor: ['first'] } },
        { id: 'cf-independent-second', type: 'cellIs', ranges: [{ startRow: 3, endRow: 4, startCol: 3, endCol: 3 }],
          operator: 'greaterThan', formula1: '0', style: { fill: '#dc3344' }, priority: 2, stopIfTrue: false },
        { id: 'cf-independent-locked', type: 'cellIs', ranges: [{ startRow: 10, endRow: 10, startCol: 13, endCol: 13 }],
          operator: 'greaterThan', formula1: '0', style: { fill: '#3344dd' }, priority: 3, stopIfTrue: false, metadata: { vendor: ['locked'] } },
      ]
    }
    replace(model); studio.editor.history.clear(model.id); executed.length = 0; emitted.length = 0; requests.length = 0; valueChanges.length = 0
  },
  protect: target => {
    const model = copy(current())
    model.sheets[0].rangeProtectionRules = [{ id: 'locked-painter', permissionId: 'locked-painter-p', ownerUserId: 'owner', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView', description: '', ranges: [target] }]
    replace(model)
  },
  removeSource: () => { const model = copy(current()); model.sheets = model.sheets.filter(item => item.id !== 'painter-main'); model.activeSheetId = 'painter-second'; replace(model) },
  pointerId: () => pointerId,
  sceneAt: (address, sheetId) => {
    const position = parseCellKey(address), rendered = view.value.sheets.find(item => item.id === (sheetId ?? current().activeSheetId)) ?? view.value.sheets[0]
    const geometry = new SpreadsheetGeometry(rendered), viewport = { width: 1600, height: 800, scrollLeft: 0, scrollTop: 0, zoom: 1 }
    const rect = geometry.cellRect(position.row, position.col, viewport)
    const scene = createSpreadsheetScene(geometry, viewport, { activeCell: '', selectedCells: [], selectedRanges: [], suppressFillHandle: true })
    const cellCommands = scene.commands.filter(item => 'x' in item && 'y' in item
      && item.x >= rect.x && item.x < rect.x + rect.width && item.y >= rect.y && item.y < rect.y + rect.height)
    return {
      rect,
      fills: cellCommands.filter(item => item.kind === 'rect' && item.x === rect.x && item.y === rect.y
        && item.width === rect.width && item.height === rect.height).map(item => item.fill),
      rectColors: cellCommands.filter(item => item.kind === 'rect').map(item => item.fill),
      text: cellCommands.filter(item => item.kind === 'text').map(item => item.text),
      fonts: cellCommands.filter(item => item.kind === 'text').map(item => item.font),
    }
  },
  rect: (address, header) => {
    const position = parseCellKey(address), element = document.querySelector('.spreadsheet-canvas'), scroll = element.querySelector('.canvas-scroll'), bounds = element.getBoundingClientRect()
    const sheet = view.value.sheets.find(item => item.id === current().activeSheetId) ?? view.value.sheets[0], geometry = new SpreadsheetGeometry(sheet)
    const rect = geometry.cellRect(position.row, position.col, { width: bounds.width, height: bounds.height, scrollLeft: scroll.scrollLeft, scrollTop: scroll.scrollTop, zoom: Number(document.querySelector('.sheet-zoom-value')?.textContent?.replace('%', '')) / 100 || 1 })
    return { x: bounds.left + (header === 'row' ? 12 : rect.x + Math.min(12, rect.width / 2)), y: bounds.top + (header === 'column' ? 12 : rect.y + rect.height / 2) }
  },
  markerPixels: async () => {
    await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))
    const canvas = document.querySelector('.canvas-surface canvas'), context = canvas.getContext('2d'), data = context.getImageData(0, 0, Math.min(canvas.width, 230), Math.min(canvas.height, 100)).data
    let hash = 2166136261; for (const value of data) hash = Math.imul(hash ^ value, 16777619)
    return hash >>> 0
  },
}
createApp({ render: () => h(ConfigProvider, { direction: direction.value, theme: { algorithm: dark.value ? theme.darkAlgorithm : theme.defaultAlgorithm, token: { colorPrimary: '#3064c7' } } }, {
  default: () => h(ExcelWorkbench, { modelValue: view.value, studio: standalone ? undefined : studio, canUndo: studio.editor.canUndo, canRedo: studio.editor.canRedo, ribbonVariant: variant.value,
    'onUpdate:modelValue': value => { emitted.push(copy(value)); if (standalone) { base = applyExcelWorkbookView(base, value); view.value = value } else studio.editor.replaceWorkbookFromView(value, studio.regionalSettings.region.value) },
    onUndo: () => api.getActiveWorkbook().undo(), onRedo: () => api.getActiveWorkbook().redo(), onNotify: value => notices.push(copy(value)),
  }),
}) }).mount('#app')
window.addEventListener('pagehide', () => studio.dispose(), { once: true })
</script></body></html>`

let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'sheet-format-painter-fixture', configureServer(vite) {
      vite.middlewares.use(fixturePath, async (_req, res, next) => {
        try { res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.end(await vite.transformIndexHtml(fixturePath, fixture)) } catch (error) { next(error) }
      })
    },
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
  await page.goto(baseURL); await expect(grid()).toBeVisible(); await expect(canvas()).toHaveCount(1)
  await canvas().evaluate(element => { element.dataset.painterTestMount = 'original' })
}, { timeout: 60_000 })
afterEach(async () => { try { assert.deepEqual(errors, []) } finally { await page?.close() } })
const grid = () => page.getByRole('grid', { name: '工作表画布', exact: true })
const canvas = () => page.locator('.canvas-surface canvas')
const editor = () => page.locator('.canvas-cell-input input, .canvas-cell-input textarea')
const modal = () => page.getByRole('dialog', { name: '合并单元格', exact: true })
const snapshot = () => page.evaluate(() => window.painterFixture.snapshot())
const state = () => page.evaluate(() => window.painterFixture.state())
const history = () => page.evaluate(() => window.painterFixture.history())
const revision = () => page.evaluate(() => window.painterFixture.revision())
const cell = async (key, sheetId = 'painter-main') => (await snapshot()).sheets.find(sheet => sheet.id === sheetId)?.cells[key]
const contents = model => Object.fromEntries(model.sheets.map(sheet => [sheet.id, Object.fromEntries(Object.entries(sheet.cells).map(([key, { value, formula, dataType, cachedDataType, richText, metadata, hyperlink }]) => [key, { value, formula, dataType, cachedDataType, richText, metadata, hyperlink }]))]))
async function select(address) {
  const input = page.getByRole('textbox', { name: '名称框', exact: true })
  await input.fill(address); await input.press('Enter'); await expect(grid()).toBeFocused()
}
async function homeItem(id) {
  const visible = page.locator('.ribbon-visible-groups [data-ribbon-item="' + id + '"]:visible, .ribbon-grid-toolbar [data-ribbon-item="' + id + '"]:visible').first()
  if (await visible.count()) return visible
  const overflow = page.locator('.ribbon-overflow-panel:visible')
  if (!await overflow.count()) await page.locator('.ribbon-more-button:visible, .ribbon-grid-more-button:visible').click()
  const item = overflow.locator('[data-ribbon-item="' + id + '"]'); await expect(item).toBeVisible(); return item
}
const brush = async () => (await homeItem('sheet.operation.format-painter')).getByRole('button')
async function arm(continuous = false) { const button = await brush(); if (continuous) await button.dblclick(); else await button.click(); await expect(grid()).toHaveClass(/format-painter-active/) }
async function clickCell(address, options) { const point = await page.evaluate(value => window.painterFixture.rect(value), address); await page.mouse.click(point.x, point.y, options) }
async function pointerDown(address) { const point = await page.evaluate(value => window.painterFixture.rect(value), address); await page.mouse.move(point.x, point.y); await page.mouse.down() }
async function command(id) { await (await homeItem(id)).getByRole('button').click() }

test('single-click brush expands the target block and shares history, undo/redo and snapshot reopen', async () => {
  await select('A1:B2'); const before = contents(await snapshot()); await arm(); assert.equal((await history()).undo, 0)
  await clickCell('D4'); await expect.poll(async () => (await cell('3:3')).style.font.size).toBe(18)
  await expect.poll(async () => (await state()).mode).toBe('off'); await expect(grid()).toBeFocused()
  await expect(page.getByRole('textbox', { name: '名称框', exact: true })).toHaveValue('D4:E5')
  assert.deepEqual(contents(await snapshot()), before); assert.equal((await cell('4:4')).style.font.bold, false)
  assert.equal((await history()).undo, 1); await expect(canvas()).toHaveAttribute('data-painter-test-mount', 'original')
  assert.deepEqual((await page.evaluate(() => window.painterFixture.commands())).map(event => event.id), ['sheet.operation.format-painter', 'sheet.command.apply-format-painter'])
  await command('sheet.command.undo'); await expect.poll(async () => (await cell('3:3')).style.font.size).toBe(11)
  await expect(page.getByRole('textbox', { name: '名称框', exact: true })).toHaveValue('D4')
  await command('sheet.command.redo'); await expect(page.getByRole('textbox', { name: '名称框', exact: true })).toHaveValue('D4:E5')
  const saved = await snapshot(); await page.evaluate(model => window.painterFixture.reopen(model), saved)
  await expect.poll(async () => (await cell('3:3')).style.font.size).toBe(18); assert.deepEqual(contents(await snapshot()), before)
})

test('run-only painting updates Canvas font and one history entry while preserving target text', async () => {
  await page.evaluate(() => window.painterFixture.seed('rich-fonts'))
  const before = await cell('3:4')
  await select('A1'); await arm(); await clickCell('E4'); await expect.poll(async () => (await state()).mode).toBe('off')
  const painted = await cell('3:4')
  assert.deepEqual(painted.richText, [{ text: '00123', id: 'target-run', vendor: { keep: true } }])
  assert.equal(painted.value, before.value); assert.equal((await history()).undo, 1); await expect(grid()).toBeFocused()
  const fonts = await page.evaluate(() => window.painterFixture.sceneAt('E4').fonts)
  assert.ok(fonts.length > 0); assert.ok(fonts.every(font => font.includes('18px') && font.includes('Courier New')))
  assert.deepEqual(await page.evaluate(() => window.painterFixture.valueChanges()), [])
  await command('sheet.command.undo'); assert.deepEqual((await cell('3:4')).richText, before.richText)
  await command('sheet.command.redo'); assert.deepEqual((await cell('3:4')).richText, painted.richText)
  const saved = await snapshot(); await page.evaluate(model => window.painterFixture.reopen(model), saved)
  assert.deepEqual((await cell('3:4')).richText, painted.richText); assert.equal((await cell('3:4')).value, before.value)
})

test('double-click brush keeps continuous mode for successive pointer targets and Esc cancels it', async () => {
  await select('A1'); await arm(true); await expect.poll(async () => (await state()).mode).toBe('continuous')
  await clickCell('D4'); await clickCell('G7'); await expect.poll(async () => (await cell('6:6')).style.font.size).toBe(18)
  assert.equal((await state()).mode, 'continuous'); assert.equal((await history()).undo, 2)
  await grid().press('Escape'); await expect.poll(async () => (await state()).mode).toBe('off'); await expect(grid()).toBeFocused()
  assert.equal((await history()).undo, 2)
})

test('pointer drag paints the completed rectangular selection with source tiling', async () => {
  await select('A1:B2'); await arm(); await pointerDown('D4')
  const end = await page.evaluate(() => window.painterFixture.rect('F6')); await page.mouse.move(end.x, end.y, { steps: 6 }); await page.mouse.up()
  await expect.poll(async () => (await cell('5:5')).style.font.bold).toBe(true)
  assert.equal((await cell('4:4')).style.font.bold, false); assert.equal((await history()).undo, 1)
  await expect(page.getByRole('textbox', { name: '名称框', exact: true })).toHaveValue('D4:F6')
})

test('right-click and keyboard or name navigation retain ONCE without painting a target', async () => {
  await select('A1'); await arm(); const before = await snapshot()
  await grid().press('ArrowRight'); await grid().press('ArrowDown'); await select('D4'); await clickCell('G7', { button: 'right' })
  await expect.poll(async () => (await state()).mode).toBe('once')
  assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0)
})

test('pointercancel releases selection without applying or consuming a single-use brush', async () => {
  await select('A1'); await arm(); const before = await snapshot(); await pointerDown('D4')
  const pointerId = await page.evaluate(() => window.painterFixture.pointerId())
  await grid().dispatchEvent('pointercancel', { pointerId, pointerType: 'mouse', bubbles: true, button: 0, buttons: 0 }); await page.mouse.up()
  assert.equal((await state()).mode, 'once'); assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0)
  await clickCell('G7'); await expect.poll(async () => (await cell('6:6')).style.font.size).toBe(18); assert.equal((await history()).undo, 1)
})

test('lost pointer capture cancels the in-flight target and leaves ONCE available for another click', async () => {
  await select('A1'); await arm(); const before = await snapshot(); await pointerDown('D4')
  const pointerId = await page.evaluate(() => window.painterFixture.pointerId())
  await grid().evaluate((element, id) => element.releasePointerCapture(id), pointerId); await page.mouse.move(1, 1); await page.mouse.up()
  await expect.poll(async () => (await state()).mode).toBe('once'); assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0)
  await clickCell('G7'); await expect.poll(async () => (await cell('6:6')).style.font.size).toBe(18)
})

test('double-clicking a target after ONCE painting does not open a cell editor', async () => {
  await select('A1'); await arm(); await clickCell('D4', { clickCount: 2 })
  await expect.poll(async () => (await state()).mode).toBe('off'); await expect(editor()).toHaveCount(0)
  assert.equal((await cell('3:3')).formula, '=1/4'); assert.equal((await history()).undo, 1)
})

test('continuous mode crosses sheets and applies the captured source while source markers stay on their source sheet', async () => {
  await page.evaluate(() => window.painterFixture.activateSheet('painter-second')); await select('D4')
  const baseline = await page.evaluate(() => window.painterFixture.markerPixels())
  await page.evaluate(() => window.painterFixture.activateSheet('painter-main')); await select('A1:B2'); await arm(true)
  await page.evaluate(() => window.painterFixture.activateSheet('painter-second')); await select('D4')
  assert.equal(await page.evaluate(() => window.painterFixture.markerPixels()), baseline)
  assert.equal((await state()).sourceSheetId, 'painter-main'); assert.equal((await state()).mode, 'continuous')
  await clickCell('D4'); await expect.poll(async () => (await cell('3:3', 'painter-second')).style.font.size).toBe(18)
  assert.equal((await cell('3:3', 'painter-second')).value, 'second target'); assert.equal((await history()).undo, 1)
})

test('source marking changes Canvas pixels without remounting the surface or writing document content', async () => {
  await select('D4'); const before = await snapshot(), baseline = await page.evaluate(() => window.painterFixture.markerPixels())
  await page.evaluate(() => window.painterFixture.command('sheet.operation.format-painter', [{ startRow: 0, startCol: 0, endRow: 1, endCol: 1 }]))
  await expect.poll(() => page.evaluate(() => window.painterFixture.markerPixels())).not.toBe(baseline)
  assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0)
  await expect(canvas()).toHaveAttribute('data-painter-test-mount', 'original')
})

test('continuous upgrade uses the immutable sample after the source formatting changes externally', async () => {
  await select('A1'); await arm(); const token = (await state()).sessionId
  await page.evaluate(() => { window.painterFixture.setSourceSize(36); window.painterFixture.command('sheet.operation.continuous-format-painter') })
  assert.equal((await state()).sessionId, token); assert.equal((await state()).mode, 'continuous')
  await clickCell('D4'); await expect.poll(async () => (await cell('3:3')).style.font.size).toBe(18)
  assert.equal((await cell('0:0')).style.font.size, 36); assert.equal((await history()).undo, 2)
})

test('redundant formatting consumes ONCE while retaining model sparsity and empty history', async () => {
  await select('B2'); await arm(); const before = await snapshot(); await clickCell('H8')
  await expect.poll(async () => (await state()).mode).toBe('off'); assert.deepEqual(await snapshot(), before)
  assert.equal((await history()).undo, 0); assert.equal((await cell('7:7')), undefined)
})

test('rejected apply commands preserve ONCE and retry through the same registered command', async () => {
  await select('A1'); await arm(); const before = await snapshot()
  await page.evaluate(() => window.painterFixture.cancel('sheet.command.apply-format-painter', true)); await clickCell('D4')
  await expect.poll(() => page.evaluate(() => window.painterFixture.notifications().length)).toBe(1)
  assert.equal((await state()).mode, 'once'); assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0)
  await page.evaluate(() => window.painterFixture.cancel('sheet.command.apply-format-painter', false)); await clickCell('D4')
  await expect.poll(async () => (await cell('3:3')).style.font.size).toBe(18); assert.equal((await history()).undo, 1)
})

test('protected targets reject painting atomically and the active brush remains cancelable', async () => {
  await select('A1'); await arm(); await page.evaluate(() => window.painterFixture.protect({ startRow: 3, endRow: 3, startCol: 3, endCol: 3 }))
  const before = await snapshot(); await clickCell('D4')
  await expect.poll(() => page.evaluate(() => window.painterFixture.notifications().length)).toBe(1)
  assert.equal((await state()).mode, 'once'); assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0)
  await expect(await brush()).toBeEnabled(); await grid().press('Escape'); assert.equal((await state()).mode, 'off')
  await expect(await brush()).toBeDisabled()
})

test('merge painting confirms discarded content and records formatting and topology in one transaction', async () => {
  await page.evaluate(() => window.painterFixture.seed('merged')); await select('B1'); await arm(); const before = await snapshot()
  await clickCell('D4'); await expect(modal()).toBeVisible(); await expect(modal()).toContainText('其余 1 项内容')
  assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0); assert.equal((await state()).mode, 'once')
  await modal().getByRole('button', { name: '继续合并', exact: true }).click()
  await expect(modal()).toBeHidden(); await expect(grid()).toBeFocused(); assert.equal((await history()).undo, 1)
  assert.equal((await cell('3:3')).formula, '=1/4'); assert.equal((await cell('3:4')).value, null)
  assert.equal((await cell('3:3')).style.font.size, 18); assert.equal((await state()).mode, 'off')
  await command('sheet.command.undo'); assert.equal((await cell('3:4')).value, '00123')
  await command('sheet.command.redo'); assert.equal((await cell('3:4')).value, null)
})

test('canceling a merge confirmation preserves the source and restores Canvas focus', async () => {
  await page.evaluate(() => window.painterFixture.seed('merged')); await select('A1'); await arm(); const before = await snapshot()
  await clickCell('D4'); await expect(modal()).toBeVisible()
  const pending = await state(), depth = await history(), baseRevision = await revision()
  assert.equal(await page.evaluate(requestId => window.painterFixture.command('sheet.operation.cancel-format-painter-confirmation', undefined,
    { unitId: 'other', requestId }), pending.confirmationId), false)
  await expect(modal()).toBeVisible(); assert.deepEqual(await state(), pending); assert.deepEqual(await history(), depth)
  assert.equal(await revision(), baseRevision); assert.deepEqual(await snapshot(), before)
  await modal().getByRole('button', { name: '取消', exact: true }).click()
  await expect(modal()).toBeHidden(); await expect(grid()).toBeFocused()
  assert.equal((await state()).mode, 'once'); assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0)
  await clickCell('G7'); await expect.poll(async () => (await cell('6:6')).style.font.size).toBe(18)
})

test('stale confirmation shows an error and preserves target data and the original brush', async () => {
  await page.evaluate(() => window.painterFixture.seed('merged')); await select('A1'); await arm(); await clickCell('D4')
  await page.evaluate(() => window.painterFixture.mutate()); const before = await snapshot()
  await modal().getByRole('button', { name: '继续合并', exact: true }).click()
  await expect(modal().getByRole('alert')).toContainText('工作簿已更改'); assert.deepEqual(await snapshot(), before)
  assert.equal((await state()).mode, 'once'); assert.equal((await history()).undo, 1)
})

test('switching sheets dismisses pending confirmation even when its user cancel command is intercepted', async () => {
  await page.evaluate(() => window.painterFixture.seed('merged')); await select('A1'); await arm(true); await clickCell('D4')
  await expect(modal()).toBeVisible()
  await page.evaluate(() => { window.painterFixture.cancel('sheet.operation.cancel-format-painter-confirmation', true); window.painterFixture.activateSheet('painter-second') })
  await expect(modal()).toBeHidden(); assert.equal((await state()).mode, 'continuous'); assert.equal((await state()).confirmationId, undefined)
  assert.equal((await history()).undo, 0)
})

test('format clear and source deletion clear temporary brush state without losing document history', async () => {
  await select('A1'); await arm(true); await clickCell('D4'); assert.equal((await history()).undo, 1)
  await command('sheet.command.clear-format'); await expect.poll(async () => (await state()).mode).toBe('off')
  assert.equal((await history()).undo, 2)
  await select('A1'); await arm(); await page.evaluate(() => window.painterFixture.removeSource())
  await expect.poll(async () => (await state()).mode).toBe('off'); await expect(grid()).toBeVisible()
})

test('workbook switching clears the brush and preserves independent document histories', async () => {
  await select('A1'); await arm(true); await clickCell('D4'); await page.evaluate(() => window.painterFixture.newWorkbook())
  await expect.poll(async () => (await state()).mode).toBe('off'); assert.equal((await history()).undo, 0)
  assert.equal((await cell('3:3')).style.font.size, 11)
  await page.evaluate(() => window.painterFixture.activateWorkbook('painter-browser'))
  await expect.poll(async () => (await cell('3:3')).style.font.size).toBe(18); assert.equal((await history()).undo, 1)
})

test('row-header selection applies formatting after its completed left pointer gesture', async () => {
  await select('A1'); await arm()
  const point = await page.evaluate(() => window.painterFixture.rect('A8', 'row')); await page.mouse.click(point.x, point.y)
  await expect.poll(async () => (await cell('7:25')).style.font.size).toBe(18)
  assert.equal((await state()).mode, 'off'); assert.equal((await history()).undo, 1); assert.equal((await cell('8:25')), undefined)
})

test('Esc cancels from a focused Ribbon brush while composing Esc leaves the session active', async () => {
  await select('A1'); await arm(); const button = await brush(); await button.focus()
  for (const info of [{ key: 'Escape', isComposing: true }, { key: 'Escape', keyCode: 229 }]) {
    assert.equal(await button.evaluate((element, options) => { const event = new KeyboardEvent('keydown', { ...options, bubbles: true, cancelable: true }); element.dispatchEvent(event); return event.defaultPrevented }, info), false)
    assert.equal((await state()).mode, 'once')
  }
  await button.press('Escape'); await expect.poll(async () => (await state()).mode).toBe('off'); await expect(grid()).toBeFocused(); assert.equal((await history()).undo, 0)
})

test('narrow grid overflow retains brush activation and Esc through a teleported focused button', async () => {
  await select('A1'); await page.evaluate(() => window.painterFixture.variant('grid')); await page.setViewportSize({ width: 180, height: 900 })
  await arm(); const button = await brush(); await button.focus(); await expect(page.locator('.ribbon-overflow-panel:visible')).toBeVisible()
  await button.press('Escape'); await expect.poll(async () => (await state()).mode).toBe('off'); await expect(grid()).toBeFocused()
  assert.equal((await history()).undo, 0)
})

test('dark theme and RTL grid layout keep formatting, theme switching and focus functional', async () => {
  await page.evaluate(() => { window.painterFixture.variant('grid'); window.painterFixture.appearance(true, 'ar-SA') })
  await select('A1'); await arm(); await clickCell('D4')
  await expect.poll(async () => (await cell('3:3')).style.font.size).toBe(18); await expect(grid()).toBeFocused()
  await expect(page.locator('.excel-workbench')).toHaveAttribute('dir', 'rtl')
  await page.evaluate(() => window.painterFixture.appearance(false, 'zh-CN')); await expect(canvas()).toHaveAttribute('data-painter-test-mount', 'original')
})

test('source sampling across frozen panes survives scrolling and pointer painting offscreen targets', async () => {
  await page.evaluate(() => window.painterFixture.seed('frozen')); await select('A1:B2'); await arm()
  await select('D40'); assert.equal((await state()).mode, 'once'); await clickCell('D40')
  await expect.poll(async () => (await cell('39:3')).style.font.size).toBe(18)
  assert.equal((await cell('40:3')).style.fill, '#00ff00'); assert.equal((await state()).mode, 'off')
  assert.equal((await history()).undo, 1); await expect(canvas()).toHaveAttribute('data-painter-test-mount', 'original')
})

test('a composing cell editor preserves its draft and brush until separate edit and brush cancellation', async () => {
  await select('A1'); await grid().press('F2'); await expect(editor()).toBeVisible(); await editor().fill('42')
  await page.evaluate(() => window.painterFixture.command('sheet.operation.format-painter', [{ startRow: 0, endRow: 0, startCol: 0, endCol: 0 }]))
  await editor().dispatchEvent('compositionstart'); await editor().press('Escape'); await expect(editor()).toHaveValue('42')
  assert.equal((await state()).mode, 'once'); assert.equal((await history()).undo, 0)
  await editor().dispatchEvent('compositionend', { data: '42' }); await editor().press('Escape'); await expect(editor()).toHaveCount(0)
  assert.equal((await state()).mode, 'once'); await grid().press('Escape'); assert.equal((await state()).mode, 'off')
  assert.equal((await cell('0:0')).value, 123.456); assert.equal((await history()).undo, 0)
})

test('input sessions block new brush sampling, preserve active drafts, and defer pending merge confirmation during composition', async () => {
  for (const standalone of [false, true]) for (const host of ['cell', 'formula']) {
    await page.goto(baseURL + (standalone ? '?standalone' : '')); await expect(grid()).toBeVisible(); await select('A1')
    const before = await snapshot(), depth = await history(), baseRevision = await revision()
    const input = host === 'cell' ? editor() : page.getByRole('textbox', { name: '公式编辑', exact: true })
    if (host === 'cell') await grid().press('F2')
    if (host === 'formula') await input.focus()
    await input.fill('42'); await expect(await brush()).toBeDisabled()
    await input.dispatchEvent('compositionstart'); await expect(await brush()).toBeDisabled()
    await (await brush()).evaluate(element => element.click())
    await expect(input).toHaveValue('42'); await expect(grid()).not.toHaveClass(/format-painter-active/)
    assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), depth); assert.equal(await revision(), baseRevision)
    await input.dispatchEvent('compositionend', { data: '42' }); await input.press('Escape'); await expect(grid()).toBeFocused()
    await arm()
    if (host === 'cell') await grid().press('F2')
    if (host === 'formula') await input.focus()
    await input.fill('77'); await input.dispatchEvent('compositionstart')
    await page.evaluate(() => window.painterFixture.appearance(true))
    const button = await brush(); await expect(button).toBeEnabled(); await button.click()
    await expect(grid()).not.toHaveClass(/format-painter-active/); await expect(grid()).not.toBeFocused()
    await expect(input).toBeVisible(); await expect(input).toHaveValue('77')
    assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), depth); assert.equal(await revision(), baseRevision)
    assert.deepEqual(await page.evaluate(() => window.painterFixture.valueChanges()), [])
    await input.dispatchEvent('compositionend', { data: '77' }); await input.focus(); await input.press('Escape')
    await expect(grid()).toBeFocused(); assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), depth)
  }
  for (const standalone of [false, true]) {
    await page.goto(baseURL + (standalone ? '?standalone' : '')); await expect(grid()).toBeVisible()
    await page.evaluate(() => window.painterFixture.seed('merged')); await select('A1'); await arm(); await clickCell('D4')
    await expect(modal()).toBeVisible()
    const before = await snapshot(), depth = await history(), baseRevision = await revision(), painterBefore = await state()
    const input = page.getByRole('textbox', { name: '公式编辑', exact: true })
    const confirm = modal().getByRole('button', { name: '继续合并', exact: true })
    await expect(confirm).toBeEnabled(); await input.dispatchEvent('compositionstart')
    await expect(confirm).toBeDisabled(); await expect(modal().getByRole('button', { name: '取消', exact: true })).toBeEnabled()
    await page.evaluate(() => window.painterFixture.appearance(true)); await expect(confirm).toBeDisabled()
    await confirm.evaluate(element => element.click()); await expect(modal()).toBeVisible()
    await expect(grid()).toHaveClass(/format-painter-active/)
    assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), depth)
    assert.equal(await revision(), baseRevision); assert.deepEqual(await state(), painterBefore)
    assert.deepEqual(await page.evaluate(() => window.painterFixture.valueChanges()), [])
    await input.dispatchEvent('compositionend', { data: '' }); await expect(confirm).toBeEnabled(); await confirm.click()
    await expect(modal()).toBeHidden(); await expect(grid()).toBeFocused()
    assert.equal((await cell('3:4')).value, null)
    assert.equal((await history()).undo, depth.undo + (standalone ? 0 : 1))
    if (standalone) assert.equal((await history()).emits, depth.emits + 1)
    await page.evaluate(() => window.painterFixture.seed('merged')); await select('A1'); await arm(); await clickCell('D4')
    await expect(modal()).toBeVisible()
    const cancelBefore = await snapshot(), cancelDepth = await history(), cancelRevision = await revision()
    await input.dispatchEvent('compositionstart'); await expect(confirm).toBeDisabled()
    await modal().getByRole('button', { name: '取消', exact: true }).click(); await expect(modal()).toBeHidden()
    await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))))
    await expect(grid()).not.toBeFocused(); await expect(grid()).toHaveClass(/format-painter-active/)
    assert.deepEqual(await snapshot(), cancelBefore); assert.deepEqual(await history(), cancelDepth); assert.equal(await revision(), cancelRevision)
    assert.deepEqual(await page.evaluate(() => window.painterFixture.valueChanges()), [])
    await input.dispatchEvent('compositionend', { data: '' }); await grid().focus(); await expect(grid()).toBeFocused()
  }
})

test('standalone workbench uses the same source block, content preservation and single change emission', async () => {
  await page.goto(baseURL + '?standalone'); await expect(grid()).toBeVisible(); await select('A1:B2'); const before = contents(await snapshot())
  await arm(); await clickCell('D4'); await expect.poll(async () => (await cell('3:3')).style.font.size).toBe(18)
  await expect(grid()).not.toHaveClass(/format-painter-active/); assert.deepEqual(contents(await snapshot()), before); assert.equal((await history()).emits, 1)
})

test('standalone merge painting confirms content loss through the shared domain command', async () => {
  await page.goto(baseURL + '?standalone'); await page.evaluate(() => window.painterFixture.seed('merged')); await select('A1'); await arm(); await clickCell('D4')
  await expect(modal()).toBeVisible(); await modal().getByRole('button', { name: '继续合并', exact: true }).click()
  await expect(modal()).toBeHidden(); assert.equal((await cell('3:4')).value, null); assert.equal((await history()).emits, 1)
})

test('standalone stale confirmation retains its immutable sample and shows an error without emitting an edit', async () => {
  await page.goto(baseURL + '?standalone'); await page.evaluate(() => window.painterFixture.seed('merged')); await select('A1'); await arm(); await clickCell('D4')
  await expect(modal()).toBeVisible(); await page.evaluate(() => window.painterFixture.mutateStandalone()); const before = await snapshot()
  await modal().getByRole('button', { name: '继续合并', exact: true }).click()
  await expect(modal().getByRole('alert')).toContainText('工作簿已更改'); assert.deepEqual(await snapshot(), before)
  await expect(grid()).toHaveClass(/format-painter-active/); assert.equal((await history()).emits, 0)
})

const conditionalRules = async (sheetId = 'painter-main') => (await snapshot()).sheets.find(sheet => sheet.id === sheetId)?.conditionalFormats
const covers = (entry, row, col) => entry.ranges.some(area => row >= area.startRow && row <= area.endRow && col >= area.startCol && col <= area.endCol)
const conditionalCommand = 'sheet.command.set-conditional-formats'
const dataBarCommand = 'sheet.command.add-data-bar-conditional-rule'
const iconSetCommand = 'sheet.command.add-icon-set-conditional-rule'
const conditionalPanel = () => page.getByRole('dialog', { name: '条件格式', exact: true })
const conditionalMatches = (sheetId = 'painter-main') => page.evaluate(id => ['A1', 'A2', 'A3', 'A4'].map(address =>
  window.painterFixture.sceneAt(address, id).fills.includes('#35a65f')), sheetId)
const valueEvents = () => page.evaluate(() => window.painterFixture.valueChanges())
async function openConditionalPanel() {
  await page.getByRole('tab', { name: '数据', exact: true }).click()
  await command('sheet.operation.open-conditional-formatting')
  await page.getByRole('menu', { name: '条件格式', exact: true }).getByRole('menuitem', { name: '管理条件格式', exact: true }).click()
  await expect(conditionalPanel()).toBeVisible()
}
async function editConditionalRule(range, formula, index = 0) {
  await conditionalPanel().locator('.conditional-rule-select').nth(index).click()
  if (range !== undefined) await conditionalPanel().getByRole('textbox', { name: '应用范围', exact: true }).fill(range)
  if (formula !== undefined) await conditionalPanel().getByRole('textbox', { name: '公式', exact: true }).fill(formula)
}
async function applyConditionalPanel() {
  await conditionalPanel().getByRole('button', { name: '应用', exact: true }).click()
}
async function expectConditionalClosed() {
  await expect(conditionalPanel()).toBeHidden()
  await expect(grid()).toBeFocused()
}

test('management menu opens all six editors and a new rule without publishing canceled drafts', async () => {
  await select('A1:A2')
  const before = await snapshot()
  await page.getByRole('tab', { name: '数据', exact: true }).click()
  for (const [action, editor] of [
    ['突出显示单元格规则', '突出显示单元格规则'], ['排名和平均值规则', '排名和平均值规则'],
    ['公式规则', '公式规则'], ['色阶规则', '色阶规则'], ['数据条规则', '数据条规则'],
    ['图标集规则', '图标集规则'], ['新建条件格式', '突出显示单元格规则'],
  ]) {
    await command('sheet.operation.open-conditional-formatting')
    await page.getByRole('menu', { name: '条件格式', exact: true }).getByRole('menuitem', { name: action, exact: true }).click()
    await expect(conditionalPanel()).toBeVisible()
    await expect(conditionalPanel().getByRole('textbox', { name: '应用范围', exact: true })).toHaveValue('A1:A2')
    await expect(conditionalPanel().getByText(editor, { exact: true })).toBeVisible()
    await conditionalPanel().getByRole('button', { name: '取消', exact: true }).click()
    await expectConditionalClosed()
  }
  assert.deepEqual(await snapshot(), before)
  assert.equal((await history()).undo, 0)
})

test('management clear distinguishes selected coverage from the entire sheet and shares history', async () => {
  await page.evaluate(() => window.painterFixture.seed('conditional-boundary'))
  await select('A2')
  const original = await conditionalRules(), before = contents(await snapshot())
  await page.getByRole('tab', { name: '数据', exact: true }).click()
  await command('sheet.operation.open-conditional-formatting')
  await page.getByRole('menu', { name: '条件格式', exact: true }).getByRole('menuitem', { name: '清除所选区域的规则', exact: true }).click()
  const partial = await conditionalRules()
  assert.equal(covers(partial[0], 1, 0), false)
  assert.equal(covers(partial[0], 2, 0), true)
  assert.equal((await history()).undo, 1)
  await command('sheet.operation.open-conditional-formatting')
  await page.getByRole('menu', { name: '条件格式', exact: true }).getByRole('menuitem', { name: '清除工作表的规则', exact: true }).click()
  assert.deepEqual(await conditionalRules(), [])
  assert.equal((await history()).undo, 2)
  await grid().press('Control+z'); assert.deepEqual(await conditionalRules(), partial)
  await grid().press('Control+z'); assert.deepEqual(await conditionalRules(), original)
  assert.deepEqual(contents(await snapshot()), before)
})

test('Insert quick visual rules render on Canvas, reflect model selection, and clear only one type in the selected part', async () => {
  await select('A1:A2')
  const before = contents(await snapshot())
  await page.getByRole('tab', { name: '插入', exact: true }).click()
  const dataBar = await homeItem(dataBarCommand)
  await dataBar.getByRole('button', { name: '数据条', exact: true }).click()
  await expect(dataBar.getByRole('button', { name: '数据条', exact: true })).toHaveAttribute('aria-pressed', 'true')
  await expect.poll(() => page.evaluate(() => window.painterFixture.sceneAt('A1').rectColors)).toContain('#ffbe38')
  const icons = await homeItem(iconSetCommand)
  await icons.getByRole('button', { name: '图标集选项', exact: true }).click()
  await page.getByRole('menu', { name: '图标集', exact: true }).getByRole('menuitem', { name: '方向', exact: true }).click()
  await expect(icons.getByRole('button', { name: '图标集', exact: true })).toHaveAttribute('aria-pressed', 'true')
  await expect.poll(() => page.evaluate(() => window.painterFixture.sceneAt('A1').text)).toContain('↑')
  const original = await conditionalRules()
  assert.deepEqual(original.map(rule => rule.type), ['iconSet', 'dataBar'])
  assert.equal((await history()).undo, 2)
  await select('A2')
  await icons.getByRole('button', { name: '图标集选项', exact: true }).click()
  const menu = page.getByRole('menu', { name: '图标集', exact: true })
  await expect(menu.getByRole('menuitem', { name: '方向', exact: true })).toHaveClass(/ant-menu-item-selected/)
  await menu.getByRole('menuitem', { name: '清除所选区域的规则', exact: true }).click()
  await expect(grid()).toBeFocused()
  const cleared = await conditionalRules()
  assert.equal(covers(cleared[0], 0, 0), true)
  assert.equal(covers(cleared[0], 1, 0), false)
  assert.deepEqual(cleared[1], original[1])
  assert.deepEqual(contents(await snapshot()), before)
  assert.equal((await history()).undo, 3)
})

test('quick custom rule opens the selected type and applies the current range in one transaction', async () => {
  await select('A1:A2')
  await page.getByRole('tab', { name: '插入', exact: true }).click()
  const icons = await homeItem(iconSetCommand)
  for (const apply of [false, true]) {
    await icons.getByRole('button', { name: '图标集选项', exact: true }).click()
    await page.getByRole('menu', { name: '图标集', exact: true }).getByRole('menuitem', { name: '新建条件格式', exact: true }).click()
    await expect(conditionalPanel()).toBeVisible()
    await expect(conditionalPanel().getByRole('textbox', { name: '应用范围', exact: true })).toHaveValue('A1:A2')
    await expect(conditionalPanel().getByText('图标集规则', { exact: true })).toBeVisible()
    if (apply) await applyConditionalPanel()
    else await conditionalPanel().getByRole('button', { name: '取消', exact: true }).click()
    await expectConditionalClosed()
    assert.equal((await history()).undo, apply ? 1 : 0)
  }
  const [rule] = await conditionalRules()
  assert.equal(rule.type, 'iconSet')
  assert.deepEqual(rule.ranges, [{ startRow: 0, startCol: 0, endRow: 1, endCol: 0 }])
  await openConditionalPanel()
  await expect(conditionalPanel().locator('.conditional-rule-select')).toContainText('图标集')
})

test('conditional-only pointer painting replaces rule scope and shares undo, redo and snapshot reopen', async () => {
  await page.evaluate(() => window.painterFixture.seed('conditional')); await select('A1')
  const before = await snapshot(); await arm(); await clickCell('D4')
  await expect.poll(async () => (await conditionalRules()).filter(entry => covers(entry, 3, 3)).map(entry => entry.id)).toEqual(['painter-cf-source'])
  assert.deepEqual(contents(await snapshot()), contents(before)); assert.equal((await cell('3:3')).style.fill, null); assert.equal((await history()).undo, 1)
  const painted = await conditionalRules(); assert.equal(painted.find(entry => entry.id === 'painter-cf-target').ranges.some(area => area.startCol === 4), true)
  await expect(grid()).toBeFocused(); await expect(canvas()).toHaveAttribute('data-painter-test-mount', 'original')
  await command('sheet.command.undo'); assert.deepEqual(await conditionalRules(), before.sheets[0].conditionalFormats)
  await command('sheet.command.redo'); assert.deepEqual(await conditionalRules(), painted)
  const saved = await snapshot(); await page.evaluate(model => window.painterFixture.reopen(model), saved)
  await expect.poll(() => conditionalRules()).toEqual(painted); assert.equal((await state()).mode, 'off')
})

test('painting a source without conditional rules clears only the selected destination rule coverage', async () => {
  await page.evaluate(() => window.painterFixture.seed('conditional')); await select('B2'); const before = contents(await snapshot())
  await arm(); await clickCell('D4')
  await expect.poll(async () => (await conditionalRules()).some(entry => covers(entry, 3, 3))).toBe(false)
  assert.equal((await conditionalRules()).some(entry => covers(entry, 4, 4)), true)
  assert.deepEqual(contents(await snapshot()), before); assert.equal((await history()).undo, 1); assert.equal((await state()).mode, 'off')
})

test('continuous conditional painting retains a repeated target rule and adds no redundant history', async () => {
  await page.evaluate(() => window.painterFixture.seed('conditional')); await select('A1'); await arm(true); await clickCell('D4')
  const painted = await conditionalRules(); await clickCell('D4')
  assert.deepEqual(await conditionalRules(), painted); assert.equal((await history()).undo, 1); assert.equal((await state()).mode, 'continuous')
  await clickCell('G7'); await expect.poll(async () => (await conditionalRules()).some(entry => covers(entry, 6, 6))).toBe(true)
  assert.equal((await history()).undo, 2); assert.equal(await cell('6:6'), undefined)
  await grid().press('Escape'); assert.equal((await state()).mode, 'off')
})

test('cross-sheet conditional painting creates an independent relative rule and keeps the source session', async () => {
  await page.evaluate(() => window.painterFixture.seed('conditional')); await select('A1'); await arm(true)
  const sourceRules = await conditionalRules(), token = (await state()).sessionId
  await page.evaluate(() => window.painterFixture.activateSheet('painter-second')); await clickCell('D4')
  await expect.poll(async () => (await conditionalRules('painter-second'))?.length ?? 0).toBe(1)
  const copied = (await conditionalRules('painter-second'))[0]
  assert.notEqual(copied.id, 'painter-cf-source'); assert.equal(copied.formula1, '=D4>0'); assert.deepEqual(copied.metadata, { vendor: { keep: true } })
  assert.deepEqual(await conditionalRules(), sourceRules); assert.equal((await state()).sessionId, token); assert.equal((await history()).undo, 1)
  await clickCell('D4'); assert.equal((await history()).undo, 1); assert.equal((await conditionalRules('painter-second'))[0].id, copied.id)
  await expect(canvas()).toHaveAttribute('data-painter-test-mount', 'original')
})

test('source conditional edits after activation preserve the immutable copied rule and current source definition', async () => {
  await page.evaluate(() => window.painterFixture.seed('conditional')); await select('A1'); await arm(true)
  await page.evaluate(() => window.painterFixture.mutateSourceConditional()); await clickCell('D4')
  await expect.poll(async () => (await conditionalRules()).find(entry => covers(entry, 3, 3))?.style.fill).toBe('#35a65f')
  assert.equal((await conditionalRules()).find(entry => entry.id === 'painter-cf-source').style.fill, '#3344dd')
  assert.equal((await state()).mode, 'continuous'); assert.equal((await history()).undo, 1)
})

test('merge confirmation commits conditional replacement only when its content-loss dialog is accepted', async () => {
  await page.evaluate(() => window.painterFixture.seed('conditional-merged')); await select('A1'); await arm(); const before = await conditionalRules()
  await clickCell('D4'); await expect(modal()).toBeVisible(); assert.deepEqual(await conditionalRules(), before); assert.equal((await history()).undo, 0)
  await modal().getByRole('button', { name: '继续合并', exact: true }).click(); await expect(modal()).toBeHidden(); await expect(grid()).toBeFocused()
  assert.equal((await conditionalRules()).some(entry => entry.id === 'painter-cf-source' && covers(entry, 3, 3)), true)
  assert.equal((await conditionalRules()).some(entry => covers(entry, 3, 4)), false)
  assert.equal((await cell('3:4')).value, null); assert.equal((await history()).undo, 1)
  await command('sheet.command.undo'); assert.deepEqual(await conditionalRules(), before); assert.equal((await cell('3:4')).value, '00123')
})

test('standalone conditional-only painting emits one complete model update with destination values retained', async () => {
  await page.goto(baseURL + '?standalone'); await page.evaluate(() => window.painterFixture.seed('conditional')); await select('A1')
  const before = contents(await snapshot()); await arm(); await clickCell('D4')
  await expect.poll(async () => (await conditionalRules()).filter(entry => covers(entry, 3, 3)).map(entry => entry.id)).toEqual(['painter-cf-source'])
  assert.deepEqual(contents(await snapshot()), before); assert.equal((await history()).emits, 1); assert.equal((await cell('3:3')).style.fill, null)
  await expect(grid()).not.toHaveClass(/format-painter-active/); await expect(grid()).toBeFocused()
})

test('same-sheet boundary painting keeps later valid references in one rule and repeated targets add no history', async () => {
  await page.evaluate(() => window.painterFixture.seed('conditional-boundary')); await select('A2:A5')
  const before = await snapshot(); await arm(true); await clickCell('A1')
  await expect.poll(() => conditionalMatches()).toEqual([false, true, false, true])
  const painted = await conditionalRules(), first = painted[0]
  assert.equal(painted.length, 1); assert.equal(first.id, 'painter-cf-boundary')
  assert.deepEqual(first.formulaAnchor, { row: 1, col: 0 }); assert.equal(first.formula1, '=A1>5')
  assert.equal(covers(first, 4, 0), true); assert.deepEqual(contents(await snapshot()), contents(before))
  assert.equal((await history()).undo, 1); assert.deepEqual(await valueEvents(), [])
  await clickCell('A1'); assert.deepEqual(await conditionalRules(), painted); assert.equal((await history()).undo, 1)
  assert.equal((await state()).mode, 'continuous'); await grid().press('Escape')
  await command('sheet.command.undo'); assert.deepEqual(await conditionalRules(), before.sheets[0].conditionalFormats)
  await command('sheet.command.redo'); await expect.poll(() => conditionalMatches()).toEqual([false, true, false, true])
  await expect(canvas()).toHaveAttribute('data-painter-test-mount', 'original')
})

test('cross-sheet boundary painting retains an independent formula origin through undo and repeated painting', async () => {
  await page.evaluate(() => window.painterFixture.seed('conditional-boundary')); await select('A2:A5'); await arm(true)
  const original = await conditionalRules(), sourceSession = (await state()).sessionId
  await page.evaluate(() => window.painterFixture.activateSheet('painter-second')); await clickCell('A1')
  await expect.poll(() => conditionalMatches('painter-second')).toEqual([false, true, false, true])
  const copied = (await conditionalRules('painter-second'))[0]
  assert.notEqual(copied.id, original[0].id); assert.equal(copied.formula1, '=A1>5')
  assert.deepEqual(copied.formulaAnchor, { row: 1, col: 0 })
  assert.deepEqual(copied.ranges, [{ startRow: 0, endRow: 3, startCol: 0, endCol: 0 }])
  assert.deepEqual(copied.config, original[0].config); assert.deepEqual(copied.vendor, original[0].vendor)
  await clickCell('A1'); assert.equal((await history()).undo, 1); assert.equal((await conditionalRules('painter-second'))[0].id, copied.id)
  assert.deepEqual(await conditionalRules(), original); assert.equal((await state()).sessionId, sourceSession)
  await grid().press('Escape'); await command('sheet.command.undo')
  await expect.poll(() => conditionalMatches('painter-second')).toEqual([false, false, false, false])
  await command('sheet.command.redo'); await expect.poll(() => conditionalMatches('painter-second')).toEqual([false, true, false, true])
  assert.deepEqual(await valueEvents(), [])
})

test('standalone boundary painting keeps per-cell reference recovery and suppresses duplicate model emissions', async () => {
  await page.goto(baseURL + '?standalone'); await page.evaluate(() => window.painterFixture.seed('conditional-boundary')); await select('A2:A5')
  const before = contents(await snapshot()); await arm(true); await clickCell('A1')
  await expect.poll(() => conditionalMatches()).toEqual([false, true, false, true])
  const painted = await conditionalRules(); assert.deepEqual(painted[0].formulaAnchor, { row: 1, col: 0 })
  assert.deepEqual(contents(await snapshot()), before); assert.equal((await history()).emits, 1)
  await clickCell('A1'); assert.deepEqual(await conditionalRules(), painted); assert.equal((await history()).emits, 1)
  await expect(grid()).toHaveClass(/format-painter-active/); await grid().press('Escape'); await expect(grid()).not.toHaveClass(/format-painter-active/)
  await expect(grid()).toBeFocused()
})

test('condition-panel formula and range edits preserve the origin, extensions and reversible conditional drawing', async () => {
  await page.evaluate(() => window.painterFixture.seed('conditional-boundary')); await select('G7'); const before = await snapshot()
  await openConditionalPanel(); await editConditionalRule('A1:A4', '=A1>8'); await applyConditionalPanel(); await expectConditionalClosed()
  await expect.poll(() => conditionalMatches()).toEqual([false, false, false, true])
  const edited = (await conditionalRules())[0]
  assert.equal(edited.id, 'painter-cf-boundary'); assert.equal(edited.formula1, 'A1>8'); assert.deepEqual(edited.formulaAnchor, { row: 1, col: 0 })
  assert.deepEqual(edited.ranges, [{ startRow: 0, endRow: 3, startCol: 0, endCol: 0 }])
  for (const field of ['vendor', 'config', 'metadata']) assert.deepEqual(edited[field], before.sheets[0].conditionalFormats[0][field])
  assert.deepEqual(edited.style.vendor, before.sheets[0].conditionalFormats[0].style.vendor)
  assert.deepEqual(contents(await snapshot()), contents(before)); assert.equal((await history()).undo, 1); assert.deepEqual(await valueEvents(), [])
  assert.equal((await page.evaluate(() => window.painterFixture.commands())).filter(event => event.id === conditionalCommand).length, 1)
  await page.getByRole('tab', { name: '开始', exact: true }).click(); await command('sheet.command.undo')
  assert.deepEqual(await conditionalRules(), before.sheets[0].conditionalFormats)
  await command('sheet.command.redo'); await expect.poll(() => conditionalMatches()).toEqual([false, false, false, true])
  const saved = await snapshot(); await page.evaluate(model => window.painterFixture.reopen(model), saved)
  await expect.poll(() => conditionalRules()).toEqual([edited]); await expect.poll(() => conditionalMatches()).toEqual([false, false, false, true])
})

test('applying an unchanged edited rule keeps the imported representation and adds no revision or history', async () => {
  await page.evaluate(() => window.painterFixture.seed('conditional-boundary')); await select('G7'); const before = await snapshot(), beforeRevision = await revision()
  await openConditionalPanel(); await editConditionalRule(); await applyConditionalPanel(); await expectConditionalClosed()
  assert.deepEqual(await snapshot(), before); assert.equal(await revision(), beforeRevision)
  assert.equal((await history()).undo, 0); assert.deepEqual(await valueEvents(), [])
  assert.equal((await page.evaluate(() => window.painterFixture.commands())).filter(event => event.id === conditionalCommand).length, 1)
})

test('canceled condition-panel submissions keep the complete local draft and retry through the registered command', async () => {
  await page.evaluate(() => window.painterFixture.seed('conditional-boundary')); await select('G7'); const before = await snapshot()
  await openConditionalPanel(); await editConditionalRule('A1:A4', '=A1>8')
  await page.evaluate(id => window.painterFixture.cancel(id, true), conditionalCommand); await applyConditionalPanel()
  await expect(conditionalPanel().getByRole('alert')).toContainText('操作已取消')
  assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0)
  await editConditionalRule(); await expect(conditionalPanel().getByRole('textbox', { name: '应用范围', exact: true })).toHaveValue('A1:A4')
  await expect(conditionalPanel().getByRole('textbox', { name: '公式', exact: true })).toHaveValue('A1>8')
  await page.evaluate(id => window.painterFixture.cancel(id, false), conditionalCommand); await applyConditionalPanel(); await expectConditionalClosed()
  await expect.poll(() => conditionalMatches()).toEqual([false, false, false, true]); assert.equal((await history()).undo, 1)
  assert.equal((await page.evaluate(() => window.painterFixture.commands())).filter(event => event.id === conditionalCommand).length, 1)
  assert.deepEqual(await valueEvents(), []); assert.deepEqual(contents(await snapshot()), contents(before))
})

test('untouched protected rule scope allows a panel no-op but prevents moving that rule from an unlocked selection', async () => {
  await page.evaluate(() => { window.painterFixture.seed('conditional-boundary'); window.painterFixture.protect({ startRow: 2, endRow: 2, startCol: 0, endCol: 0 }) })
  await select('G7'); const before = await snapshot(); await openConditionalPanel(); await applyConditionalPanel(); await expectConditionalClosed()
  assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0)
  await openConditionalPanel(); await editConditionalRule('G7', '=A1>8'); await applyConditionalPanel()
  await expect(conditionalPanel().getByRole('alert')).toContainText('受保护')
  assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0); assert.deepEqual(await valueEvents(), [])
  await editConditionalRule(); await expect(conditionalPanel().getByRole('textbox', { name: '应用范围', exact: true })).toHaveValue('G7')
  await conditionalPanel().getByRole('button', { name: '取消', exact: true }).click(); await expectConditionalClosed()
})

test('deleting a rule from the panel validates its former protected coverage before modifying the model', async () => {
  await page.evaluate(() => { window.painterFixture.seed('conditional-boundary'); window.painterFixture.protect({ startRow: 2, endRow: 2, startCol: 0, endCol: 0 }) })
  await select('G7'); const before = await snapshot(); await openConditionalPanel()
  await conditionalPanel().getByRole('button', { name: '删除规则', exact: true }).click(); await applyConditionalPanel()
  await expect(conditionalPanel().getByRole('alert')).toContainText('受保护')
  await expect(conditionalPanel().locator('.conditional-rule-select')).toHaveCount(0)
  assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0)
  await conditionalPanel().getByRole('button', { name: '取消', exact: true }).click(); await expectConditionalClosed()
})

test('condition-panel list actions keep independent protected priorities and reject a move across the protected rule', async () => {
  for (const action of ['move', 'remove', 'prepend', 'cross-locked']) {
    await page.evaluate(() => {
      window.painterFixture.seed('conditional-independent')
      window.painterFixture.protect({ startRow: 10, endRow: 10, startCol: 13, endCol: 13 })
    })
    await select('G7'); const before = await snapshot(), beforeRules = before.sheets[0].conditionalFormats, beforeRevision = await revision()
    await openConditionalPanel()
    const rows = conditionalPanel().locator('.conditional-rule-row')
    if (action === 'move' || action === 'cross-locked') {
      await rows.nth(action === 'move' ? 0 : 1).getByRole('button', { name: '下移规则', exact: true }).click()
    } else if (action === 'remove') {
      await rows.nth(0).getByRole('button', { name: '删除规则', exact: true }).click()
    } else {
      await conditionalPanel().getByRole('button', { name: '新建规则', exact: true }).click()
      await expect(conditionalPanel().getByRole('textbox', { name: '应用范围', exact: true })).toHaveValue('G7')
      await conditionalPanel().getByRole('textbox', { name: '数值或公式', exact: true }).fill('5')
      await conditionalPanel().getByRole('button', { name: '保存规则', exact: true }).click()
      await expect(rows).toHaveCount(4)
    }
    assert.deepEqual(await snapshot(), before); assert.equal(await revision(), beforeRevision); assert.equal((await history()).undo, 0)
    await applyConditionalPanel()
    if (action === 'cross-locked') {
      await expect(conditionalPanel().getByRole('alert')).toContainText('受保护')
      assert.deepEqual(await snapshot(), before); assert.equal(await revision(), beforeRevision); assert.equal((await history()).undo, 0)
      await editConditionalRule(undefined, undefined, 1)
      await expect(conditionalPanel().getByRole('textbox', { name: '应用范围', exact: true })).toHaveValue('N11')
      await conditionalPanel().getByRole('button', { name: '取消', exact: true }).click(); await expectConditionalClosed()
      assert.deepEqual(await valueEvents(), [])
      assert.equal((await page.evaluate(() => window.painterFixture.commands())).filter(event => event.id === conditionalCommand).length, 0)
      continue
    }
    await expectConditionalClosed(); const afterRules = await conditionalRules()
    assert.deepEqual(afterRules.find(rule => rule.id === 'cf-independent-locked'), beforeRules[2])
    if (action === 'move') assert.deepEqual(afterRules, [{ ...beforeRules[1], priority: 1 }, { ...beforeRules[0], priority: 2 }, beforeRules[2]])
    else if (action === 'remove') assert.deepEqual(afterRules, beforeRules.slice(1))
    else {
      assert.deepEqual(afterRules.slice(1), beforeRules); assert.equal(afterRules[0].priority, 1); assert.equal(afterRules[0].formula1, '5')
      assert.equal(beforeRules.some(rule => rule.id === afterRules[0].id), false)
      assert.deepEqual(afterRules[0].ranges, [{ startRow: 6, endRow: 6, startCol: 6, endCol: 6 }])
    }
    assert.deepEqual(contents(await snapshot()), contents(before)); assert.equal((await history()).undo, 1); assert.deepEqual(await valueEvents(), [])
    assert.notEqual(await revision(), beforeRevision)
    assert.equal((await page.evaluate(() => window.painterFixture.commands())).filter(event => event.id === conditionalCommand).length, 1)
    await page.getByRole('tab', { name: '开始', exact: true }).click(); await command('sheet.command.undo')
    assert.deepEqual(await conditionalRules(), beforeRules)
    await command('sheet.command.redo'); assert.deepEqual(await conditionalRules(), afterRules)
  }
})

test('condition-panel revisions reject external edits while preserving both the new document and the unsaved draft', async () => {
  await page.evaluate(() => window.painterFixture.seed('conditional-boundary')); await select('G7'); await openConditionalPanel(); await editConditionalRule('A1:A4', '=A1>8')
  await page.evaluate(() => window.painterFixture.mutate()); const changed = await snapshot(), changedValues = await valueEvents()
  await applyConditionalPanel(); await expect(conditionalPanel().getByRole('alert')).toContainText('工作簿已更改')
  assert.deepEqual(await snapshot(), changed); assert.equal((await history()).undo, 1); assert.deepEqual(await valueEvents(), changedValues)
  await editConditionalRule(); await expect(conditionalPanel().getByRole('textbox', { name: '应用范围', exact: true })).toHaveValue('A1:A4')
  await expect(conditionalPanel().getByRole('textbox', { name: '公式', exact: true })).toHaveValue('A1>8')
  await conditionalPanel().getByRole('button', { name: '取消', exact: true }).click(); await expectConditionalClosed()
})

test('standalone condition-panel edits emit one complete model and unchanged reapplication emits nothing', async () => {
  await page.goto(baseURL + '?standalone'); await page.evaluate(() => window.painterFixture.seed('conditional-boundary')); await select('G7')
  const before = contents(await snapshot()); await openConditionalPanel(); await editConditionalRule('A1:A4', '=A1>8'); await applyConditionalPanel(); await expectConditionalClosed()
  await expect.poll(() => conditionalMatches()).toEqual([false, false, false, true]); assert.equal((await history()).emits, 1)
  assert.deepEqual(contents(await snapshot()), before); const edited = await snapshot()
  await openConditionalPanel(); await editConditionalRule(); await applyConditionalPanel(); await expectConditionalClosed()
  assert.deepEqual(await snapshot(), edited); assert.equal((await history()).emits, 1)
  assert.deepEqual((await conditionalRules())[0].formulaAnchor, { row: 1, col: 0 })
})

test('standalone condition-panel drafts reject an externally replaced workbook without emitting an edit', async () => {
  await page.goto(baseURL + '?standalone'); await page.evaluate(() => window.painterFixture.seed('conditional-boundary')); await select('G7')
  await openConditionalPanel(); await editConditionalRule('A1:A4', '=A1>8'); await page.evaluate(() => window.painterFixture.mutateStandalone())
  const changed = await snapshot(); await applyConditionalPanel(); await expect(conditionalPanel().getByRole('alert')).toContainText('工作簿已更改')
  assert.deepEqual(await snapshot(), changed); assert.equal((await history()).emits, 0)
  await conditionalPanel().getByRole('button', { name: '取消', exact: true }).click(); await expectConditionalClosed()
})

test('condition-panel library inputs retain a draft across theme changes and cancel restores Canvas focus', async () => {
  await page.evaluate(() => { window.painterFixture.seed('conditional-boundary'); window.painterFixture.appearance(true) }); await select('G7')
  const before = await snapshot(); await openConditionalPanel(); await editConditionalRule('A1:A4', '=A1>8')
  const rangeInput = conditionalPanel().getByRole('textbox', { name: '应用范围', exact: true })
  await expect(rangeInput).toHaveClass(/ant-input/); const darkColor = await rangeInput.evaluate(element => getComputedStyle(element).color)
  await page.evaluate(() => window.painterFixture.appearance(false))
  await expect.poll(() => rangeInput.evaluate(element => getComputedStyle(element).color)).not.toBe(darkColor)
  await expect(rangeInput).toHaveValue('A1:A4'); await expect(conditionalPanel().getByRole('textbox', { name: '公式', exact: true })).toHaveValue('=A1>8')
  await conditionalPanel().getByRole('button', { name: '取消', exact: true }).click(); await expectConditionalClosed()
  assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0)
})

test('condition-panel style edits preserve unedited imported expression and numeric grammar', async () => {
  await page.evaluate(() => window.painterFixture.seed('conditional-imported')); await select('G7')
  for (const [index, name] of [[0, '公式'], [1, '数值或公式']]) {
    const original = (await conditionalRules())[index]
    await openConditionalPanel(); await editConditionalRule(undefined, undefined, index)
    await expect(conditionalPanel().getByRole('textbox', { name, exact: true })).toHaveValue('=UNSUPPORTED([VendorColumn])')
    await conditionalPanel().getByRole('checkbox', { name: '加粗', exact: true }).check(); await applyConditionalPanel(); await expectConditionalClosed()
    const edited = (await conditionalRules())[index]
    assert.equal(edited.formula1, original.formula1); assert.equal(edited.type, original.type); assert.equal(edited.subType, original.subType)
    assert.equal(edited.style.bold, true); assert.deepEqual(edited.metadata, original.metadata)
  }
  assert.equal((await history()).undo, 2); assert.deepEqual(await valueEvents(), [])
})

test('imported formula-value rules expose the expression editor and retain their storage contract after a range edit', async () => {
  await page.evaluate(() => window.painterFixture.seed('conditional-value-formula')); await select('G7'); const before = contents(await snapshot())
  await openConditionalPanel(); await editConditionalRule()
  await expect(conditionalPanel().getByRole('textbox', { name: '公式', exact: true })).toHaveValue('=A1>5')
  await conditionalPanel().getByRole('textbox', { name: '应用范围', exact: true }).fill('A1:A4')
  await conditionalPanel().getByRole('textbox', { name: '公式', exact: true }).fill('=A1>8'); await applyConditionalPanel(); await expectConditionalClosed()
  const edited = (await conditionalRules())[0]
  assert.equal(edited.type, 'highlightCell'); assert.equal(edited.subType, 'formula'); assert.equal(edited.formula1, ''); assert.equal(edited.value, 'A1>8')
  assert.deepEqual(edited.formulaAnchor, { row: 1, col: 0 }); await expect.poll(() => conditionalMatches()).toEqual([false, false, false, true])
  assert.deepEqual(contents(await snapshot()), before); assert.equal((await history()).undo, 1)
})

test('switching sheets closes condition-panel drafts before they can be submitted to a different target', async () => {
  await page.evaluate(() => window.painterFixture.seed('conditional-boundary')); await select('G7'); const before = await conditionalRules()
  await openConditionalPanel(); await editConditionalRule('A1:A4', '=A1>8')
  await page.evaluate(() => window.painterFixture.activateSheet('painter-second')); await expect(conditionalPanel()).toBeHidden()
  assert.deepEqual(await conditionalRules(), before); assert.deepEqual(await conditionalRules('painter-second') ?? [], []); assert.equal((await history()).undo, 0)
  await openConditionalPanel(); await expect(conditionalPanel().locator('.conditional-rule-select')).toHaveCount(0)
  await conditionalPanel().getByRole('button', { name: '取消', exact: true }).click(); await expectConditionalClosed()
})

test.each(['classic', 'grid'].flatMap(variant => ['list', 'listMultiple'].flatMap(type =>
  ['inline', 'reference'].map(source => ({ variant, type, source })),
)))('$variant $type $source management preserves option colors through edits, history and snapshot reopen', async ({ variant, type, source }) => {
  await page.evaluate(value => window.painterFixture.variant(value), variant)
  await page.evaluate(rule => window.painterFixture.command('sheet.command.set-data-validation', undefined, { rule }), {
    id: 'colored-validation', type,
    ranges: [{ startRow: 3, endRow: 4, startCol: 3, endCol: 3 }],
    formula1: source === 'reference' ? '=$A$1:$C$1' : '"Red,Green,Blue"',
    formula2: '#dc3344,,#336699', renderMode: 'custom', allowBlank: true,
    showDropDown: true, showErrorMessage: true, errorStyle: 'stop', showInputMessage: true,
    promptTitle: '选项颜色', prompt: '保留现有选项颜色', metadata: { vendor: { keep: ['colors'] } },
  })
  await select('D4')
  const before = await snapshot()
  const original = before.sheets[0].dataValidations[0]
  assert.equal((await history()).undo, 1)
  await page.getByRole('tab', { name: '数据', exact: true }).click()
  await command('sheet.menu.data-validation')
  await page.getByRole('menuitem', { name: '管理验证规则', exact: true }).click()
  const panel = page.getByRole('complementary', { name: '数据验证', exact: true })
  await panel.locator('.validation-rule-main').first().click()
  const rangeInput = panel.getByRole('textbox', { name: '数据验证应用区域', exact: true })
  await rangeInput.fill('D4:E5')
  await rangeInput.focus()
  await expect(rangeInput).toHaveClass(/ant-input/)
  const lightColor = await rangeInput.evaluate(element => getComputedStyle(element).color)
  await page.evaluate(() => window.painterFixture.appearance(true))
  await expect.poll(() => rangeInput.evaluate(element => getComputedStyle(element).color)).not.toBe(lightColor)
  await expect(rangeInput).toHaveValue('D4:E5')
  await expect(rangeInput).toBeFocused()
  await panel.getByRole('button', { name: /应\s*用/ }).click()
  await expect(panel.locator('.validation-rule-main')).toHaveCount(1)
  const edited = (await snapshot()).sheets[0].dataValidations[0]
  assert.equal(edited.id, original.id)
  assert.equal(edited.formula1, original.formula1)
  assert.equal(edited.formula2, original.formula2)
  assert.deepEqual(edited.metadata, original.metadata)
  assert.deepEqual(edited.ranges, [{ startRow: 3, endRow: 4, startCol: 3, endCol: 4 }])
  assert.equal((await history()).undo, 2)
  await panel.locator('.validation-rule-main').first().click()
  await panel.locator('#sheet-validation-type').click()
  await page.locator('.ant-select-dropdown:visible').getByText(type === 'list' ? '下拉多选' : '下拉列表', { exact: true }).click()
  await panel.getByRole('button', { name: /应\s*用/ }).click()
  await expect(panel.locator('.validation-rule-main')).toHaveCount(1)
  const converted = (await snapshot()).sheets[0].dataValidations[0]
  assert.equal(converted.type, type === 'list' ? 'listMultiple' : 'list')
  assert.equal(converted.id, original.id)
  assert.equal(converted.formula1, original.formula1)
  assert.equal(converted.formula2, original.formula2)
  assert.deepEqual(converted.metadata, original.metadata)
  assert.equal((await history()).undo, 3)
  assert.equal((await page.evaluate(() => window.painterFixture.commands())).filter(event => event.id === 'sheet.command.replace-data-validation').length, 2)
  await panel.getByRole('button', { name: '关闭', exact: true }).click()
  await page.getByRole('tab', { name: '开始', exact: true }).click()
  await command('sheet.command.undo')
  assert.deepEqual((await snapshot()).sheets[0].dataValidations[0], edited)
  await command('sheet.command.undo')
  assert.deepEqual((await snapshot()).sheets[0].dataValidations[0], original)
  await command('sheet.command.redo')
  await command('sheet.command.redo')
  assert.deepEqual((await snapshot()).sheets[0].dataValidations[0], converted)
  assert.deepEqual(contents(await snapshot()), contents(before))
  await page.evaluate(model => window.painterFixture.reopen(model), await snapshot())
  await expect.poll(async () => (await snapshot()).sheets[0].dataValidations[0]).toEqual(converted)
})

test.each(['classic', 'grid'].flatMap(variant => ['whole', 'checkbox', 'list', 'listMultiple'].map(type => ({ variant, type }))))(
  '$variant $type conversion does not invent list colors from a numeric or checkbox operand', async ({ variant, type }) => {
    await page.evaluate(value => window.painterFixture.variant(value), variant)
    await page.evaluate(rule => window.painterFixture.command('sheet.command.set-data-validation', undefined, { rule }), {
      id: 'uncolored-validation', type,
      ranges: [{ startRow: 3, endRow: 4, startCol: 3, endCol: 3 }],
      ...(type === 'whole' ? { operator: 'between', formula1: '1', formula2: '10' } : {}),
      ...(type === 'checkbox' ? { formula1: 'TRUE', formula2: 'FALSE' } : {}),
      ...(type === 'list' || type === 'listMultiple' ? { formula1: '"Red,Green,Blue"' } : {}),
    })
    await select('D4')
    await page.getByRole('tab', { name: '数据', exact: true }).click()
    await command('sheet.menu.data-validation')
    await page.getByRole('menuitem', { name: '管理验证规则', exact: true }).click()
    const panel = page.getByRole('complementary', { name: '数据验证', exact: true })
    await panel.locator('.validation-rule-main').first().click()
    await panel.locator('#sheet-validation-type').click()
    await page.locator('.ant-select-dropdown:visible').getByText(type === 'list' ? '下拉多选' : '下拉列表', { exact: true }).click()
    await panel.getByRole('button', { name: /应\s*用/ }).click()
    await expect(panel.locator('.validation-rule-main')).toHaveCount(1)
    const converted = (await snapshot()).sheets[0].dataValidations[0]
    assert.equal(converted.type, type === 'list' ? 'listMultiple' : 'list')
    assert.equal(converted.formula2, undefined)
    assert.equal((await history()).undo, 2)
  },
)
