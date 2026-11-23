import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

// Deferred DOC-06 browser drafts. ClipboardEvent and CompositionEvent assertions
// cover event wiring only; system clipboard permissions and real Windows IME need separate acceptance.
const fixturePath = '/__document-table-workbench-test__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>Docs table contracts</title><style>body{margin:0}#app{height:100vh}</style></head><body><div id="app"></div><script type="module">
import { createApp, h, ref, shallowRef } from 'vue'
import { ConfigProvider, theme } from 'antdv-next'
import { OfficeStudio } from '/packages/office-studio/src/api/office-studio'
import { FDocument } from '/packages/office-studio/src/api/facade'
import { FDocumentTable } from '/packages/office-studio/src/api/document-tables'
import DocumentWorkbench from '/packages/office-studio/src/components/DocumentWorkbench.vue'
import { createBlankDocument } from '/packages/office-studio/src/domain/model'
import { blocksFromHtml, createParagraph, createRun, createTable, documentHtml, normalizeDocument, tableForBlock, runText } from '/packages/office-studio/src/features/document/model/document-model'
import { documentTableGrid } from '/packages/office-studio/src/features/document/model/document-table-grid'
import { createTextMeasurer, layoutDocument } from '/packages/office-studio/src/features/document/layout/document-layout'
import { sceneOffset } from '/packages/office-studio/src/features/document/rendering/document-scene'
const copy = value => JSON.parse(JSON.stringify(value))
const initial = createBlankDocument(); initial.id = 'table-browser'; initial.name = '表格行为草稿'
initial.page = { width: 620, height: 460, marginTop: 40, marginBottom: 40, marginLeft: 40, marginRight: 40 }
const block = createParagraph(); block.id = 'table'; block.type = 'table'; block.runs = []
const table = createTable(10, 3); table.columnWidths = [140, 140, 140]; table.headerRowCount = 1
table.style = { layout: 'fixed', width: 420, cellMargin: { top: 4, right: 4, bottom: 4, left: 4 } }
table.rows.forEach((row, r) => { row.id = 'row-' + r; row.height = r % 2 ? 74 : 54; row.heightRule = r % 2 ? 'atLeast' : 'exact'; row.allowSplit = r % 2 === 1
  row.cells.forEach((cell, c) => { cell.id = 'cell-' + r + '-' + c; cell.blocks = [createParagraph('R' + r + 'C' + c)]; cell.blocks[0].id = 'p-' + r + '-' + c
    cell.style.margin = { top: 3 + r, right: 4 + c, bottom: 5 + r, left: 6 + c }
    cell.style.borders = { top: { width: c + 1, color: c ? '#008800' : '#880000', style: c ? 'dash' : 'dot' } }
  })
})
table.columnTypes = { 2: { type: 'dropdown', dropdownOptions: [{ id: 'first', label: '同名', color: '#ff0000' }, { id: 'second', label: '同名', color: '#0000ff' }] } }
table.rows[0].cells[2].blocks[0].runs = [createRun('同名')]; table.rows[0].cells[2].style.selectedDropdownOptionId = 'first'
block.metadata = { table }; initial.blocks = [createParagraph('Before'), block, createParagraph('After')]; initial.blocks[0].id = 'before'; initial.blocks[2].id = 'after'
let studio = new OfficeStudio({ document: normalizeDocument(initial), currentUserId: 'editor' }); studio.activateDocument(initial.id)
const api = shallowRef(new FDocument(studio, initial.id)), model = shallowRef(api.value.getModel()), dark = ref(false), variant = ref('classic'), notices = [], commands = []
let unsubscribe
function observe() { unsubscribe?.(); unsubscribe = api.value.onModelChanged(event => { model.value = api.value.getModel(); commands.push(event.commandId) }) }
observe()
const tableApi = () => new FDocumentTable(api.value, 'table')
function geometry(row, column, page) {
  const layout = layoutDocument(model.value, createTextMeasurer()), scroll = document.querySelector('.document-scroll'), bounds = scroll.getBoundingClientRect()
  const owner = documentTableGrid(tableForBlock(model.value.blocks.find(block => block.id === 'table'))).cellAt(row, column)
  const cell = layout.cells.find(cell => cell.cellId === owner.cell.id && (page === undefined ? !cell.repeatedHeader : cell.page === page))
  if (!cell) throw new Error('Missing fixture cell')
  const visible = cell.clip ? { x: Math.max(cell.x, cell.clip.x), y: Math.max(cell.y, cell.clip.y), width: Math.min(cell.x + cell.width, cell.clip.x + cell.clip.width) - Math.max(cell.x, cell.clip.x), height: Math.min(cell.y + cell.height, cell.clip.y + cell.clip.height) - Math.max(cell.y, cell.clip.y) } : cell
  const offset = sceneOffset(layout, { width: scroll.clientWidth, height: scroll.clientHeight, scrollTop: scroll.scrollTop, scrollLeft: scroll.scrollLeft, zoom: 100 })
  return { x: bounds.left + offset.x + visible.x + visible.width / 2, y: bounds.top + offset.y + visible.y + visible.height / 2 }
}
window.documentTableFixture = {
  snapshot: () => copy(model.value), commands: () => [...commands], notices: () => [...notices], point: geometry,
  history: () => ({ undo: studio.editor.history.units[api.value.id]?.undo.length ?? 0, redo: studio.editor.history.units[api.value.id]?.redo.length ?? 0 }),
  select: range => tableApi().selectRange(range),
  selectText: (row, column, offset = 0) => { const source = tableForBlock(model.value.blocks.find(block => block.id === 'table')), owner = documentTableGrid(source).cellAt(row, column), paragraph = owner.cell.blocks[0]; api.value.setInputSelection({ anchor: { blockId: 'table', row: owner.row, col: owner.cellIndex, cellBlockId: paragraph.id, offset }, focus: { blockId: 'table', row: owner.row, col: owner.cellIndex, cellBlockId: paragraph.id, offset } }) },
  appearance: (isDark, layout) => { dark.value = isDark; variant.value = layout },
  scrollTo: (row, column) => { const layout = layoutDocument(model.value, createTextMeasurer()), source = tableForBlock(model.value.blocks.find(block => block.id === 'table')), owner = documentTableGrid(source).cellAt(row, column), cell = layout.cells.find(cell => cell.cellId === owner.cell.id && !cell.repeatedHeader), scroll = document.querySelector('.document-scroll'); scroll.scrollTop = Math.max(0, cell.y - 100); scroll.dispatchEvent(new Event('scroll')) },
  clipboard: (type, text = '', html = '') => { const input = document.querySelector('.document-input-bridge textarea') ?? document.querySelector('textarea.document-input-bridge'), data = new DataTransfer(); if (text) data.setData('text/plain', text); if (html) data.setData('text/html', html); const event = new ClipboardEvent(type, { clipboardData: data, bubbles: true, cancelable: true }); input.dispatchEvent(event); return { prevented: event.defaultPrevented, text: data.getData('text/plain'), html: data.getData('text/html') } },
  htmlRoundtrip: html => { const blocks = blocksFromHtml(html), first = tableForBlock(blocks[0]); const serialized = documentHtml({ ...model.value, blocks }); const reopened = blocksFromHtml(serialized); return { source: copy(first), html: serialized, reopened: copy(tableForBlock(reopened[0])) } },
  saveReopen: () => { const saved = copy(api.value.save()); unsubscribe?.(); studio.dispose(); studio = new OfficeStudio({ document: saved, currentUserId: 'editor' }); studio.activateDocument(saved.id); api.value = new FDocument(studio, saved.id); model.value = api.value.getModel(); observe(); return saved },
}
const app = createApp({ render: () => h(ConfigProvider, { theme: { algorithm: dark.value ? theme.darkAlgorithm : theme.defaultAlgorithm, token: { colorPrimary: '#3064c7' } } }, { default: () => h(DocumentWorkbench, { modelValue: model.value, documentApi: api.value, canUndo: studio.editor.canUndo, canRedo: studio.editor.canRedo, ribbonVariant: variant.value,
  'onUpdate:modelValue': value => { model.value = value }, onUndo: () => api.value.undo(), onRedo: () => api.value.redo(), onNotify: message => notices.push(message) }) }) })
app.mount('#app')
window.addEventListener('pagehide', () => { unsubscribe?.(); app.unmount(); studio.dispose() }, { once: true })
</script></body></html>`

let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'document-table-workbench-fixture',
    configureServer(vite) { vite.middlewares.use(fixturePath, async (_req, res, next) => { try { res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.end(await vite.transformIndexHtml(fixturePath, fixture)) } catch (error) { next(error) } }) },
  }], server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen()
  const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate an HTTP port')
  baseURL = 'http://127.0.0.1:' + address.port + fixturePath
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, { timeout: 60_000 })
afterAll(async () => { await browser?.close(); await server?.close() })
beforeEach(async () => { errors = []; page = await browser.newPage({ viewport: { width: 1440, height: 900 }, deviceScaleFactor: 1.25 }); page.on('pageerror', error => errors.push(error.message)); await page.goto(baseURL); await page.waitForFunction(() => !!window.documentTableFixture); await expect(page.getByRole('textbox', { name: '文档文字输入', exact: true })).toBeAttached() })
afterEach(async () => { await page?.close(); assert.deepEqual(errors, []) })

const rect = (startRow, startColumn, endRow = startRow, endColumn = startColumn) => ({ startRow, endRow, startColumn, endColumn })
const snapshot = () => page.evaluate(() => window.documentTableFixture.snapshot())
const history = () => page.evaluate(() => window.documentTableFixture.history())
const selected = range => page.evaluate(range => window.documentTableFixture.select(range), range)
async function hoverCell(row = 0, column = 0, pageIndex) { const point = await page.evaluate(({ row, column, pageIndex }) => window.documentTableFixture.point(row, column, pageIndex), { row, column, pageIndex }); await page.mouse.move(point.x, point.y); return point }
async function contextProperty(title, row = 0, column = 0) { const point = await hoverCell(row, column); await page.mouse.click(point.x, point.y, { button: 'right' }); await page.getByRole('menuitem', { name: title, exact: true }).click(); return page.getByRole('dialog') }
const rowSource = document => document.blocks.find(block => block.id === 'table').metadata.table

test('row and column handles select real structural ranges and insert dots create one history item', async () => {
  await selected(rect(0, 0)); await hoverCell()
  await page.getByRole('button', { name: '选择第 1 列', exact: true }).first().click()
  await expect(page.getByRole('tab', { name: '表格', exact: true })).toBeVisible()
  const before = await snapshot()
  await page.getByRole('button', { name: '在第 1 行上方插入', exact: true }).first().click()
  await expect.poll(async () => rowSource(await snapshot()).rows.length).toBe(11)
  assert.equal((await history()).undo, 1)
  assert.equal(rowSource(await snapshot()).rows[1].cells[0].id, rowSource(before).rows[0].cells[0].id)
})

test('resize previews defer publication, commit once, and Escape cancels a later drag', async () => {
  await selected(rect(0, 0)); await hoverCell()
  const source = rowSource(await snapshot()), button = page.getByRole('button', { name: '调整第 1 列宽度', exact: true }).first(), box = await button.boundingBox()
  await page.mouse.move(box.x + box.width / 2, box.y + 30); await page.mouse.down(); await page.mouse.move(box.x + box.width / 2 + 30, box.y + 30, { steps: 5 })
  assert.deepEqual(rowSource(await snapshot()).columnWidths, source.columnWidths)
  await page.mouse.up()
  await expect.poll(async () => rowSource(await snapshot()).columnWidths[0]).toBe(source.columnWidths[0] + 30)
  assert.equal((await history()).undo, 1)
  await hoverCell()
  const second = await button.boundingBox(), saved = await snapshot()
  await page.mouse.move(second.x + second.width / 2, second.y + 30); await page.mouse.down(); await page.mouse.move(second.x + 50, second.y + 30, { steps: 4 }); await page.keyboard.press('Escape'); await page.mouse.up()
  assert.deepEqual(await snapshot(), saved); assert.equal((await history()).undo, 1)
})

test('dragging at the viewport edge scrolls across pages and cancellation retains the table model', async () => {
  await selected(rect(0, 0, 0, 2)); await hoverCell()
  const before = await snapshot(), handle = await page.getByRole('button', { name: '移动第 1 行', exact: true }).first().boundingBox(), viewport = await page.locator('.document-scroll').boundingBox()
  const scrollTop = await page.locator('.document-scroll').evaluate(element => element.scrollTop)
  await page.mouse.move(handle.x + handle.width / 2, handle.y + handle.height / 2); await page.mouse.down(); await page.mouse.move(viewport.x + viewport.width / 2, viewport.y + viewport.height - 5, { steps: 10 })
  await expect.poll(() => page.locator('.document-scroll').evaluate(element => element.scrollTop)).toBeGreaterThan(scrollTop)
  await page.keyboard.press('Escape'); await page.mouse.up()
  assert.deepEqual(await snapshot(), before)
})

test('a row move across physical pages preserves paragraph IDs and creates one content transaction', async () => {
  await selected(rect(1, 0, 1, 2)); await hoverCell(1, 0)
  const before = rowSource(await snapshot()), handle = await page.getByRole('button', { name: '移动第 2 行', exact: true }).first().boundingBox()
  await page.mouse.move(handle.x + handle.width / 2, handle.y + handle.height / 2); await page.mouse.down()
  await page.evaluate(() => window.documentTableFixture.scrollTo(7, 0))
  const point = await page.evaluate(() => window.documentTableFixture.point(7, 0))
  await page.mouse.move(point.x, point.y + 10, { steps: 8 }); await page.mouse.up()
  await expect.poll(async () => rowSource(await snapshot()).rows[1].id).not.toBe(before.rows[1].id)
  const moved = rowSource(await snapshot()).rows.find(row => row.id === before.rows[1].id)
  assert.deepEqual(moved.cells.map(cell => cell.blocks[0].id), before.rows[1].cells.map(cell => cell.blocks[0].id))
  assert.equal((await history()).undo, 1)
})

test('mixed row and cell properties edit only the changed field and return focus to the input bridge', async () => {
  await selected(rect(0, 0, 1, 1))
  const before = rowSource(await snapshot()), dialog = await contextProperty('行高与跨页')
  await expect(dialog.getByPlaceholder('混合值（保持）')).toHaveCount(2)
  await dialog.getByLabel('行高（px）', { exact: true }).fill('90')
  await dialog.getByRole('button', { name: '应用', exact: true }).click()
  await expect(dialog).toBeHidden()
  assert.deepEqual(rowSource(await snapshot()).rows.slice(0, 2).map(row => [row.height, row.heightRule, row.allowSplit]), before.rows.slice(0, 2).map(row => [90, row.heightRule, row.allowSplit]))
  await expect(page.getByRole('textbox', { name: '文档文字输入', exact: true })).toBeFocused()
  await selected(rect(0, 0, 1, 1))
  const marginDialog = await contextProperty('单元格边距')
  await marginDialog.getByLabel('左内边距（px）', { exact: true }).fill('12')
  await marginDialog.getByRole('button', { name: '应用', exact: true }).click()
  const result = rowSource(await snapshot())
  for (let r = 0; r < 2; r++) for (let c = 0; c < 2; c++) assert.deepEqual(result.rows[r].cells[c].style.margin, { ...before.rows[r].cells[c].style.margin, left: 12 })
})

test('duplicate dropdown labels retain option IDs and theme changes keep the ant overlay interactive', async () => {
  await selected(rect(0, 2)); await hoverCell(0, 2)
  await page.getByRole('combobox', { name: '第 1 行下拉选项，当前 同名', exact: true }).click()
  await page.getByRole('option', { name: '同名', exact: true }).nth(1).click()
  await expect.poll(async () => rowSource(await snapshot()).rows[0].cells[2].style.selectedDropdownOptionId).toBe('second')
  const before = await snapshot(), handle = page.getByRole('button', { name: '选择整个表格', exact: true }).first()
  const light = await handle.evaluate(element => getComputedStyle(element).backgroundColor)
  await page.evaluate(() => window.documentTableFixture.appearance(true, 'grid')); await hoverCell(0, 2)
  await expect.poll(() => handle.evaluate(element => getComputedStyle(element).backgroundColor)).not.toBe(light)
  await handle.click(); await expect(page.getByRole('tab', { name: '表格', exact: true })).toBeVisible()
  assert.deepEqual(await snapshot(), before)
})

test('table copy emits TSV and HTML and rich paste commits a merged table footprint with independent continuation text', async () => {
  await selected(rect(0, 0, 1, 1))
  const copied = await page.evaluate(() => window.documentTableFixture.clipboard('copy'))
  assert.equal(copied.prevented, true); assert.match(copied.text, /R0C0/); assert.match(copied.html, /<table/)
  const html = '<table><tbody><tr><td rowspan="2"><p>Owner</p><div data-office-table-continuation="v1" data-office-row-offset="1"><p>Continued</p></div></td><td><p>Right</p><table><tr><td><p>Nested</p></td></tr></table></td></tr><tr><td><p>Bottom</p></td></tr></tbody></table>'
  await selected(rect(0, 0)); await page.evaluate(html => window.documentTableFixture.clipboard('paste', '', html), html)
  await expect.poll(async () => rowSource(await snapshot()).rows[0].cells[0].rowSpan).toBe(2)
  const result = rowSource(await snapshot()), owner = result.rows[0].cells[0]
  assert.equal(owner.style.nativeContinuationContent[1].blocks[0].runs[0].text, 'Continued')
  assert.equal(owner.style.nativeContinuationContent[1].rowId, result.rows[1].id)
  assert.equal(result.rows[0].cells[1].blocks.some(block => block.type === 'table'), true)
  assert.equal((await history()).undo, 1)
  const saved = await page.evaluate(() => window.documentTableFixture.saveReopen())
  assert.deepEqual((await snapshot()).blocks, saved.blocks)
})

test('HTML rowspan zero, multiple paragraphs and wrapped nested content survive the HTML model roundtrip', async () => {
  const html = '<table><colgroup><col style="width:80px"><col style="width:120px"></colgroup><tbody><tr><td rowspan="0" style="padding-left:6px;border-top:2px dashed #123456"><div><p>A</p><p>B</p><table><tr><td>Nested</td></tr></table><p>C</p></div></td><td>D</td></tr><tr><td>E</td></tr></tbody></table>'
  const result = await page.evaluate(html => window.documentTableFixture.htmlRoundtrip(html), html)
  assert.equal(result.source.rows[0].cells[0].rowSpan, 2)
  assert.deepEqual(result.source.rows[0].cells[0].blocks.map(block => block.type), ['paragraph', 'paragraph', 'table', 'paragraph'])
  assert.deepEqual(result.reopened.columnWidths, [80, 120])
  assert.equal(result.reopened.rows[0].cells[0].style.margin.left, 6)
  assert.deepEqual(result.reopened.rows[0].cells[0].style.borders.top, result.source.rows[0].cells[0].style.borders.top)
})

test('synthetic composition in a cell produces one committed edit and restores the Canvas input focus', async () => {
  await page.evaluate(() => window.documentTableFixture.selectText(0, 0, 2))
  const input = page.getByRole('textbox', { name: '文档文字输入', exact: true }); await input.focus()
  await input.evaluate(element => { element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' })); element.dispatchEvent(new CompositionEvent('compositionupdate', { bubbles: true, data: '中' })); element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '中文' })) })
  await expect.poll(async () => rowSource(await snapshot()).rows[0].cells[0].blocks[0].runs.map(run => run.text).join('')).toBe('R0中文C0')
  assert.equal((await history()).undo, 1)
  await expect(input).toBeFocused()
})
