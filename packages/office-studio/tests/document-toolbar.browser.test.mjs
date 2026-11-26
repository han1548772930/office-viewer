import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

const fixturePath = '/__document-toolbar-test__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>Docs toolbar interaction fixture</title><style>body{margin:0}#app{height:100vh;min-width:0}</style></head><body><div id="app"></div><script type="module">
import { createApp, h, ref } from 'vue'
import DocumentWorkbench from '/packages/office-studio/src/components/DocumentWorkbench.vue'
import { createBlankDocument } from '/packages/office-studio/src/domain/model'
import { applyNamedParagraphStyle, blocksFromHtml, createParagraph, documentHtml } from '/packages/office-studio/src/features/document/model/document-model'
import { createDocumentContentControl } from '/packages/office-studio/src/features/document/services/content-controls'
import { createTextMeasurer, hitTestChecklist, layoutDocument } from '/packages/office-studio/src/features/document/layout/document-layout'
import { sceneOffset } from '/packages/office-studio/src/features/document/rendering/document-scene'
const initial = createBlankDocument()
initial.id = 'toolbar-doc'; initial.name = '工具栏交互文档'
initial.blocks = [createParagraph('甲乙丙丁')]
const model = ref(initial)
const history = ref([])
const redo = ref([])
const copy = (value) => JSON.parse(JSON.stringify(value))
const update = (value) => { history.value.push(copy(model.value)); redo.value = []; model.value = value }
window.documentToolbarFixture = {
  snapshot: () => copy(model.value),
  namedStyleFixture: () => {
    const next = copy(model.value)
    next.blocks = [createParagraph('Title text'), createParagraph('Heading text')]
    applyNamedParagraphStyle(next.blocks[0], 'title')
    applyNamedParagraphStyle(next.blocks[1], 'heading2')
    model.value = next
  },
  paragraphFixture: () => {
    const next = copy(model.value)
    next.blocks = [createParagraph('First paragraph'), createParagraph('Second paragraph')]
    next.blocks[0].style = { ...next.blocks[0].style, align: 'right', lineHeight: 2, indentFirstLine: 20, spaceBefore: 6, spaceAfter: 12 }
    next.blocks[1].style = { ...next.blocks[1].style, align: 'center', lineHeight: 1, indentFirstLine: 80, spaceBefore: 36, spaceAfter: 0 }
    model.value = next
  },
  tablePageFixture: (kind) => {
    const next = copy(model.value)
    next.page = { width: 900, height: 600, marginTop: 60, marginBottom: 60, marginLeft: 60, marginRight: 40 }
    next.sectionSetup = { columnCount: kind === 'columns' ? 2 : 1, columnGap: 20 }
    const target = createParagraph('Table target')
    next.blocks = kind === 'continuous' ? [createParagraph('Previous section'),
      { ...createParagraph(), id: 'table-continuous', type: 'sectionBreak', sectionBreak: { type: 'continuous', setup: { marginLeft: 120, marginRight: 120, columnCount: 2, columnGap: 20 } } }, target] : [target]
    model.value = next
  },
  lockNamedStyleTail: () => {
    const next = copy(model.value)
    createDocumentContentControl(next, { id: 'locked-tail', type: 'text', blockId: next.blocks[0].id, startOffset: 2, endOffset: 4, lock: 'contentLocked' })
    model.value = next
  },
  reopenSnapshot: () => { model.value = JSON.parse(JSON.stringify(model.value)) },
  painterFixture: () => { const next = copy(model.value); next.blocks = [createParagraph('source', { bold: true }), createParagraph('target word')]; model.value = next },
  paragraphPoint: (index) => {
    const scroll = document.querySelector('.document-scroll'), bounds = scroll.getBoundingClientRect()
    const layout = layoutDocument(model.value, createTextMeasurer())
    const offset = sceneOffset(layout, { width: scroll.clientWidth, height: scroll.clientHeight, scrollLeft: scroll.scrollLeft, scrollTop: scroll.scrollTop, zoom: 100 })
    const line = layout.lines.find(line => line.target.blockId === model.value.blocks[index].id)
    return { x: bounds.left + offset.x + line.x + 3, y: bounds.top + offset.y + line.y + line.height / 2 }
  },
  heading: () => { const next = copy(model.value); next.blocks[0].type = 'heading'; next.blocks[0].level = 1; model.value = next },
  checklistPosition: (blockId) => {
    const scroll = document.querySelector('.document-scroll')
    const bounds = scroll.getBoundingClientRect()
    const layout = layoutDocument(model.value, createTextMeasurer())
    const offset = sceneOffset(layout, { width: scroll.clientWidth, height: scroll.clientHeight, scrollLeft: scroll.scrollLeft, scrollTop: scroll.scrollTop, zoom: 100 })
    const line = layout.lines.find((item) => item.checklist?.blockId === blockId)
    const x = line.x - 15
    const y = line.y + line.height / 2
    return { x: bounds.left + offset.x + x, y: bounds.top + offset.y + y, hit: hitTestChecklist(layout, x, y), line, bounds: { left: bounds.left, top: bounds.top }, offset }
  },
  checklistHtmlRoundtrip: () => blocksFromHtml(documentHtml(model.value)).map((block) => ({ type: block.type, checked: block.checklist?.checked, listId: block.checklist?.listId, text: block.runs.map((run) => run.text).join('') })),
  verticalAlignHtmlRoundtrip: () => {
    const source = createParagraph()
    source.runs = [
      { ...source.runs[0], id: 'formula-h', text: 'H', verticalAlign: 'baseline' },
      { ...source.runs[0], id: 'formula-two', text: '2', verticalAlign: 'subscript' },
      { ...source.runs[0], id: 'formula-o', text: 'O', verticalAlign: 'superscript' },
    ]
    const html = documentHtml({ ...initial, blocks: [source] })
    return { html, runs: blocksFromHtml(html)[0].runs.map((run) => [run.text, run.verticalAlign]) }
  },
}
createApp({ render: () => h(DocumentWorkbench, {
  modelValue: model.value, canUndo: !!history.value.length, canRedo: !!redo.value.length,
  'onUpdate:modelValue': update,
  onUndo: () => { if (history.value.length) { redo.value.push(copy(model.value)); model.value = history.value.pop() } },
  onRedo: () => { if (redo.value.length) { history.value.push(copy(model.value)); model.value = redo.value.pop() } },
}) }).mount('#app')
</script></body></html>`

let server
let browser
let page
let baseURL
let errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'document-toolbar-interaction-fixture',
    configureServer(vite) {
      vite.middlewares.use(fixturePath, async (_req, res, next) => {
        try { res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.end(await vite.transformIndexHtml(fixturePath, fixture)) } catch (error) { next(error) }
      })
    },
  }], server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen()
  const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate an HTTP port')
  baseURL = `http://127.0.0.1:${address.port}${fixturePath}`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, { timeout: 60_000 })
afterAll(async () => { await browser?.close(); await server?.close() })
beforeEach(async () => {
  errors = []
  const diagnostics = []
  page = await browser.newPage({ viewport: { width: 1440, height: 900 } })
  page.on('pageerror', (error) => errors.push(error.message))
  page.on('console', (message) => { if (message.type() === 'error') diagnostics.push(`console: ${message.text()}`) })
  page.on('requestfailed', (request) => diagnostics.push(`request: ${request.url()} ${request.failure()?.errorText ?? ''}`))
  page.on('response', (response) => { if (response.status() >= 400) diagnostics.push(`HTTP ${response.status()}: ${response.url()}`) })
  await page.goto(baseURL)
  try {
    await expect(page.getByRole('textbox', { name: '文档画布编辑区', exact: true })).toBeVisible()
  } catch (cause) {
    const overlay = await page.evaluate(() => document.querySelector('vite-error-overlay')?.shadowRoot?.textContent ?? document.body.innerText)
    throw new Error(`Docs fixture failed to mount.\n${[...errors, ...diagnostics, overlay].join('\n').slice(0, 8000)}`, { cause })
  }
}, { timeout: 60_000 })
afterEach(async () => { await page?.close() })

test('Docs composes a classic Ribbon and a single footer without duplicating the application header', async () => {
  assert.equal(await page.locator('.document-header').count(), 0)
  assert.equal(await page.locator('.document-statusbar').count(), 1)
  assert.equal(await page.locator('.workbench-ribbon-tabs').evaluate((element) => element.getBoundingClientRect().height), 36)
  assert.equal(await page.locator('.workbench-ribbon-toolbar').evaluate((element) => element.getBoundingClientRect().height), 40)
  assert.equal(await page.locator('.document-statusbar').evaluate((element) => element.getBoundingClientRect().height), 36)
  assert.equal(await page.locator('.document-outline').count(), 0)
  await expect(page.locator('.document-statusbar')).toContainText('无未保存更改')
  assert.deepEqual(errors, [])
})

test('the table Ribbon operation opens the size dialog and retains an undoable table after snapshot reopening', async () => {
  const before = await page.evaluate(() => window.documentToolbarFixture.snapshot())
  const bridge = page.getByLabel('文档文字输入', { exact: true })
  await bridge.focus(); await bridge.press('Control+End')
  await page.getByRole('tab', { name: '插入', exact: true }).click()
  await page.locator('.ribbon-visible-groups [data-ribbon-item="doc-table"] button').click()
  await page.getByRole('menuitem', { name: '插入表格', exact: true }).click()
  const dialog = page.getByRole('dialog', { name: '插入表格', exact: true })
  await expect(dialog.getByLabel('行数', { exact: true })).toHaveValue('3')
  await expect(dialog.getByLabel('列数', { exact: true })).toHaveValue('5')
  await dialog.getByLabel('行数', { exact: true }).fill('4')
  await dialog.getByLabel('列数', { exact: true }).fill('7')
  await dialog.getByRole('button', { name: '插入', exact: true }).click()
  await expect(dialog).toBeHidden()
  const tables = () => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks.filter(block => block.type === 'table').map(block => [block.metadata.table.rows.length, block.metadata.table.rows[0].cells.length]))
  await expect.poll(tables).toEqual([[4, 7]])
  const expectedWidth = before.page.width - before.page.marginLeft - before.page.marginRight
  const insertedWidth = await page.evaluate(() => window.documentToolbarFixture.snapshot().blocks.find(block => block.type === 'table').metadata.table)
  assert.ok(Math.abs(insertedWidth.style.width - expectedWidth) < 0.001)
  assert.ok(insertedWidth.columnWidths.every(width => Math.abs(width - expectedWidth / 7) < 0.001))
  assert.equal(await page.evaluate(() => window.documentToolbarFixture.snapshot().blocks.flatMap(block => block.runs).map(run => run.text).join('')), '甲乙丙丁')
  await page.getByRole('tab', { name: '开始', exact: true }).click()
  await page.locator('.ribbon-visible-groups [data-ribbon-item="doc-undo"] button').click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot())).toEqual(before)
  await page.locator('.ribbon-visible-groups [data-ribbon-item="doc-redo"] button').click()
  await expect.poll(tables).toEqual([[4, 7]])
  await page.evaluate(() => window.documentToolbarFixture.reopenSnapshot())
  await expect.poll(tables).toEqual([[4, 7]])
  assert.deepEqual(errors, [])
})

test.each([['columns', false, 800], ['continuous', false, 800], ['continuous', true, 640]])('table insertion uses the physical page in %s sections and respects explicit width %s', async (kind, specifyWidth, expectedWidth) => {
  await page.evaluate(value => window.documentToolbarFixture.tablePageFixture(value), kind)
  const bridge = page.getByLabel('文档文字输入', { exact: true })
  await bridge.focus(); await bridge.press('Control+End')
  await page.getByRole('tab', { name: '插入', exact: true }).click()
  await page.locator('.ribbon-visible-groups [data-ribbon-item="doc-table"] button').click()
  await page.getByRole('menuitem', { name: '插入表格', exact: true }).click()
  const dialog = page.getByRole('dialog', { name: '插入表格', exact: true })
  if (specifyWidth) {
    await dialog.getByRole('switch', { name: '指定表格宽度', exact: true }).click()
    await dialog.getByLabel('表格宽度（px）', { exact: true }).fill('640')
  }
  await dialog.getByRole('button', { name: '插入', exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks.filter(block => block.type === 'table').length)).toBe(1)
  const table = await page.evaluate(() => window.documentToolbarFixture.snapshot().blocks.find(block => block.type === 'table').metadata.table)
  assert.ok(Math.abs(table.style.width - expectedWidth) < 0.001)
  assert.ok(table.columnWidths.every(width => Math.abs(width - expectedWidth / 5) < 0.001))
  assert.deepEqual(errors, [])
})

test('Docs once painter applies on pointer release, retains target text and undoes as one edit', async () => {
  await page.evaluate(() => window.documentToolbarFixture.painterFixture())
  const clickParagraph = async index => { const p = await page.evaluate(i => window.documentToolbarFixture.paragraphPoint(i), index); await page.mouse.click(p.x, p.y) }
  await clickParagraph(0)
  const brush = page.getByRole('button', { name: '格式刷', exact: true })
  await brush.click(); await expect(brush).toHaveAttribute('aria-pressed', 'true')
  await clickParagraph(1); await expect(brush).toHaveAttribute('aria-pressed', 'false')
  const target = () => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks[1].runs.map(run => [run.text, run.bold]))
  await expect.poll(target).toEqual([['target', true], [' word', false]])
  await page.getByRole('button', { name: '撤销', exact: true }).click()
  await expect.poll(target).toEqual([['target word', false]])
})

test('font selection preserves the logical Canvas range and uses the external document history', async () => {
  const bridge = page.getByLabel('文档文字输入', { exact: true })
  await bridge.focus()
  await bridge.press('Control+A')
  const font = page.getByRole('combobox', { name: '字体', exact: true })
  await font.fill('georg'); await font.press('Enter')
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks[0].runs[0].fontFamily)).toBe('Georgia')
  await expect(page.locator('.document-statusbar')).toContainText('未保存更改')
  await page.getByRole('button', { name: '撤销', exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks[0].runs[0].fontFamily)).toBe('Arial')
  await page.getByRole('button', { name: '重做', exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks[0].runs[0].fontFamily)).toBe('Georgia')
  assert.deepEqual(errors, [])
})

test('typing a size keeps the Canvas range, cancels on blur or Escape, and commits one undo step on Enter', async () => {
  const bridge = page.getByLabel('文档文字输入', { exact: true })
  await bridge.focus(); await bridge.press('Control+A')
  const size = page.getByRole('combobox', { name: '字号', exact: true })
  const current = () => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks[0].runs[0].fontSize)
  const undo = page.getByRole('button', { name: '撤销', exact: true })
  await size.fill('24'); await expect(undo).toBeDisabled()
  await size.press('Escape'); await expect(size).toHaveValue('11')
  await size.fill('37'); await size.blur(); await expect(size).toHaveValue('11'); await expect(undo).toBeDisabled()
  await size.fill('13.5'); await size.press('Enter'); await expect.poll(current).toBe(13.5)
  await undo.click(); await expect.poll(current).toBe(11); await expect(undo).toBeDisabled()
})

test('subscript and superscript retain the selection, toggle to baseline, and replay through undo and redo', async () => {
  const bridge = page.getByLabel('文档文字输入', { exact: true })
  await bridge.focus()
  await bridge.press('Control+A')

  const subscript = page.getByRole('button', { name: '下标', exact: true })
  const superscript = page.getByRole('button', { name: '上标', exact: true })
  await subscript.click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks[0].runs.every(run => run.verticalAlign === 'subscript'))).toBe(true)
  await expect(subscript).toHaveAttribute('aria-pressed', 'true')

  await superscript.click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks[0].runs.every(run => run.verticalAlign === 'superscript'))).toBe(true)
  await expect(superscript).toHaveAttribute('aria-pressed', 'true')
  await subscript.click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks[0].runs.every(run => run.verticalAlign === 'subscript'))).toBe(true)

  await subscript.click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks[0].runs.every(run => run.verticalAlign === undefined))).toBe(true)
  await page.getByRole('button', { name: '撤销', exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks[0].runs.every(run => run.verticalAlign === 'subscript'))).toBe(true)
  await page.getByRole('button', { name: '重做', exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks[0].runs.every(run => run.verticalAlign === undefined))).toBe(true)
  assert.deepEqual(errors, [])
})

test('collapsed caret subscript stays pending until text insertion creates a reversible run', async () => {
  const bridge = page.getByLabel('文档文字输入', { exact: true })
  await bridge.focus()
  await bridge.press('Control+End')
  const before = await page.evaluate(() => window.documentToolbarFixture.snapshot())

  await page.getByRole('button', { name: '下标', exact: true }).click()
  assert.deepEqual(await page.evaluate(() => window.documentToolbarFixture.snapshot()), before)

  await bridge.press('2')
  await expect.poll(() => page.evaluate(() => {
    const run = window.documentToolbarFixture.snapshot().blocks[0].runs.find((item) => item.text.includes('2'))
    return run && [run.text, run.verticalAlign]
  })).toEqual(['2', 'subscript'])
  await page.getByRole('button', { name: '撤销', exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks[0].runs.map((run) => run.text).join(''))).toBe('甲乙丙丁')
  await page.getByRole('button', { name: '重做', exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks[0].runs.find((run) => run.text.includes('2')).verticalAlign)).toBe('subscript')
  assert.deepEqual(errors, [])
})

test('HTML serialization and DOM import preserve superscript and subscript semantics', async () => {
  const result = await page.evaluate(() => window.documentToolbarFixture.verticalAlignHtmlRoundtrip())
  assert.match(result.html, /<sub>/)
  assert.match(result.html, /<sup>/)
  assert.deepEqual(result.runs, [['H', undefined], ['2', 'subscript'], ['O', 'superscript']])
  assert.deepEqual(errors, [])
})

for (const backwards of [false, true]) test('paragraph controls and the Ant settings panel read the first selected paragraph in a ' + (backwards ? 'backwards' : 'forwards') + ' range', async () => {
  await page.evaluate(() => window.documentToolbarFixture.paragraphFixture())
  const bridge = page.getByLabel('文档文字输入', { exact: true })
  await bridge.focus()
  await bridge.press(backwards ? 'Control+End' : 'Control+Home')
  await bridge.press(backwards ? 'Control+Shift+Home' : 'Control+Shift+End')
  const before = await page.evaluate(() => window.documentToolbarFixture.snapshot())
  await expect(page.locator('.ribbon-visible-groups [data-ribbon-item="doc-line-height"] .ant-select-selection-item')).toHaveText('2 倍')
  await page.locator('.ribbon-visible-groups [data-ribbon-item="doc-align"]').getByRole('button').click()
  await expect(page.locator('.ribbon-command-menu:visible').getByRole('menuitem', { name: '右对齐', exact: true })).toHaveClass(/ant-menu-item-selected/)
  await page.keyboard.press('Escape')
  const point = await page.evaluate(() => window.documentToolbarFixture.paragraphPoint(0))
  await page.mouse.click(point.x, point.y, { button: 'right' })
  await page.locator('.document-context-menu').getByRole('menuitem', { name: '段落设置', exact: true }).click()
  const panel = page.locator('[data-document-paragraph-settings]')
  await expect(panel.getByRole('button', { name: '右对齐', exact: true })).toHaveAttribute('aria-pressed', 'true')
  await expect(panel.getByRole('spinbutton', { name: '首行缩进 (px)', exact: true })).toHaveValue('20')
  await expect(panel.getByRole('spinbutton', { name: '段前 (px)', exact: true })).toHaveValue('6')
  await expect(panel.getByRole('spinbutton', { name: '段后 (px)', exact: true })).toHaveValue('12')
  assert.deepEqual(await page.evaluate(() => window.documentToolbarFixture.snapshot()), before)
  await panel.getByRole('spinbutton', { name: '段后 (px)', exact: true }).fill('18')
  await panel.getByRole('spinbutton', { name: '段后 (px)', exact: true }).press('Tab')
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks.map(block => block.style.spaceAfter))).toEqual([18, 18])
  const after = await page.evaluate(() => window.documentToolbarFixture.snapshot())
  assert.deepEqual(after.blocks.map(block => [block.id, block.runs]), before.blocks.map(block => [block.id, block.runs]))
  await panel.getByRole('button', { name: '关闭段落设置', exact: true }).click()
  await expect(bridge).toBeFocused()
  await page.getByRole('button', { name: '撤销', exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot())).toEqual(before)
  await page.getByRole('button', { name: '重做', exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot())).toEqual(after)
  await page.evaluate(() => window.documentToolbarFixture.reopenSnapshot())
  assert.deepEqual(await page.evaluate(() => window.documentToolbarFixture.snapshot()), after)
})

test('the named-style picker follows the starting paragraph of a mixed backwards selection and retains one undoable change', async () => {
  await page.evaluate(() => window.documentToolbarFixture.namedStyleFixture())
  const bridge = page.getByLabel('文档文字输入', { exact: true })
  await bridge.focus(); await bridge.press('Control+End'); await bridge.press('Control+Shift+Home')
  const styleCell = page.locator('.ribbon-visible-groups [data-ribbon-item="doc-block-style"]')
  await expect(styleCell.locator('.ant-select-selection-item')).toHaveText('标题')
  const before = await page.evaluate(() => window.documentToolbarFixture.snapshot())
  const style = page.getByRole('combobox', { name: '段落样式', exact: true })
  await style.click()
  await page.getByRole('option', { name: '标题 3', exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks.map(block => block.metadata.namedStyleType))).toEqual(['heading3', 'heading3'])
  const after = await page.evaluate(() => window.documentToolbarFixture.snapshot())
  assert.deepEqual(after.blocks.map(block => [block.id, block.runs.map(run => run.text).join('')]), before.blocks.map(block => [block.id, block.runs.map(run => run.text).join('')]))
  await expect(bridge).toBeFocused()
  await page.getByRole('button', { name: '撤销', exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot())).toEqual(before)
  await page.getByRole('button', { name: '重做', exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot())).toEqual(after)
  await page.evaluate(() => window.documentToolbarFixture.reopenSnapshot())
  await bridge.focus(); await bridge.press('Control+Home')
  await expect(styleCell.locator('.ant-select-selection-item')).toHaveText('标题 3')
  assert.deepEqual(await page.evaluate(() => window.documentToolbarFixture.snapshot()), after)
  assert.deepEqual(errors, [])
})

test('named paragraph styles retain checklist identity, Canvas toggles and reversible style changes', async () => {
  await page.getByRole('button', { name: '清单', exact: true }).click()
  const before = await page.evaluate(() => window.documentToolbarFixture.snapshot())
  const style = page.getByRole('combobox', { name: '段落样式', exact: true })
  await style.click(); await page.getByRole('option', { name: '标题', exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks[0].metadata.namedStyleType)).toBe('title')
  const styled = await page.evaluate(() => window.documentToolbarFixture.snapshot())
  assert.equal(styled.blocks[0].type, 'checklist')
  assert.deepEqual(styled.blocks[0].checklist, before.blocks[0].checklist)
  await expect(page.getByRole('button', { name: '清单', exact: true })).toHaveAttribute('aria-pressed', 'true')
  const position = await page.evaluate(id => window.documentToolbarFixture.checklistPosition(id), styled.blocks[0].id)
  assert.equal(position.hit, styled.blocks[0].id)
  await page.mouse.click(position.x, position.y)
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks[0].checklist.checked)).toBe(true)
  await page.getByRole('button', { name: '撤销', exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot())).toEqual(styled)
  await page.getByRole('button', { name: '撤销', exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot())).toEqual(before)
  await page.getByRole('button', { name: '重做', exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot())).toEqual(styled)
  assert.deepEqual(errors, [])
})

test('a content lock outside the selected word disables the full-paragraph named-style command', async () => {
  await page.evaluate(() => window.documentToolbarFixture.lockNamedStyleTail())
  const before = await page.evaluate(() => window.documentToolbarFixture.snapshot())
  const bridge = page.getByLabel('文档文字输入', { exact: true })
  await bridge.focus(); await bridge.press('Control+Home'); await bridge.press('Shift+ArrowRight')
  await expect(page.getByRole('combobox', { name: '段落样式', exact: true })).toBeDisabled()
  await expect(page.getByRole('button', { name: '撤销', exact: true })).toBeDisabled()
  assert.deepEqual(await page.evaluate(() => window.documentToolbarFixture.snapshot()), before)
  assert.deepEqual(errors, [])
})

test('section break Ribbon menu inserts an odd-page section and replays through undo and redo', async () => {
  await page.getByRole('tab', { name: '插入', exact: true }).click()
  await page.getByRole('button', { name: '分页符和分节符', exact: true }).click()
  await page.getByRole('button', { name: '分节符', exact: true }).click()
  await page.getByRole('button', { name: '奇数页', exact: true }).click()

  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks.map(block => block.type))).toEqual(['paragraph', 'sectionBreak', 'paragraph'])
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks[1].sectionBreak.type)).toBe('oddPage')
  await page.locator('.workbench-ribbon-tab[data-tab-id="home"]').click()
  const undo = page.locator('.ribbon-visible-groups [data-ribbon-item="doc-undo"] button')
  await expect(undo).toBeVisible()
  await undo.click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks.map(block => block.type))).toEqual(['paragraph'])
  await page.locator('.ribbon-visible-groups [data-ribbon-item="doc-redo"] button').click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks[1].sectionBreak.type)).toBe('oddPage')
  assert.deepEqual(errors, [])
})

test('checklist Ribbon command creates editable task items with Canvas toggles and undoable state', async () => {
  await expect.poll(() => page.locator('.document-scroll').evaluate(element => element.__vueParentComponent.setupState.viewport.width), { timeout: 15000 }).toBeGreaterThan(100)
  await page.getByRole('button', { name: '清单', exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks[0].type)).toBe('checklist')
  let model = await page.evaluate(() => window.documentToolbarFixture.snapshot())
  const taskId = model.blocks[0].id
  assert.equal(model.blocks[0].checklist.checked, false)
  let position = await page.evaluate((id) => window.documentToolbarFixture.checklistPosition(id), taskId)
  assert.equal(position.hit, taskId)
  await page.mouse.click(position.x, position.y)
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks[0].checklist.checked)).toBe(true)
  await page.getByRole('button', { name: '撤销', exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks[0].checklist.checked)).toBe(false)
  await page.getByRole('button', { name: '重做', exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks[0].checklist.checked)).toBe(true)
  const roundtrip = await page.evaluate(() => window.documentToolbarFixture.checklistHtmlRoundtrip())
  assert.deepEqual(roundtrip, [{ type: 'checklist', checked: true, listId: model.blocks[0].checklist.listId, text: '甲乙丙丁' }])

  const bridge = page.getByLabel('文档文字输入', { exact: true })
  await bridge.focus()
  await bridge.press('Control+End')
  await bridge.press('Enter')
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks.length)).toBe(2)
  model = await page.evaluate(() => window.documentToolbarFixture.snapshot())
  assert.deepEqual(model.blocks.map((block) => block.type), ['checklist', 'checklist'])
  assert.equal(model.blocks[1].checklist.checked, false)
  position = await page.evaluate((id) => window.documentToolbarFixture.checklistPosition(id), model.blocks[1].id)
  await page.mouse.click(position.x, position.y)
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().blocks[1].checklist.checked)).toBe(true)
  assert.deepEqual(errors, [])
})

test('outline, zoom and focus view operations keep document content and dirty state unchanged', async () => {
  await page.evaluate(() => window.documentToolbarFixture.heading())
  const before = await page.evaluate(() => window.documentToolbarFixture.snapshot())
  await page.getByRole('tab', { name: '开始', exact: true }).click()
  if (!await page.getByRole('button', { name: '文档大纲', exact: true }).isVisible()) await page.getByRole('button', { name: '更多功能', exact: true }).click()
  await page.getByRole('button', { name: '文档大纲', exact: true }).click()
  await expect(page.locator('.document-outline')).toBeVisible()
  await page.setViewportSize({ width: 680, height: 900 })
  await expect(page.locator('.document-outline')).toHaveClass(/document-outline--float/)
  await page.locator('.document-statusbar').getByRole('button', { name: '放大', exact: true }).click()
  await page.getByRole('tab', { name: '视图', exact: true }).click()
  await page.getByRole('button', { name: '全屏', exact: true }).click()
  await expect.poll(() => page.evaluate(() => document.fullscreenElement !== null)).toBe(true)
  await page.getByRole('button', { name: '退出全屏', exact: true }).click()
  await expect.poll(() => page.evaluate(() => document.fullscreenElement === null)).toBe(true)
  await page.getByRole('tab', { name: '开始', exact: true }).click()
  if (!await page.getByRole('button', { name: '专注模式', exact: true }).isVisible()) await page.getByRole('button', { name: '更多功能', exact: true }).click()
  await page.getByRole('button', { name: '专注模式', exact: true }).click()
  await expect(page.locator('.workbench-ribbon')).toHaveCount(0)
  await page.getByRole('button', { name: '退出专注模式', exact: true }).click()
  await expect(page.locator('.workbench-ribbon')).toBeVisible()
  assert.deepEqual(await page.evaluate(() => window.documentToolbarFixture.snapshot()), before)
  assert.deepEqual(errors, [])
})

test('view tab page settings edits real geometry and margins in one undoable document update', async () => {
  const before = await page.evaluate(() => window.documentToolbarFixture.snapshot())
  await page.getByRole('tab', { name: '视图', exact: true }).click()
  await page.getByRole('button', { name: '页面设置', exact: true }).click()
  const dialog = page.getByRole('dialog', { name: '页面设置', exact: true })
  await dialog.getByRole('spinbutton', { name: '上边距', exact: true }).fill('48')
  await dialog.getByRole('spinbutton', { name: '纸张宽度', exact: true }).fill('900')
  await dialog.getByRole('button', { name: /应\s*用/ }).click()
  await expect(dialog).toBeHidden()
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().page.marginTop)).toBe(48)
  await expect.poll(() => page.evaluate(() => window.documentToolbarFixture.snapshot().page.width)).toBe(900)
  await page.getByRole('button', { name: '页面设置', exact: true }).click()
  await expect(dialog.getByRole('spinbutton', { name: '纸张宽度', exact: true })).toHaveValue('900')
  await expect(dialog.getByRole('spinbutton', { name: '上边距', exact: true })).toHaveValue('48')
  await dialog.getByRole('button', { name: /取\s*消/ }).click()
  await page.getByRole('tab', { name: '开始', exact: true }).click()
  await page.getByRole('button', { name: '撤销', exact: true }).click()
  assert.deepEqual((await page.evaluate(() => window.documentToolbarFixture.snapshot())).page, before.page)
  assert.deepEqual(errors, [])
})
