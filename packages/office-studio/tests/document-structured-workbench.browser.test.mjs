import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

// Deferred browser drafts mount the actual Canvas workbench and antdv-next dialogs.
// Synthetic composition guards are not acceptance for the real Windows IME.
const fixturePath = '/__document-structured-workbench-test__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>Docs structured blocks</title><style>body{margin:0}#app{height:100vh}</style></head><body><div id="app"></div><script type="module">
import { computed, createApp, h, ref, shallowRef } from 'vue'
import { ConfigProvider, theme } from 'antdv-next'
import 'antdv-next/dist/reset.css'
import { OfficeStudio } from '/packages/office-studio/src/api/office-studio'
import { FDocument } from '/packages/office-studio/src/api/facade'
import { createBlankDocument } from '/packages/office-studio/src/domain/model'
import { createParagraph, createTable, findDocumentBlock, textTargets } from '/packages/office-studio/src/features/document/model/document-model'
import { createTextMeasurer, layoutDocument } from '/packages/office-studio/src/features/document/layout/document-layout'
import { documentLayoutForBlockHost } from '/packages/office-studio/src/features/document/layout/document-auxiliary-layout'
import { documentModelForBlockHost } from '/packages/office-studio/src/features/document/services/document-block-host'
import { projectDocumentStoryBlocks } from '/packages/office-studio/src/features/document/model/document-story-blocks'
import { initializeDocumentQuotePagination } from '/packages/office-studio/src/features/document/model/document-paragraph-pagination'
import { registerDocumentCodeFontFamilies } from '/packages/office-studio/src/features/document/code-blocks/document-code-fonts'
import { EMPTY_SECTION_HEADER_FOOTER } from '/packages/office-studio/src/features/document/services/section-header-footer'
import { sceneOffset } from '/packages/office-studio/src/features/document/rendering/document-scene'
import { createOfficeI18n, createOfficeTranslator, OfficeLocaleKey, OfficeTranslatorKey } from '/packages/office-studio/src/i18n/office-i18n'
import DocumentWorkbench from '/packages/office-studio/src/components/DocumentWorkbench.vue'
const query = new URLSearchParams(location.search), copy = value => JSON.parse(JSON.stringify(value)), standalone = query.has('standalone')
const codeControlValues = { fontFamily: 'Registered Office Code Font '.repeat(6).trim(), language: 'author-code-language-'.repeat(8) }
const paragraph = (id, text = '') => ({ ...createParagraph(text), id })
const columnBlock = (id, count = 3) => ({ ...paragraph(id), type: 'columns', runs: [], columns: { version: 1, gap: 16, horizontalPadding: 8, minHeight: 72,
  layout: 'fixed', responsive: 'shrink', columns: Array.from({ length: count }, (_, index) => ({ id: id + '-column-' + index, widthRatio: 1, minWidth: 32,
    blocks: [paragraph(id + '-p-' + index, ['Left', 'Middle', 'Right'][index] ?? 'Column')] })) } })
const initial = createBlankDocument(); initial.id = 'structured-browser'; initial.name = '结构块交互草稿'
initial.page = { width: 620, height: 540, marginTop: 42, marginBottom: 42, marginLeft: 48, marginRight: 48 }
initial.blocks = [{ ...paragraph('code'), type: 'code', runs: [], children: [paragraph('code-p', 'a😀b'), paragraph('code-last', '  c')],
  code: { language: 'typescript', showLineNumbers: true, tabSize: 4, wrap: true } },
  { ...paragraph('quote'), type: 'quote', runs: [], children: [paragraph('quote-p', 'Quoted text')] },
  columnBlock('columns-three'), columnBlock('columns-two', 2),
  { ...paragraph('long-code'), type: 'code', runs: [], children: [paragraph('long-code-p', 'abcdefgh'.repeat(100))],
    code: { language: 'typescript', showLineNumbers: true, tabSize: 4, wrap: false } }, paragraph('tail', 'Outside')]
if (query.has('long-controls')) {
  initial.page = { ...initial.page, width: 900, marginLeft: 680 }
  initial.blocks = [initial.blocks[0], paragraph('tail', 'Outside')]
  initial.blocks[0].code.language = codeControlValues.language
  registerDocumentCodeFontFamilies([{ value: codeControlValues.fontFamily, category: 'monospace' }])
}
if (new URLSearchParams(location.search).has('blank')) initial.blocks = [paragraph('empty'), paragraph('tail', 'Outside')]
if (new URLSearchParams(location.search).has('table')) {
  const table = createTable(1, 1); table.rows[0].cells[0].blocks = [paragraph('table-cell', 'Retained table')]
  table.style = { layout: 'fixed', width: 420 }; table.columnWidths = [420]
  initial.blocks = [{ ...paragraph('table'), type: 'table', runs: [], metadata: { table } }, paragraph('after-table'), paragraph('tail', 'Outside')]
}
if (new URLSearchParams(location.search).has('auxiliary')) {
  initial.page = { ...initial.page, height: 780, marginTop: 170, marginBottom: 170 }
  const header = projectDocumentStoryBlocks({ blocks: [columnBlock('header-columns', 2),
    { ...paragraph('header-code'), type: 'code', runs: [], children: [paragraph('header-code-p', 'const header = 1')],
      code: { language: 'typescript', showLineNumbers: true, tabSize: 4, wrap: false } }], runs: [], paragraphs: [], metadata: {} }).content
  initial.sectionHeaderFooters = { first: { ...EMPTY_SECTION_HEADER_FOOTER, header: header.runs.map(run => run.text).join(''), storyContent: { header } } }
}
if (query.has('pagination')) {
  initial.page = { ...initial.page, height: 280, marginTop: 30, marginBottom: 30 }
  initial.metadata = { ...initial.metadata, documentFlavor: Number(query.get('flavor') ?? 1), documentFlavorContract: 'univer' }
  const exact = (id, text, height = 24) => ({ ...paragraph(id, text), style: { align: 'left', indent: 0,
    spacingRule: 'exact', lineHeight: height, spaceBefore: 0, spaceAfter: 0 } })
  const children = query.get('pagination') === 'oversized'
    ? [exact('pagination-first', Array.from({ length: 15 }, (_, index) => 'Quote line ' + index).join('\\n'))]
    : [exact('pagination-first', 'Quote 😀 first'), exact('pagination-last', 'Quote last')]
  const quote = { ...exact('pagination-quote', ''), type: 'quote', runs: [], children }
  quote.style.spaceBefore = 8; quote.style.spaceAfter = 8; initializeDocumentQuotePagination(quote)
  initial.blocks = [exact('pagination-before', 'Before quote', 145), quote, exact('tail', 'Outside')]
}
if (standalone) delete initial.id
let studio = standalone ? undefined : new OfficeStudio({ document: initial, currentUserId: 'editor' })
studio?.activateDocument(initial.id)
const api = shallowRef(studio ? new FDocument(studio, initial.id) : undefined), model = shallowRef(api.value?.getModel() ?? initial)
const commandEvents = []
for (const [eventName, phase] of [['BeforeCommandExecute', 'before'], ['CommandExecuted', 'after']]) {
  studio?.events.addEvent(eventName, event => commandEvents.push({ phase, id: event.id, value: event.params?.value, unitId: event.unitId, kind: event.kind }))
}
if (query.has('long-controls')) {
  api.value.getCode('code').setFontFamily(codeControlValues.fontFamily)
  model.value = api.value.getModel()
}
const dark = ref(false), variant = ref('classic'), notices = [], locale = ref(query.get('locale') ?? 'zh-CN')
const translator = computed(() => createOfficeTranslator(locale.value))
function sync() { if (api.value) model.value = api.value.getModel() }
function select(blockId, offset = 0, endId = blockId, endOffset = offset, host) {
  if (!api.value) return false
  const root = api.value.getModel(), scoped = host ? documentModelForBlockHost(root, host) : root, targets = textTargets(scoped)
  const anchor = targets.find(target => (target.cellBlockId ?? target.blockId) === blockId), focus = targets.find(target => (target.cellBlockId ?? target.blockId) === endId)
  if (!anchor || !focus) throw new Error('Selection must use real current text targets')
  return api.value.setInputSelection({ anchor: { ...anchor, offset }, focus: { ...focus, offset: endOffset } }, undefined, host)
}
window.documentStructuredFixture = {
  codeControlValues,
  snapshot: () => copy(api.value?.getModel() ?? model.value), notices,
  commandEvents: () => copy(commandEvents),
  source: () => findDocumentBlock(api.value?.getModel() ?? model.value, 'code').children.map(block => block.runs.map(run => run.text).join('')).join('\\n'),
  history: () => studio ? { undo: studio.editor.history.units[api.value.id]?.undo.length ?? 0, redo: studio.editor.history.units[api.value.id]?.redo.length ?? 0 } : undefined,
  select, dark: value => dark.value = value, variant: value => variant.value = value,
  selection: () => copy(api.value.getInputState().location),
  publicSnapshot: () => copy(api.value.getSnapshot()),
  unwrapQuote: id => { const result = api.value.getQuote(id).unwrap(); sync(); return result },
  paginationLayout: () => {
    const layout = layoutDocument(api.value?.getModel() ?? model.value, createTextMeasurer())
    return copy({ pages: layout.pages.map(page => ({ index: page.index, content: page.content })),
      lines: layout.lines.map(line => ({ blockId: line.target.cellBlockId ?? line.target.blockId, page: line.page,
        column: line.column, y: line.y, height: line.height, start: line.start, end: line.end })), containers: layout.containers })
  },
  selectHost: (id, host) => select(id, 0, id, 0, host),
  block: (id, host) => { const root = api.value?.getModel() ?? model.value; return copy(findDocumentBlock(host ? documentModelForBlockHost(root, host) : root, id)) },
  group: (id, host) => { const root = api.value?.getModel() ?? model.value; return copy(findDocumentBlock(host ? documentModelForBlockHost(root, host) : root, id).columns) },
  remoteRatios: ratios => { api.value.getColumnGroup('columns-three').setWidthRatios(ratios); sync() },
  undo: () => { const result = api.value.undo(); sync(); return result }, redo: () => { const result = api.value.redo(); sync(); return result },
  longSource: () => findDocumentBlock(api.value?.getModel() ?? model.value, 'long-code').children.map(block => block.runs.map(run => run.text).join('')).join('\\n'),
  changeSource: value => { api.value.getCode('code').setText(value); sync() },
  replaceSameSnapshot: () => { model.value = copy(model.value) },
  saveAndReopen: (paragraphId = 'code-p', host) => { const saved = api.value.save(); studio.dispose(); studio = new OfficeStudio({ document: copy(saved), currentUserId: 'editor' }); studio.activateDocument(saved.id); api.value = new FDocument(studio, saved.id); sync(); select(paragraphId, 0, paragraphId, 0, host); return copy(saved) },
  point: (id, offset = 0, codeOffset = 0, host) => {
    const viewport = document.querySelector('.document-scroll'), rect = viewport.getBoundingClientRect()
    const layout = documentLayoutForBlockHost(layoutDocument(api.value?.getModel() ?? model.value, createTextMeasurer()), host)
    const line = layout.lines.find(line => (line.target.cellBlockId ?? line.target.blockId) === id && offset >= line.start && offset <= line.end)
    if (!line) throw new Error('Paragraph does not have current Canvas geometry')
    const glyph = line.glyphs.find(glyph => glyph.offset >= offset), shift = sceneOffset(layout, { width: rect.width, height: rect.height, scrollLeft: viewport.scrollLeft, scrollTop: viewport.scrollTop, zoom: 100 })
    return { x: rect.x + shift.x + (glyph?.x ?? line.x + line.width) - codeOffset, y: rect.y + shift.y + line.y + line.height / 2 }
  },
}
const app = createApp({ render: () => h(ConfigProvider, { theme: { algorithm: dark.value ? theme.darkAlgorithm : theme.defaultAlgorithm } }, {
  default: () => h(DocumentWorkbench, { modelValue: model.value, documentApi: api.value, canUndo: studio?.editor.canUndo, canRedo: studio?.editor.canRedo, ribbonVariant: variant.value, showOutline: query.has('long-controls') ? false : undefined,
    'onUpdate:modelValue': next => model.value = next, onNotify: message => notices.push(message) }),
}) })
const i18n = createOfficeI18n(); i18n.global.locale.value = locale.value
app.use(i18n); app.provide(OfficeLocaleKey, locale); app.provide(OfficeTranslatorKey, (key, fallback) => translator.value(key, fallback)); app.mount('#app')
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
  await page.getByRole('textbox', { name: '文档文字输入', exact: true }).focus()
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

test('symbol and emoji popups execute the fixed value command in body and header Code with focus, themes, geometry and persistence', async () => {
  const header = { kind: 'story', sectionId: 'first', story: 'header' }
  for (const [variant, dark, query, paragraphId, host] of [['classic', false, '', 'code-p', undefined], ['grid', true, '?auxiliary', 'header-code-p', header]]) {
    for (const [itemId, title, value] of [['doc-emoji', '表情', '👨‍👩‍👧‍👦'], ['doc-symbol', '符号', '|a|e\u0301'], ['doc-symbol', '符号', ' '], ['doc-symbol', '符号', '\u00a0']]) {
      await page.goto(baseURL + query); await page.waitForFunction(() => !!window.documentStructuredFixture)
      await page.evaluate(({ variant, dark, paragraphId, host }) => {
        window.documentStructuredFixture.variant(variant); window.documentStructuredFixture.dark(dark)
        window.documentStructuredFixture.select(paragraphId, 1, paragraphId, 3, host)
      }, { variant, dark, paragraphId, host })
      const before = await page.evaluate(({ paragraphId, host }) => window.documentStructuredFixture.block(paragraphId, host), { paragraphId, host })
      const previous = await history(), input = page.getByRole('textbox', { name: '文档文字输入', exact: true })
      await input.focus()
      await page.getByRole('tab', { name: '插入', exact: true }).click()
      const button = page.locator(`.ribbon-visible-groups [data-ribbon-item="${itemId}"] button`).first()
      const background = await button.evaluate(element => getComputedStyle(element).backgroundColor)
      await button.hover()
      await expect.poll(() => button.evaluate(element => getComputedStyle(element).backgroundColor)).not.toBe(background)
      const geometry = await button.evaluate(element => {
        const button = element.getBoundingClientRect(), arrow = element.querySelector('.ribbon-dropdown-chevron').getBoundingClientRect()
        return { height: button.height, inside: arrow.left >= button.left && arrow.right <= button.right,
          center: Math.abs(arrow.top + arrow.height / 2 - button.top - button.height / 2) }
      })
      assert.equal(geometry.height, variant === 'classic' ? 28 : 24); assert.equal(geometry.inside, true); assert.ok(geometry.center <= 0.05)
      await button.click()
      const popup = page.getByRole('dialog', { name: title, exact: true })
      await expect(popup.locator('.ant-form')).toHaveCount(2)
      const cell = popup.locator('.character-grid .ant-btn').first(), character = await cell.textContent()
      assert.equal(await cell.evaluate(element => element.getBoundingClientRect().height), 42)
      await cell.hover(); await expect(page.getByRole('tooltip')).toContainText(character)
      await popup.getByRole('textbox', { name: '自定义字符', exact: true }).fill(value)
      await popup.getByRole('button', { name: '插入字符', exact: true }).click()
      await expect(popup).toBeHidden(); await expect(input).toBeFocused()
      const originalText = before.runs.map(run => run.text).join(''), expected = originalText.slice(0, 1) + value + originalText.slice(3)
      await expect.poll(() => page.evaluate(({ paragraphId, host }) => window.documentStructuredFixture.block(paragraphId, host).runs.map(run => run.text).join(''), { paragraphId, host })).toBe(expected)
      const events = await page.evaluate(() => window.documentStructuredFixture.commandEvents().filter(event => event.id === 'doc.command.insert-special-character' || event.id === 'doc.command.insert-text'))
      assert.deepEqual(events, ['before', 'after'].map(phase => ({ phase, id: 'doc.command.insert-special-character', value, unitId: 'structured-browser', kind: 'document' })))
      assert.equal((await history()).undo, previous.undo + 1)
      const location = await page.evaluate(() => window.documentStructuredFixture.selection())
      assert.equal(location.kind, 'body'); assert.equal(location.selection.focus.offset, 1 + value.length)
      assert.deepEqual(location.host, host)
      await page.evaluate(() => window.documentStructuredFixture.undo())
      assert.deepEqual(await page.evaluate(({ paragraphId, host }) => window.documentStructuredFixture.block(paragraphId, host), { paragraphId, host }), before)
      await page.evaluate(() => window.documentStructuredFixture.redo())
      await page.evaluate(({ paragraphId, host }) => window.documentStructuredFixture.saveAndReopen(paragraphId, host), { paragraphId, host })
      assert.equal(await page.evaluate(({ paragraphId, host }) => window.documentStructuredFixture.block(paragraphId, host).runs.map(run => run.text).join(''), { paragraphId, host }), expected)
    }
  }
})

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

const group = (id = 'columns-three') => page.evaluate(id => window.documentStructuredFixture.group(id), id)
async function boundary(id = 'columns-three') {
  await page.evaluate(id => window.documentStructuredFixture.select(id + '-p-0'), id)
  const handle = page.locator('[data-document-column-group="' + id + '"][data-document-column-boundary="0"]').first()
  await expect(handle).toBeVisible(); return handle
}
async function drag(handle, delta) {
  const box = await handle.boundingBox()
  assert.ok(box)
  const x = box.x + box.width / 2, y = box.y + Math.min(20, box.height / 2)
  await page.mouse.move(x, y); await page.mouse.down(); await page.mouse.move(x + delta, y, { steps: 4 })
}

test('column drag previews without content history, commits once on release, and Escape cancels a later gesture', async () => {
  const handle = await boundary(), before = await group(), previous = await history()
  await drag(handle, 36)
  await expect(page.locator('.structured-column-guide')).toBeVisible()
  assert.deepEqual(await group(), before); assert.deepEqual(await history(), previous)
  await page.mouse.up()
  await expect.poll(group).not.toEqual(before)
  assert.equal((await history()).undo, previous.undo + 1)
  const after = await group(), committedHistory = await history(), later = await boundary()
  await drag(later, -28); await page.keyboard.press('Escape'); await page.mouse.up()
  await expect(page.locator('.structured-column-guide')).toHaveCount(0)
  assert.deepEqual(await group(), after); assert.deepEqual(await history(), committedHistory)
  await page.evaluate(() => window.documentStructuredFixture.undo()); await expect.poll(group).toEqual(before)
  await page.evaluate(() => window.documentStructuredFixture.redo()); await expect.poll(group).toEqual(after)
})

test('keyboard boundary edits keep the other column weight and use the same one-command history', async () => {
  const handle = await boundary(), before = await group(), previous = await history()
  await handle.press('Shift+ArrowRight')
  await expect.poll(group).not.toEqual(before)
  const next = await group()
  assert.equal(next.columns[2].widthRatio, before.columns[2].widthRatio)
  assert.ok(Math.abs(next.columns[0].widthRatio + next.columns[1].widthRatio - 2) < 1e-8)
  assert.equal((await history()).undo, previous.undo + 1)
})

test('column property dimensions use ant number fields in both themes and persist through save/reopen', async () => {
  for (const dark of [false, true]) {
    await page.evaluate(dark => window.documentStructuredFixture.dark(dark), dark)
    await boundary('columns-two')
    await page.locator('[data-document-structured-toolbar]').getByRole('button', { name: '属性', exact: true }).click()
    const dialog = page.getByRole('dialog'), previous = await history(), padding = dark ? 12 : 10, minimum = dark ? 108 : 96
    const paddingInput = dialog.getByRole('spinbutton', { name: '栏水平内边距', exact: true }), heightInput = dialog.getByRole('spinbutton', { name: '栏最小高度', exact: true })
    await expect(paddingInput.locator('xpath=ancestor::div[contains(@class,"ant-input-number")][1]')).toBeVisible()
    await paddingInput.fill(String(padding)); await heightInput.fill(String(minimum))
    await dialog.getByRole('button', { name: '应用', exact: true }).click(); await expect(dialog).toBeHidden()
    const config = await group('columns-two')
    assert.equal(config.horizontalPadding, padding); assert.equal(config.minHeight, minimum); assert.equal((await history()).undo, previous.undo + 1)
    await expect(page.getByRole('textbox', { name: '文档文字输入', exact: true })).toBeFocused()
    await page.evaluate(() => window.documentStructuredFixture.saveAndReopen())
    assert.equal((await group('columns-two')).horizontalPadding, padding); assert.equal((await group('columns-two')).minHeight, minimum)
  }
})

test('nowrap code Slider keeps a stable maximum and view-only history, then allows real editing at the scrolled end', async () => {
  for (const dark of [false, true]) {
    await page.evaluate(dark => { window.documentStructuredFixture.dark(dark); window.documentStructuredFixture.select('long-code-p') }, dark)
    const slider = page.locator('[data-document-code-block="long-code"]').getByRole('slider', { name: '代码块水平滚动', exact: true }).first()
    await expect(slider).toBeVisible()
    const before = await page.evaluate(() => window.documentStructuredFixture.snapshot()), previous = await history(), maximum = await slider.getAttribute('aria-valuemax')
    assert.ok(Number(maximum) > 0)
    await slider.press('End'); await expect.poll(async () => Number(await slider.getAttribute('aria-valuenow'))).toBeGreaterThan(Number(maximum) - 2)
    assert.equal(await slider.getAttribute('aria-valuemax'), maximum)
    await slider.press('Home'); await expect(slider).toHaveAttribute('aria-valuenow', '0'); assert.equal(await slider.getAttribute('aria-valuemax'), maximum)
    await slider.press('End')
    assert.deepEqual(await page.evaluate(() => window.documentStructuredFixture.snapshot()), before); assert.deepEqual(await history(), previous)
    const offset = Number(await slider.getAttribute('aria-valuenow')), text = await page.evaluate(() => window.documentStructuredFixture.longSource())
    const point = await page.evaluate(({ length, offset }) => window.documentStructuredFixture.point('long-code-p', length, offset), { length: text.length, offset })
    await page.mouse.click(point.x - 1, point.y); await page.keyboard.press('End'); await page.keyboard.insertText('Z')
    await expect.poll(() => page.evaluate(() => window.documentStructuredFixture.longSource())).toBe(text + 'Z')
    assert.equal((await history()).undo, previous.undo + 1)
  }
})

test('floating language and paragraph insertion menus publish actual commands and preserve source paragraphs', async () => {
  await page.evaluate(() => window.documentStructuredFixture.select('code-p'))
  const before = await source(), previous = await history(), toolbar = page.locator('[data-document-structured-toolbar]')
  await toolbar.getByRole('button', { name: '代码块语言选择', exact: true }).click()
  const controls = page.locator('.document-code-controls:visible')
  await controls.getByRole('combobox', { name: '代码块语言', exact: true }).click()
  await page.getByRole('option', { name: 'Rust', exact: true }).click()
  await controls.getByRole('button', { name: '应用', exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.documentStructuredFixture.snapshot().blocks.find(block => block.id === 'code').code.language)).toBe('rust')
  assert.equal(await source(), before); assert.equal((await history()).undo, previous.undo + 1)
  await page.evaluate(() => window.documentStructuredFixture.select('tail'))
  const count = await page.evaluate(() => window.documentStructuredFixture.snapshot().blocks.filter(block => block.type === 'quote').length), earlier = await history()
  await page.locator('[data-document-paragraph-insert]').getByRole('button', { name: '段落操作', exact: true }).click()
  await page.getByRole('menuitem', { name: '在段落下方插入', exact: true }).hover()
  await page.getByRole('menuitem', { name: '引用块', exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.documentStructuredFixture.snapshot().blocks.filter(block => block.type === 'quote').length)).toBe(count + 1)
  const snapshot = await page.evaluate(() => window.documentStructuredFixture.snapshot()), tail = snapshot.blocks.findIndex(block => block.id === 'tail')
  assert.equal(snapshot.blocks[tail].runs.map(run => run.text).join(''), 'Outside')
  assert.equal(snapshot.blocks[tail + 1].type, 'quote'); assert.ok(snapshot.blocks[tail + 1].children[0].runs)
  assert.equal((await history()).undo, earlier.undo + 1)
})

test('source replacement closes a floating menu and a detached stale action cannot overwrite the newer ratios', async () => {
  await boundary()
  await page.locator('[data-document-structured-toolbar]').getByRole('button', { name: '高级块操作', exact: true }).click()
  const stale = await page.getByRole('menuitem', { name: '各栏等宽', exact: true }).elementHandle()
  assert.ok(stale)
  await page.evaluate(() => window.documentStructuredFixture.remoteRatios([2, 1, 1]))
  const newer = await group(), previous = await history()
  await expect(page.getByRole('menuitem', { name: '各栏等宽', exact: true })).toHaveCount(0)
  await stale.evaluate(element => element.click())
  assert.deepEqual(await group(), newer); assert.deepEqual(await history(), previous)
})

test('real header column geometry supports drag, property edits and save/reopen through the same document history', async () => {
  await page.goto(baseURL + '?auxiliary=1'); await page.waitForFunction(() => !!window.documentStructuredFixture)
  await page.getByRole('textbox', { name: '文档文字输入', exact: true }).focus()
  const host = { kind: 'story', sectionId: 'first', story: 'header' }
  await page.evaluate(host => window.documentStructuredFixture.selectHost('header-columns-p-0', host), host)
  const handle = page.locator('[data-document-column-group="header-columns"][data-document-column-boundary="0"]').first()
  await expect(handle).toBeVisible()
  const initial = await page.evaluate(host => window.documentStructuredFixture.group('header-columns', host), host), body = await page.evaluate(() => window.documentStructuredFixture.snapshot().blocks), previous = await history()
  await drag(handle, 28)
  assert.deepEqual(await page.evaluate(host => window.documentStructuredFixture.group('header-columns', host), host), initial); assert.deepEqual(await history(), previous)
  await page.mouse.up()
  await expect.poll(() => page.evaluate(host => window.documentStructuredFixture.group('header-columns', host), host)).not.toEqual(initial)
  assert.equal((await history()).undo, previous.undo + 1)
  await page.locator('[data-document-structured-toolbar]').getByRole('button', { name: '属性', exact: true }).click()
  const dialog = page.getByRole('dialog')
  await dialog.getByRole('spinbutton', { name: '栏水平内边距', exact: true }).fill('12')
  await dialog.getByRole('spinbutton', { name: '栏最小高度', exact: true }).fill('96')
  await dialog.getByRole('button', { name: '应用', exact: true }).click(); await expect(dialog).toBeHidden()
  const updated = await page.evaluate(host => window.documentStructuredFixture.group('header-columns', host), host)
  assert.equal(updated.horizontalPadding, 12); assert.equal(updated.minHeight, 96); assert.equal((await history()).undo, previous.undo + 2)
  assert.deepEqual(await page.evaluate(() => window.documentStructuredFixture.snapshot().blocks), body)
  await page.evaluate(() => window.documentStructuredFixture.saveAndReopen())
  assert.deepEqual(await page.evaluate(host => window.documentStructuredFixture.group('header-columns', host), host), updated)
  await page.evaluate(host => window.documentStructuredFixture.selectHost('header-columns-p-0', host), host)
  await expect(page.locator('[data-document-column-group="header-columns"][data-document-column-boundary="0"]').first()).toBeVisible()
})

const block = (id = 'code', host) => page.evaluate(({ id, host }) => window.documentStructuredFixture.block(id, host), { id, host })
const controls = () => page.locator('.document-code-controls:visible')
async function chooseCodeOption(name, label) {
  await controls().getByRole('combobox', { name, exact: true }).click()
  await page.getByRole('option', { name: label, exact: true }).click()
}

test('desktop Code long font and author language remain inside the right edge with full Ant tooltips', async () => {
  await page.setViewportSize({ width: 900, height: 900 })
  await page.goto(baseURL + '?long-controls=1'); await page.waitForFunction(() => !!window.documentStructuredFixture)
  const values = await page.evaluate(() => window.documentStructuredFixture.codeControlValues), before = await block(), previous = await history()
  const input = page.getByRole('textbox', { name: '文档文字输入', exact: true })
  for (const [dark, variant] of [[false, 'classic'], [true, 'grid']]) {
    await page.evaluate(({ dark, variant }) => {
      window.documentStructuredFixture.dark(dark); window.documentStructuredFixture.variant(variant); window.documentStructuredFixture.select('code-p')
    }, { dark, variant })
    await input.focus()
    const toolbar = page.locator('[data-document-structured-toolbar]'); await expect(toolbar).toBeVisible()
    const bounds = await toolbar.evaluate(element => {
      const rectangle = element.getBoundingClientRect(), viewport = element.closest('.document-canvas-viewport').getBoundingClientRect()
      return { left: rectangle.left, right: rectangle.right, viewportLeft: viewport.left, viewportRight: viewport.right, viewportWidth: viewport.width }
    })
    assert.ok(bounds.viewportWidth > 640); assert.ok(bounds.left >= bounds.viewportLeft + 8 - 1); assert.ok(bounds.right <= bounds.viewportRight - 8 + 1)
    for (const [name, value, field] of [['代码块字体', values.fontFamily, '代码块等宽字体'], ['代码块语言选择', values.language, '代码块语言']]) {
      const trigger = toolbar.getByRole('button', { name, exact: true })
      await expect(trigger.locator('.structured-code-control__label')).toHaveText(value)
      const geometry = await trigger.evaluate(element => {
        const button = element.getBoundingClientRect(), label = element.querySelector('.structured-code-control__label'), text = label.getBoundingClientRect()
        const style = getComputedStyle(element)
        return { width: button.width, left: button.left, right: button.right, textLeft: text.left, textRight: text.right,
          innerLeft: button.left + parseFloat(style.borderLeftWidth) + parseFloat(style.paddingLeft),
          innerRight: button.right - parseFloat(style.borderRightWidth) - parseFloat(style.paddingRight), clipped: label.scrollWidth > label.clientWidth }
      })
      assert.ok(geometry.width <= 119); assert.ok(geometry.left >= bounds.left + 4 - 1); assert.ok(geometry.right <= bounds.right - 4 + 1)
      assert.ok(geometry.textLeft >= geometry.innerLeft - 1); assert.ok(geometry.textRight <= geometry.innerRight + 1); assert.equal(geometry.clipped, true)
      await trigger.hover(); await expect(page.getByRole('tooltip').filter({ hasText: value })).toHaveText(value)
      const hovered = await trigger.boundingBox(); assert.ok(hovered); assert.ok(Math.abs(hovered.width - geometry.width) <= 1)
      await expect(input).toBeFocused(); await trigger.click()
      await expect(controls().getByRole('combobox', { name: field, exact: true })).toBeVisible()
      await controls().getByRole('button', { name: '取消', exact: true }).click(); await expect(controls()).toHaveCount(0); await expect(input).toBeFocused()
    }
    assert.deepEqual(await block(), before); assert.deepEqual(await history(), previous)
  }
  await page.evaluate(() => window.documentStructuredFixture.saveAndReopen())
  assert.equal((await block()).code.language, values.language); assert.equal((await block()).children[0].runs[0].fontFamily, values.fontFamily)
})

test('desktop Code font Popover uses Ant controls in both themes, cancels without history, and applies real font formatting', async () => {
  for (const dark of [false, true]) {
    await page.evaluate(dark => { window.documentStructuredFixture.dark(dark); window.documentStructuredFixture.select('code-p') }, dark)
    const sourceBefore = await source(), before = await block(), previous = await history(), font = dark ? 'Courier New' : 'Menlo'
    await page.locator('[data-document-structured-toolbar]').getByRole('button', { name: '代码块字体', exact: true }).click()
    await expect(page.locator('.ant-popover:visible .ant-form')).toBeVisible()
    await chooseCodeOption('代码块等宽字体', font)
    await controls().getByRole('button', { name: '取消', exact: true }).click()
    assert.deepEqual(await block(), before); assert.deepEqual(await history(), previous)
    await expect(page.getByRole('textbox', { name: '文档文字输入', exact: true })).toBeFocused()
    await page.locator('[data-document-structured-toolbar]').getByRole('button', { name: '代码块字体', exact: true }).click()
    await chooseCodeOption('代码块等宽字体', font); await controls().getByRole('button', { name: '应用', exact: true }).click()
    await expect.poll(async () => (await block()).children[0].runs[0].fontFamily).toBe(font)
    const current = await block()
    assert.ok(current.children.every(child => child.metadata.documentCodeFontFamily === font && child.runs.every(run => run.fontFamily === font && run.eastAsiaFontFamily === font)))
    assert.equal(await source(), sourceBefore); assert.equal((await history()).undo, previous.undo + 1)
    await page.evaluate(() => window.documentStructuredFixture.undo()); assert.deepEqual(await block(), before)
    await page.evaluate(() => window.documentStructuredFixture.redo()); assert.deepEqual(await block(), current)
    await page.evaluate(() => window.documentStructuredFixture.saveAndReopen())
    const reopened = await block(); assert.deepEqual(reopened.code, current.code)
    assert.ok(reopened.children.every(child => child.runs.every(run => run.fontFamily === font && run.eastAsiaFontFamily === font)))
    assert.equal(await source(), sourceBefore)
  }
})

test('desktop Code More edits all four settings once and preserves source and font runs', async () => {
  await page.evaluate(() => window.documentStructuredFixture.select('code-p'))
  const before = await block(), previous = await history()
  await page.locator('[data-document-structured-toolbar]').getByRole('button', { name: '高级块操作', exact: true }).click()
  await chooseCodeOption('代码块语言', 'Ruby')
  await controls().getByRole('switch', { name: '代码块自动换行', exact: true }).click()
  await controls().getByRole('switch', { name: '代码块显示行号', exact: true }).click()
  await controls().getByRole('spinbutton', { name: '代码块缩进空格数', exact: true }).fill('2')
  await controls().getByRole('button', { name: '应用', exact: true }).click()
  await expect.poll(async () => (await block()).code).toEqual({ language: 'ruby', wrap: false, showLineNumbers: false, tabSize: 2 })
  assert.deepEqual((await block()).children, before.children); assert.equal((await history()).undo, previous.undo + 1)
  await expect(page.getByRole('textbox', { name: '文档文字输入', exact: true })).toBeFocused()
  await page.evaluate(() => window.documentStructuredFixture.undo()); assert.deepEqual(await block(), before)
  await page.evaluate(() => window.documentStructuredFixture.redo()); assert.equal((await block()).code.tabSize, 2)
})

test('mobile Code font and More use bottom Ant Drawers and return focus to the current Canvas input', async () => {
  await page.setViewportSize({ width: 390, height: 844 })
  await page.evaluate(() => { window.documentStructuredFixture.dark(true); window.documentStructuredFixture.select('code-p') })
  const previous = await history(), toolbar = page.locator('[data-document-mobile-toolbar="true"]')
  await toolbar.getByRole('button', { name: '代码块字体', exact: true }).click()
  const drawer = page.getByRole('dialog'); await expect(drawer.locator('.ant-form')).toBeVisible()
  await chooseCodeOption('代码块等宽字体', 'Consolas'); await controls().getByRole('button', { name: '应用', exact: true }).click()
  await expect(drawer).toBeHidden(); assert.equal((await block()).children[0].runs[0].fontFamily, 'Consolas')
  await expect(page.getByRole('textbox', { name: '文档文字输入', exact: true })).toBeFocused()
  await toolbar.getByRole('button', { name: '高级块操作', exact: true }).click()
  await expect(drawer).toBeVisible(); await controls().getByRole('spinbutton', { name: '代码块缩进空格数', exact: true }).fill('3')
  await controls().getByRole('button', { name: '应用', exact: true }).click(); await expect(drawer).toBeHidden()
  assert.equal((await block()).code.tabSize, 3); assert.equal((await history()).undo, previous.undo + 2)
  await expect(page.getByRole('textbox', { name: '文档文字输入', exact: true })).toBeFocused()
  await toolbar.getByRole('button', { name: '代码块语言选择', exact: true }).click()
  await expect(drawer).toBeVisible(); await controls().getByRole('button', { name: '取消', exact: true }).click(); await expect(drawer).toBeHidden()
  assert.equal((await history()).undo, previous.undo + 2)
})

test('detached Code property controls cannot submit after selection or source has changed', async () => {
  await page.evaluate(() => window.documentStructuredFixture.select('code-p'))
  await page.locator('[data-document-structured-toolbar]').getByRole('button', { name: '代码块字体', exact: true }).click()
  await chooseCodeOption('代码块等宽字体', 'Menlo')
  const stale = await controls().getByRole('button', { name: '应用', exact: true }).elementHandle(); assert.ok(stale)
  await page.evaluate(() => window.documentStructuredFixture.select('quote-p'))
  await expect(controls()).toHaveCount(0)
  const before = await block(), previous = await history()
  await stale.evaluate(element => element.click())
  assert.deepEqual(await block(), before); assert.deepEqual(await history(), previous)
  await page.evaluate(() => window.documentStructuredFixture.select('code-p'))
  await page.locator('[data-document-structured-toolbar]').getByRole('button', { name: '高级块操作', exact: true }).click()
  const oldApply = await controls().getByRole('button', { name: '应用', exact: true }).elementHandle(); assert.ok(oldApply)
  await page.evaluate(() => window.documentStructuredFixture.changeSource('remote value'))
  await expect(controls()).toHaveCount(0)
  const newer = await block(), at = await history(); await oldApply.evaluate(element => element.click())
  assert.deepEqual(await block(), newer); assert.deepEqual(await history(), at)
})

test('paragraph T converts Code without a nested wrapper and mobile empty-paragraph insertion creates real content', async () => {
  await page.evaluate(() => window.documentStructuredFixture.select('code-p'))
  const before = await block(), previous = await history()
  await page.locator('[data-document-paragraph-insert]').getByRole('button', { name: '段落操作', exact: true }).click()
  await page.getByRole('menuitem', { name: '转换段落', exact: true }).hover()
  await page.getByRole('menuitem', { name: '引用块', exact: true }).click()
  await expect.poll(async () => (await block()).type).toBe('quote')
  const converted = await block()
  assert.deepEqual(converted.children.map(child => ({ id: child.id, runs: child.runs })), before.children.map(child => ({ id: child.id, runs: child.runs })))
  assert.deepEqual(converted.children.map(child => [child.style.keepLines, child.style.keepNext, child.style.widowControl]), [[true, true, true], [true, false, true]])
  assert.equal(converted.id, before.id)
  assert.equal((await history()).undo, previous.undo + 1)
  await page.evaluate(() => window.documentStructuredFixture.undo()); assert.deepEqual(await block(), before)
  await page.setViewportSize({ width: 390, height: 844 }); await page.goto(baseURL + '?blank=1')
  await page.waitForFunction(() => !!window.documentStructuredFixture); await page.evaluate(() => window.documentStructuredFixture.select('empty'))
  await page.locator('[data-document-paragraph-menu="quick-insert"]').getByRole('button', { name: '段落操作', exact: true }).click()
  const drawer = page.getByRole('dialog'); await expect(drawer).toBeVisible()
  await drawer.getByRole('menuitem', { name: '插入', exact: true }).click()
  await drawer.getByRole('menuitem', { name: '代码块', exact: true }).click(); await expect(drawer).toBeHidden()
  const current = await page.evaluate(() => window.documentStructuredFixture.snapshot())
  assert.equal(current.blocks[0].type, 'code'); assert.equal(current.blocks[0].children[0].id, 'empty'); assert.equal((await history()).undo, 1)
  await expect(page.getByRole('textbox', { name: '文档文字输入', exact: true })).toBeFocused()
})

test('table-below Canvas menu inserts a real Quote or three-column group after the table and preserves its cells', async () => {
  await page.goto(baseURL + '?table=1'); await page.waitForFunction(() => !!window.documentStructuredFixture)
  await page.evaluate(() => window.documentStructuredFixture.select('table-cell'))
  const before = await block('table'), previous = await history(), trigger = page.locator('[data-document-table-below-insert][data-document-table-id="table"]')
  await expect(trigger).toBeVisible(); await trigger.getByRole('button', { name: '在表格下方插入', exact: true }).click()
  await page.getByRole('menuitem', { name: '引用块', exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.documentStructuredFixture.snapshot().blocks[1].type)).toBe('quote')
  assert.deepEqual(await block('table'), before); assert.equal((await history()).undo, previous.undo + 1)
  await page.evaluate(() => { window.documentStructuredFixture.undo(); window.documentStructuredFixture.select('table-cell') })
  await trigger.getByRole('button', { name: '在表格下方插入', exact: true }).click()
  await page.getByRole('menuitem', { name: '分栏', exact: true }).hover(); await page.getByRole('menuitem', { name: '3 栏', exact: true }).click()
  const current = await page.evaluate(() => window.documentStructuredFixture.snapshot())
  assert.equal(current.blocks[1].type, 'columns'); assert.equal(current.blocks[1].columns.columns.length, 3)
  assert.deepEqual(await block('table'), before); assert.equal((await history()).undo, previous.undo + 1)
})

test('auxiliary content-insert menu stays in the header and shares the root document history', async () => {
  await page.goto(baseURL + '?auxiliary=1'); await page.waitForFunction(() => !!window.documentStructuredFixture)
  const host = { kind: 'story', sectionId: 'first', story: 'header' }
  await page.evaluate(host => window.documentStructuredFixture.selectHost('header-code-p', host), host)
  const body = await page.evaluate(() => window.documentStructuredFixture.snapshot().blocks), sourceBefore = await block('header-code', host), previous = await history()
  await page.locator('[data-document-content-insert]').getByRole('button', { name: '在下方插入内容', exact: true }).click()
  await page.getByRole('menuitem', { name: '标注块', exact: true }).click()
  const current = await page.evaluate(() => window.documentStructuredFixture.snapshot().sectionHeaderFooters.first.storyContent.header.blocks)
  const index = current.findIndex(value => value.id === 'header-code')
  assert.equal(current[index + 1].type, 'callout'); assert.deepEqual(await block('header-code', host), sourceBefore)
  assert.deepEqual(await page.evaluate(() => window.documentStructuredFixture.snapshot().blocks), body)
  assert.equal((await history()).undo, previous.undo + 1)
  await page.evaluate(() => window.documentStructuredFixture.undo())
  assert.deepEqual(await block('header-code', host), sourceBefore)
})

const paginationLayout = () => page.evaluate(() => window.documentStructuredFixture.paginationLayout())
const pagesFor = (layout, id) => [...new Set(layout.lines.filter(line => line.blockId === id).map(line => line.page))]
const paginationCheckbox = (panel, field) => panel.locator('[data-document-pagination-field="' + field + '"]')
async function openPaginationSettings(locale = 'zh-CN') {
  const point = await page.evaluate(() => window.documentStructuredFixture.point('pagination-first'))
  await page.mouse.click(point.x + 1, point.y, { button: 'right' })
  await page.getByRole('menuitem', { name: locale === 'en-US' ? 'Paragraph settings' : '段落设置', exact: true }).click()
  const panel = page.locator('[data-document-paragraph-settings]'); await expect(panel).toBeVisible(); return panel
}

test('a fitting Quote moves its real Canvas content and decoration together while retaining UTF-16 input positions', async () => {
  await page.goto(baseURL + '?pagination=fitting'); await page.waitForFunction(() => !!window.documentStructuredFixture)
  const layout = await paginationLayout(), quote = await block('pagination-quote')
  assert.deepEqual(pagesFor(layout, 'pagination-before'), [0]); assert.deepEqual(pagesFor(layout, 'pagination-first'), [1])
  assert.deepEqual(pagesFor(layout, 'pagination-last'), [1]); assert.equal(quote.children[1].style.keepNext, false)
  assert.equal((await block('tail')).style.keepLines, undefined)
  const frames = layout.containers.filter(frame => frame.blockId === quote.id)
  assert.equal(frames.length, 1); assert.equal(frames[0].page, 1)
  const first = layout.lines.find(line => line.blockId === 'pagination-first')
  assert.ok(frames[0].y < first.y); await expect(page.locator('.document-scroll canvas').first()).toBeVisible()
  const point = await page.evaluate(() => window.documentStructuredFixture.point('pagination-first', 8))
  await page.mouse.click(point.x + 0.1, point.y)
  const location = await page.evaluate(() => window.documentStructuredFixture.selection())
  assert.equal(location.kind, 'body'); assert.equal(location.selection.focus.blockId, 'pagination-first'); assert.equal(location.selection.focus.offset, 8)
  assert.equal((await history()).undo, 0)
})

test('four pagination settings use Ant checkboxes and mixed selections share one command in both themes and locales', async () => {
  for (const [locale, dark] of [['zh-CN', false], ['en-US', true]]) {
    await page.goto(baseURL + '?pagination=fitting&locale=' + locale); await page.waitForFunction(() => !!window.documentStructuredFixture)
    await page.evaluate(dark => window.documentStructuredFixture.dark(dark), dark)
    const panel = await openPaginationSettings(locale)
    await expect(panel.locator('.ant-form')).toBeVisible(); await expect(panel.locator('[data-document-pagination-field]')).toHaveCount(4)
    for (const field of ['keepLines', 'keepNext', 'widowControl', 'pageBreakBefore']) {
      const checkbox = paginationCheckbox(panel, field)
      assert.ok(await checkbox.evaluate(element => !!element.closest('.ant-checkbox-wrapper') || !!element.querySelector('.ant-checkbox')))
    }
    await expect(panel.getByText(locale === 'en-US' ? 'Page break before' : '段前分页', { exact: true })).toBeVisible()
    await page.evaluate(() => window.documentStructuredFixture.select('pagination-first', 0, 'tail', 7))
    const mixed = panel.getByRole('checkbox', { name: locale === 'en-US' ? 'Keep lines together' : '段中不分页', exact: true })
    await expect(mixed).toHaveAttribute('aria-checked', 'mixed')
    const before = await page.evaluate(() => window.documentStructuredFixture.snapshot().blocks), previous = await history()
    await mixed.click()
    await expect.poll(async () => [(await block('pagination-first')).style.keepLines, (await block('pagination-last')).style.keepLines, (await block('tail')).style.keepLines]).toEqual([false, false, false])
    assert.equal((await history()).undo, previous.undo + 1)
    const after = await page.evaluate(() => window.documentStructuredFixture.snapshot().blocks)
    await page.evaluate(() => window.documentStructuredFixture.undo()); assert.deepEqual(await page.evaluate(() => window.documentStructuredFixture.snapshot().blocks), before)
    await page.evaluate(() => window.documentStructuredFixture.redo()); assert.deepEqual(await page.evaluate(() => window.documentStructuredFixture.snapshot().blocks), after)
    await panel.getByRole('button', { name: locale === 'en-US' ? 'Close paragraph settings' : '关闭段落设置', exact: true }).click()
    await expect(page.getByRole('textbox', { name: '文档文字输入', exact: true })).toBeFocused()
    await page.evaluate(() => window.documentStructuredFixture.saveAndReopen('pagination-first'))
    assert.equal((await block('pagination-first')).style.keepLines, false); assert.equal((await block('pagination-last')).style.keepLines, false)
  }
})

test('synthetic paragraph-panel composition defers a pagination patch until one committed history item', async () => {
  await page.goto(baseURL + '?pagination=fitting'); await page.waitForFunction(() => !!window.documentStructuredFixture)
  const panel = await openPaginationSettings(), previous = await history(), before = await block('pagination-first')
  await panel.dispatchEvent('compositionstart', { data: '中' })
  await panel.getByRole('checkbox', { name: '与下段同页', exact: true }).click()
  assert.deepEqual(await block('pagination-first'), before); assert.deepEqual(await history(), previous)
  await panel.dispatchEvent('keydown', { key: 'Escape', isComposing: true, keyCode: 229 }); await expect(panel).toBeVisible()
  await panel.dispatchEvent('compositionend', { data: '' })
  await expect.poll(async () => (await block('pagination-first')).style.keepNext).toBe(false)
  assert.equal((await history()).undo, previous.undo + 1)
  await panel.getByRole('button', { name: '关闭段落设置', exact: true }).click()
  await expect(page.getByRole('textbox', { name: '文档文字输入', exact: true })).toBeFocused()
})

test('public modes preserve Quote rules through body save/reopen and only Traditional applies keeping', async () => {
  for (const flavor of [0, 1, 2]) {
    await page.goto(baseURL + '?pagination=fitting&flavor=' + flavor); await page.waitForFunction(() => !!window.documentStructuredFixture)
    const snapshot = await page.evaluate(() => window.documentStructuredFixture.publicSnapshot()), expectedPage = flavor === 1 ? 1 : 0
    assert.equal(snapshot.documentStyle.documentFlavor, flavor)
    assert.deepEqual(pagesFor(await paginationLayout(), 'pagination-first'), [expectedPage])
    const saved = await page.evaluate(() => window.documentStructuredFixture.saveAndReopen('pagination-first'))
    assert.equal(saved.documentStyle.documentFlavor, flavor)
    assert.deepEqual((await block('pagination-quote')).children.map(child => [child.style.keepLines, child.style.keepNext, child.style.widowControl]), [[true, true, true], [true, false, true]])
    assert.deepEqual(pagesFor(await paginationLayout(), 'pagination-first'), [expectedPage])
    await page.evaluate(() => window.documentStructuredFixture.unwrapQuote('pagination-quote'))
    const reopened = await page.evaluate(() => window.documentStructuredFixture.publicSnapshot())
    const first = reopened.body.paragraphs.find(paragraph => paragraph.paragraphId === 'pagination-first').paragraphStyle
    assert.equal(first.keepLines, undefined); assert.equal(first.keepNext, undefined); assert.equal(first.widowControl, undefined)
  }
})

test('an oversized Quote splits into visible Canvas fragments without generating empty decorated pages', async () => {
  await page.goto(baseURL + '?pagination=oversized'); await page.waitForFunction(() => !!window.documentStructuredFixture)
  const before = await block('pagination-quote'), layout = await paginationLayout(), lines = layout.lines.filter(line => line.blockId === 'pagination-first')
  assert.equal(lines.length, 15); assert.ok(pagesFor(layout, 'pagination-first').length > 1)
  for (const frame of layout.containers.filter(frame => frame.blockId === 'pagination-quote')) assert.ok(lines.some(line => line.page === frame.page))
  for (const current of layout.pages) assert.ok(layout.lines.some(line => line.page === current.index))
  assert.deepEqual(await block('pagination-quote'), before); assert.equal((await history()).undo, 0)
  await page.evaluate(() => window.documentStructuredFixture.saveAndReopen('pagination-first'))
  assert.equal((await paginationLayout()).lines.filter(line => line.blockId === 'pagination-first').length, 15)
})

test('real Canvas Enter updates the old and new last Quote rules with shared undo, redo and body persistence', async () => {
  await page.goto(baseURL + '?pagination=fitting'); await page.waitForFunction(() => !!window.documentStructuredFixture)
  await page.evaluate(() => window.documentStructuredFixture.select('pagination-last', 10))
  await page.getByRole('textbox', { name: '文档文字输入', exact: true }).focus()
  const before = await block('pagination-quote'), previous = await history()
  await page.keyboard.press('Enter')
  await expect.poll(async () => (await block('pagination-quote')).children.length).toBe(3)
  const current = await block('pagination-quote')
  assert.deepEqual(current.children.map(child => child.style.keepNext), [true, true, false])
  assert.equal((await block('tail')).style.keepNext, undefined); assert.equal((await history()).undo, previous.undo + 1)
  await page.evaluate(() => window.documentStructuredFixture.undo()); assert.deepEqual(await block('pagination-quote'), before)
  await page.evaluate(() => window.documentStructuredFixture.redo()); assert.deepEqual(await block('pagination-quote'), current)
  await page.evaluate(id => window.documentStructuredFixture.saveAndReopen(id), current.children[2].id)
  assert.deepEqual((await block('pagination-quote')).children.map(child => child.style.keepNext), [true, true, false])
})
