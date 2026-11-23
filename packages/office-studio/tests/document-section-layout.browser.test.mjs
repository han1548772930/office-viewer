import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

// Deferred Canvas/antdv-next/public facade drafts, 2026-11-22. No browser run.
// This fixture does not cover the production application history adapter or a real Windows IME.
const fixturePath = '/__document-section-layout-test__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>Docs 节布局</title><style>body{margin:0}#app{height:100vh;min-width:0}</style></head><body><div id="app"></div><script type="module">
import { computed, createApp, h, ref, shallowRef, watch } from 'vue'
import { ConfigProvider, theme } from 'antdv-next'
import { OfficeStudio } from '/packages/office-studio/src/api/office-studio'
import { FDocument } from '/packages/office-studio/src/api/facade'
import DocumentWorkbench from '/packages/office-studio/src/components/DocumentWorkbench.vue'
import { createBlankDocument } from '/packages/office-studio/src/domain/model'
import { createParagraph, createRun, normalizeDocument, runText } from '/packages/office-studio/src/features/document/model/document-model'
import { EMPTY_SECTION_HEADER_FOOTER, resolveSectionHeaderFooter, sectionStoryLinkedToPrevious } from '/packages/office-studio/src/features/document/services/section-header-footer'
import { DOCUMENT_STORY_KEYS, documentStoryDefaultStyle } from '/packages/office-studio/src/features/document/model/document-story'
import { replaceStoryRange, storyBuffer } from '/packages/office-studio/src/features/document/services/story-editor'
import { createTextMeasurer, layoutDocument } from '/packages/office-studio/src/features/document/layout/document-layout'
import { sceneOffset } from '/packages/office-studio/src/features/document/rendering/document-scene'
import { createOfficeMessages, OfficeDirectionKey, OfficeLocaleKey, OfficeTranslatorKey } from '/packages/office-studio/src/i18n/office-i18n'
const copy = value => JSON.parse(JSON.stringify(value))
const initial = createBlankDocument(); initial.id = 'section-layout-browser'; initial.name = '节布局编辑'
initial.page = { width: 440, height: 360, marginLeft: 40, marginRight: 40, marginTop: 40, marginBottom: 40 }
initial.sectionSetup = { columnCount: 2, columnGap: 24, columnSeparator: 'line' }
const first = createParagraph('列'.repeat(240), { fontSize: 12 }); first.id = 'body'
const boundary = { ...createParagraph(), id: 'section-two', type: 'sectionBreak', runs: [], sectionBreak: { type: 'evenPage', setup: { width: 620, height: 420, pageNumberStart: 3 } } }
const last = createParagraph('Second section', { fontSize: 12 }); last.id = 'after'
for (const block of [first, last]) Object.assign(block.style, { spaceBefore: 0, spaceAfter: 0, lineHeight: 20, spacingRule: 'exact', wordWrap: true })
initial.blocks = [first, boundary, last]
initial.sectionHeaderFooters = { first: { ...EMPTY_SECTION_HEADER_FOOTER, header: 'Odd', evenPageHeader: 'Even', differentOddEven: true } }
let studio = new OfficeStudio({ document: normalizeDocument(initial), currentUserId: 'local-user' })
studio.activateDocument(initial.id)
const documentApi = shallowRef(new FDocument(studio, initial.id)), model = ref(documentApi.value.getModel())
const dark = ref(false), variant = ref('classic'), locale = ref('zh-CN'), notices = [], commands = []
let reentrantReplacements = 0
let stopWatching
function observe() {
  stopWatching?.()
  stopWatching = watch(() => studio.editor.document.model, value => { model.value = value }, { flush: 'sync' })
  studio.events.addEvent('CommandExecuted', event => commands.push(copy(event)))
}
observe()
const bodySelection = () => ({ anchor: { blockId: 'body', offset: 0 }, focus: { blockId: 'body', offset: 0 } })
const layout = () => layoutDocument(model.value, createTextMeasurer())
function position(pageIndex, x, y) {
  const scroll = document.querySelector('.document-scroll'), bounds = scroll.getBoundingClientRect(), current = layout()
  scroll.scrollTop = Math.max(0, y - 120); scroll.dispatchEvent(new Event('scroll'))
  const offset = sceneOffset(current, { width: scroll.clientWidth, height: scroll.clientHeight, scrollLeft: scroll.scrollLeft, scrollTop: scroll.scrollTop, zoom: 100 })
  return { page: pageIndex, x: bounds.left + offset.x + x, y: bounds.top + offset.y + y }
}
window.documentSectionFixture = {
  snapshot: () => copy(model.value), input: () => copy(documentApi.value.getInputState()),
  section: (sectionId = 'section-two') => copy(resolveSectionHeaderFooter(model.value, sectionId)),
  links: (sectionId = 'section-two') => {
    const section = resolveSectionHeaderFooter(model.value, sectionId)
    return Object.fromEntries(DOCUMENT_STORY_KEYS.map(story => [story, sectionStoryLinkedToPrevious(section, story)]))
  },
  text: () => runText(model.value.blocks.find(block => block.id === 'body').runs),
  history: () => { const unit = studio.editor.history.units[documentApi.value.id]; return { undo: unit?.undo.length ?? 0, redo: unit?.redo.length ?? 0, dirty: studio.editor.history.isDirty(documentApi.value.id) } },
  geometry: () => layout().pages.map(page => ({ x: page.x, y: page.y, width: page.width, height: page.height, number: page.pageNumber, blank: page.isParityBlank, sectionId: page.sectionId })),
  storyGlyphs: (pageIndex, area = 'header') => layout().pageHeaderFooters[pageIndex][area + 'LayoutLines'].flatMap(line => line.glyphs)
    .map(glyph => ({ text: glyph.text, start: glyph.offset, end: glyph.end, color: glyph.run.color, bold: glyph.run.bold, italic: glyph.run.italic, underline: glyph.run.underline, href: glyph.run.href })),
  pointInColumn: column => {
    const line = layout().lines.find(line => line.target.blockId === 'body' && line.column === column && line.glyphs.length)
    if (!line) throw new Error('Requested column has no text')
    const glyph = line.glyphs[0]
    return { ...position(line.page, glyph.x + 1, line.y + line.height / 2), offset: glyph.offset }
  },
  pageHeaderPoint: pageIndex => { const page = layout().pages[pageIndex]; return position(pageIndex, page.x + page.width / 2, page.y + 12) },
  setup: (sectionId, patch) => documentApi.value.executeCommand({ id: 'doc.command.set-section-setup', sectionId, patch }, bodySelection()),
  permissions: rules => documentApi.value.executeCommand({ id: 'doc.command.set-document-permission-rules', rules }, bodySelection()),
  unlinkEvenStories: () => documentApi.value.executeCommand({ id: 'doc.command.set-section-header-footer', sectionId: 'section-two', patch: { storyLinkToPrevious: { evenPageHeader: false, evenPageFooter: false } } }, bodySelection()),
  selectStory: (story, start = 0, end = start) => documentApi.value.setInputSelection(bodySelection(), { sectionId: 'section-two', story, start, end }),
  seedMixedPageField: () => {
    documentApi.value.executeCommand({ id: 'doc.command.set-section-setup', sectionId: 'section-two', patch: { pageNumberStart: 19 } }, bodySelection())
    documentApi.value.executeCommand({ id: 'doc.command.set-section-header-footer', sectionId: 'section-two', patch: { storyLinkToPrevious: { evenPageHeader: false } } }, bodySelection())
    const source = documentApi.value.getModel(), section = resolveSectionHeaderFooter(source, 'section-two'), defaults = documentStoryDefaultStyle(source)
    const field = { field: 'PAGE', instruction: 'PAGE', cachedResult: '19', metadata: { retained: { source: 'mixed-cache' }, officeStoryFieldCache: { version: 1, runs: [
      createRun('1', { ...defaults.text, bold: true, color: '#aa0000' }), createRun('9', { ...defaults.text, italic: true, color: '#0000aa', href: 'https://example.org/field' }),
    ] } } }
    return documentApi.value.executeCommand({ id: 'doc.command.set-section-header-footer', sectionId: 'section-two', patch: replaceStoryRange(section, 'evenPageHeader', 0, storyBuffer(section, 'evenPageHeader').value.length, field, source) }, bodySelection())
  },
  replaceOnNextStorySelection: () => {
    const api = documentApi.value
    const subscription = api.onInputStateChanged(event => {
      const state = event.state
      if (!state?.valid || state.location.kind !== 'story' || state.location.range.story !== 'evenPageHeader' || state.location.range.start !== 0 || state.location.range.end !== 1) return
      subscription.dispose(); reentrantReplacements++
      const source = api.getModel(), section = resolveSectionHeaderFooter(source, 'section-two')
      api.executeCommand({ id: 'doc.command.set-section-header-footer', sectionId: 'section-two', patch: replaceStoryRange(section, 'evenPageHeader', 0, storyBuffer(section, 'evenPageHeader').value.length, 'Observer replacement', source) }, bodySelection())
    })
  },
  reentrantReplacements: () => reentrantReplacements,
  appearance: (isDark, ribbonVariant) => { dark.value = isDark; variant.value = ribbonVariant },
  saveReopen: () => {
    const saved = copy(documentApi.value.save()), next = new OfficeStudio({ document: saved, currentUserId: 'local-user' })
    stopWatching?.(); studio.dispose(); studio = next; studio.activateDocument(saved.id)
    documentApi.value = new FDocument(studio, saved.id); model.value = documentApi.value.getModel(); observe()
    return saved
  },
  notices: () => [...notices], commands: () => copy(commands),
}
const direction = computed(() => 'ltr'), messages = createOfficeMessages(), translate = (key, fallback) => messages[locale.value][key] ?? fallback
const app = createApp({ render: () => h(ConfigProvider, { direction: direction.value, theme: { algorithm: dark.value ? theme.darkAlgorithm : theme.defaultAlgorithm, token: { colorPrimary: '#3064c7' } } }, {
  default: () => h(DocumentWorkbench, { modelValue: model.value, documentApi: documentApi.value, canUndo: studio.editor.canUndo, canRedo: studio.editor.canRedo, ribbonVariant: variant.value,
    'onUpdate:modelValue': value => { model.value = value }, onUndo: () => documentApi.value.undo(), onRedo: () => documentApi.value.redo(), onNotify: message => notices.push(message) }),
}) })
app.provide(OfficeLocaleKey, locale); app.provide(OfficeDirectionKey, direction); app.provide(OfficeTranslatorKey, translate)
app.mount('#app')
window.addEventListener('pagehide', () => { stopWatching?.(); app.unmount(); studio.dispose() }, { once: true })
</script></body></html>`

let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'document-section-layout-fixture', configureServer(vite) {
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
  errors = []; page = await browser.newPage({ viewport: { width: 1440, height: 900 }, deviceScaleFactor: 1.25 })
  page.on('pageerror', error => errors.push(error.message))
  await page.goto(baseURL)
  await expect(page.getByRole('textbox', { name: '文档画布编辑区', exact: true })).toBeVisible()
}, { timeout: 60_000 })
afterEach(async () => { try { assert.deepEqual(errors, []) } finally { await page?.close() } })
const inputState = () => page.evaluate(() => window.documentSectionFixture.input())

test('Canvas input in the second column retains logical positions through edit, keyboard history and internal reopen', async () => {
  const before = await page.evaluate(() => window.documentSectionFixture.text())
  const point = await page.evaluate(() => window.documentSectionFixture.pointInColumn(1))
  await page.mouse.click(point.x, point.y)
  const bridge = page.getByLabel('文档文字输入', { exact: true })
  await expect(bridge).toBeFocused(); await expect(bridge).toHaveClass(/ant-input/)
  await expect.poll(async () => (await inputState())?.location.selection.focus.offset).toBe(point.offset)
  await page.keyboard.insertText('编')
  const edited = before.slice(0, point.offset) + '编' + before.slice(point.offset)
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.text())).toBe(edited)
  await page.keyboard.press('Control+z')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.text())).toBe(before)
  await page.keyboard.press('Control+Shift+z')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.text())).toBe(edited)
  await page.evaluate(() => window.documentSectionFixture.saveReopen())
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.text())).toBe(edited)
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.history()), { undo: 0, redo: 0, dirty: false })
})

test('mixed page geometry stays clickable after a section change and theme/Ribbon switching', async () => {
  await page.evaluate(() => { window.documentSectionFixture.setup('first', { columnCount: 1, marginLeft: 50, marginRight: 50 }); window.documentSectionFixture.appearance(true, 'grid') })
  await page.setViewportSize({ width: 980, height: 760 })
  const geometry = await page.evaluate(() => window.documentSectionFixture.geometry())
  assert.equal(geometry[0].width, 440)
  assert.equal(geometry.at(-1).width, 620)
  assert.equal(geometry[0].x - geometry.at(-1).x, 90)
  const point = await page.evaluate(() => window.documentSectionFixture.pointInColumn(0))
  await page.mouse.click(point.x, point.y)
  await expect(page.getByLabel('文档文字输入', { exact: true })).toBeFocused()
  await expect.poll(async () => (await inputState())?.location.selection.focus.offset).toBe(point.offset)
  await page.keyboard.insertText('改')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.text())).toMatch(/^改列/)
})

test('parity filler pages reject story activation and logical even pages open the even-header story', async () => {
  const geometry = await page.evaluate(() => window.documentSectionFixture.geometry()), blank = geometry.findIndex(page => page.blank)
  assert.ok(blank > 0)
  const point = await page.evaluate(index => window.documentSectionFixture.pageHeaderPoint(index), blank)
  await page.mouse.dblclick(point.x, point.y)
  await expect(page.getByLabel('编辑页眉', { exact: true })).toHaveCount(0)
  const actual = geometry.findIndex((entry, index) => index > blank && !entry.blank), header = await page.evaluate(index => window.documentSectionFixture.pageHeaderPoint(index), actual)
  assert.equal(geometry[actual].number, 4)
  await page.mouse.dblclick(header.x, header.y)
  await expect(page.getByLabel('编辑页眉', { exact: true })).toBeFocused()
  await expect(page.getByLabel('编辑页眉', { exact: true })).toHaveValue('Even')
})

test('section setup survives save/reopen with the same page numbers, filler pages and physical rectangles', async () => {
  await page.evaluate(() => window.documentSectionFixture.setup('section-two', { width: 500, height: 460, marginTop: 50, pageNumberStart: 9 }))
  const before = await page.evaluate(() => window.documentSectionFixture.geometry())
  const saved = await page.evaluate(() => window.documentSectionFixture.saveReopen())
  const after = await page.evaluate(() => window.documentSectionFixture.geometry())
  assert.deepEqual(after, before)
  assert.equal(saved.blocks.find(block => block.id === 'section-two').sectionBreak.setup.pageNumberStart, 9)
  assert.equal(after.at(-1).number, 10)
  assert.equal(after.at(-1).height, 460)
})

async function openSecondSectionHeader() {
  const geometry = await page.evaluate(() => window.documentSectionFixture.geometry())
  const index = geometry.findIndex(entry => entry.sectionId === 'section-two' && !entry.blank)
  assert.ok(index >= 0)
  const point = await page.evaluate(index => window.documentSectionFixture.pageHeaderPoint(index), index)
  await page.mouse.dblclick(point.x, point.y)
  await expect(page.getByLabel('编辑页眉', { exact: true })).toBeFocused()
  return index
}

test('header/footer navigation shows independent associations and Canvas edits survive undo and internal reopen', async () => {
  await openSecondSectionHeader()
  const header = page.getByLabel('编辑页眉', { exact: true }), link = page.getByRole('checkbox', { name: '链接到前一节', exact: true })
  await expect(header).toHaveValue('Even'); await expect(header).toHaveJSProperty('readOnly', true)
  const linkedHistory = await page.evaluate(() => window.documentSectionFixture.history())
  await page.keyboard.press('Control+b')
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.history()), linkedHistory)
  await expect(link).toBeChecked(); await link.uncheck()
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.links().evenPageHeader)).toBe(false)
  await expect(header).toHaveJSProperty('readOnly', false); await expect(header).toBeFocused()
  await page.keyboard.insertText('Own ')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.section().evenPageHeader)).toBe('Own Even')
  await page.getByRole('button', { name: '转到页脚', exact: true }).click()
  const footer = page.getByLabel('编辑页脚', { exact: true })
  await expect(footer).toBeFocused(); await expect(footer).toHaveJSProperty('readOnly', true); await expect(link).toBeChecked()
  await link.uncheck(); await expect(footer).toHaveJSProperty('readOnly', false); await expect(footer).toBeFocused()
  await page.keyboard.insertText('Own footer')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.section().evenPageFooter)).toBe('Own footer')
  const links = await page.evaluate(() => window.documentSectionFixture.links())
  assert.deepEqual(links, { header: true, footer: true, firstPageHeader: true, firstPageFooter: true, evenPageHeader: false, evenPageFooter: false })
  await page.getByRole('button', { name: '转到页眉', exact: true }).click()
  await expect(header).toHaveValue('Own Even'); await expect(link).not.toBeChecked()
  await link.check(); await expect(header).toHaveValue('Even'); await expect(header).toHaveJSProperty('readOnly', true)
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.section().evenPageFooter)).toBe('Own footer')
  await expect(header).toBeFocused(); await page.keyboard.press('Control+z')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.section().evenPageHeader)).toBe('Own Even')
  const saved = await page.evaluate(() => window.documentSectionFixture.saveReopen())
  assert.equal(saved.sectionHeaderFooters['section-two'].storyLinkToPrevious.evenPageHeader, false)
  assert.equal(saved.sectionHeaderFooters['section-two'].storyLinkToPrevious.evenPageFooter, false)
  await page.evaluate(() => window.documentSectionFixture.appearance(true, 'grid'))
  await openSecondSectionHeader(); await expect(header).toHaveValue('Own Even'); await expect(link).not.toBeChecked()
  await expect(header).toHaveClass(/ant-input/)
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.section().evenPageFooter)).toBe('Own footer')
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.history()), { undo: 0, redo: 0, dirty: false })
})

test('first-page options retarget the public input story and keep variant associations separate', async () => {
  await openSecondSectionHeader()
  const header = page.getByLabel('编辑页眉', { exact: true }), first = page.getByRole('checkbox', { name: '首页不同', exact: true })
  const link = page.getByRole('checkbox', { name: '链接到前一节', exact: true })
  await expect.poll(async () => (await inputState())?.location.range.story).toBe('evenPageHeader')
  await first.check()
  await expect.poll(async () => (await inputState())?.location.range.story).toBe('firstPageHeader')
  await expect(header).toHaveValue(''); await expect(link).toBeChecked()
  await link.uncheck(); await expect(header).toHaveJSProperty('readOnly', false); await expect(header).toBeFocused()
  await page.keyboard.insertText('Cover')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.section().firstPageHeader)).toBe('Cover')
  let links = await page.evaluate(() => window.documentSectionFixture.links())
  assert.equal(links.firstPageHeader, false); assert.equal(links.evenPageHeader, true); assert.equal(links.header, true)
  await first.uncheck()
  await expect.poll(async () => (await inputState())?.location.range.story).toBe('evenPageHeader')
  await expect(header).toHaveValue('Even'); await expect(header).toHaveJSProperty('readOnly', true); await expect(link).toBeChecked()
  await first.check()
  await expect.poll(async () => (await inputState())?.location.range.story).toBe('firstPageHeader')
  await expect(header).toHaveValue('Cover'); await expect(link).not.toBeChecked()
  await page.evaluate(() => window.documentSectionFixture.saveReopen())
  await openSecondSectionHeader(); await expect(header).toHaveValue('Cover')
  links = await page.evaluate(() => window.documentSectionFixture.links())
  assert.equal(links.firstPageHeader, false); assert.equal(links.evenPageHeader, true)
})

test('Canvas live field glyphs keep mixed cache formats through a story shortcut, undo/redo and internal reopen', async () => {
  await page.evaluate(() => window.documentSectionFixture.seedMixedPageField())
  const index = await openSecondSectionHeader(), header = page.getByLabel('编辑页眉', { exact: true })
  const geometry = await page.evaluate(() => window.documentSectionFixture.geometry())
  assert.equal(geometry[index].number, 20)
  await expect(header).toHaveValue('\uFFFC'); await expect(header).toHaveJSProperty('readOnly', false)
  const before = await page.evaluate(() => window.documentSectionFixture.section().pageFields)
  const glyphs = () => page.evaluate(index => window.documentSectionFixture.storyGlyphs(index), index)
  assert.deepEqual((await glyphs()).map(glyph => [glyph.text, glyph.start, glyph.end, glyph.color, glyph.bold, glyph.italic]), [
    ['2', 0, 1, '#aa0000', true, false], ['0', 0, 1, '#0000aa', false, true],
  ])
  await page.evaluate(() => window.documentSectionFixture.selectStory('evenPageHeader', 0, 1))
  await expect.poll(async () => (await inputState())?.location.range.end).toBe(1); await expect(header).toBeFocused()
  await page.keyboard.press('Control+b')
  await expect.poll(async () => (await glyphs()).map(glyph => glyph.bold)).toEqual([true, true])
  const formatted = await page.evaluate(() => window.documentSectionFixture.section().pageFields)
  assert.equal(formatted[0].cachedResult, '19'); assert.equal(formatted[0].instruction, 'PAGE')
  assert.equal(formatted[0].metadata.officeStoryFieldCache.runs.map(run => run.text).join(''), '19')
  assert.equal((await glyphs())[1].href, 'https://example.org/field')
  await expect(header).toHaveValue('\uFFFC')
  await page.keyboard.press('Control+z')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.section().pageFields)).toEqual(before)
  await page.evaluate(() => window.documentSectionFixture.selectStory('evenPageHeader', 0, 1)); await expect(header).toBeFocused()
  await page.keyboard.press('Control+y')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.section().pageFields)).toEqual(formatted)
  await page.evaluate(() => window.documentSectionFixture.saveReopen())
  await page.evaluate(() => window.documentSectionFixture.appearance(true, 'grid')); await openSecondSectionHeader()
  await expect(header).toHaveValue('\uFFFC'); await expect(header).toHaveClass(/ant-input/)
  assert.deepEqual((await glyphs()).map(glyph => [glyph.text, glyph.color, glyph.bold, glyph.italic]), [
    ['2', '#aa0000', true, false], ['0', '#0000aa', true, true],
  ])
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.section().pageFields), formatted)
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.history()), { undo: 0, redo: 0, dirty: false })
})

test('contextual PAGE and NUMPAGES insertion writes the current layout values into new field caches', async () => {
  const index = await openSecondSectionHeader(), geometry = await page.evaluate(() => window.documentSectionFixture.geometry())
  await page.getByRole('checkbox', { name: '链接到前一节', exact: true }).uncheck()
  const pageButton = page.getByRole('button', { name: '页码', exact: true })
  await expect(pageButton).toBeEnabled(); await pageButton.click()
  await page.getByRole('menuitem', { name: '当前页码', exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.section().pageFields.find(field => field.field === 'PAGE')?.cachedResult)).toBe(String(geometry[index].number))
  await pageButton.click(); await page.getByRole('menuitem', { name: '总页数', exact: true }).click()
  const fields = await page.evaluate(() => window.documentSectionFixture.section().pageFields.filter(field => field.story === 'evenPageHeader'))
  assert.deepEqual(fields.map(field => [field.field, field.instruction, field.cachedResult]), [
    ['PAGE', 'PAGE', String(geometry[index].number)], ['NUMPAGES', 'NUMPAGES', String(geometry.length)],
  ])
  await expect(page.getByLabel('编辑页眉', { exact: true })).toHaveValue('\uFFFC\uFFFCEven')
})

for (const [shortcut, property] of [['Control+i', 'italic'], ['Control+u', 'underline']]) test(`story ${shortcut} updates field formats and ignores composition key events`, async () => {
  await page.evaluate(() => window.documentSectionFixture.seedMixedPageField()); await openSecondSectionHeader()
  const header = page.getByLabel('编辑页眉', { exact: true })
  await page.evaluate(() => window.documentSectionFixture.selectStory('evenPageHeader', 0, 1)); await expect(header).toBeFocused()
  const before = await page.evaluate(() => window.documentSectionFixture.history())
  await page.evaluate(() => document.querySelector('textarea[aria-label="编辑页眉"]').dispatchEvent(new KeyboardEvent('keydown', { key: 'b', ctrlKey: true, isComposing: true, bubbles: true })))
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.history()), before)
  await page.keyboard.press(shortcut)
  await expect.poll(() => page.evaluate(property => window.documentSectionFixture.section().pageFields[0].metadata.officeStoryFieldCache.runs.map(run => run[property]), property)).toEqual([true, true])
  assert.equal((await page.evaluate(() => window.documentSectionFixture.section())).pageFields[0].cachedResult, '19')
  await expect(header).toHaveValue('\uFFFC')
  assert.equal((await page.evaluate(() => window.documentSectionFixture.history())).undo, before.undo + 1)
})

const deniedEvenHeader = [{ id: 'deny-current-header', scope: 'entity', targetId: 'section-two:evenPageHeader', action: 'edit', effect: 'deny', userIds: [] }]

test('read-only header retains section position controls and can navigate to an editable footer', async () => {
  await page.evaluate(rules => { window.documentSectionFixture.unlinkEvenStories(); window.documentSectionFixture.permissions(rules) }, deniedEvenHeader)
  await openSecondSectionHeader()
  const header = page.getByLabel('编辑页眉', { exact: true })
  await expect(header).toHaveJSProperty('readOnly', true)
  const beforeShortcut = await page.evaluate(() => window.documentSectionFixture.history())
  await page.keyboard.press('Control+b')
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.history()), beforeShortcut)
  await expect(page.getByRole('button', { name: '页码', exact: true })).toBeDisabled()
  await expect(page.getByRole('checkbox', { name: '首页不同', exact: true })).toBeDisabled()
  const margin = page.locator('[data-control-kind="position"]:visible').getByRole('spinbutton').first()
  await expect(margin).toBeEnabled(); await margin.fill('32'); await margin.press('Tab')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.snapshot().blocks.find(block => block.id === 'section-two').sectionBreak.setup.marginHeader)).toBe(32)
  const navigate = page.getByRole('button', { name: '转到页脚', exact: true })
  await expect(navigate).toBeEnabled(); await navigate.click()
  const footer = page.getByLabel('编辑页脚', { exact: true })
  await expect(footer).toHaveJSProperty('readOnly', false); await expect(footer).toBeFocused()
  await page.keyboard.insertText('Allowed footer')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.section().evenPageFooter)).toBe('Allowed footer')
  assert.equal((await page.evaluate(() => window.documentSectionFixture.section())).evenPageHeader, 'Even')
})

test('narrow classic More keeps read-only story navigation available with the shared dark theme', async () => {
  await page.evaluate(rules => { window.documentSectionFixture.unlinkEvenStories(); window.documentSectionFixture.permissions(rules); window.documentSectionFixture.appearance(true, 'classic') }, deniedEvenHeader)
  await page.setViewportSize({ width: 240, height: 900 }); await openSecondSectionHeader()
  const more = page.locator('.ribbon-more-button')
  await expect(more).toBeVisible(); await more.click()
  const overflow = page.locator('.ribbon-overflow-panel')
  await expect(overflow).toBeVisible()
  const navigate = overflow.getByRole('button', { name: '转到页脚', exact: true })
  await expect(navigate).toBeEnabled(); await navigate.click()
  const footer = page.getByLabel('编辑页脚', { exact: true })
  await expect(footer).toHaveJSProperty('readOnly', false); await expect(footer).toBeFocused()
  await page.keyboard.insertText('More footer')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.section().evenPageFooter)).toBe('More footer')
})

test('date and page popups close on story or model changes and reject queued actions from their old scope', async () => {
  await page.evaluate(() => window.documentSectionFixture.unlinkEvenStories()); await openSecondSectionHeader()
  const date = page.getByRole('button', { name: '日期和时间', exact: true }), dialog = page.getByRole('dialog', { name: '日期和时间', exact: true })
  const before = await page.evaluate(() => window.documentSectionFixture.section()), history = await page.evaluate(() => window.documentSectionFixture.history())
  await date.click(); await expect(dialog).toBeVisible()
  await page.evaluate(() => window.documentSectionFixture.selectStory('evenPageFooter'))
  await expect(dialog).not.toBeVisible(); await expect(page.getByLabel('编辑页脚', { exact: true })).toBeFocused()
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.section()), before)
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.history()), history)
  await date.click(); await expect(dialog).toBeVisible()
  await page.evaluate(() => {
    const oldOk = document.querySelector('.ant-modal [role="button"].ant-btn-primary, .ant-modal button.ant-btn-primary')
    if (!oldOk) throw new Error('Date modal has no OK button')
    window.documentSectionFixture.setup('section-two', { columnGap: 32 }); oldOk.click()
  })
  await expect(dialog).not.toBeVisible()
  let currentHistory = await page.evaluate(() => window.documentSectionFixture.history())
  assert.equal(currentHistory.undo, history.undo + 1)
  assert.deepEqual((await page.evaluate(() => window.documentSectionFixture.section())).pageFields, before.pageFields)
  await page.getByRole('button', { name: '页码', exact: true }).click()
  await expect(page.getByRole('menuitem', { name: '当前页码', exact: true })).toBeVisible()
  await page.evaluate(() => {
    const oldItem = Array.from(document.querySelectorAll('.ant-menu-item')).find(item => item.textContent.trim() === '当前页码')
    if (!oldItem) throw new Error('Page popup has no PAGE item')
    window.documentSectionFixture.setup('section-two', { columnGap: 36 }); oldItem.click()
  })
  await expect(page.getByRole('menuitem', { name: '当前页码', exact: true })).not.toBeVisible()
  currentHistory = await page.evaluate(() => window.documentSectionFixture.history())
  assert.equal(currentHistory.undo, history.undo + 2)
  assert.deepEqual((await page.evaluate(() => window.documentSectionFixture.section())).pageFields, before.pageFields)
})

test('a synchronous public selection observer replaces a story without letting an older date action overwrite it', async () => {
  await page.evaluate(() => window.documentSectionFixture.seedMixedPageField()); await openSecondSectionHeader()
  await page.getByRole('button', { name: '日期和时间', exact: true }).click()
  await expect(page.getByRole('dialog', { name: '日期和时间', exact: true })).toBeVisible()
  const history = await page.evaluate(() => window.documentSectionFixture.history())
  await page.evaluate(() => {
    window.documentSectionFixture.replaceOnNextStorySelection()
    const header = document.querySelector('textarea[aria-label="编辑页眉"]'), oldOk = document.querySelector('.ant-modal button.ant-btn-primary')
    if (!header || !oldOk) throw new Error('Reentrancy fixture is missing its story input or date button')
    header.setSelectionRange(0, 1); oldOk.click()
  })
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.reentrantReplacements())).toBe(1)
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.section().evenPageHeader)).toBe('Observer replacement')
  const section = await page.evaluate(() => window.documentSectionFixture.section()), currentHistory = await page.evaluate(() => window.documentSectionFixture.history())
  assert.deepEqual(section.pageFields?.filter(field => field.story === 'evenPageHeader') ?? [], [])
  assert.equal(currentHistory.undo, history.undo + 1)
})

test('an older date popup cannot rebind to a native story selection that has not yet emitted its select event', async () => {
  await page.evaluate(() => window.documentSectionFixture.seedMixedPageField()); await openSecondSectionHeader()
  await page.getByRole('button', { name: '日期和时间', exact: true }).click()
  const dialog = page.getByRole('dialog', { name: '日期和时间', exact: true })
  await expect(dialog).toBeVisible()
  const section = await page.evaluate(() => window.documentSectionFixture.section()), history = await page.evaluate(() => window.documentSectionFixture.history())
  await page.evaluate(() => {
    const header = document.querySelector('textarea[aria-label="编辑页眉"]'), oldOk = document.querySelector('.ant-modal button.ant-btn-primary')
    if (!header || !oldOk) throw new Error('Popup fixture is missing its story input or date button')
    header.setSelectionRange(0, 1); oldOk.click()
  })
  await expect(dialog).not.toBeVisible()
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.section()), section)
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.history()), history)
})

test('a story format shortcut rejects the old source after a synchronous public selection observer replaces it', async () => {
  await page.evaluate(() => window.documentSectionFixture.seedMixedPageField()); await openSecondSectionHeader()
  const history = await page.evaluate(() => window.documentSectionFixture.history())
  await page.evaluate(() => {
    window.documentSectionFixture.replaceOnNextStorySelection()
    const header = document.querySelector('textarea[aria-label="编辑页眉"]')
    if (!header) throw new Error('Shortcut fixture has no story input')
    header.setSelectionRange(0, 1)
    header.dispatchEvent(new KeyboardEvent('keydown', { key: 'b', ctrlKey: true, bubbles: true, cancelable: true }))
  })
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.reentrantReplacements())).toBe(1)
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.section().evenPageHeader)).toBe('Observer replacement')
  const section = await page.evaluate(() => window.documentSectionFixture.section()), currentHistory = await page.evaluate(() => window.documentSectionFixture.history())
  assert.equal(currentHistory.undo, history.undo + 1)
  assert.deepEqual(section.pageFields?.filter(field => field.story === 'evenPageHeader') ?? [], [])
})
