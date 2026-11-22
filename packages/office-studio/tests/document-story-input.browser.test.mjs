import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

// Real ant TextArea / Workbench drafts. DOM composition events below do not
// claim coverage of a real Windows Chinese IME. This turn ran no browser.
const fixturePath = '/__document-story-input-test__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>Docs story input contract</title><style>body{margin:0}#app{height:100vh;min-width:0}.fixture-away{height:100vh}</style></head><body><div id="app"></div><script type="module">
import { computed, createApp, h, KeepAlive, ref, shallowRef, watch } from 'vue'
import { ConfigProvider, theme } from 'antdv-next'
import { OfficeStudio } from '/packages/office-studio/src/api/office-studio'
import { FDocument } from '/packages/office-studio/src/api/facade'
import DocumentWorkbench from '/packages/office-studio/src/components/DocumentWorkbench.vue'
import { createBlankDocument } from '/packages/office-studio/src/domain/model'
import { createParagraph, createRun, defaultParagraphStyle, normalizeDocument, runText } from '/packages/office-studio/src/features/document/model/document-model'
import { documentStoryParagraphs } from '/packages/office-studio/src/features/document/model/document-story'
import { EMPTY_SECTION_HEADER_FOOTER } from '/packages/office-studio/src/features/document/services/section-header-footer'
import { createTextMeasurer, documentStoryCaretRect, documentStoryLayoutLines, caretRect, layoutDocument } from '/packages/office-studio/src/features/document/layout/document-layout'
import { sceneOffset } from '/packages/office-studio/src/features/document/rendering/document-scene'
import { createOfficeMessages, OfficeDirectionKey, OfficeLocaleKey, OfficeTranslatorKey } from '/packages/office-studio/src/i18n/office-i18n'
const copy = value => JSON.parse(JSON.stringify(value))
const initial = createBlankDocument(); initial.id = 'story-browser'; initial.name = '页眉页脚输入'
initial.defaultStyle.text = { ...initial.defaultStyle.text, fontFamily: 'Georgia', eastAsiaFontFamily: '宋体', fontSize: 15, color: 'auto' }
initial.defaultStyle.paragraph = { ...initial.defaultStyle.paragraph, align: 'right', spaceAfter: 12 }
initial.blocks = [createParagraph('body untouched')]; initial.blocks[0].id = 'body'
const text = 'A😀 word\\nTail'
const run = createRun(text, { bold: true, italic: true, fontSize: 22, color: '#FF0000', href: 'https://example.org/story' })
run.metadata = { source: { sourceFormat: { version: 1, propertiesXml: '<w:rPr><w:caps/><w:spacing w:val="24"/></w:rPr>', namespaces: { w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' } } } }
initial.sectionHeaderFooters = { first: { ...EMPTY_SECTION_HEADER_FOOTER, header: text, footer: 'Footer',
  pageFields: [{ story: 'header', field: 'PAGE', offset: 3, cachedResult: '12', textStyle: { ...initial.defaultStyle.text, fontSize: 18 } }],
  storyContent: { header: { runs: [run], paragraphs: documentStoryParagraphs(text, { ...defaultParagraphStyle(), align: 'center' }) } } } }
let studio = new OfficeStudio({ document: normalizeDocument(initial), currentUserId: 'local-user' })
studio.activateDocument(initial.id)
const documentApi = shallowRef(new FDocument(studio, initial.id)), model = ref(documentApi.value.getModel())
const visible = ref(true), dark = ref(false), variant = ref('classic'), locale = ref('zh-CN'), commands = [], notices = []
let cancelEdits = false, stopWatching
function observe() {
  stopWatching?.()
  stopWatching = watch(() => studio.editor.document.model, value => { model.value = value }, { flush: 'sync' })
  studio.events.addEvent('CommandExecuted', event => commands.push(copy(event)))
  studio.events.addEvent('BeforeCommandExecute', event => { if (cancelEdits && event.id === 'doc.command.set-section-header-footer') event.cancel = true })
}
observe()
function replace(next, clearHistory = false) {
  next = normalizeDocument(next)
  if (!studio.getDocumentModel(next.id)) studio.createDocument(next)
  studio.activateDocument(next.id); studio.editor.document.replace(next)
  documentApi.value = new FDocument(studio, next.id); model.value = documentApi.value.getModel()
  if (clearHistory) studio.editor.history.clear(next.id)
}
function geometry(area, offset = 0) {
  const scroll = document.querySelector('.document-scroll'), bounds = scroll.getBoundingClientRect(), layout = layoutDocument(model.value, createTextMeasurer())
  const transform = sceneOffset(layout, { width: scroll.clientWidth, height: scroll.clientHeight, scrollLeft: scroll.scrollLeft, scrollTop: scroll.scrollTop, zoom: 100 })
  const caret = area === 'body' ? caretRect(layout, { blockId: model.value.blocks[0].id, offset }) : documentStoryCaretRect(documentStoryLayoutLines(layout, { page: 0, area }), offset)
  if (!caret) throw new Error('Fixture did not find the logical caret')
  return { x: bounds.left + transform.x + caret.x, y: bounds.top + transform.y + caret.y + caret.height / 2 }
}
window.documentStoryFixture = {
  snapshot: () => copy(model.value),
  bodyText: () => runText(model.value.blocks[0].runs),
  history: () => { const unit = studio.editor.history.units[documentApi.value.id]; return { undo: unit?.undo.length ?? 0, redo: unit?.redo.length ?? 0, dirty: studio.editor.history.isDirty(documentApi.value.id) } },
  commands: () => copy(commands), notices: () => [...notices],
  appearance: (isDark, layout) => { dark.value = isDark; variant.value = layout },
  cancelEdits: value => { cancelEdits = value },
  point: geometry,
  scrollToFooter: () => { const scroll = document.querySelector('.document-scroll'); const layout = layoutDocument(model.value, createTextMeasurer()); scroll.scrollTop = Math.max(0, layout.pageHeaderFooters[0].footerBottom - scroll.clientHeight / 2); scroll.dispatchEvent(new Event('scroll')) },
  replaceSame: () => { const next = copy(model.value); next.name += ' 替换'; replace(next) },
  replaceId: () => { const next = copy(initial); next.id = 'replacement-story'; next.name = '新文档'; replace(next, true) },
  readonly: () => { const next = copy(model.value); next.permissionRules = [{ id: 'deny-story', scope: 'entity', targetId: 'first:header', action: 'edit', effect: 'deny', userIds: [] }]; replace(next) },
  show: value => { visible.value = value },
  saveReopen: () => {
    const saved = copy(documentApi.value.save()), next = new OfficeStudio({ document: saved, currentUserId: 'local-user' })
    stopWatching?.(); studio.dispose(); studio = next; studio.activateDocument(saved.id)
    documentApi.value = new FDocument(studio, saved.id); model.value = documentApi.value.getModel(); observe()
    return saved
  },
}
const direction = computed(() => 'ltr'), messages = createOfficeMessages(), translate = (key, fallback) => messages[locale.value][key] ?? fallback
const app = createApp({ render: () => h(ConfigProvider, { direction: direction.value, theme: { algorithm: dark.value ? theme.darkAlgorithm : theme.defaultAlgorithm, token: { colorPrimary: '#3064c7' } } }, {
  default: () => h(KeepAlive, { max: 2 }, { default: () => visible.value ? h(DocumentWorkbench, {
    key: 'workbench', modelValue: model.value, documentApi: documentApi.value, canUndo: studio.editor.canUndo, canRedo: studio.editor.canRedo, ribbonVariant: variant.value,
    'onUpdate:modelValue': value => { model.value = value },
    onUndo: () => documentApi.value.undo(), onRedo: () => documentApi.value.redo(), onNotify: message => notices.push(message),
  }) : h('div', { key: 'away', class: 'fixture-away' }, '工作台暂时隐藏') }),
}) })
app.provide(OfficeLocaleKey, locale); app.provide(OfficeDirectionKey, direction); app.provide(OfficeTranslatorKey, translate)
app.mount('#app')
window.addEventListener('pagehide', () => { stopWatching?.(); app.unmount(); studio.dispose() }, { once: true })
</script></body></html>`

let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'document-story-input-fixture',
    configureServer(vite) {
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
  errors = []
  page = await browser.newPage({ viewport: { width: 1440, height: 900 }, deviceScaleFactor: 1.25 })
  page.on('pageerror', error => errors.push(error.message))
  await page.goto(baseURL)
  await expect(page.getByRole('textbox', { name: '文档画布编辑区', exact: true })).toBeVisible()
}, { timeout: 60_000 })
afterEach(async () => { try { assert.deepEqual(errors, []) } finally { await page?.close() } })

const snapshot = () => page.evaluate(() => window.documentStoryFixture.snapshot())
const history = () => page.evaluate(() => window.documentStoryFixture.history())
const storyBridge = () => page.getByLabel('编辑页眉', { exact: true })
const bodyBridge = () => page.getByLabel('文档文字输入', { exact: true })
const buffer = 'A😀\uFFFC word\nTail'
async function homeItem(id) {
  const visible = page.locator('.ribbon-visible-groups [data-ribbon-item="' + id + '"]')
  if (await visible.count()) return visible
  const panel = page.locator('.ribbon-overflow-panel:visible')
  if (!await panel.isVisible()) await page.locator('.ribbon-more-button').click()
  return panel.locator('[data-ribbon-item="' + id + '"]')
}
async function openHeader() {
  await (await homeItem('doc-header-footer')).getByRole('button').click()
  await expect(storyBridge()).toBeFocused()
  await expect(storyBridge()).toHaveValue(buffer)
  await expect(storyBridge()).toHaveClass(/ant-input/)
}
async function nativeSelection(start, end = start) {
  await storyBridge().evaluate((element, value) => {
    element.setSelectionRange(value.start, value.end)
    element.dispatchEvent(new Event('select', { bubbles: true }))
  }, { start, end })
}
async function dragStory(start, end) {
  const from = await page.evaluate(offset => window.documentStoryFixture.point('header', offset), start)
  const to = await page.evaluate(offset => window.documentStoryFixture.point('header', offset), end)
  await page.mouse.move(from.x, from.y); await page.mouse.down()
  await page.mouse.move(to.x, to.y, { steps: 6 }); await page.mouse.up()
  await expect.poll(() => storyBridge().evaluate(element => [element.selectionStart, element.selectionEnd])).toEqual([Math.min(start, end), Math.max(start, end)])
}
async function startComposition() {
  await nativeSelection(5, 9)
  await storyBridge().evaluate(element => {
    window.staleStoryElement = element
    element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' }))
    element.dispatchEvent(new InputEvent('beforeinput', { bubbles: true, cancelable: true, inputType: 'insertCompositionText', data: '拼', isComposing: true }))
    element.value = element.value.slice(0, 5) + '拼' + element.value.slice(9)
    element.setSelectionRange(6, 6)
    element.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertCompositionText', data: '拼', isComposing: true }))
  })
}
async function staleStoryEvents() {
  await page.evaluate(() => {
    const element = window.staleStoryElement
    element.value = 'stale corrupt input'
    element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' }))
    element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '错误' }))
    element.dispatchEvent(new InputEvent('beforeinput', { bubbles: true, cancelable: true, inputType: 'historyUndo' }))
    element.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: '错误' }))
    element.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Enter' }))
  })
}

for (const variant of ['classic', 'grid']) test('Canvas drag, clear, focus and internal reopen with the real story TextArea in ' + variant, async () => {
  const before = await snapshot()
  await page.evaluate(layout => window.documentStoryFixture.appearance(layout === 'grid', layout), variant)
  await openHeader(); await dragStory(9, 5)
  await page.getByRole('tab', { name: '开始', exact: true }).click()
  await (await homeItem('doc-clear-format')).getByRole('button').click()
  await expect(storyBridge()).toBeFocused()
  await expect.poll(() => storyBridge().evaluate(element => [element.selectionStart, element.selectionEnd])).toEqual([5, 9])
  const cleared = await snapshot()
  assert.deepEqual(cleared.blocks, before.blocks)
  assert.deepEqual(cleared.sectionHeaderFooters.first.pageFields, before.sectionHeaderFooters.first.pageFields)
  assert.deepEqual(cleared.sectionHeaderFooters.first.storyContent.header.runs.map(run => [run.text, run.bold, run.fontSize]), [['A😀 ', true, 22], ['word', false, 9], ['\nTail', true, 22]])
  assert.deepEqual(await history(), { undo: 1, redo: 0, dirty: true })
  await storyBridge().press('Control+z'); await expect.poll(() => snapshot().then(value => value.sectionHeaderFooters)).toEqual(before.sectionHeaderFooters)
  await storyBridge().press('Control+y'); await expect.poll(() => snapshot().then(value => value.sectionHeaderFooters)).toEqual(cleared.sectionHeaderFooters)
  const saved = await page.evaluate(() => window.documentStoryFixture.saveReopen())
  assert.deepEqual((await snapshot()).sectionHeaderFooters, saved.sectionHeaderFooters)
  assert.deepEqual(await history(), { undo: 0, redo: 0, dirty: false })
  await expect(storyBridge()).toHaveValue(buffer)
  assert.deepEqual(await page.evaluate(() => window.documentStoryFixture.commands().filter(event => event.id === 'doc.command.clear-formatting').map(event => [event.unitId, event.kind])), [['story-browser', 'document']])
})

test('a rejected native replacement restores buffer and selection and permits an identical retry', async () => {
  await openHeader(); await nativeSelection(5, 9)
  const before = await snapshot()
  await page.evaluate(() => window.documentStoryFixture.cancelEdits(true))
  await page.keyboard.insertText('新词')
  await expect(storyBridge()).toHaveValue(buffer)
  await expect.poll(() => storyBridge().evaluate(element => [element.selectionStart, element.selectionEnd])).toEqual([5, 9])
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), { undo: 0, redo: 0, dirty: false })
  await page.evaluate(() => window.documentStoryFixture.cancelEdits(false))
  await page.keyboard.insertText('新词')
  await expect(storyBridge()).toHaveValue('A😀\uFFFC 新词\nTail')
  assert.equal((await snapshot()).sectionHeaderFooters.first.header, 'A😀 新词\nTail')
  assert.deepEqual(await history(), { undo: 1, redo: 0, dirty: true })
  assert.ok((await page.evaluate(() => window.documentStoryFixture.notices())).some(message => message.includes('canceled')))
})

test('synthetic composition on the real TextArea previews without history and commits once at compositionend', async () => {
  await openHeader()
  const before = await snapshot(); await startComposition()
  await expect(storyBridge()).toHaveValue('A😀\uFFFC 拼\nTail')
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), { undo: 0, redo: 0, dirty: false })
  await storyBridge().evaluate(element => {
    element.value = 'A😀\uFFFC 中文\nTail'; element.setSelectionRange(7, 7)
    element.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertCompositionText', data: '中文', isComposing: true }))
    element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '中文' }))
    element.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: '中文', isComposing: false }))
  })
  await expect(storyBridge()).toHaveValue('A😀\uFFFC 中文\nTail')
  assert.equal((await snapshot()).sectionHeaderFooters.first.header, 'A😀 中文\nTail')
  assert.deepEqual(await history(), { undo: 1, redo: 0, dirty: true })
  assert.equal((await page.evaluate(() => window.documentStoryFixture.commands().filter(event => event.id === 'doc.command.set-section-header-footer'))).length, 1)
  await storyBridge().press('Control+z'); await expect(storyBridge()).toHaveValue(buffer)
})

for (const reason of ['same-model', 'document-switch', 'readonly', 'story-switch', 'deactivate']) test('composition cancellation ignores old TextArea events after ' + reason, async () => {
  await openHeader(); await startComposition()
  if (reason === 'same-model') await page.evaluate(() => window.documentStoryFixture.replaceSame())
  else if (reason === 'document-switch') await page.evaluate(() => window.documentStoryFixture.replaceId())
  else if (reason === 'readonly') await page.evaluate(() => window.documentStoryFixture.readonly())
  else if (reason === 'deactivate') {
    await page.evaluate(() => window.documentStoryFixture.show(false))
    await expect(page.locator('.fixture-away')).toBeVisible()
  } else {
    await page.evaluate(() => window.documentStoryFixture.scrollToFooter())
    const footer = await page.evaluate(() => window.documentStoryFixture.point('footer', 0))
    await page.mouse.dblclick(footer.x, footer.y)
    await expect(page.getByLabel('编辑页脚', { exact: true })).toBeFocused()
  }
  const before = await snapshot(), beforeHistory = await history()
  await expect.poll(() => page.evaluate(() => window.staleStoryElement !== document.querySelector('.document-story-bridge textarea') && window.staleStoryElement !== document.querySelector('textarea.document-story-bridge'))).toBe(true)
  await staleStoryEvents()
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), beforeHistory)
  assert.equal(before.sectionHeaderFooters.first.header, 'A😀 word\nTail')
  if (reason === 'readonly') await expect(storyBridge()).toHaveAttribute('readonly', '')
  if (reason === 'deactivate') {
    await page.evaluate(() => window.documentStoryFixture.show(true))
    await expect(page.getByRole('textbox', { name: '文档画布编辑区', exact: true })).toBeVisible()
    await storyBridge().focus(); await expect(storyBridge()).toHaveValue(buffer)
  }
})

test('body composition cancellation rejects stale native beforeinput, input, keyboard and clipboard events', async () => {
  await bodyBridge().focus()
  const before = await snapshot()
  await bodyBridge().evaluate(element => {
    window.staleBodyElement = element
    element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' }))
    element.dispatchEvent(new CompositionEvent('compositionupdate', { bubbles: true, data: '拼' }))
  })
  assert.deepEqual(await snapshot(), before)
  await page.evaluate(() => window.documentStoryFixture.replaceSame())
  await expect.poll(() => page.evaluate(() => !window.staleBodyElement.isConnected)).toBe(true)
  const replaced = await snapshot()
  await page.evaluate(() => {
    const element = window.staleBodyElement
    element.value = '错误'
    element.dispatchEvent(new InputEvent('beforeinput', { bubbles: true, cancelable: true, inputType: 'insertText', data: '错误' }))
    element.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: '错误' }))
    element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '错误' }))
    element.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Enter' }))
    const transfer = new DataTransfer(); transfer.setData('text/plain', '错误')
    element.dispatchEvent(new ClipboardEvent('paste', { bubbles: true, cancelable: true, clipboardData: transfer }))
  })
  assert.deepEqual(await snapshot(), replaced); assert.deepEqual(await history(), { undo: 0, redo: 0, dirty: false })
  assert.equal(await page.evaluate(() => window.documentStoryFixture.bodyText()), 'body untouched')
})
