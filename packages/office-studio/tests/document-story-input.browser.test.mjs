import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

// Real ant TextArea / Workbench drafts. DOM composition events below do not
// claim coverage of a real Windows Chinese IME. Draft date: 2026-11-22
// Asia/Shanghai. This turn ran no browser or other validation.
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
import { projectDocumentStoryBlocks } from '/packages/office-studio/src/features/document/model/document-story-blocks'
import { presetNumberingDefinition } from '/packages/office-studio/src/features/document/model/document-lists'
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
const visible = ref(true), dark = ref(false), variant = ref('classic'), locale = ref('zh-CN'), commands = [], notices = [], historyRequests = []
let cancelEdits = false, cancelUndo = false, navigateAfterStoryInput = false, closeAfterStoryInput = false, stopWatching
const closedDocumentIds = []
function bodySelection(start = 0, end = start) {
  const blockId = model.value.blocks[0].id
  return { anchor: { blockId, offset: start }, focus: { blockId, offset: end } }
}
function formatBody(patch, clearFields = []) {
  const state = documentApi.value.getInputState()
  const selection = state?.location.kind === 'body' ? state.location.selection : bodySelection()
  return documentApi.value.executeCommand({ id: 'doc.command.format-text', patch, clearFields }, selection)
}
function formatStory(patch, clearFields = []) {
  const state = documentApi.value.getInputState()
  const story = state?.location.kind === 'story' ? state.location.range : { sectionId: 'first', story: 'header', start: 0, end: 0 }
  return documentApi.value.executeCommand({ id: 'doc.command.format-text', story, patch, clearFields }, bodySelection())
}
function observe() {
  stopWatching?.()
  stopWatching = watch(() => studio.editor.document.model, value => { model.value = value }, { flush: 'sync' })
  studio.events.addEvent('CommandExecuted', event => commands.push(copy(event)))
  studio.events.addEvent('BeforeCommandExecute', event => { if (cancelEdits && event.id === 'doc.command.set-section-header-footer') event.cancel = true })
  studio.events.addEvent('BeforeUndo', event => { if (cancelUndo) event.cancel = true })
  studio.events.addEvent('DocumentInputStateChanged', event => {
    if (event.unitId !== documentApi.value.id || event.state?.valid !== false || event.state.location.kind !== 'story') return
    if (closeAfterStoryInput) {
      closeAfterStoryInput = false; closedDocumentIds.push(event.unitId)
      studio.disposeUnit(event.unitId)
      const next = studio.editor.document.model
      documentApi.value = new FDocument(studio, next.id); model.value = documentApi.value.getModel()
    } else if (navigateAfterStoryInput) {
      navigateAfterStoryInput = false
      documentApi.value.setInputSelection(bodySelection(2))
      formatBody({ fontSize: 31 })
    }
  })
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
  commands: () => copy(commands), notices: () => [...notices], historyRequests: () => copy(historyRequests),
  inputState: () => {
    const state = documentApi.value.getInputState()
    return state && { ...copy(state),
      clearedColor: Object.hasOwn(state.marks, 'color') && state.marks.color === undefined,
      clearedHighlight: Object.hasOwn(state.marks, 'highlight') && state.marks.highlight === undefined }
  },
  externalBodySelection: (start, end = start) => documentApi.value.setInputSelection(bodySelection(start, end)),
  externalStorySelection: (start, end = start, story = 'header') => documentApi.value.setInputSelection(bodySelection(), { sectionId: 'first', story, start, end }),
  formatBody, formatStory,
  rebindSame: () => { documentApi.value = new FDocument(studio, documentApi.value.id) },
  navigateOnStoryCommit: () => { navigateAfterStoryInput = true },
  closeOnStoryCommit: () => { closeAfterStoryInput = true },
  closedDocumentIds: () => [...closedDocumentIds],
  mixedField: () => {
    const next = copy(model.value), section = next.sectionHeaderFooters.first
    const primary = { ...next.defaultStyle.text, fontFamily: 'Georgia', eastAsiaFontFamily: '宋体', fontSize: 18, bold: true, italic: true, color: '#AA0000' }
    const source = copy(section.storyContent.header.runs[0].metadata.source)
    section.pageFields[0] = { ...section.pageFields[0], field: 'DATE', instruction: 'DATE', cachedResult: '中Z', textStyle: primary,
      metadata: { officeStoryStyleBaseline: { ...primary }, officeStoryFieldCache: { version: 1, runs: [
        createRun('中', { ...primary, metadata: { source: copy(source) } }),
        createRun('Z', { fontFamily: 'Aptos', fontSize: 12, color: '#0000AA', href: 'https://example.org/cache', metadata: { source: copy(source) } }),
        createRun('', { fontFamily: 'Verdana', fontSize: 28, color: '#00AA00', metadata: { entityId: 'empty-cache' } }),
      ] } } }
    replace(next, true)
  },
  simpleStory: () => {
    const next = copy(model.value), section = next.sectionHeaderFooters.first
    section.header = 'aaa'; section.pageFields = []
    section.storyContent.header = { runs: [
      createRun('a', { fontFamily: 'Georgia', fontSize: 24, bold: true }),
      createRun('aa', { fontFamily: 'Aptos', fontSize: 12, italic: true }),
    ], paragraphs: documentStoryParagraphs('aaa', next.defaultStyle.paragraph) }
    replace(next, true)
  },
  namedStyles: () => {
    const next = copy(model.value), paragraphs = next.sectionHeaderFooters.first.storyContent.header.paragraphs
    paragraphs[0].metadata = { namedStyleType: 'title' }
    paragraphs[1].metadata = { namedStyleType: 'heading4' }
    replace(next, true)
  },
  richStories: () => {
    const next = copy(model.value), section = next.sectionHeaderFooters.first
    section.header = 'Rich header'; section.footer = 'Rich footer'; section.pageFields = []
    section.storyContent = {
      header: projectDocumentStoryBlocks({ blocks: [createParagraph(section.header)], runs: [], paragraphs: [] }, 'display').content,
      footer: projectDocumentStoryBlocks({ blocks: [createParagraph(section.footer)], runs: [], paragraphs: [] }, 'display').content,
    }
    replace(next, true)
  },
  numberingStory: (kind, bodyNumbered, headerNumbered) => {
    const next = copy(initial), section = next.sectionHeaderFooters.first
    if (bodyNumbered) { next.blocks[0].type = 'numbered'; next.blocks[0].list = { id: 'body-numbering', preset: 'ORDER_LIST' } }
    const blocks = ['Previous', 'Gap', 'Current'].map(value => createParagraph(value, next.defaultStyle.text))
    if (headerNumbered) for (const index of [0, 2]) {
      const id = 'header-numbering-' + index, definition = presetNumberingDefinition('ORDER_LIST')
      blocks[index].type = 'numbered'; blocks[index].list = { id, preset: 'ORDER_LIST' }
      blocks[index].metadata = { numberingListId: id, numberingDefinition: { ...definition, instance: { numId: id } } }
    }
    section.header = blocks.map(block => runText(block.runs)).join('\\n'); section.pageFields = []
    const content = projectDocumentStoryBlocks({ blocks, runs: [], paragraphs: [] }, 'display').content
    if (kind === 'flat') delete content.blocks
    section.storyContent = { ...section.storyContent, header: content }
    replace(next, true)
  },
  paragraphStory: kind => {
    const next = copy(initial), section = next.sectionHeaderFooters.first
    next.blocks[0].style.indentFirstLine = 100
    const styles = [{ indentFirstLine: 0 }, { indentFirstLine: 10 }, { indentFirstLine: 20, spacingRule: 'exact', lineHeight: 24, spaceBefore: 6, spaceAfter: 18 }]
    const blocks = ['Previous', 'Gap', 'Current'].map((text, index) => {
      const block = createParagraph(text, next.defaultStyle.text)
      block.style = { ...block.style, ...styles[index] }
      return block
    })
    section.header = blocks.map(block => runText(block.runs)).join('\\n'); section.pageFields = []
    const content = projectDocumentStoryBlocks({ blocks, runs: [], paragraphs: [] }, 'display').content
    if (kind === 'flat') delete content.blocks
    section.storyContent = { ...section.storyContent, header: content }
    replace(next, true)
  },
  appearance: (isDark, layout) => { dark.value = isDark; variant.value = layout },
  cancelEdits: value => { cancelEdits = value },
  cancelUndo: value => { cancelUndo = value },
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
    onUndo: unitId => { historyRequests.push({ direction: 'undo', unitId }); try { documentApi.value.undo() } catch (error) { notices.push(error instanceof Error ? error.message : String(error)) } },
    onRedo: unitId => { historyRequests.push({ direction: 'redo', unitId }); try { documentApi.value.redo() } catch (error) { notices.push(error instanceof Error ? error.message : String(error)) } },
    onNotify: message => notices.push(message),
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
const inputState = () => page.evaluate(() => window.documentStoryFixture.inputState())
const storyBridge = () => page.getByLabel('编辑页眉', { exact: true })
const bodyBridge = () => page.getByLabel('文档文字输入', { exact: true })
const buffer = 'A😀\uFFFC word\nTail'
async function homeItem(id) {
  const selector = '[data-ribbon-item="' + id + '"]'
  const visible = page.locator('.ribbon-visible-groups ' + selector + ', .ribbon-grid-group:not(.is-hidden) ' + selector)
  if (await visible.count()) return visible
  const panel = page.locator('.ribbon-overflow-panel:visible')
  if (!await panel.isVisible()) await page.locator('.ribbon-more-button:visible, .ribbon-grid-more-button:visible').click()
  return panel.locator(selector)
}
async function openHeader(expected = buffer) {
  await (await homeItem('doc-header-footer')).getByRole('button').click()
  await expect(storyBridge()).toBeFocused()
  await expect(storyBridge()).toHaveValue(expected)
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
async function startComposition(start = 5, end = 9, text = '拼') {
  await nativeSelection(start, end)
  await storyBridge().evaluate((element, value) => {
    window.staleStoryElement = element
    element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' }))
    element.dispatchEvent(new InputEvent('beforeinput', { bubbles: true, cancelable: true, inputType: 'insertCompositionText', data: value.text, isComposing: true }))
    element.value = element.value.slice(0, value.start) + value.text + element.value.slice(value.end)
    const caret = value.start + value.text.length
    element.setSelectionRange(caret, caret)
    element.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertCompositionText', data: value.text, isComposing: true }))
  }, { start, end, text })
}
async function characterControl(id) {
  await page.getByRole('tab', { name: '开始', exact: true }).click()
  return (await homeItem(id)).getByRole('combobox')
}
async function setCharacterControl(id, value) {
  const control = await characterControl(id)
  await control.fill(String(value)); await control.press('Enter')
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
    element.dispatchEvent(new Event('select', { bubbles: true }))
    element.dispatchEvent(new FocusEvent('focus', { bubbles: false }))
    element.dispatchEvent(new FocusEvent('blur', { bubbles: false }))
  })
}

for (const variant of ['classic', 'grid']) test('rich header and footer navigation switches the real block host without editing content and restores the body range in ' + variant, async () => {
  await page.evaluate(layout => {
    window.documentStoryFixture.richStories()
    window.documentStoryFixture.appearance(layout === 'grid', layout)
    window.documentStoryFixture.externalBodySelection(2, 6)
  }, variant)
  const before = await snapshot(), previousHistory = await history()
  await (await homeItem('doc-header-footer')).getByRole('button').click()
  await expect.poll(async () => (await inputState()).location.host).toEqual({ kind: 'story', sectionId: 'first', story: 'header' })
  await expect(bodyBridge()).toBeFocused()
  const footer = page.getByRole('button', { name: '转到页脚', exact: true })
  await expect(footer).toBeEnabled()
  await footer.click()
  await expect.poll(async () => (await inputState()).location.host).toEqual({ kind: 'story', sectionId: 'first', story: 'footer' })
  await expect(footer).toBeDisabled()
  await expect(bodyBridge()).toBeFocused()
  const header = page.getByRole('button', { name: '转到页眉', exact: true })
  await expect(header).toBeEnabled()
  await header.click()
  await expect.poll(async () => (await inputState()).location.host).toEqual({ kind: 'story', sectionId: 'first', story: 'header' })
  await expect(header).toBeDisabled()
  assert.deepEqual(await snapshot(), before)
  assert.deepEqual(await history(), previousHistory)
  await page.getByRole('button', { name: '关闭页眉和页脚', exact: true }).click()
  await expect.poll(async () => (await inputState()).location).toEqual({
    kind: 'body', selection: { anchor: { blockId: 'body', offset: 2 }, focus: { blockId: 'body', offset: 6 } },
  })
  await expect(bodyBridge()).toBeFocused()
  const saved = await page.evaluate(() => window.documentStoryFixture.saveReopen())
  assert.deepEqual(saved.sectionHeaderFooters, before.sectionHeaderFooters)
  await (await homeItem('doc-header-footer')).getByRole('button').click()
  await expect.poll(async () => (await inputState()).location.host).toEqual({ kind: 'story', sectionId: 'first', story: 'header' })
  await page.getByRole('button', { name: '转到页脚', exact: true }).click()
  await expect.poll(async () => (await inputState()).location.host).toEqual({ kind: 'story', sectionId: 'first', story: 'footer' })
  assert.deepEqual((await snapshot()).sectionHeaderFooters, saved.sectionHeaderFooters)
  assert.deepEqual(await history(), { undo: 0, redo: 0, dirty: false })
})

test.each([['flat', 'classic', false], ['flat', 'grid', true], ['rich', 'classic', false], ['rich', 'grid', true]])('header composition disables Home character formats and clear in %s, %s, dark %s', async (kind, variant, dark) => {
  await page.evaluate(({ kind, variant, dark }) => {
    if (kind === 'rich') window.documentStoryFixture.richStories()
    window.documentStoryFixture.appearance(dark, variant)
  }, { kind, variant, dark })
  await (await homeItem('doc-header-footer')).getByRole('button').click()
  const input = kind === 'flat' ? storyBridge() : bodyBridge()
  await expect(input).toBeFocused()
  await page.getByRole('tab', { name: '开始', exact: true }).click()
  await input.focus()
  if (kind === 'rich') await input.press('Control+A')
  const before = await snapshot(), previousCommands = await page.evaluate(() => window.documentStoryFixture.commands())
  if (kind === 'flat') await startComposition()
  if (kind === 'rich') await input.evaluate(element => {
    element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' }))
    element.value = '拼'
    element.dispatchEvent(new CompositionEvent('compositionupdate', { bubbles: true, data: '拼' }))
    element.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertCompositionText', data: '拼', isComposing: true }))
  })
  await input.evaluate(element => {
    element.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: ',', code: 'Comma', ctrlKey: true }))
    element.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: '.', code: 'Period', metaKey: true }))
    for (const [code, key] of [['KeyX', 'X'], ['KeyL', 'L'], ['KeyE', 'E'], ['KeyR', 'R'], ['KeyJ', 'J'], ['Digit7', '&'], ['Digit8', '*']]) {
      for (const modifier of ['ctrlKey', 'metaKey']) element.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key, code, shiftKey: true, [modifier]: true }))
    }
  })
  for (const id of ['doc-bold', 'doc-italic', 'doc-underline', 'doc-strike', 'doc-subscript', 'doc-superscript', 'doc-clear-format']) {
    await expect((await homeItem(id)).getByRole('button')).toBeDisabled()
  }
  for (const id of ['doc-font-family', 'doc-font-size']) await expect((await homeItem(id)).getByRole('combobox')).toBeDisabled()
  for (const id of ['doc-align', 'doc-numbered', 'doc-bullet']) {
    for (const button of await (await homeItem(id)).getByRole('button').all()) await expect(button).toBeDisabled()
  }
  for (const id of ['doc-indent-less', 'doc-indent-more', 'doc-paragraph-spacing']) await expect((await homeItem(id)).getByRole('button')).toBeDisabled()
  await expect((await homeItem('doc-line-height')).getByRole('combobox')).toBeDisabled()
  for (const id of ['doc-text-color', 'doc-highlight']) for (const selector of ['.document-color-main', '.document-color-arrow button']) {
    await expect((await homeItem(id)).locator(selector)).toBeDisabled()
  }
  await page.locator('[data-testid="workbench-ribbon"]').evaluate(element => {
    element.__vueParentComponent.emit('execute', 'doc.command.format-text', { patch: { bold: true } })
    element.__vueParentComponent.emit('execute', 'doc.command.format-paragraph', { patch: { align: 'center' } })
    element.__vueParentComponent.emit('execute', 'doc.command.set-block-type', { type: 'numbered' })
    element.__vueParentComponent.emit('execute', 'doc.command.change-list-type', { listType: 'BULLET_LIST' })
    element.__vueParentComponent.emit('execute', 'doc.command.restart-list-numbering', { start: 7 })
    element.__vueParentComponent.emit('execute', 'doc.command.continue-list-numbering')
    element.__vueParentComponent.emit('execute', 'doc.command.change-paragraph-indent', { direction: 1 })
    element.__vueParentComponent.emit('execute', 'doc.command.change-paragraph-indent', { direction: -1 })
    element.__vueParentComponent.emit('execute', 'doc.command.format-paragraph', { patch: { spacingRule: 'auto', lineHeight: 2 } })
    element.__vueParentComponent.emit('execute', 'doc.command.format-paragraph', { patch: { spaceBefore: 18, spaceAfter: 24 } })
    for (const id of ['doc.command.horizontal-line', 'doc.command.insert-divider', 'doc.command.insert-horizontal-line-bellow']) {
      element.__vueParentComponent.emit('execute', id)
    }
    element.__vueParentComponent.emit('execute', 'ui.command.clear-formatting')
  })
  assert.deepEqual(await snapshot(), before)
  assert.deepEqual(await history(), { undo: 0, redo: 0, dirty: false })
  assert.deepEqual(await page.evaluate(() => window.documentStoryFixture.commands()), previousCommands)
  await expect(input).toBeFocused()
  await input.evaluate((element, rich) => {
    element.value = rich ? '中文' : 'A😀\uFFFC 中文\nTail'
    element.setSelectionRange(rich ? 2 : 7, rich ? 2 : 7)
    element.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertCompositionText', data: '中文', isComposing: true }))
    element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '中文' }))
    element.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: '中文' }))
  }, kind === 'rich')
  await expect.poll(async () => (await snapshot()).sectionHeaderFooters.first.header).toBe(kind === 'rich' ? '中文' : 'A😀 中文\nTail')
  assert.deepEqual((await snapshot()).blocks, before.blocks)
  assert.deepEqual(await history(), { undo: 1, redo: 0, dirty: true })
  for (const id of ['doc-bold', 'doc-italic', 'doc-underline', 'doc-strike', 'doc-subscript', 'doc-superscript', 'doc-clear-format']) {
    await expect((await homeItem(id)).getByRole('button')).toBeEnabled()
  }
  await expect(input).toBeFocused()
  const state = await inputState()
  assert.equal(state.location.kind, kind === 'flat' ? 'story' : 'body')
  if (kind === 'rich') assert.deepEqual(state.location.host, { kind: 'story', sectionId: 'first', story: 'header' })
  if (kind === 'flat') await nativeSelection(5, 7)
  if (kind === 'rich') await input.press('Control+A')
  await (await homeItem('doc-subscript')).getByRole('button').click()
  await expect.poll(async () => (await snapshot()).sectionHeaderFooters.first.storyContent.header.runs.some(run => run.text.includes('中文') && run.verticalAlign === 'subscript')).toBe(true)
  await expect(input).toBeFocused()
  await (await homeItem('doc-clear-format')).getByRole('button').click()
  await expect.poll(async () => (await snapshot()).sectionHeaderFooters.first.storyContent.header.runs.some(run => run.text.includes('中文') && run.verticalAlign === undefined && run.italic === false)).toBe(true)
  assert.deepEqual((await snapshot()).blocks, before.blocks)
  assert.deepEqual(await history(), { undo: 3, redo: 0, dirty: true })
  await expect(input).toBeFocused()
})

test.each([['flat', 'classic', false], ['flat', 'grid', true], ['rich', 'classic', false], ['rich', 'grid', true]])('numbering capabilities use the active header and keep restart, continuation and delayed composition callbacks in %s, %s, dark %s', async (kind, variant, dark) => {
  await page.evaluate(({ kind, variant, dark }) => {
    window.documentStoryFixture.numberingStory(kind, true, false)
    window.documentStoryFixture.appearance(dark, variant)
  }, { kind, variant, dark })
  await (await homeItem('doc-header-footer')).getByRole('button').click()
  const initialInput = kind === 'flat' ? storyBridge() : bodyBridge()
  await expect(initialInput).toBeFocused()
  await page.getByRole('tab', { name: '开始', exact: true }).click()
  await initialInput.focus()
  if (kind === 'flat') await nativeSelection(13)
  if (kind === 'rich') await initialInput.press('Control+End')
  const plainHeader = await snapshot()
  await (await homeItem('doc-numbered')).locator('.ribbon-split-arrow').click()
  const unavailable = page.locator('.doc-list-picker:visible')
  await expect(unavailable.getByRole('spinbutton', { name: '起始编号', exact: true })).toBeDisabled()
  await expect(unavailable.getByRole('button', { name: '重新编号', exact: true })).toBeDisabled()
  await expect(unavailable.getByRole('button', { name: '继续前一编号', exact: true })).toBeDisabled()
  assert.deepEqual(await snapshot(), plainHeader)
  assert.deepEqual(await history(), { undo: 0, redo: 0, dirty: false })

  await page.goto(baseURL)
  await page.evaluate(({ kind, variant, dark }) => {
    window.documentStoryFixture.numberingStory(kind, false, true)
    window.documentStoryFixture.appearance(dark, variant)
  }, { kind, variant, dark })
  await (await homeItem('doc-header-footer')).getByRole('button').click()
  await expect(initialInput).toBeFocused()
  await page.getByRole('tab', { name: '开始', exact: true }).click()
  await initialInput.focus()
  if (kind === 'flat') await nativeSelection(13)
  if (kind === 'rich') await initialInput.press('Control+End')
  const before = await snapshot()
  await (await homeItem('doc-numbered')).locator('.ribbon-split-arrow').click()
  const picker = page.locator('.doc-list-picker:visible')
  const start = picker.getByRole('spinbutton', { name: '起始编号', exact: true })
  await expect(start).toBeEnabled()
  await expect(start).toHaveClass(/ant-input-number-input/)
  await expect(picker.locator('.doc-list-card')).toHaveCount(6)
  await expect(picker.getByRole('button', { name: '继续前一编号', exact: true })).toBeEnabled()
  await start.fill('7')
  await picker.getByRole('button', { name: '重新编号', exact: true }).click()
  await expect.poll(async () => (await snapshot()).sectionHeaderFooters.first.storyContent.header.blocks[2].metadata.numberingDefinition.instance.levelStartOverrides['0']).toBe(7)
  await expect.poll(async () => (await inputState()).location.host).toEqual({ kind: 'story', sectionId: 'first', story: 'header' })
  await expect(bodyBridge()).toBeFocused()
  assert.deepEqual((await snapshot()).blocks, before.blocks)
  assert.equal((await snapshot()).sectionHeaderFooters.first.header, before.sectionHeaderFooters.first.header)
  assert.equal((await snapshot()).sectionHeaderFooters.first.footer, before.sectionHeaderFooters.first.footer)
  assert.deepEqual(await history(), { undo: 1, redo: 0, dirty: true })
  await bodyBridge().press('Control+z')
  await expect.poll(async () => (await snapshot()).sectionHeaderFooters).toEqual(before.sectionHeaderFooters)
  await bodyBridge().press('Control+y')
  await expect.poll(async () => (await snapshot()).sectionHeaderFooters.first.storyContent.header.blocks[2].metadata.numberingDefinition.instance.levelStartOverrides['0']).toBe(7)
  await (await homeItem('doc-numbered')).locator('.ribbon-split-arrow').click()
  await expect(picker.getByRole('button', { name: '继续前一编号', exact: true })).toBeEnabled()
  await picker.getByRole('button', { name: '继续前一编号', exact: true }).click()
  await expect.poll(async () => {
    const blocks = (await snapshot()).sectionHeaderFooters.first.storyContent.header.blocks
    return blocks[2].metadata.numberingListId === blocks[0].metadata.numberingListId
  }).toBe(true)
  await expect(bodyBridge()).toBeFocused()
  assert.deepEqual(await history(), { undo: 2, redo: 0, dirty: true })
  const numbered = await snapshot(), previousHistory = await history(), previousCommands = await page.evaluate(() => window.documentStoryFixture.commands())
  await bodyBridge().dispatchEvent('compositionstart', { data: '' })
  for (const button of await (await homeItem('doc-numbered')).getByRole('button').all()) await expect(button).toBeDisabled()
  await page.locator('[data-testid="workbench-ribbon"]').evaluate(element => {
    element.__vueParentComponent.emit('execute', 'doc.command.restart-list-numbering', { start: 11 })
    element.__vueParentComponent.emit('execute', 'doc.command.continue-list-numbering')
  })
  assert.deepEqual(await snapshot(), numbered)
  assert.deepEqual(await history(), previousHistory)
  assert.deepEqual(await page.evaluate(() => window.documentStoryFixture.commands()), previousCommands)
  await expect(bodyBridge()).toBeFocused()
  await bodyBridge().dispatchEvent('compositionend', { data: '' })
  await (await homeItem('doc-numbered')).locator('.ribbon-split-arrow').click()
  await expect(picker.getByRole('button', { name: '继续前一编号', exact: true })).toBeDisabled()
  const saved = await page.evaluate(() => window.documentStoryFixture.saveReopen())
  assert.deepEqual(saved.blocks, before.blocks)
  assert.deepEqual(saved.sectionHeaderFooters, numbered.sectionHeaderFooters)
  assert.deepEqual((await snapshot()).sectionHeaderFooters, saved.sectionHeaderFooters)
  assert.deepEqual(await history(), { undo: 0, redo: 0, dirty: false })
})

test.each([['flat', 'classic', false], ['flat', 'grid', true], ['rich', 'classic', false], ['rich', 'grid', true]])('paragraph indent and spacing use the active header, real menus, history and internal reopen in %s, %s, dark %s', async (kind, variant, dark) => {
  await page.evaluate(({ kind, variant, dark }) => {
    window.documentStoryFixture.paragraphStory(kind)
    window.documentStoryFixture.appearance(dark, variant)
  }, { kind, variant, dark })
  await (await homeItem('doc-header-footer')).getByRole('button').click()
  const initialInput = kind === 'flat' ? storyBridge() : bodyBridge()
  await expect(initialInput).toBeFocused()
  await page.getByRole('tab', { name: '开始', exact: true }).click()
  await initialInput.focus()
  if (kind === 'flat') await nativeSelection(13)
  if (kind === 'rich') await initialInput.press('Control+End')
  const before = await snapshot()
  await expect((await homeItem('doc-indent-more')).getByRole('button')).toBeEnabled()
  await expect((await homeItem('doc-indent-less')).getByRole('button')).toBeEnabled()
  await expect(await homeItem('doc-line-height')).toContainText('固定值 24 px')
  await (await homeItem('doc-indent-more')).getByRole('button').click()
  await expect.poll(async () => (await snapshot()).sectionHeaderFooters.first.storyContent.header.blocks[2].style.indentFirstLine).toBe(40)
  await expect.poll(async () => (await inputState()).location.host).toEqual({ kind: 'story', sectionId: 'first', story: 'header' })
  await expect(bodyBridge()).toBeFocused()
  assert.deepEqual(await history(), { undo: 1, redo: 0, dirty: true })
  await bodyBridge().press('Control+z')
  await expect.poll(async () => (await snapshot()).sectionHeaderFooters).toEqual(before.sectionHeaderFooters)
  await bodyBridge().press('Control+y')
  await expect.poll(async () => (await snapshot()).sectionHeaderFooters.first.storyContent.header.blocks[2].style.indentFirstLine).toBe(40)
  await (await homeItem('doc-indent-less')).getByRole('button').click()
  await expect.poll(async () => (await snapshot()).sectionHeaderFooters.first.storyContent.header.blocks[2].style.indentFirstLine).toBe(20)
  await (await homeItem('doc-line-height')).getByRole('combobox').click()
  await page.locator('.ant-select-dropdown:visible').getByText('1.5 倍', { exact: true }).click()
  await expect.poll(async () => {
    const style = (await snapshot()).sectionHeaderFooters.first.storyContent.header.blocks[2].style
    return [style.spacingRule, style.lineHeight]
  }).toEqual(['auto', 1.5])
  for (const [key, title, value] of [['spaceBefore', '段前间距', 18], ['spaceAfter', '段后间距', 24]]) {
    await (await homeItem('doc-paragraph-spacing')).getByRole('button').click()
    await page.getByRole('menuitem', { name: title, exact: true }).click()
    await page.locator('[data-ribbon-item="doc-' + key + '-' + value + '"]:visible').click()
    await expect.poll(async () => (await snapshot()).sectionHeaderFooters.first.storyContent.header.blocks[2].style[key]).toBe(value)
    await expect(bodyBridge()).toBeFocused()
  }
  const formatted = await snapshot(), previousHistory = await history(), previousCommands = await page.evaluate(() => window.documentStoryFixture.commands())
  assert.deepEqual(previousHistory, { undo: 5, redo: 0, dirty: true })
  assert.deepEqual(formatted.blocks, before.blocks)
  assert.equal(formatted.sectionHeaderFooters.first.header, before.sectionHeaderFooters.first.header)
  assert.equal(formatted.sectionHeaderFooters.first.footer, before.sectionHeaderFooters.first.footer)
  assert.deepEqual(formatted.sectionHeaderFooters.first.storyContent.header.paragraphs.slice(0, 2).map(paragraph => paragraph.style),
    before.sectionHeaderFooters.first.storyContent.header.paragraphs.slice(0, 2).map(paragraph => paragraph.style))
  await bodyBridge().dispatchEvent('compositionstart', { data: '' })
  for (const id of ['doc-indent-less', 'doc-indent-more', 'doc-paragraph-spacing']) await expect((await homeItem(id)).getByRole('button')).toBeDisabled()
  await expect((await homeItem('doc-line-height')).getByRole('combobox')).toBeDisabled()
  await page.locator('[data-testid="workbench-ribbon"]').evaluate(element => {
    element.__vueParentComponent.emit('execute', 'doc.command.change-paragraph-indent', { direction: 1 })
    element.__vueParentComponent.emit('execute', 'doc.command.change-paragraph-indent', { direction: -1 })
    element.__vueParentComponent.emit('execute', 'doc.command.format-paragraph', { patch: { spacingRule: 'auto', lineHeight: 2 } })
    element.__vueParentComponent.emit('execute', 'doc.command.format-paragraph', { patch: { spaceBefore: 0, spaceAfter: 0 } })
  })
  assert.deepEqual(await snapshot(), formatted)
  assert.deepEqual(await history(), previousHistory)
  assert.deepEqual(await page.evaluate(() => window.documentStoryFixture.commands()), previousCommands)
  await expect(bodyBridge()).toBeFocused()
  await bodyBridge().dispatchEvent('compositionend', { data: '' })
  await expect((await homeItem('doc-indent-more')).getByRole('button')).toBeEnabled()
  await expect((await homeItem('doc-line-height')).getByRole('combobox')).toBeEnabled()
  const saved = await page.evaluate(() => window.documentStoryFixture.saveReopen())
  assert.deepEqual(saved.blocks, before.blocks)
  assert.deepEqual(saved.sectionHeaderFooters, formatted.sectionHeaderFooters)
  assert.deepEqual((await snapshot()).sectionHeaderFooters, saved.sectionHeaderFooters)
  assert.deepEqual(await history(), { undo: 0, redo: 0, dirty: false })
})

test.each([['flat', 'classic', false], ['flat', 'grid', true], ['rich', 'classic', false], ['rich', 'grid', true]])('fixed strike, alignment and list shortcuts retain the active header host, history and internal reopen in %s, %s, dark %s', async (kind, variant, dark) => {
  await page.evaluate(({ kind, variant, dark }) => {
    if (kind === 'rich') window.documentStoryFixture.richStories()
    window.documentStoryFixture.appearance(dark, variant)
  }, { kind, variant, dark })
  await (await homeItem('doc-header-footer')).getByRole('button').click()
  const initialInput = kind === 'flat' ? storyBridge() : bodyBridge()
  await expect(initialInput).toBeFocused()
  await page.getByRole('tab', { name: '开始', exact: true }).click()
  await initialInput.focus()
  if (kind === 'flat') await nativeSelection(5, 9)
  if (kind === 'rich') await initialInput.press('Control+A')
  const before = await snapshot(), selectedText = kind === 'flat' ? 'word' : 'Rich header'
  const previousCommands = await page.evaluate(() => window.documentStoryFixture.commands())
  await initialInput.evaluate(element => {
    for (const [code, key] of [['KeyX', 'X'], ['KeyL', 'L'], ['KeyE', 'E'], ['KeyR', 'R'], ['KeyJ', 'J'], ['Digit7', '&'], ['Digit8', '*']]) {
      element.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key, code, ctrlKey: true }))
      element.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key, code, ctrlKey: true, shiftKey: true, altKey: true }))
    }
  })
  await page.getByRole('tab', { name: '开始', exact: true }).focus()
  await initialInput.evaluate(element => {
    for (const [code, key] of [['KeyX', 'X'], ['KeyL', 'L'], ['KeyE', 'E'], ['KeyR', 'R'], ['KeyJ', 'J'], ['Digit7', '&'], ['Digit8', '*']]) element.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key, code, metaKey: true, shiftKey: true }))
  })
  assert.deepEqual(await snapshot(), before)
  assert.deepEqual(await history(), { undo: 0, redo: 0, dirty: false })
  assert.deepEqual(await page.evaluate(() => window.documentStoryFixture.commands()), previousCommands)
  await initialInput.focus()
  for (const [shortcut, strike] of [['Control+Shift+X', true], ['Meta+Shift+X', false]]) {
    await initialInput.press(shortcut)
    await expect.poll(async () => (await snapshot()).sectionHeaderFooters.first.storyContent.header.runs.some(run => run.text === selectedText && run.strike === strike)).toBe(true)
    await expect(initialInput).toBeFocused()
  }
  for (const [index, [shortcut, align]] of [['Control+Shift+R', 'right'], ['Meta+Shift+L', 'left'], ['Control+Shift+E', 'center'], ['Meta+Shift+J', 'justify']].entries()) {
    await (index === 0 ? initialInput : bodyBridge()).press(shortcut)
    await expect.poll(async () => (await snapshot()).sectionHeaderFooters.first.storyContent.header.blocks[0].style.align).toBe(align)
    await expect.poll(async () => (await inputState()).location.host).toEqual({ kind: 'story', sectionId: 'first', story: 'header' })
    await expect(bodyBridge()).toBeFocused()
  }
  const aligned = await snapshot()
  await bodyBridge().press('Control+Shift+J')
  assert.deepEqual(await snapshot(), aligned)
  assert.deepEqual(await history(), { undo: 6, redo: 0, dirty: true })
  for (const [shortcut, type, preset] of [['Control+Shift+7', 'numbered', 'ORDER_LIST'], ['Meta+Shift+7', 'paragraph', undefined], ['Meta+Shift+8', 'bullet', 'BULLET_LIST'], ['Control+Shift+8', 'paragraph', undefined]]) {
    await bodyBridge().press(shortcut)
    await expect.poll(async () => (await snapshot()).sectionHeaderFooters.first.storyContent.header.blocks[0].type).toBe(type)
    const block = (await snapshot()).sectionHeaderFooters.first.storyContent.header.blocks[0]
    if (preset) assert.equal(block.list.preset, preset)
    if (!preset) assert.equal(block.list, undefined)
    await expect(bodyBridge()).toBeFocused()
  }
  const formatted = await snapshot()
  assert.deepEqual(formatted.blocks, before.blocks)
  assert.equal(formatted.sectionHeaderFooters.first.header, before.sectionHeaderFooters.first.header)
  assert.equal(formatted.sectionHeaderFooters.first.footer, before.sectionHeaderFooters.first.footer)
  assert.deepEqual(formatted.sectionHeaderFooters.first.storyContent.footer, before.sectionHeaderFooters.first.storyContent.footer)
  assert.deepEqual(formatted.sectionHeaderFooters.first.storyContent.header.paragraphs.slice(1).map(paragraph => paragraph.style), before.sectionHeaderFooters.first.storyContent.header.paragraphs.slice(1).map(paragraph => paragraph.style))
  for (const [index, field] of formatted.sectionHeaderFooters.first.pageFields.entries()) {
    const original = before.sectionHeaderFooters.first.pageFields[index]
    assert.deepEqual([field.story, field.field, field.offset, field.cachedResult, field.textStyle.fontFamily, field.textStyle.fontSize], [original.story, original.field, original.offset, original.cachedResult, original.textStyle.fontFamily, original.textStyle.fontSize])
  }
  assert.equal(formatted.sectionHeaderFooters.first.storyContent.header.runs.some(run => run.text === selectedText && run.fontSize === before.sectionHeaderFooters.first.storyContent.header.runs[0].fontSize), true)
  assert.deepEqual(await history(), { undo: 10, redo: 0, dirty: true })
  await bodyBridge().press('Control+z')
  await expect.poll(async () => (await snapshot()).sectionHeaderFooters.first.storyContent.header.blocks[0].type).toBe('bullet')
  await bodyBridge().press('Control+y')
  await expect.poll(async () => (await snapshot()).sectionHeaderFooters.first.storyContent.header.blocks[0].type).toBe('paragraph')
  await bodyBridge().press('Meta+z')
  await expect.poll(async () => (await snapshot()).sectionHeaderFooters.first.storyContent.header.blocks[0].type).toBe('bullet')
  assert.deepEqual(await history(), { undo: 9, redo: 1, dirty: true })
  await page.evaluate(() => window.documentStoryFixture.readonly())
  const readonly = await snapshot(), readonlyHistory = await history(), readonlyCommands = await page.evaluate(() => window.documentStoryFixture.commands())
  await bodyBridge().focus()
  await bodyBridge().evaluate(element => {
    for (const [code, key] of [['KeyX', 'X'], ['KeyL', 'L'], ['KeyE', 'E'], ['KeyR', 'R'], ['KeyJ', 'J'], ['Digit7', '&'], ['Digit8', '*']]) element.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key, code, ctrlKey: true, shiftKey: true }))
  })
  assert.deepEqual(await snapshot(), readonly)
  assert.deepEqual(await history(), readonlyHistory)
  assert.deepEqual(await page.evaluate(() => window.documentStoryFixture.commands()), readonlyCommands)
  const saved = await page.evaluate(() => window.documentStoryFixture.saveReopen())
  assert.deepEqual(saved.blocks, before.blocks)
  assert.equal(saved.sectionHeaderFooters.first.storyContent.header.blocks[0].type, 'bullet')
  assert.equal(saved.sectionHeaderFooters.first.storyContent.header.blocks[0].list.preset, 'BULLET_LIST')
  assert.equal(saved.sectionHeaderFooters.first.storyContent.header.blocks[0].style.align, 'justify')
  assert.deepEqual((await snapshot()).sectionHeaderFooters, saved.sectionHeaderFooters)
  assert.deepEqual(await history(), { undo: 0, redo: 0, dirty: false })
})

test.each([['flat', 'classic', false], ['flat', 'grid', true], ['rich', 'classic', false], ['rich', 'grid', true]])('fixed vertical shortcuts keep the active header, body, history and saved input format in %s, %s, dark %s', async (kind, variant, dark) => {
  await page.evaluate(({ kind, variant, dark }) => {
    if (kind === 'rich') window.documentStoryFixture.richStories()
    window.documentStoryFixture.appearance(dark, variant)
  }, { kind, variant, dark })
  await (await homeItem('doc-header-footer')).getByRole('button').click()
  const input = kind === 'flat' ? storyBridge() : bodyBridge()
  await expect(input).toBeFocused()
  await page.getByRole('tab', { name: '开始', exact: true }).click()
  await input.focus()
  if (kind === 'flat') await nativeSelection(5, 9)
  if (kind === 'rich') await input.press('Control+A')
  const before = await snapshot(), selectedText = kind === 'flat' ? 'word' : 'Rich header'
  for (const [shortcut, verticalAlign] of [['Control+Comma', 'subscript'], ['Meta+Period', 'superscript'], ['Control+Period', undefined], ['Meta+Comma', 'subscript'], ['Control+Comma', undefined]]) {
    await input.press(shortcut)
    await expect.poll(async () => (await snapshot()).sectionHeaderFooters.first.storyContent.header.runs.some(run => run.text === selectedText && run.verticalAlign === verticalAlign)).toBe(true)
    await expect(input).toBeFocused()
  }
  assert.deepEqual((await snapshot()).blocks, before.blocks)
  assert.deepEqual((await snapshot()).sectionHeaderFooters.first.pageFields, before.sectionHeaderFooters.first.pageFields)
  assert.deepEqual(await history(), { undo: 5, redo: 0, dirty: true })
  const formatted = await snapshot(), previousCommands = await page.evaluate(() => window.documentStoryFixture.commands())
  await input.evaluate(element => {
    element.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: '<', code: 'Comma', ctrlKey: true, shiftKey: true }))
    element.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: '.', code: 'Period', ctrlKey: true, altKey: true }))
  })
  await page.getByRole('tab', { name: '开始', exact: true }).focus()
  await input.evaluate(element => element.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: '.', code: 'Period', ctrlKey: true })))
  assert.deepEqual(await snapshot(), formatted)
  assert.deepEqual(await page.evaluate(() => window.documentStoryFixture.commands()), previousCommands)
  assert.deepEqual(await history(), { undo: 5, redo: 0, dirty: true })
  await input.focus()
  await input.press('Control+z')
  await expect.poll(async () => (await snapshot()).sectionHeaderFooters.first.storyContent.header.runs.some(run => run.text === selectedText && run.verticalAlign === 'subscript')).toBe(true)
  await input.press('Control+y')
  await expect.poll(async () => (await snapshot()).sectionHeaderFooters.first.storyContent.header.runs.some(run => run.text === selectedText && run.verticalAlign === undefined)).toBe(true)
  await input.press('ArrowRight')
  const collapsed = await snapshot()
  await input.press('Control+Comma')
  assert.deepEqual(await snapshot(), collapsed)
  assert.deepEqual(await history(), { undo: 5, redo: 0, dirty: true })
  await input.press('!')
  await expect.poll(async () => (await snapshot()).sectionHeaderFooters.first.header).toBe(kind === 'flat' ? 'A😀 word!\nTail' : 'Rich header!')
  assert.equal((await snapshot()).sectionHeaderFooters.first.storyContent.header.runs.some(run => run.text === '!' && run.verticalAlign === 'subscript'), true)
  assert.deepEqual((await snapshot()).blocks, before.blocks)
  assert.deepEqual(await history(), { undo: 6, redo: 0, dirty: true })
  await page.evaluate(() => window.documentStoryFixture.readonly())
  await expect((await homeItem('doc-subscript')).getByRole('button')).toBeDisabled()
  const readonly = await snapshot(), readonlyHistory = await history(), readonlyCommands = await page.evaluate(() => window.documentStoryFixture.commands())
  await input.focus()
  await input.press('Control+Period')
  assert.deepEqual(await snapshot(), readonly)
  assert.deepEqual(await history(), readonlyHistory)
  assert.deepEqual(await page.evaluate(() => window.documentStoryFixture.commands()), readonlyCommands)
  const saved = await page.evaluate(() => window.documentStoryFixture.saveReopen())
  assert.deepEqual(saved.blocks, before.blocks)
  assert.equal(saved.sectionHeaderFooters.first.storyContent.header.runs.some(run => run.text === '!' && run.verticalAlign === 'subscript'), true)
  assert.deepEqual((await snapshot()).sectionHeaderFooters, saved.sectionHeaderFooters)
  assert.deepEqual(await history(), { undo: 0, redo: 0, dirty: false })
})

for (const variant of ['classic', 'grid']) test('legacy header named styles use the story selection, real history and saved reopen in ' + variant, async () => {
  await page.evaluate(layout => { window.documentStoryFixture.namedStyles(); window.documentStoryFixture.appearance(false, layout) }, variant)
  await openHeader()
  await nativeSelection(11, 12)
  await page.getByRole('tab', { name: '开始', exact: true }).click()
  const style = await homeItem('doc-block-style')
  await expect(style.locator('.ant-select-selection-item')).toHaveText('标题 4')
  await nativeSelection(1, 9)
  await expect(style.locator('.ant-select-selection-item')).toHaveText('标题')
  const before = await snapshot()
  assert.equal((await inputState()).location.kind, 'story')
  await style.getByRole('combobox').click()
  await page.getByRole('option', { name: '标题 3', exact: true }).click()
  await expect.poll(() => snapshot().then(value => value.sectionHeaderFooters.first.storyContent.header.paragraphs.map(paragraph => paragraph.metadata.namedStyleType))).toEqual(['heading3', 'heading4'])
  await expect(bodyBridge()).toBeFocused()
  const applied = await snapshot(), field = applied.sectionHeaderFooters.first.pageFields[0]
  assert.deepEqual(applied.blocks, before.blocks)
  assert.equal(applied.sectionHeaderFooters.first.header, before.sectionHeaderFooters.first.header)
  assert.equal(applied.sectionHeaderFooters.first.footer, before.sectionHeaderFooters.first.footer)
  assert.deepEqual([field.field, field.offset, field.cachedResult], ['PAGE', 3, before.sectionHeaderFooters.first.pageFields[0].cachedResult])
  assert.deepEqual((await inputState()).location.host, { kind: 'story', sectionId: 'first', story: 'header' })
  assert.deepEqual(await history(), { undo: 1, redo: 0, dirty: true })
  assert.equal((await page.evaluate(() => window.documentStoryFixture.commands().filter(event => event.id === 'doc.command.set-paragraph-named-style'))).length, 1)
  await page.keyboard.press('Control+z')
  await expect.poll(() => snapshot().then(value => value.sectionHeaderFooters)).toEqual(before.sectionHeaderFooters)
  await page.keyboard.press('Control+y')
  await expect.poll(() => snapshot().then(value => value.sectionHeaderFooters)).toEqual(applied.sectionHeaderFooters)
  const saved = await page.evaluate(() => window.documentStoryFixture.saveReopen())
  assert.deepEqual(saved.sectionHeaderFooters, applied.sectionHeaderFooters)
  assert.deepEqual((await snapshot()).sectionHeaderFooters, applied.sectionHeaderFooters)
  assert.deepEqual(await history(), { undo: 0, redo: 0, dirty: false })
})

for (const variant of ['classic', 'grid']) test('legacy header paragraph settings migrate the selected story and preserve body, history and saved reopen in ' + variant, async () => {
  await page.evaluate(layout => window.documentStoryFixture.appearance(layout === 'grid', layout), variant)
  await openHeader()
  await nativeSelection(1, 9)
  const before = await snapshot(), point = await page.evaluate(() => window.documentStoryFixture.point('header', 1))
  await page.mouse.click(point.x, point.y, { button: 'right' })
  await page.locator('.document-context-menu').getByRole('menuitem', { name: '段落设置', exact: true }).click()
  const panel = page.locator('[data-document-paragraph-settings]')
  await expect(panel.getByRole('button', { name: '居中', exact: true })).toHaveAttribute('aria-pressed', 'true')
  await expect(panel.getByRole('spinbutton', { name: '首行缩进 (px)', exact: true })).toBeEnabled()
  assert.deepEqual(await snapshot(), before)
  assert.deepEqual(await history(), { undo: 0, redo: 0, dirty: false })
  assert.deepEqual((await inputState()).location.host, { kind: 'story', sectionId: 'first', story: 'header' })
  await panel.getByRole('spinbutton', { name: '段前 (px)', exact: true }).fill('6')
  await panel.getByRole('spinbutton', { name: '段前 (px)', exact: true }).press('Tab')
  await expect.poll(() => snapshot().then(value => value.sectionHeaderFooters.first.storyContent.header.paragraphs[0].style.spaceBefore)).toBe(6)
  const applied = await snapshot()
  assert.deepEqual(applied.blocks, before.blocks)
  assert.equal(applied.sectionHeaderFooters.first.header, before.sectionHeaderFooters.first.header)
  assert.equal(applied.sectionHeaderFooters.first.footer, before.sectionHeaderFooters.first.footer)
  assert.deepEqual(await history(), { undo: 1, redo: 0, dirty: true })
  await panel.getByRole('button', { name: '关闭段落设置', exact: true }).click()
  await expect(bodyBridge()).toBeFocused()
  await page.keyboard.press('Control+z')
  await expect.poll(() => snapshot().then(value => value.sectionHeaderFooters)).toEqual(before.sectionHeaderFooters)
  await page.keyboard.press('Control+y')
  await expect.poll(() => snapshot().then(value => value.sectionHeaderFooters)).toEqual(applied.sectionHeaderFooters)
  const saved = await page.evaluate(() => window.documentStoryFixture.saveReopen())
  assert.deepEqual(saved.sectionHeaderFooters, applied.sectionHeaderFooters)
  assert.deepEqual((await snapshot()).sectionHeaderFooters, applied.sectionHeaderFooters)
  assert.deepEqual(await history(), { undo: 0, redo: 0, dirty: false })
})

test('legacy header composition disables named paragraph styles without adding history', async () => {
  await openHeader()
  await startComposition()
  await page.getByRole('tab', { name: '开始', exact: true }).click()
  await expect((await homeItem('doc-block-style')).locator('.ant-select')).toHaveClass(/ant-select-disabled/)
  assert.deepEqual(await history(), { undo: 0, redo: 0, dirty: false })
})

for (const variant of ['classic', 'grid']) test('Canvas drag, clear, focus and internal reopen with the real story TextArea in ' + variant, async () => {
  const before = await snapshot()
  await page.evaluate(layout => window.documentStoryFixture.appearance(layout === 'grid', layout), variant)
  await openHeader(); await dragStory(5, 9); await dragStory(9, 5)
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
  assert.deepEqual(await page.evaluate(() => window.documentStoryFixture.historyRequests()), [{ direction: 'undo', unitId: 'story-browser' }, { direction: 'redo', unitId: 'story-browser' }])
  const saved = await page.evaluate(() => window.documentStoryFixture.saveReopen())
  assert.deepEqual((await snapshot()).sectionHeaderFooters, saved.sectionHeaderFooters)
  assert.deepEqual(await history(), { undo: 0, redo: 0, dirty: false })
  await expect(storyBridge()).toHaveValue(buffer)
  assert.deepEqual(await page.evaluate(() => window.documentStoryFixture.commands().filter(event => event.id === 'doc.command.clear-formatting').map(event => [event.unitId, event.kind])), [['story-browser', 'document']])
  assert.equal(await inputState(), undefined)
  await storyBridge().focus()
  await expect.poll(() => storyBridge().evaluate(element => [element.selectionStart, element.selectionEnd])).toEqual([5, 9])
  await page.keyboard.insertText('新词')
  await expect(storyBridge()).toHaveValue('A😀\uFFFC 新词\nTail')
  const reopenedInput = await snapshot(), typed = reopenedInput.sectionHeaderFooters.first.storyContent.header.runs.find(run => run.text === '新词')
  assert.deepEqual([typed.fontFamily, typed.fontSize, typed.bold, typed.color], ['Georgia', 9, false, 'auto'])
  assert.equal(typed.metadata.officeStoryTextFormatCleared, true)
  assert.deepEqual(reopenedInput.blocks, before.blocks)
  assert.deepEqual((await inputState()).location.range, { sectionId: 'first', story: 'header', start: 7, end: 7 })
  assert.deepEqual((await inputState()).marks, {})
  assert.deepEqual(await history(), { undo: 1, redo: 0, dirty: true })
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

test('a story composition renews its invalidated session after clear without moving the restored range', async () => {
  await openHeader(); await nativeSelection(5, 9)
  await page.getByRole('tab', { name: '开始', exact: true }).click()
  await (await homeItem('doc-clear-format')).getByRole('button').click()
  await expect(storyBridge()).toBeFocused()
  const cleared = await snapshot()
  assert.equal((await inputState()).valid, false)
  await startComposition()
  const source = await inputState()
  assert.equal(source.valid, true)
  assert.deepEqual(source.location.range, { sectionId: 'first', story: 'header', start: 5, end: 9 })
  assert.deepEqual(await snapshot(), cleared); assert.deepEqual(await history(), { undo: 1, redo: 0, dirty: true })
  await storyBridge().evaluate(element => {
    element.value = 'A😀\uFFFC 中文\nTail'; element.setSelectionRange(7, 7)
    element.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertCompositionText', data: '中文', isComposing: true }))
    element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '中文' }))
    element.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: '中文' }))
  })
  await expect(storyBridge()).toHaveValue('A😀\uFFFC 中文\nTail')
  const after = await snapshot(), typed = after.sectionHeaderFooters.first.storyContent.header.runs.find(run => run.text === '中文'), state = await inputState()
  assert.deepEqual([typed.fontFamily, typed.fontSize, typed.bold, typed.color], ['Georgia', 9, false, 'auto'])
  assert.deepEqual(after.blocks, cleared.blocks)
  assert.equal(state.valid, true); assert.equal(state.modelRevision, source.modelRevision + 1); assert.equal(state.generation, source.generation + 1)
  assert.deepEqual(state.location.range, { sectionId: 'first', story: 'header', start: 7, end: 7 })
  assert.deepEqual(await history(), { undo: 2, redo: 0, dirty: true })
})

for (const area of ['body', 'header']) test('a composition begun before Canvas receives a newer model blocks its old native events in ' + area, async () => {
  if (area === 'header') await openHeader()
  await page.evaluate(target => {
    const element = document.querySelector(target === 'body' ? '.document-input-bridge textarea, textarea.document-input-bridge' : '.document-story-bridge textarea, textarea.document-story-bridge')
    if (!(element instanceof HTMLTextAreaElement)) throw new Error('Fixture did not find the native TextArea')
    window.rejectedCompositionElement = element
    if (target === 'body') {
      window.documentStoryFixture.externalBodySelection(0, 4)
      window.documentStoryFixture.formatBody({ fontSize: 32 })
    } else {
      element.setSelectionRange(5, 9)
      window.documentStoryFixture.externalStorySelection(5, 9)
      window.documentStoryFixture.formatStory({ fontSize: 32 })
    }
    element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' }))
    element.dispatchEvent(new CompositionEvent('compositionupdate', { bubbles: true, data: '错误' }))
    element.value = '错误'
    element.dispatchEvent(new InputEvent('beforeinput', { bubbles: true, cancelable: true, inputType: 'insertCompositionText', data: '错误', isComposing: true }))
    element.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertCompositionText', data: '错误', isComposing: true }))
    element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '错误' }))
    element.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: '错误' }))
    element.dispatchEvent(new InputEvent('beforeinput', { bubbles: true, cancelable: true, inputType: 'historyUndo' }))
    element.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'z', ctrlKey: true }))
  }, area)
  await expect.poll(() => page.evaluate(() => !window.rejectedCompositionElement.isConnected)).toBe(true)
  const after = await snapshot(), state = await inputState()
  assert.equal(await page.evaluate(() => window.documentStoryFixture.bodyText()), 'body untouched')
  assert.equal(after.sectionHeaderFooters.first.header, 'A😀 word\nTail')
  const formatted = area === 'body' ? after.blocks[0].runs.find(run => run.text === 'body') : after.sectionHeaderFooters.first.storyContent.header.runs.find(run => run.text === 'word')
  assert.equal(formatted.fontSize, 32)
  assert.equal(state.valid, false); assert.deepEqual(state.marks, {})
  assert.deepEqual(await history(), { undo: 1, redo: 0, dirty: true })
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
    element.dispatchEvent(new FocusEvent('focus', { bubbles: false }))
    element.dispatchEvent(new FocusEvent('blur', { bubbles: false }))
    const transfer = new DataTransfer(); transfer.setData('text/plain', '错误')
    element.dispatchEvent(new ClipboardEvent('paste', { bubbles: true, cancelable: true, clipboardData: transfer }))
  })
  assert.deepEqual(await snapshot(), replaced); assert.deepEqual(await history(), { undo: 0, redo: 0, dirty: false })
  assert.equal(await page.evaluate(() => window.documentStoryFixture.bodyText()), 'body untouched')
})

for (const variant of ['classic', 'grid']) test('body Ribbon family and size remain transient until consecutive native input in ' + variant, async () => {
  await page.evaluate(layout => window.documentStoryFixture.appearance(layout === 'grid', layout), variant)
  await page.evaluate(() => window.documentStoryFixture.externalBodySelection(2))
  await expect(bodyBridge()).toBeFocused()
  const before = await snapshot()
  await setCharacterControl('doc-font-family', 'Tahoma')
  await setCharacterControl('doc-font-size', 25)
  await expect(bodyBridge()).toBeFocused()
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), { undo: 0, redo: 0, dirty: false })
  assert.deepEqual((await inputState()).marks, { fontFamily: 'Tahoma', fontSize: 25 })
  await page.keyboard.insertText('X'); await page.keyboard.insertText('Y')
  await expect.poll(() => page.evaluate(() => window.documentStoryFixture.bodyText())).toBe('boXYdy untouched')
  const after = await snapshot(), typed = after.blocks[0].runs.filter(run => /[XY]/.test(run.text))
  assert.ok(typed.length > 0)
  assert.ok(typed.every(run => run.fontFamily === 'Tahoma' && run.fontSize === 25))
  assert.deepEqual(after.sectionHeaderFooters, before.sectionHeaderFooters)
  assert.deepEqual(await history(), { undo: 2, redo: 0, dirty: true })
  assert.equal((await inputState()).location.selection.focus.offset, 4)
  await expect(await characterControl('doc-font-size')).toHaveValue('25')
  await bodyBridge().focus(); await bodyBridge().press('Control+z'); await bodyBridge().press('Control+z')
  await expect.poll(() => snapshot().then(value => value.blocks)).toEqual(before.blocks)
  await bodyBridge().press('Control+y'); await bodyBridge().press('Control+y')
  await expect.poll(() => snapshot().then(value => value.blocks)).toEqual(after.blocks)
  const saved = await page.evaluate(() => window.documentStoryFixture.saveReopen())
  assert.deepEqual((await snapshot()).blocks, saved.blocks)
  assert.equal(await inputState(), undefined)
  assert.deepEqual(await history(), { undo: 0, redo: 0, dirty: false })
})

for (const variant of ['classic', 'grid']) test('story Ribbon family and size drive captured typing and internal reopen in ' + variant, async () => {
  await page.evaluate(layout => window.documentStoryFixture.appearance(layout === 'grid', layout), variant)
  await openHeader(); await nativeSelection(5)
  const before = await snapshot()
  await setCharacterControl('doc-font-family', 'Tahoma')
  await setCharacterControl('doc-font-size', 25)
  await expect(storyBridge()).toBeFocused()
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), { undo: 0, redo: 0, dirty: false })
  assert.deepEqual((await inputState()).marks, { fontFamily: 'Tahoma', fontSize: 25 })
  await page.keyboard.insertText('X'); await page.keyboard.insertText('Y')
  await expect(storyBridge()).toHaveValue('A😀\uFFFC XYword\nTail')
  const after = await snapshot(), typed = after.sectionHeaderFooters.first.storyContent.header.runs.filter(run => /[XY]/.test(run.text))
  assert.ok(typed.length > 0)
  assert.ok(typed.every(run => run.fontFamily === 'Tahoma' && run.fontSize === 25 && run.bold && run.italic && run.color === '#FF0000'))
  assert.deepEqual(after.blocks, before.blocks)
  assert.deepEqual(after.sectionHeaderFooters.first.pageFields, before.sectionHeaderFooters.first.pageFields)
  assert.deepEqual(await history(), { undo: 2, redo: 0, dirty: true })
  assert.deepEqual((await inputState()).location.range, { sectionId: 'first', story: 'header', start: 7, end: 7 })
  await expect(await characterControl('doc-font-family')).toHaveValue('Tahoma')
  await expect(await characterControl('doc-font-size')).toHaveValue('25')
  await storyBridge().focus(); await storyBridge().press('Control+z'); await storyBridge().press('Control+z')
  await expect.poll(() => snapshot().then(value => value.sectionHeaderFooters)).toEqual(before.sectionHeaderFooters)
  await storyBridge().press('Control+y'); await storyBridge().press('Control+y')
  await expect.poll(() => snapshot().then(value => value.sectionHeaderFooters)).toEqual(after.sectionHeaderFooters)
  const saved = await page.evaluate(() => window.documentStoryFixture.saveReopen())
  assert.deepEqual((await snapshot()).sectionHeaderFooters, saved.sectionHeaderFooters)
  assert.equal(await inputState(), undefined)
  assert.deepEqual(await history(), { undo: 0, redo: 0, dirty: false })
})

test('external body selection and partial formatting drive Ribbon values and captured synthetic composition', async () => {
  await page.evaluate(() => {
    window.documentStoryFixture.externalBodySelection(0, 4)
    window.documentStoryFixture.formatBody({ fontFamily: 'Georgia', fontSize: 28 })
  })
  await expect(await characterControl('doc-font-family')).toHaveValue('Georgia')
  await expect(await characterControl('doc-font-size')).toHaveValue('28')
  await page.evaluate(() => {
    window.documentStoryFixture.externalBodySelection(1)
    window.documentStoryFixture.formatBody({ fontSize: 29 }, ['color', 'highlight'])
  })
  await expect(bodyBridge()).toBeFocused()
  const before = await snapshot(), captured = await inputState()
  assert.equal(captured.clearedColor, true); assert.equal(captured.clearedHighlight, true)
  await bodyBridge().evaluate(element => {
    element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' }))
    element.dispatchEvent(new CompositionEvent('compositionupdate', { bubbles: true, data: '中文' }))
  })
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), { undo: 1, redo: 0, dirty: true })
  await bodyBridge().evaluate(element => {
    element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '中文' }))
    element.value = '中文'
    element.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: '中文' }))
  })
  const after = await snapshot(), typed = after.blocks[0].runs.find(run => run.text === '中文')
  assert.equal(await page.evaluate(() => window.documentStoryFixture.bodyText()), 'b中文ody untouched')
  assert.deepEqual([typed.fontFamily, typed.fontSize, typed.color, typed.highlight], ['Georgia', 29, 'auto', undefined])
  assert.deepEqual(after.sectionHeaderFooters, before.sectionHeaderFooters)
  assert.deepEqual(await history(), { undo: 2, redo: 0, dirty: true })
  assert.equal((await inputState()).location.selection.focus.offset, 3)
})

test('the field-only Ribbon shows mixed values and changes only the requested family and size across cached runs', async () => {
  await page.evaluate(() => window.documentStoryFixture.mixedField())
  await openHeader(); await nativeSelection(3, 4)
  const before = await snapshot(), fieldBefore = before.sectionHeaderFooters.first.pageFields[0]
  await expect(await characterControl('doc-font-family')).toHaveValue('')
  await expect(await characterControl('doc-font-size')).toHaveValue('')
  await setCharacterControl('doc-font-family', 'Tahoma')
  await setCharacterControl('doc-font-size', 20)
  const after = await snapshot(), field = after.sectionHeaderFooters.first.pageFields[0], cache = field.metadata.officeStoryFieldCache.runs
  const original = fieldBefore.metadata.officeStoryFieldCache.runs
  const otherProperties = runs => runs.map(({ fontFamily: _family, eastAsiaFontFamily: _eastAsia, fontSize: _size, ...run }) => run)
  assert.deepEqual(otherProperties(cache), otherProperties(original))
  assert.ok(cache.every(run => run.fontFamily === 'Tahoma' && run.eastAsiaFontFamily === 'Tahoma' && run.fontSize === 20))
  assert.equal(cache[2].metadata.entityId, 'empty-cache')
  assert.deepEqual(after.blocks, before.blocks)
  assert.deepEqual(after.sectionHeaderFooters.first.storyContent, before.sectionHeaderFooters.first.storyContent)
  assert.deepEqual([field.field, field.instruction, field.offset, field.cachedResult], [fieldBefore.field, fieldBefore.instruction, fieldBefore.offset, fieldBefore.cachedResult])
  await expect(await characterControl('doc-font-family')).toHaveValue('Tahoma')
  await expect(await characterControl('doc-font-size')).toHaveValue('20')
  assert.deepEqual(await history(), { undo: 2, redo: 0, dirty: true })
})

test('story synthetic composition captures size and explicit color clears and advances once to its exact caret', async () => {
  await openHeader(); await nativeSelection(5)
  await page.evaluate(() => window.documentStoryFixture.formatStory({ fontSize: 27 }, ['color', 'highlight']))
  await expect(storyBridge()).toBeFocused()
  const before = await snapshot(), source = await inputState()
  assert.equal(source.clearedColor, true); assert.equal(source.clearedHighlight, true)
  await startComposition(5, 5)
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), { undo: 0, redo: 0, dirty: false })
  await storyBridge().evaluate(element => {
    element.value = 'A😀\uFFFC 中文word\nTail'; element.setSelectionRange(7, 7)
    element.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertCompositionText', data: '中文', isComposing: true }))
    element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '中文' }))
    element.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: '中文' }))
  })
  await expect(storyBridge()).toHaveValue('A😀\uFFFC 中文word\nTail')
  const after = await snapshot(), typed = after.sectionHeaderFooters.first.storyContent.header.runs.find(run => run.text === '中文'), state = await inputState()
  assert.deepEqual([typed.fontSize, typed.color, typed.highlight, typed.bold, typed.italic], [27, 'auto', undefined, true, true])
  assert.deepEqual(after.blocks, before.blocks)
  assert.deepEqual(state.location.range, { sectionId: 'first', story: 'header', start: 7, end: 7 })
  assert.equal(state.modelRevision, source.modelRevision + 1); assert.equal(state.generation, source.generation + 1)
  assert.equal(state.clearedColor, true); assert.equal(state.clearedHighlight, true)
  assert.deepEqual(await history(), { undo: 1, redo: 0, dirty: true })
  assert.equal((await page.evaluate(() => window.documentStoryFixture.commands().filter(event => event.id === 'doc.command.set-section-header-footer'))).length, 1)
})

for (const reason of ['body-navigation', 'story-navigation', 'same-location-format', 'same-id-facade']) test('public input changes cancel the old story composition after ' + reason, async () => {
  await openHeader()
  await startComposition(5, reason === 'same-location-format' ? 5 : 9)
  await page.evaluate(value => {
    if (value === 'body-navigation') window.documentStoryFixture.externalBodySelection(2)
    else if (value === 'story-navigation') window.documentStoryFixture.externalStorySelection(1, 1, 'footer')
    else if (value === 'same-location-format') window.documentStoryFixture.formatStory({ fontSize: 31 })
    else window.documentStoryFixture.rebindSame()
  }, reason)
  if (reason === 'body-navigation') await expect(bodyBridge()).toBeFocused()
  else if (reason === 'story-navigation') await expect(page.getByLabel('编辑页脚', { exact: true })).toBeFocused()
  else await expect(storyBridge()).toBeFocused()
  await expect.poll(() => page.evaluate(() => !window.staleStoryElement.isConnected)).toBe(true)
  const before = await snapshot(), state = await inputState(), beforeHistory = await history()
  await staleStoryEvents()
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await inputState(), state); assert.deepEqual(await history(), beforeHistory)
  assert.equal(before.sectionHeaderFooters.first.header, 'A😀 word\nTail')
  if (reason === 'same-location-format') assert.equal(state.marks.fontSize, 31)
})

for (const reason of ['same-location-format', 'same-id-facade']) test('body public session changes cancel a draft without accepting its stale composition end after ' + reason, async () => {
  await page.evaluate(() => {
    window.documentStoryFixture.externalBodySelection(2)
    window.documentStoryFixture.formatBody({ fontSize: 24 })
  })
  await expect(bodyBridge()).toBeFocused()
  await bodyBridge().evaluate(element => {
    window.staleBodyElement = element
    element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' }))
    element.dispatchEvent(new CompositionEvent('compositionupdate', { bubbles: true, data: '拼' }))
  })
  await page.evaluate(value => {
    if (value === 'same-location-format') window.documentStoryFixture.formatBody({ fontSize: 31 })
    else window.documentStoryFixture.rebindSame()
  }, reason)
  await expect.poll(() => page.evaluate(() => !window.staleBodyElement.isConnected)).toBe(true)
  const before = await snapshot(), state = await inputState()
  await page.evaluate(() => {
    const element = window.staleBodyElement
    element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '错误' }))
    element.value = '错误'
    element.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: '错误' }))
  })
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await inputState(), state)
  assert.equal(state.marks.fontSize, reason === 'same-location-format' ? 31 : 24)
  assert.deepEqual(await history(), { undo: 0, redo: 0, dirty: false })
})

test('navigation and a new caret format inside story commit publication survive the older result and advance', async () => {
  await openHeader(); await nativeSelection(5, 9)
  await page.evaluate(() => window.documentStoryFixture.navigateOnStoryCommit())
  await page.keyboard.insertText('新词')
  await expect(bodyBridge()).toBeFocused()
  const after = await snapshot(), state = await inputState()
  assert.equal(after.sectionHeaderFooters.first.header, 'A😀 新词\nTail')
  assert.equal(state.location.kind, 'body'); assert.equal(state.location.selection.focus.offset, 2); assert.equal(state.marks.fontSize, 31)
  await expect(await characterControl('doc-font-size')).toHaveValue('31')
  assert.deepEqual(await history(), { undo: 1, redo: 0, dirty: true })
  assert.equal((await page.evaluate(() => window.documentStoryFixture.commands().filter(event => event.id === 'doc.operation.advance-input-marks'))).length, 0)
})

test('a same-text native rewrite restores a rejected range and then inherits its first character on an identical retry', async () => {
  await page.evaluate(() => window.documentStoryFixture.simpleStory())
  await openHeader('aaa'); await nativeSelection(0, 3)
  const before = await snapshot()
  await page.evaluate(() => window.documentStoryFixture.cancelEdits(true))
  await page.keyboard.insertText('aaa')
  await expect(storyBridge()).toHaveValue('aaa')
  await expect.poll(() => storyBridge().evaluate(element => [element.selectionStart, element.selectionEnd])).toEqual([0, 3])
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), { undo: 0, redo: 0, dirty: false })
  await page.evaluate(() => window.documentStoryFixture.cancelEdits(false))
  await page.keyboard.insertText('aaa')
  await expect.poll(() => storyBridge().evaluate(element => [element.selectionStart, element.selectionEnd])).toEqual([3, 3])
  const after = await snapshot(), runs = after.sectionHeaderFooters.first.storyContent.header.runs.filter(run => run.text)
  assert.deepEqual(runs.map(run => [run.text, run.fontFamily, run.fontSize, run.bold, run.italic]), [['aaa', 'Georgia', 24, true, false]])
  assert.deepEqual(await history(), { undo: 1, redo: 0, dirty: true })
  assert.deepEqual((await inputState()).location.range, { sectionId: 'first', story: 'header', start: 3, end: 3 })
})

test('closing the document from story input publication does not read or focus the disposed facade after a successful commit', async () => {
  await openHeader(); await nativeSelection(5, 9)
  await page.evaluate(() => window.documentStoryFixture.closeOnStoryCommit())
  await page.keyboard.insertText('新词')
  await expect.poll(() => snapshot().then(value => value.id)).not.toBe('story-browser')
  assert.deepEqual(await page.evaluate(() => window.documentStoryFixture.closedDocumentIds()), ['story-browser'])
  assert.ok((await page.evaluate(() => window.documentStoryFixture.notices())).every(message => !message.includes('Document not found')))
  assert.deepEqual(await history(), { undo: 0, redo: 0, dirty: false })
})

test('a canceled UI undo keeps its captured input preferences and a successful retry clears them', async () => {
  await page.evaluate(() => window.documentStoryFixture.externalBodySelection(2))
  await expect(bodyBridge()).toBeFocused()
  await setCharacterControl('doc-font-size', 25)
  await expect(bodyBridge()).toBeFocused(); await page.keyboard.insertText('X')
  const before = await snapshot(), state = await inputState()
  await page.evaluate(() => window.documentStoryFixture.cancelUndo(true))
  await bodyBridge().press('Control+z')
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await inputState(), state)
  assert.deepEqual(await page.evaluate(() => window.documentStoryFixture.historyRequests()), [{ direction: 'undo', unitId: 'story-browser' }])
  assert.deepEqual(await history(), { undo: 1, redo: 0, dirty: true })
  assert.ok((await page.evaluate(() => window.documentStoryFixture.notices())).some(message => message.includes('canceled')))
  await page.evaluate(() => window.documentStoryFixture.cancelUndo(false))
  await bodyBridge().press('Control+z')
  await expect.poll(() => page.evaluate(() => window.documentStoryFixture.bodyText())).toBe('body untouched')
  assert.deepEqual((await inputState()).marks, {})
  assert.deepEqual(await page.evaluate(() => window.documentStoryFixture.historyRequests()), [{ direction: 'undo', unitId: 'story-browser' }, { direction: 'undo', unitId: 'story-browser' }])
  assert.deepEqual(await history(), { undo: 0, redo: 1, dirty: false })
})
