import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

const fixturePath = '/__document-home-colors-test__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>Docs Home command fixture</title><style>body{margin:0}#app{height:100vh;min-width:0}</style></head><body><div id="app"></div><script type="module">
import { computed, createApp, h, KeepAlive, ref, shallowRef, watch } from 'vue'
import { ConfigProvider, theme } from 'antdv-next'
import { OfficeStudio } from '/packages/office-studio/src/api/office-studio'
import { FDocument } from '/packages/office-studio/src/api/facade'
import DocumentWorkbench from '/packages/office-studio/src/components/DocumentWorkbench.vue'
import { createBlankDocument } from '/packages/office-studio/src/domain/model'
import { blocksFromHtml, createParagraph, documentHtml, normalizeDocument } from '/packages/office-studio/src/features/document/model/document-model'
import { paragraphListMarkers } from '/packages/office-studio/src/features/document/model/document-lists'
import { numberingDefinitionFromParts } from '/packages/office-studio/src/features/document/model/document-numbering'
import { officeNumberingMarkerBaseline } from '/packages/office-studio/src/features/document/model/document-office-numbering-marker'
import { createTextMeasurer, layoutDocument } from '/packages/office-studio/src/features/document/layout/document-layout'
import { sceneOffset } from '/packages/office-studio/src/features/document/rendering/document-scene'
import { createOfficeMessages, OfficeDirectionKey, OfficeLocaleKey, OfficeTranslatorKey } from '/packages/office-studio/src/i18n/office-i18n'
const copy = value => JSON.parse(JSON.stringify(value))
const initial = normalizeDocument({ ...createBlankDocument(), id: 'home-colors', name: 'Docs 开始命令', blocks: [createParagraph('甲乙丙丁', { color: '#414657' })] })
const apiMode = new URLSearchParams(location.search).has('api')
const studio = apiMode ? new OfficeStudio({ document: initial, currentUserId: 'local-user' }) : undefined
if (studio) studio.activateDocument(initial.id)
const documentApi = shallowRef(studio ? new FDocument(studio, initial.id) : undefined)
const model = ref(documentApi.value?.getModel() ?? initial), history = ref([]), redo = ref([]), locale = ref('zh-CN'), dark = ref(false), variant = ref('classic'), visible = ref(true)
const apiCommands = []
let cancelClear = false
if (studio) {
  watch(() => studio.editor.document.model, value => { model.value = value }, { flush: 'sync' })
  studio.events.addEvent('CommandExecuted', event => { apiCommands.push(copy(event)) })
  studio.events.addEvent('BeforeCommandExecute', event => { if (cancelClear && event.id === 'doc.command.clear-formatting') event.cancel = true })
  window.addEventListener('pagehide', () => studio.dispose(), { once: true })
}
const direction = computed(() => locale.value === 'ar-SA' ? 'rtl' : 'ltr')
const messages = createOfficeMessages()
const translate = (key, fallback) => messages[locale.value][key] ?? fallback
const update = value => {
  if (studio) { studio.editor.replaceDocumentFromView(value); model.value = documentApi.value.getModel() }
  else { history.value.push(copy(model.value)); redo.value = []; model.value = value }
}
function replaceFixtureModel(next, clearHistory = false) {
  next = normalizeDocument(next)
  if (studio) {
    if (!studio.getDocumentModel(next.id)) studio.createDocument(next)
    else { studio.activateDocument(next.id); studio.editor.document.replace(next) }
    documentApi.value = new FDocument(studio, next.id)
    if (clearHistory) studio.editor.history.clear(next.id)
    model.value = documentApi.value.getModel()
  } else {
    model.value = next
    if (clearHistory) { history.value = []; redo.value = [] }
  }
}
window.documentHomeFixture = {
  snapshot: () => copy(model.value),
  history: () => {
    const unit = studio?.editor.history.units[documentApi.value?.id]
    return unit ? { undo: unit.undo.length, redo: unit.redo.length } : { undo: history.value.length, redo: redo.value.length }
  },
  apiCommands: () => copy(apiCommands),
  cancelClear: value => { cancelClear = value },
  appearance: (isDark, language) => { dark.value = isDark; locale.value = language },
  variant: value => { variant.value = value },
  show: value => { visible.value = value },
  setColor: color => { const next = copy(model.value); next.blocks[0].runs.forEach(run => { run.color = color }); replaceFixtureModel(next) },
  denyEdit: () => { const next = copy(model.value); next.permissionRules = [{ id: 'edit-denied', scope: 'document', action: 'edit', effect: 'deny', userIds: [] }]; replaceFixtureModel(next) },
  reopen: snapshot => { replaceFixtureModel(JSON.parse(snapshot)) },
  newDocument: () => { const next = createBlankDocument(); next.id = 'new-home-doc'; next.blocks = [createParagraph('新文档', { color: '#414657' })]; replaceFixtureModel(next, true) },
  paragraphs: styles => {
    const next = copy(initial)
    next.blocks = styles.map((style, index) => { const paragraph = createParagraph('段落' + (index + 1)); paragraph.style = { ...paragraph.style, ...style }; return paragraph })
    replaceFixtureModel(next, true)
  },
  codeRange: () => {
    const next = copy(initial)
    next.blocks = [createParagraph('Before'), createParagraph('Code'), createParagraph('After')]
    const point = { blockId: next.blocks[1].id, offset: 1 }
    replaceFixtureModel(next, true)
    documentApi.value.executeCommand({ id: 'doc.command.wrap-structured-block', blockType: 'code' }, { anchor: point, focus: { ...point } })
    studio.editor.history.clear(documentApi.value.id)
  },
  clearFormatting: () => {
    const next = createBlankDocument(); next.id = 'clear-formatting-home-doc'
    next.defaultStyle.text = { ...next.defaultStyle.text, fontFamily: 'Georgia', eastAsiaFontFamily: '宋体', fontSize: 15, color: 'auto' }
    next.defaultStyle.paragraph = { ...next.defaultStyle.paragraph, align: 'right', spaceAfter: 12 }
    const block = createParagraph('Hello world', { bold: false, italic: true, underline: true, strike: true, fontSize: 26, color: '#FF0000', highlight: '#FFFF00', verticalAlign: 'superscript', href: 'https://example.org' })
    block.type = 'heading'; block.level = 3; block.style.align = 'center'; block.list = { id: 'clear-test-list', preset: 'ORDER_LIST' }
    next.blocks = [block]; replaceFixtureModel(next, true)
  },
  htmlColors: () => copy(blocksFromHtml(documentHtml(model.value)).flatMap(block => block.runs.map(run => ({ color: run.color, highlight: run.highlight })))),
  mixedLists: kind => {
    const next = copy(initial)
    next.blocks = [kind === 'numbered' ? 'ORDER_LIST' : 'BULLET_LIST', kind === 'numbered' ? 'ORDER_LIST_1' : 'BULLET_LIST_1'].map((preset, index) => {
      const block = createParagraph('项目' + (index + 1)); block.type = kind; block.list = { id: 'mixed-list-' + index, preset }; return block
    })
    replaceFixtureModel(next, true)
  },
  numbering: () => {
    const next = createBlankDocument(); next.id = 'numbering-home-doc'
    next.blocks = ['前一项', '前二项', '间隔段落', '后一项', '后二项'].map(text => createParagraph(text))
    for (const index of [0, 1, 3, 4]) { next.blocks[index].type = 'numbered'; next.blocks[index].list = { id: index < 2 ? 'previous-list' : 'current-list', preset: 'ORDER_LIST' } }
    replaceFixtureModel(next, true)
  },
  consecutiveNumbering: options => {
    const next = createBlankDocument(); next.id = 'consecutive-home-doc'
    const definition = numberingDefinitionFromParts({ abstractNumId: 7, consecutiveNumbering: true, levels: [
      { ilvl: 0, start: 1, format: 'decimal', text: '%1.' },
      { ilvl: 1, start: 1, format: 'decimal', text: '%1.%2.' },
    ] }, { numId: 9, abstractNumId: 7, consecutiveListId: 9, levelStartOverrides: { 0: 0 } })
    next.blocks = ['第一项', '第二项', '第三项', '第四项'].map((text, index) => {
      const block = createParagraph(text); block.type = 'numbered'; block.level = index % 2
      block.metadata = { numberingListId: 'shared', numberingDefinition: copy(definition) }
      if (options?.headings && index < 2) {
        block.type = 'heading'; block.metadata.numberingLevel = index; block.level = index + 2
      }
      if (options?.linked && index === 1) {
        block.metadata.numberingListId = 'secondary-instance'; block.metadata.numberingDefinition.instance.numId = 10
        block.metadata.numberingDefinition.instance.levelStartOverrides = {}
      }
      return block
    })
    replaceFixtureModel(next, true)
  },
  cachedNumberingMarker: options => {
    const next = createBlankDocument(); next.id = 'cached-office-numbering-doc'
    const block = createParagraph('[0] 原始正文'); block.id = 'cached-office-paragraph'
    const run = block.runs[0]
    block.runs = [{ ...run, id: 'cached-label', text: '[0] ', color: '#123456' }, { ...run, id: 'editable-body', text: '原始正文' }]
    block.metadata.officeNumberingMarker = { version: 1, bookmark: 'office_list_v1_fixture', bookmarkId: 99, text: '[0] ', sourceNumbering: null,
      xml: '<w:bookmarkStart w:id="99" w:name="office_list_v1_fixture"/><w:r><w:t>[0] </w:t></w:r><w:bookmarkEnd w:id="99"/>', baseline: officeNumberingMarkerBaseline(block, '[0] ') }
    next.blocks = [block]
    const startOffset = options?.labelOnly ? 1 : 2, endOffset = options?.labelOnly ? 3 : 8
    const anchorText = block.runs.map(run => run.text).join('').slice(startOffset, endOffset)
    next.comments = { review: { id: 'review', unitId: next.id, referenceText: anchorText, resolved: false,
      anchor: [{ blockId: block.id, startOffset, endOffset, text: anchorText }],
      messages: [{ id: 'review-message', text: '保留批注', authorId: 'local-user', authorName: '作者', createdAt: '2026-11-23T00:00:00Z' }],
      createdAt: '2026-11-23T00:00:00Z', updatedAt: '2026-11-23T00:00:00Z' } }
    next.contentControls = [{ id: 'body-control', type: 'richText', blockId: block.id, startOffset: 4, endOffset: 8, lock: 'unlocked' }]
    replaceFixtureModel(next, true)
  },
  markers: () => [...paragraphListMarkers(model.value.blocks).values()].map(value => value.marker),
  paragraphPoint: index => {
    const scroll = document.querySelector('.document-scroll'), bounds = scroll.getBoundingClientRect()
    const layout = layoutDocument(model.value, createTextMeasurer())
    const offset = sceneOffset(layout, { width: scroll.clientWidth, height: scroll.clientHeight, scrollLeft: scroll.scrollLeft, scrollTop: scroll.scrollTop, zoom: 100 })
    const line = layout.lines.find(value => value.target.blockId === model.value.blocks[index].id)
    return { x: bounds.left + offset.x + line.x + 3, y: bounds.top + offset.y + line.y + line.height / 2 }
  },
}
const app = createApp({ render: () => h(ConfigProvider, {
  direction: direction.value, theme: { algorithm: dark.value ? theme.darkAlgorithm : theme.defaultAlgorithm, token: { colorPrimary: '#3064c7' } },
}, { default: () => h(KeepAlive, { max: 2 }, { default: () => visible.value ? h(DocumentWorkbench, {
  key: 'workbench',
  modelValue: model.value, documentApi: documentApi.value, canUndo: studio ? studio.editor.canUndo : !!history.value.length, canRedo: studio ? studio.editor.canRedo : !!redo.value.length, ribbonVariant: variant.value,
  'onUpdate:modelValue': update,
  onUndo: () => { if (documentApi.value) documentApi.value.undo(); else if (history.value.length) { redo.value.push(copy(model.value)); model.value = history.value.pop() } },
  onRedo: () => { if (documentApi.value) documentApi.value.redo(); else if (redo.value.length) { history.value.push(copy(model.value)); model.value = redo.value.pop() } },
}) : h('div', { key: 'away', class: 'fixture-away' }, '工作台暂时隐藏') }) }) })
app.provide(OfficeLocaleKey, locale); app.provide(OfficeDirectionKey, direction); app.provide(OfficeTranslatorKey, translate)
app.mount('#app')
</script></body></html>`

let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'document-home-command-fixture',
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
  page = await browser.newPage({ viewport: { width: 1440, height: 900 } })
  page.on('pageerror', error => errors.push(error.message))
  await page.goto(baseURL)
  await expect(page.getByRole('textbox', { name: '文档画布编辑区', exact: true })).toBeVisible()
}, { timeout: 60_000 })
afterEach(async () => { try { assert.deepEqual(errors, []) } finally { await page?.close() } })

const snapshot = () => page.evaluate(() => window.documentHomeFixture.snapshot())
const history = () => page.evaluate(() => window.documentHomeFixture.history())
const colors = field => page.evaluate(key => window.documentHomeFixture.snapshot().blocks[0].runs.map(run => run[key]), field)
const ribbonItem = id => page.locator(`.ribbon-visible-groups [data-ribbon-item="${id}"]`)
const bridge = () => page.getByLabel('文档文字输入', { exact: true })
async function selectAll() { await bridge().focus(); await bridge().press('Control+A') }
async function openColor(field = 'color', keyboard = false) {
  const arrow = ribbonItem(field === 'color' ? 'doc-text-color' : 'doc-highlight').locator('.document-color-arrow button')
  if (keyboard) { await arrow.focus(); await arrow.press('ArrowDown') } else await arrow.click()
  const panel = page.locator('.document-color-panel:visible')
  await expect(panel).toBeVisible()
  return panel
}
async function openCustom(panel) {
  await panel.locator('.document-color-more button').click()
  const input = panel.getByRole('textbox', { name: 'HEX', exact: true })
  await expect(input).toBeFocused()
  return input
}
async function clickParagraph(index) {
  const point = await page.evaluate(value => window.documentHomeFixture.paragraphPoint(value), index)
  await page.mouse.click(point.x, point.y)
}
async function openParagraphSettings(index = 0) {
  const point = await page.evaluate(value => window.documentHomeFixture.paragraphPoint(value), index)
  await page.mouse.click(point.x, point.y, { button: 'right' })
  await page.getByRole('menuitem', { name: '段落设置', exact: true }).click()
  const panel = page.locator('.paragraph-settings-panel')
  await expect(panel).toBeVisible()
  return panel
}
async function homeItem(id) {
  const item = page.locator('.ribbon-visible-groups [data-ribbon-item="' + id + '"], .ribbon-grid-group:not(.is-hidden) [data-ribbon-item="' + id + '"]')
  if (await item.count()) return item
  let overflow = page.locator('.ribbon-overflow-panel:visible')
  if (!await overflow.isVisible()) await page.locator('.ribbon-more-button:visible, .ribbon-grid-more-button:visible').click()
  overflow = page.locator('.ribbon-overflow-panel:visible')
  const result = overflow.locator('[data-ribbon-item="' + id + '"]')
  await expect(result).toBeVisible()
  return result
}
async function chooseOption(select, label) {
  await select.click()
  await page.locator('.ant-select-dropdown:visible').getByText(label, { exact: true }).click()
}

test('45 preset colors apply to the Canvas selection, undo, redo and reopen', async () => {
  const before = await snapshot()
  await selectAll()
  const panel = await openColor('color', true)
  assert.deepEqual(await panel.locator('.document-color-preset-row').evaluateAll(rows => rows.map(row => row.querySelectorAll('button').length)), [9, 9, 9, 9, 9])
  await expect(panel.getByRole('button', { name: '#414657', exact: true })).toHaveAttribute('aria-pressed', 'true')
  await panel.getByRole('button', { name: '#F05252', exact: true }).click()
  await expect.poll(() => colors('color')).toEqual(['#f05252'])
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  await expect(bridge()).toBeFocused()
  await ribbonItem('doc-undo').getByRole('button').click()
  assert.deepEqual(await snapshot(), before)
  await ribbonItem('doc-redo').getByRole('button').click()
  await expect.poll(() => colors('color')).toEqual(['#f05252'])
  const saved = JSON.stringify(await snapshot())
  await page.evaluate(value => window.documentHomeFixture.reopen(value), saved)
  await expect.poll(() => colors('color')).toEqual(['#f05252'])
  assert.deepEqual(await page.evaluate(() => window.documentHomeFixture.htmlColors()), [{ color: '#f05252' }])
})

test('custom color changes stay draft until confirmation and cancel restores both value and More focus', async () => {
  await selectAll()
  const before = await snapshot()
  const panel = await openColor()
  let input = await openCustom(panel)
  const initialHex = await input.inputValue()
  await input.fill('abcdef')
  assert.deepEqual(await snapshot(), before)
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  await panel.getByRole('button', { name: /取\s*消/ }).click()
  await expect(panel.locator('.document-color-more button')).toBeFocused()
  await expect(panel.locator('.document-color-presets')).toBeVisible()
  input = await openCustom(panel)
  await expect(input).toHaveValue(initialHex)
  await input.fill('abcdef')
  await panel.getByRole('button', { name: /确\s*定/ }).click()
  await expect.poll(() => colors('color')).toEqual(['#abcdef'])
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  await expect(bridge()).toBeFocused()
})

test('Enter rejects an incomplete HEX draft and commits a valid value as a single undo step', async () => {
  await selectAll()
  const panel = await openColor()
  const input = await openCustom(panel)
  await input.fill('12'); await input.press('Enter')
  await expect(panel.locator('.document-color-custom')).toBeVisible()
  await expect(panel).toContainText('请输入有效的颜色值')
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  await input.fill('0a1b2c'); await input.press('Enter')
  await expect.poll(() => colors('color')).toEqual(['#0a1b2c'])
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  await ribbonItem('doc-undo').getByRole('button').click()
  await expect.poll(() => colors('color')).toEqual(['#414657'])
})

test('custom Escape returns to presets, while a second Escape closes and focuses the arrow', async () => {
  await selectAll()
  const before = await snapshot()
  const panel = await openColor()
  let input = await openCustom(panel)
  const initialHex = await input.inputValue()
  await input.fill('112233'); await input.press('Escape')
  await expect(panel.locator('.document-color-more button')).toBeFocused()
  input = await openCustom(panel)
  await expect(input).toHaveValue(initialHex)
  await input.press('Escape')
  await panel.locator('.document-color-more button').press('Escape')
  await expect(page.locator('.document-color-panel')).toBeHidden()
  await expect(ribbonItem('doc-text-color').locator('.document-color-arrow button')).toBeFocused()
  assert.deepEqual(await snapshot(), before)
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
})

test('IME Enter and Escape keep browser defaults and leave the custom draft open', async () => {
  await selectAll()
  const panel = await openColor()
  const input = await openCustom(panel)
  await input.fill('112233')
  for (const options of [
    { key: 'Enter', isComposing: true }, { key: 'Escape', isComposing: true },
    { key: 'Enter', keyCode: 229 }, { key: 'Escape', keyCode: 229 },
  ]) {
    const prevented = await input.evaluate((element, info) => {
      const event = new KeyboardEvent('keydown', { ...info, bubbles: true, cancelable: true })
      element.dispatchEvent(event)
      return event.defaultPrevented
    }, options)
    assert.equal(prevented, false)
    await expect(panel.locator('.document-color-custom')).toBeVisible()
    assert.deepEqual(await history(), { undo: 0, redo: 0 })
  }
  await input.press('Enter')
  await expect.poll(() => colors('color')).toEqual(['#112233'])
})

test('picker keyboard changes and RGB channels remain draft until Enter', async () => {
  await selectAll()
  const before = await snapshot()
  const panel = await openColor()
  const input = await openCustom(panel)
  await input.fill('336699')
  const slider = panel.getByRole('slider').first()
  await slider.focus(); await slider.press('ArrowRight')
  await expect(input).not.toHaveValue('336699')
  assert.deepEqual(await snapshot(), before)
  await panel.getByRole('combobox', { name: '颜色格式', exact: true }).click()
  await panel.getByRole('option', { name: 'RGB', exact: true }).click()
  const red = panel.getByRole('spinbutton', { name: 'R', exact: true })
  await red.fill('10')
  await panel.getByRole('spinbutton', { name: 'G', exact: true }).fill('20')
  await panel.getByRole('spinbutton', { name: 'B', exact: true }).fill('30')
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  await red.press('Enter')
  await expect.poll(() => colors('color')).toEqual(['#0a141e'])
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
})

test('the main button remembers an applied color while the panel reads the current and automatic selection', async () => {
  await selectAll()
  let panel = await openColor()
  await panel.getByRole('button', { name: '#3F83F8', exact: true }).click()
  await page.evaluate(() => window.documentHomeFixture.setColor('#f05252'))
  panel = await openColor()
  await expect(panel.getByRole('button', { name: '#F05252', exact: true })).toHaveAttribute('aria-pressed', 'true')
  const main = ribbonItem('doc-text-color').locator('.document-color-main')
  assert.equal(await main.locator('.document-color-bar').evaluate(element => getComputedStyle(element).backgroundColor), 'rgb(63, 131, 248)')
  await panel.getByRole('button', { name: '#F05252', exact: true }).press('Escape')
  await main.click()
  await expect.poll(() => colors('color')).toEqual(['#3f83f8'])
  panel = await openColor()
  await panel.getByRole('button', { name: '自动文字颜色', exact: true }).click()
  await expect.poll(() => colors('color')).toEqual(['auto'])
  assert.equal(await main.locator('.document-color-bar').evaluate(element => getComputedStyle(element).backgroundColor), 'rgb(63, 131, 248)')
  panel = await openColor()
  await expect(panel.locator('.document-color-state')).toHaveText('自动文字颜色')
  await expect(panel.locator('.document-color-preset[aria-pressed="true"]')).toHaveCount(0)
  await panel.getByRole('button', { name: '自动文字颜色', exact: true }).press('Escape')
  await main.click()
  await expect.poll(() => colors('color')).toEqual(['#3f83f8'])
})

test('No Background clears the model without forgetting the last highlight, while Reset applies the theme color', async () => {
  await selectAll()
  let panel = await openColor('highlight')
  await panel.getByRole('button', { name: '#FAC815', exact: true }).click()
  panel = await openColor('highlight')
  await panel.getByRole('button', { name: '无背景色', exact: true }).click()
  await expect.poll(() => colors('highlight')).toEqual([undefined])
  const main = ribbonItem('doc-highlight').locator('.document-color-main')
  assert.equal(await main.locator('.document-color-bar').evaluate(element => getComputedStyle(element).backgroundColor), 'rgb(250, 200, 21)')
  await main.click()
  await expect.poll(() => colors('highlight')).toEqual(['#fac815'])
  panel = await openColor('highlight')
  await panel.getByRole('button', { name: '重置颜色', exact: true }).click()
  await expect.poll(() => colors('highlight')).toEqual(['#3064c7'])
  await ribbonItem('doc-undo').getByRole('button').click()
  await expect.poll(() => colors('highlight')).toEqual(['#fac815'])
})

test('reapplying the current color creates no content history and pending caret colors are isolated by document', async () => {
  await selectAll()
  const before = await snapshot()
  let panel = await openColor()
  await panel.getByRole('button', { name: '#414657', exact: true }).click()
  assert.deepEqual(await snapshot(), before)
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  await bridge().focus(); await bridge().press('Control+End')
  panel = await openColor()
  await panel.getByRole('button', { name: '#F05252', exact: true }).click()
  assert.deepEqual(await snapshot(), before)
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  await page.evaluate(() => window.documentHomeFixture.newDocument())
  await bridge().focus(); await bridge().press('Control+End'); await bridge().press('!')
  await expect.poll(async () => (await snapshot()).blocks[0].runs.map(run => run.text).join('')).toBe('新文档!')
  await expect.poll(() => colors('color')).toEqual(['#414657'])
})

test('permission changes discard an open draft, use disabled theme tokens and block subsequent palette commands', async () => {
  await selectAll()
  const panel = await openColor()
  const input = await openCustom(panel)
  await input.fill('112233')
  await page.evaluate(() => window.documentHomeFixture.denyEdit())
  await expect(page.locator('.document-color-panel')).toBeHidden()
  const main = ribbonItem('doc-text-color').locator('.document-color-main')
  const arrow = ribbonItem('doc-text-color').locator('.document-color-arrow button')
  await expect(main).toBeDisabled()
  await expect(arrow).toBeDisabled()
  for (const dark of [false, true]) {
    await page.evaluate(value => window.documentHomeFixture.appearance(value, 'zh-CN'), dark)
    for (const button of [main, arrow]) {
      await expect(button).toHaveCSS('border-width', '0px')
      await expect.poll(() => button.evaluate(element => getComputedStyle(element).color === getComputedStyle(element).getPropertyValue('--ribbon-disabled').trim())).toBe(true)
      await expect.poll(() => button.evaluate(element => getComputedStyle(element).backgroundColor)).toBe('rgba(0, 0, 0, 0)')
    }
  }
  await main.evaluate(element => element.dispatchEvent(new MouseEvent('click', { bubbles: true })))
  assert.deepEqual(await colors('color'), ['#414657'])
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
})

test('a portaled palette updates theme and RTL direction while retaining its usable presets', async () => {
  await selectAll()
  const panel = await openColor()
  const preset = panel.getByRole('button', { name: '#414657', exact: true })
  assert.equal(await preset.getAttribute('title'), null)
  await preset.focus()
  await expect(page.getByRole('tooltip').getByText('#414657', { exact: true })).toBeVisible()
  assert.deepEqual(await preset.evaluate(element => {
    const bounds = element.getBoundingClientRect()
    return [bounds.width, bounds.height]
  }), [32, 32])
  const background = await panel.evaluate(element => getComputedStyle(element).getPropertyValue('--ribbon-background').trim())
  await page.evaluate(() => window.documentHomeFixture.appearance(true, 'ar-SA'))
  await expect(panel).toHaveAttribute('dir', 'rtl')
  await expect.poll(() => panel.evaluate(element => getComputedStyle(element).getPropertyValue('--ribbon-background').trim())).not.toBe(background)
  await expect(panel.locator('.document-color-preset')).toHaveCount(45)
  await panel.getByRole('button', { name: '#C81E1E', exact: true }).click()
  await expect.poll(() => colors('color')).toEqual(['#c81e1e'])
  await expect(bridge()).toBeFocused()
})

for (const variant of ['classic', 'grid']) test(`${variant} Docs color arrows stay centered in both directions and hover only within the button`, async () => {
  await selectAll()
  const before = await snapshot()
  await page.evaluate(value => window.documentHomeFixture.variant(value), variant)
  const controls = page.locator('.ribbon-visible-groups .document-color-control, .ribbon-grid-group:not(.is-hidden) .document-color-control')
  await expect(controls).toHaveCount(2)
  for (const appearance of [{ dark: false, locale: 'zh-CN' }, { dark: true, locale: 'ar-SA' }]) {
    await page.evaluate(value => window.documentHomeFixture.appearance(value.dark, value.locale), appearance)
    for (const control of await controls.all()) {
      for (const button of [control.locator('.document-color-main'), control.locator('.document-color-arrow button')]) {
        await expect(button).toHaveCSS('border-width', '0px')
        await button.hover()
        await expect(button).toHaveCSS('border-width', '0px')
        await page.mouse.down()
        assert.equal(await button.evaluate(element => element.matches(':active')), true)
        await expect(button).toHaveCSS('border-width', '0px')
        await page.mouse.move(1400, 800)
        await page.mouse.up()
      }
      const arrow = control.locator('.document-color-arrow button')
      const metrics = await arrow.evaluate(element => {
        const button = element.getBoundingClientRect(), svg = element.querySelector('svg').getBoundingClientRect()
        return { width: button.width, height: button.height, svgWidth: svg.width, svgHeight: svg.height,
          offsetX: svg.x + svg.width / 2 - button.x - button.width / 2,
          offsetY: svg.y + svg.height / 2 - button.y - button.height / 2 }
      })
      assert.equal(metrics.width, 16)
      assert.equal(metrics.height, variant === 'classic' ? 28 : 24)
      assert.deepEqual([metrics.svgWidth, metrics.svgHeight], [12, 12])
      assert.ok(Math.abs(metrics.offsetX) <= 0.05 && Math.abs(metrics.offsetY) <= 0.05)
      await arrow.hover()
      await expect.poll(() => arrow.evaluate(element => getComputedStyle(element).backgroundColor === getComputedStyle(element).getPropertyValue('--ribbon-hover-background').trim())).toBe(true)
      await page.mouse.move(1400, 800)
      assert.equal(await arrow.evaluate(element => element.matches(':hover')), false)
      await expect.poll(() => arrow.evaluate(element => getComputedStyle(element).backgroundColor)).toBe('rgba(0, 0, 0, 0)')
    }
  }
  assert.deepEqual(await snapshot(), before)
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
})

test('grid overflow retains custom color editing and applies a single document command', async () => {
  await selectAll()
  await page.evaluate(() => window.documentHomeFixture.variant('grid'))
  await page.setViewportSize({ width: 260, height: 900 })
  await page.locator('.ribbon-more-button').click()
  const overflow = page.locator('.ribbon-overflow-panel:visible')
  const color = overflow.locator('[data-ribbon-item="doc-text-color"]')
  await expect(color).toBeVisible()
  const arrow = color.locator('.document-color-arrow button')
  const bounds = await arrow.evaluate(element => {
    const button = element.getBoundingClientRect(), svg = element.querySelector('svg').getBoundingClientRect()
    return { width: button.width, height: button.height, offsetX: svg.x + svg.width / 2 - button.x - button.width / 2,
      offsetY: svg.y + svg.height / 2 - button.y - button.height / 2 }
  })
  assert.deepEqual([bounds.width, bounds.height], [16, 28])
  assert.ok(Math.abs(bounds.offsetX) <= 0.05 && Math.abs(bounds.offsetY) <= 0.05)
  await expect(arrow).toHaveCSS('border-width', '0px')
  await arrow.hover()
  await expect(arrow).toHaveCSS('border-width', '0px')
  await page.mouse.down()
  assert.equal(await arrow.evaluate(element => element.matches(':active')), true)
  await expect(arrow).toHaveCSS('border-width', '0px')
  await page.mouse.up()
  const panel = page.locator('.document-color-panel:visible')
  const input = await openCustom(panel)
  await input.fill('aabbcc'); await input.press('Enter')
  await expect.poll(() => colors('color')).toEqual(['#aabbcc'])
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  await expect(bridge()).toBeFocused()
})

test('restart at zero and continue numbering share the popup command path with external history and reopen', async () => {
  await page.evaluate(() => window.documentHomeFixture.numbering())
  const markers = () => page.evaluate(() => window.documentHomeFixture.markers())
  await expect.poll(markers).toEqual(['1.', '2.', '1.', '2.'])
  await clickParagraph(3)
  const arrow = ribbonItem('doc-numbered').locator('.ribbon-split-arrow')
  await arrow.click()
  let picker = page.locator('.doc-list-picker:visible')
  await picker.getByRole('spinbutton', { name: '起始编号', exact: true }).fill('0')
  await picker.getByRole('button', { name: '重新编号', exact: true }).click()
  await expect.poll(markers).toEqual(['1.', '2.', '0.', '1.'])
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  await ribbonItem('doc-undo').getByRole('button').click()
  await expect.poll(markers).toEqual(['1.', '2.', '1.', '2.'])
  await ribbonItem('doc-redo').getByRole('button').click()
  await expect.poll(markers).toEqual(['1.', '2.', '0.', '1.'])
  await arrow.click()
  picker = page.locator('.doc-list-picker:visible')
  const continuation = picker.getByRole('button', { name: '继续前一编号', exact: true })
  await expect(continuation).toBeEnabled(); await continuation.click()
  await expect.poll(markers).toEqual(['1.', '2.', '3.', '4.'])
  assert.deepEqual(await history(), { undo: 2, redo: 0 })
  await ribbonItem('doc-undo').getByRole('button').click()
  await expect.poll(markers).toEqual(['1.', '2.', '0.', '1.'])
  await ribbonItem('doc-redo').getByRole('button').click()
  const saved = JSON.stringify(await snapshot())
  await page.evaluate(value => window.documentHomeFixture.reopen(value), saved)
  await expect.poll(markers).toEqual(['1.', '2.', '3.', '4.'])
})

test.each([['classic', false, 'numbered'], ['grid', true, 'numbered'], ['classic', false, 'bullet'], ['grid', true, 'bullet']])('mixed list presets follow the real gallery, toggle, history and internal reopen in %s, dark %s, %s', async (variant, dark, kind) => {
  await page.goto(baseURL + '?api=1')
  await page.evaluate(({ variant, dark, kind }) => {
    window.documentHomeFixture.mixedLists(kind)
    window.documentHomeFixture.variant(variant)
    window.documentHomeFixture.appearance(dark, 'zh-CN')
  }, { variant, dark, kind })
  await selectAll()
  const item = await homeItem(kind === 'numbered' ? 'doc-numbered' : 'doc-bullet')
  const main = item.locator('.ribbon-button:not(.ribbon-split-arrow)')
  await expect(main).toHaveAttribute('aria-pressed', 'false')
  const before = await snapshot(), beforeHistory = await history(), beforeCommands = await page.evaluate(() => window.documentHomeFixture.apiCommands())
  const textIdentity = before.blocks.map(block => [block.id, block.runs.map(run => [run.id, run.text])])
  await item.locator('.ribbon-split-arrow').click()
  const picker = page.locator('.doc-list-picker:visible')
  await expect(picker.locator('.doc-list-card')).toHaveCount(6)
  await expect(picker.locator('.doc-list-card[aria-pressed="true"]')).toHaveCount(0)
  assert.deepEqual(await snapshot(), before)
  assert.deepEqual(await history(), beforeHistory)
  assert.deepEqual(await page.evaluate(() => window.documentHomeFixture.apiCommands()), beforeCommands)
  const preset = kind === 'numbered' ? 'ORDER_LIST_2' : 'BULLET_LIST_2'
  await picker.locator('[data-list-preset="' + preset + '"]').click()
  await expect.poll(async () => (await snapshot()).blocks.map(block => block.list.preset)).toEqual([preset, preset])
  await expect(main).toHaveAttribute('aria-pressed', 'true')
  await expect(bridge()).toBeFocused()
  const edited = await snapshot()
  assert.deepEqual(edited.blocks.map(block => [block.id, block.runs.map(run => [run.id, run.text])]), textIdentity)
  assert.deepEqual(await history(), { undo: beforeHistory.undo + 1, redo: 0 })
  await bridge().press('Control+z')
  await expect.poll(async () => (await snapshot()).blocks).toEqual(before.blocks)
  await expect(main).toHaveAttribute('aria-pressed', 'false')
  await bridge().press('Control+y')
  await expect.poll(async () => (await snapshot()).blocks).toEqual(edited.blocks)
  await expect(main).toHaveAttribute('aria-pressed', 'true')
  await page.evaluate(value => window.documentHomeFixture.reopen(value), JSON.stringify(edited))
  await expect.poll(async () => (await snapshot()).blocks).toEqual(edited.blocks)
  await selectAll()
  await expect(main).toHaveAttribute('aria-pressed', 'true')
  await main.click()
  await expect.poll(async () => (await snapshot()).blocks.map(block => block.type)).toEqual(['paragraph', 'paragraph'])
  await expect(main).toHaveAttribute('aria-pressed', 'false')
  const removed = await snapshot()
  assert.deepEqual(removed.blocks.map(block => [block.id, block.runs.map(run => [run.id, run.text])]), textIdentity)
  assert.equal(removed.blocks.every(block => block.list === undefined), true)
  assert.deepEqual(await history(), { undo: beforeHistory.undo + 2, redo: 0 })
  await expect(bridge()).toBeFocused()
})

test('an open numbering popup disables every preset and numbering action when edit permission changes', async () => {
  await page.goto(baseURL + '?api=1')
  await page.evaluate(() => window.documentHomeFixture.numbering())
  await clickParagraph(3)
  await ribbonItem('doc-numbered').locator('.ribbon-split-arrow').click()
  const picker = page.locator('.doc-list-picker').filter({ has: page.locator('.doc-numbering-actions') })
  await expect(picker).toBeVisible()
  const cards = picker.locator('.doc-list-card')
  await expect(cards).toHaveCount(6)
  assert.deepEqual(await cards.evaluateAll(elements => elements.map(element => {
    const bounds = element.getBoundingClientRect()
    return [bounds.width, bounds.height, element.disabled]
  })), Array.from({ length: 6 }, () => [72, 80, false]))
  const compound = picker.locator('[data-list-preset="ORDER_LIST_2"]')
  await expect(compound).toHaveAttribute('aria-label', /1\.1\.1\./)
  await expect(compound.locator('.doc-list-row > span').nth(2)).toHaveText('1.1.1.')
  const previewGeometry = () => compound.evaluate(element => {
    const button = element.getBoundingClientRect()
    const rows = [...element.querySelectorAll('.doc-list-row')].map(row => {
      const marker = row.firstElementChild
      const markerBounds = marker.getBoundingClientRect()
      const line = row.lastElementChild.getBoundingClientRect()
      return { markerWidth: markerBounds.width, markerRight: markerBounds.right,
        clipped: marker.scrollWidth > marker.clientWidth, flexShrink: getComputedStyle(marker).flexShrink,
        gap: parseFloat(getComputedStyle(row).columnGap), lineLeft: line.left, lineRight: line.right,
        rowRight: row.getBoundingClientRect().right }
    })
    return { width: button.width, height: button.height, rows }
  })
  const beforeHover = await previewGeometry()
  assert.equal(beforeHover.rows[2].clipped, true)
  for (const row of beforeHover.rows) {
    assert.equal(row.markerWidth, 12)
    assert.equal(row.flexShrink, '0')
    assert.ok(row.markerRight + row.gap <= row.lineLeft + 0.05)
    assert.ok(row.lineRight <= row.rowRight + 0.05)
  }
  await compound.hover()
  assert.deepEqual(await previewGeometry(), beforeHover)
  await expect(picker.getByRole('button', { name: '继续前一编号', exact: true })).toBeEnabled()
  await page.evaluate(() => window.documentHomeFixture.denyEdit())
  await expect(picker).toBeHidden()
  for (const button of await picker.locator('button').all()) await expect(button).toBeDisabled()
  await expect(picker.locator('input')).toBeDisabled()
  const before = await snapshot()
  for (const card of await cards.all()) await card.evaluate(element => element.click())
  assert.deepEqual(await snapshot(), before)
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  assert.deepEqual(await page.evaluate(() => window.documentHomeFixture.apiCommands()), [])
})

test('replacing an imported cached label keeps comment and control anchors through public undo, redo and reopen', async () => {
  await page.goto(baseURL + '?api=1')
  await page.evaluate(() => window.documentHomeFixture.cachedNumberingMarker())
  await clickParagraph(0)
  const before = await snapshot()
  await ribbonItem('doc-numbered').locator('.ribbon-split-arrow').click()
  await page.locator('.doc-list-picker:visible [data-list-preset="ORDER_LIST"]').click()
  await expect.poll(async () => (await snapshot()).blocks[0].runs.map(run => run.text).join('')).toBe('原始正文')
  await expect(bridge()).toBeFocused()
  const after = await snapshot(), id = after.blocks[0].id
  assert.equal(after.blocks[0].metadata.officeNumberingMarker, undefined)
  assert.deepEqual(after.comments.review.anchor, [{ blockId: id, startOffset: 0, endOffset: 4, text: '原始正文' }])
  assert.equal(after.comments.review.messages[0].text, '保留批注')
  assert.deepEqual(after.contentControls.map(control => [control.startOffset, control.endOffset]), [[0, 4]])
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  assert.deepEqual((await page.evaluate(() => window.documentHomeFixture.apiCommands())).map(event => event.id), ['doc.command.change-list-type'])
  await ribbonItem('doc-undo').getByRole('button').click()
  assert.deepEqual(await snapshot(), before)
  await ribbonItem('doc-redo').getByRole('button').click()
  assert.deepEqual(await snapshot(), after)
  await page.evaluate(value => window.documentHomeFixture.reopen(value), JSON.stringify(after))
  const reopened = await snapshot()
  assert.deepEqual(reopened.comments.review.anchor, after.comments.review.anchor)
  assert.deepEqual(reopened.contentControls, after.contentControls)
  assert.equal(reopened.blocks[0].runs.map(run => run.text).join(''), '原始正文')
  await expect(page.locator('.document-surface canvas')).toHaveCount(1)
})

test('deleting an annotated cached label retains its collapsed comment through the Ribbon, public history and reopen', async () => {
  await page.goto(baseURL + '?api=1')
  await page.evaluate(() => window.documentHomeFixture.cachedNumberingMarker({ labelOnly: true }))
  await clickParagraph(0)
  const before = await snapshot()
  assert.deepEqual(before.comments.review.anchor.map(anchor => [anchor.startOffset, anchor.endOffset, anchor.text]), [[1, 3, '0]']])
  await ribbonItem('doc-numbered').locator('.ribbon-split-arrow').click()
  await page.locator('.doc-list-picker:visible [data-list-preset="ORDER_LIST"]').click()
  await expect.poll(async () => (await snapshot()).blocks[0].runs.map(run => run.text).join('')).toBe('原始正文')
  const after = await snapshot()
  assert.deepEqual(after.comments.review.anchor, [{ blockId: after.blocks[0].id, startOffset: 0, endOffset: 0, text: '' }])
  assert.equal(after.comments.review.referenceText, '')
  assert.equal(after.comments.review.messages[0].text, '保留批注')
  assert.deepEqual(after.contentControls.map(control => [control.startOffset, control.endOffset]), [[0, 4]])
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  await expect(bridge()).toBeFocused()
  await ribbonItem('doc-undo').getByRole('button').click()
  assert.deepEqual(await snapshot(), before)
  await ribbonItem('doc-redo').getByRole('button').click()
  assert.deepEqual(await snapshot(), after)
  await page.evaluate(value => window.documentHomeFixture.reopen(value), JSON.stringify(after))
  assert.deepEqual((await snapshot()).comments.review, after.comments.review)
  await clickParagraph(0)
  await expect(bridge()).toBeFocused()
})

for (const headings of [false, true]) test(`consecutive ${headings ? 'numbered headings' : 'paragraphs'} restart through the public history and retain their own levels after reopen`, async () => {
  await page.goto(baseURL + '?api=1')
  await page.evaluate(value => window.documentHomeFixture.consecutiveNumbering({ headings: value }), headings)
  const before = await snapshot()
  const markers = () => page.evaluate(() => window.documentHomeFixture.markers())
  await expect.poll(markers).toEqual(['0.', '0.1.', '2.', '2.3.'])
  await clickParagraph(1)
  await ribbonItem('doc-numbered').locator('.ribbon-split-arrow').click()
  const picker = page.locator('.doc-list-picker:visible')
  const start = picker.getByRole('spinbutton', { name: '起始编号', exact: true })
  await expect(start).toBeEnabled()
  await start.fill('7')
  await picker.getByRole('button', { name: '重新编号', exact: true }).click()
  await expect.poll(markers).toEqual(['0.', '0.7.', '8.', '8.9.'])
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  await expect(bridge()).toBeFocused()
  assert.deepEqual((await snapshot()).blocks.map(block => [block.type, block.level, block.metadata.numberingLevel, block.runs]), before.blocks.map(block => [block.type, block.level, block.metadata.numberingLevel, block.runs]))
  const commands = await page.evaluate(() => window.documentHomeFixture.apiCommands())
  assert.deepEqual(commands.map(event => [event.id, event.kind]), [['doc.command.restart-list-numbering', 'document']])
  await ribbonItem('doc-undo').getByRole('button').click()
  assert.deepEqual(await snapshot(), before)
  await ribbonItem('doc-redo').getByRole('button').click()
  await expect.poll(markers).toEqual(['0.', '0.7.', '8.', '8.9.'])
  const saved = JSON.stringify(await snapshot())
  await page.evaluate(value => window.documentHomeFixture.reopen(value), saved)
  await expect.poll(markers).toEqual(['0.', '0.7.', '8.', '8.9.'])
  await expect(page.locator('.document-surface canvas')).toHaveCount(1)
})

test('shared consecutive groups disable redundant continuation across imported instance IDs', async () => {
  await page.goto(baseURL + '?api=1')
  await page.evaluate(() => window.documentHomeFixture.consecutiveNumbering({ linked: true }))
  await clickParagraph(1)
  const before = await snapshot()
  await ribbonItem('doc-numbered').locator('.ribbon-split-arrow').click()
  const picker = page.locator('.doc-list-picker:visible')
  await expect(picker.getByRole('button', { name: '继续前一编号', exact: true })).toBeDisabled()
  assert.deepEqual(await snapshot(), before)
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
})

test('numbering InputNumber guards IME confirmation and follows ConfigProvider theme changes', async () => {
  await page.goto(baseURL + '?api=1')
  await page.evaluate(() => window.documentHomeFixture.consecutiveNumbering())
  await clickParagraph(1)
  await ribbonItem('doc-numbered').locator('.ribbon-split-arrow').click()
  const picker = page.locator('.doc-list-picker:visible')
  const start = picker.getByRole('spinbutton', { name: '起始编号', exact: true })
  const submit = picker.getByRole('button', { name: '重新编号', exact: true })
  await start.fill('7')
  await start.dispatchEvent('compositionstart')
  await expect(submit).toBeDisabled()
  const cancelled = await start.evaluate(element => !element.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', keyCode: 229, isComposing: true, bubbles: true, cancelable: true })))
  assert.equal(cancelled, true)
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  await start.dispatchEvent('compositionend')
  await expect(submit).toBeEnabled()
  const cancelledWithoutStart = await start.evaluate(element => !element.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', keyCode: 229, bubbles: true, cancelable: true })))
  assert.equal(cancelledWithoutStart, true)
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  const background = await start.evaluate(element => getComputedStyle(element.closest('.ant-input-number')).backgroundColor)
  await page.evaluate(() => window.documentHomeFixture.appearance(true, 'ar-SA'))
  await expect.poll(() => start.evaluate(element => getComputedStyle(element.closest('.ant-input-number')).backgroundColor)).not.toBe(background)
  await page.evaluate(() => window.documentHomeFixture.appearance(false, 'zh-CN'))
  await start.press('Enter')
  await expect.poll(() => page.evaluate(() => window.documentHomeFixture.markers())).toEqual(['0.', '0.7.', '8.', '8.9.'])
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  await expect(bridge()).toBeFocused()
})

for (const reverse of [false, true]) test(`paragraph settings preserve a ${reverse ? 'reverse' : 'forward'} mixed selection on right-click and share public history`, async () => {
  await page.goto(baseURL + '?api=1')
  await page.evaluate(() => window.documentHomeFixture.paragraphs([
    { indentStart: 10, indentEnd: 3, wordWrap: true },
    { indentStart: 20, indentEnd: 7, wordWrap: false },
  ]))
  await bridge().focus()
  if (reverse) { await bridge().press('Control+End'); await bridge().press('Control+Shift+Home') }
  else await bridge().press('Control+A')
  const before = await snapshot()
  const panel = await openParagraphSettings(1)
  const left = panel.getByRole('spinbutton', { name: /左.*\(px\)/ })
  await expect(left).toHaveValue('')
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  await left.fill('25'); await left.press('Tab')
  await expect.poll(async () => (await snapshot()).blocks.map(block => block.style.indentStart)).toEqual([25, 25])
  assert.deepEqual((await snapshot()).blocks.map(block => [block.style.indentEnd, block.style.wordWrap]), [[3, true], [7, false]])
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  const commands = await page.evaluate(() => window.documentHomeFixture.apiCommands())
  assert.deepEqual(commands.map(event => [event.id, event.kind]), [['doc.command.format-paragraph', 'document']])
  await ribbonItem('doc-undo').getByRole('button').click()
  assert.deepEqual(await snapshot(), before)
  await ribbonItem('doc-redo').getByRole('button').click()
  await expect.poll(async () => (await snapshot()).blocks.map(block => block.style.indentStart)).toEqual([25, 25])
  const background = await panel.evaluate(element => getComputedStyle(element).backgroundColor)
  await page.evaluate(() => window.documentHomeFixture.appearance(true, 'ar-SA'))
  await expect.poll(() => panel.evaluate(element => getComputedStyle(element).backgroundColor)).not.toBe(background)
  await page.evaluate(() => window.documentHomeFixture.appearance(false, 'zh-CN'))
  const saved = JSON.stringify(await snapshot())
  await page.evaluate(value => window.documentHomeFixture.reopen(value), saved)
  await expect.poll(async () => (await snapshot()).blocks.map(block => block.style.indentStart)).toEqual([25, 25])
  await panel.getByRole('button', { name: /关闭/ }).click()
  await expect(panel).toBeHidden()
  await expect(bridge()).toBeFocused()
  await expect(page.locator('.document-surface canvas')).toHaveCount(1)
})

test('right-clicking outside the selected paragraph retargets settings to the clicked paragraph', async () => {
  await page.goto(baseURL + '?api=1')
  await page.evaluate(() => window.documentHomeFixture.paragraphs([{ indentStart: 10 }, { indentStart: 20 }, { indentStart: 30 }]))
  await clickParagraph(0)
  await bridge().press('Home'); await bridge().press('Shift+End')
  const panel = await openParagraphSettings(2)
  const left = panel.getByRole('spinbutton', { name: /左.*\(px\)/ })
  await expect(left).toHaveValue('30')
  await left.fill('44'); await left.press('Tab')
  await expect.poll(async () => (await snapshot()).blocks.map(block => block.style.indentStart)).toEqual([10, 20, 44])
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
})

test.each([false, true])('the divider button disables across a CODE block in a backwards %s selection and becomes usable outside it', async backwards => {
  await page.goto(baseURL + '?api=1')
  await page.evaluate(() => window.documentHomeFixture.codeRange())
  await bridge().focus()
  await bridge().press(backwards ? 'Control+End' : 'Control+Home')
  await bridge().press(backwards ? 'Control+Shift+Home' : 'Control+Shift+End')
  const before = await snapshot()
  await page.getByRole('tab', { name: '插入', exact: true }).click()
  const divider = page.locator('.ribbon-visible-groups [data-ribbon-item="doc-divider"] button')
  await expect(divider).toBeDisabled()
  assert.deepEqual(await snapshot(), before)
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  await bridge().focus()
  await bridge().press('Control+End')
  await expect(divider).toBeEnabled()
  await divider.click()
  await expect.poll(async () => (await snapshot()).blocks.filter(block => block.type === 'divider').length).toBe(1)
  assert.deepEqual((await snapshot()).blocks.find(block => block.type === 'code'), before.blocks.find(block => block.type === 'code'))
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  await expect(bridge()).toBeFocused()
  await page.getByRole('tab', { name: '开始', exact: true }).click()
  await (await homeItem('doc-undo')).getByRole('button').click()
  await expect.poll(snapshot).toEqual(before)
  assert.deepEqual(await history(), { undo: 0, redo: 1 })
})

test.each([['classic', false], ['grid', true]])('body composition disables Home character formats and clear until it commits in %s, dark %s', async (variant, dark) => {
  await page.goto(baseURL + '?api=1')
  await page.evaluate(({ variant, dark }) => {
    window.documentHomeFixture.clearFormatting()
    window.documentHomeFixture.variant(variant)
    window.documentHomeFixture.appearance(dark, 'zh-CN')
  }, { variant, dark })
  await selectAll()
  const before = await snapshot(), previousCommands = await page.evaluate(() => window.documentHomeFixture.apiCommands())
  await bridge().evaluate(element => {
    element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' }))
    element.value = '拼'
    element.dispatchEvent(new CompositionEvent('compositionupdate', { bubbles: true, data: '拼' }))
    element.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertCompositionText', data: '拼', isComposing: true }))
    element.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: ',', code: 'Comma', ctrlKey: true }))
    element.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: '.', code: 'Period', metaKey: true }))
    for (const [code, key] of [['KeyX', 'X'], ['KeyL', 'L'], ['KeyE', 'E'], ['KeyR', 'R'], ['KeyJ', 'J'], ['Digit7', '&'], ['Digit8', '*']]) {
      for (const modifier of ['ctrlKey', 'metaKey']) element.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key, code, shiftKey: true, [modifier]: true }))
    }
  })
  const groups = page.locator('.ribbon-visible-groups, .ribbon-grid-group:not(.is-hidden)')
  for (const id of ['doc-bold', 'doc-italic', 'doc-underline', 'doc-strike', 'doc-subscript', 'doc-superscript', 'doc-clear-format']) {
    await expect(groups.locator('[data-ribbon-item="' + id + '"]').getByRole('button')).toBeDisabled()
  }
  for (const id of ['doc-font-family', 'doc-font-size']) await expect(groups.locator('[data-ribbon-item="' + id + '"]').getByRole('combobox')).toBeDisabled()
  for (const id of ['doc-align', 'doc-numbered', 'doc-bullet']) {
    for (const button of await groups.locator('[data-ribbon-item="' + id + '"]').getByRole('button').all()) await expect(button).toBeDisabled()
  }
  for (const id of ['doc-indent-less', 'doc-indent-more', 'doc-paragraph-spacing']) await expect((await homeItem(id)).getByRole('button')).toBeDisabled()
  await expect((await homeItem('doc-line-height')).getByRole('combobox')).toBeDisabled()
  for (const id of ['doc-text-color', 'doc-highlight']) for (const selector of ['.document-color-main', '.document-color-arrow button']) {
    await expect(groups.locator('[data-ribbon-item="' + id + '"]').locator(selector)).toBeDisabled()
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
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  assert.deepEqual(await page.evaluate(() => window.documentHomeFixture.apiCommands()), previousCommands)
  await expect(bridge()).toBeFocused()
  await bridge().evaluate(element => {
    element.value = '中文'
    element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '中文' }))
    element.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: '中文' }))
  })
  await expect.poll(async () => (await snapshot()).blocks[0].runs.map(run => run.text).join('')).toBe('中文')
  assert.equal((await snapshot()).blocks[0].runs.every(run => run.bold === false), true)
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  for (const id of ['doc-bold', 'doc-italic', 'doc-underline', 'doc-strike', 'doc-subscript', 'doc-superscript', 'doc-clear-format']) {
    await expect(groups.locator('[data-ribbon-item="' + id + '"]').getByRole('button')).toBeEnabled()
  }
  await expect(bridge()).toBeFocused()
  await selectAll()
  await groups.locator('[data-ribbon-item="doc-subscript"]').getByRole('button').click()
  await expect.poll(async () => (await snapshot()).blocks[0].runs.every(run => run.verticalAlign === 'subscript')).toBe(true)
  await expect(bridge()).toBeFocused()
  await groups.locator('[data-ribbon-item="doc-clear-format"]').getByRole('button').click()
  await expect.poll(async () => (await snapshot()).blocks[0].runs.every(run => run.verticalAlign === undefined && run.italic === false)).toBe(true)
  assert.deepEqual(await history(), { undo: 3, redo: 0 })
  await expect(bridge()).toBeFocused()
})

test.each([['classic', false], ['grid', true]])('Home character controls follow KeepAlive activation before input focus returns in %s, dark %s', async (variant, dark) => {
  await page.goto(baseURL + '?api=1')
  await page.evaluate(({ variant, dark }) => {
    window.documentHomeFixture.variant(variant)
    window.documentHomeFixture.appearance(dark, 'zh-CN')
  }, { variant, dark })
  await selectAll()
  const characterItems = ['doc-font-family', 'doc-font-size', 'doc-bold', 'doc-italic', 'doc-underline', 'doc-strike', 'doc-subscript', 'doc-superscript', 'doc-text-color', 'doc-highlight']
  const controls = []
  for (const id of characterItems) {
    const item = await homeItem(id)
    if (id === 'doc-font-family' || id === 'doc-font-size') {
      controls.push(await item.getByRole('combobox').elementHandle())
      continue
    }
    if (id === 'doc-text-color' || id === 'doc-highlight') {
      controls.push(await item.locator('.document-color-main').elementHandle(), await item.locator('.document-color-arrow button').elementHandle())
      continue
    }
    controls.push(await item.getByRole('button').elementHandle())
  }
  assert.equal(controls.length, 12)
  const ribbon = await page.locator('[data-testid="workbench-ribbon"]').elementHandle()
  const before = await snapshot(), beforeHistory = await history(), beforeCommands = await page.evaluate(() => window.documentHomeFixture.apiCommands())
  await page.evaluate(() => window.documentHomeFixture.show(false))
  await expect(page.locator('.fixture-away')).toBeVisible()
  await expect.poll(() => ribbon.evaluate((element, ids) => element.__vueParentComponent.props.tabs.flatMap(tab => tab.groups).flatMap(group => group.items)
    .filter(item => ids.includes(item.id)).map(item => item.disabled), characterItems)).toEqual(Array(characterItems.length).fill(true))
  for (const control of controls) await expect.poll(() => control.evaluate(element => element.disabled)).toBe(true)
  await expect.poll(() => ribbon.evaluate(element => element.__vueParentComponent.props.tabs.flatMap(tab => tab.groups).flatMap(group => group.items)
    .find(item => item.id === 'doc-divider').disabled)).toBe(true)
  await ribbon.evaluate(element => {
    for (const patch of [{ fontFamily: 'Georgia' }, { fontSize: 20 }, { bold: true }, { italic: true }, { underline: true }, { strike: true },
      { verticalAlign: 'subscript' }, { verticalAlign: 'superscript' }, { color: '#123456' }, { highlight: '#aabbcc' }]) {
      element.__vueParentComponent.emit('execute', 'doc.command.format-text', { patch })
    }
    for (const id of ['doc.command.horizontal-line', 'doc.command.insert-divider', 'doc.command.insert-horizontal-line-bellow']) {
      element.__vueParentComponent.emit('execute', id)
    }
  })
  assert.deepEqual(await snapshot(), before)
  assert.deepEqual(await history(), beforeHistory)
  assert.deepEqual(await page.evaluate(() => window.documentHomeFixture.apiCommands()), beforeCommands)
  await page.evaluate(() => window.documentHomeFixture.show(true))
  await expect(page.locator('.fixture-away')).toBeHidden()
  await expect.poll(() => ribbon.evaluate((element, ids) => element.__vueParentComponent.props.tabs.flatMap(tab => tab.groups).flatMap(group => group.items)
    .filter(item => ids.includes(item.id)).map(item => item.disabled), characterItems)).toEqual(Array(characterItems.length).fill(false))
  for (const control of controls) await expect.poll(() => control.evaluate(element => element.disabled)).toBe(false)
  await expect.poll(() => ribbon.evaluate(element => Boolean(element.__vueParentComponent.props.tabs.flatMap(tab => tab.groups).flatMap(group => group.items)
    .find(item => item.id === 'doc-divider').disabled))).toBe(false)
  assert.deepEqual(await snapshot(), before)
  assert.deepEqual(await history(), beforeHistory)
  assert.deepEqual(await page.evaluate(() => window.documentHomeFixture.apiCommands()), beforeCommands)
  await selectAll()
  const size = (await homeItem('doc-font-size')).getByRole('combobox')
  await size.fill('20')
  await size.press('Enter')
  await expect.poll(async () => (await snapshot()).blocks[0].runs.every(run => run.fontSize === 20)).toBe(true)
  const edited = await snapshot()
  assert.deepEqual(await history(), { undo: beforeHistory.undo + 1, redo: 0 })
  await expect(bridge()).toBeFocused()
  await bridge().press('Control+z')
  await expect.poll(async () => (await snapshot()).blocks).toEqual(before.blocks)
  await bridge().press('Control+y')
  await expect.poll(async () => (await snapshot()).blocks).toEqual(edited.blocks)
  await page.evaluate(value => window.documentHomeFixture.reopen(value), JSON.stringify(edited))
  await expect.poll(async () => (await snapshot()).blocks).toEqual(edited.blocks)
})

test.each([['classic', false], ['grid', true]])('fixed strike, alignment and list shortcuts share body commands, history and internal reopen in %s, dark %s', async (variant, dark) => {
  await page.goto(baseURL + '?api=1')
  await page.evaluate(({ variant, dark }) => {
    window.documentHomeFixture.paragraphs([{ align: 'center' }, { align: 'center' }])
    window.documentHomeFixture.variant(variant)
    window.documentHomeFixture.appearance(dark, 'zh-CN')
  }, { variant, dark })
  await selectAll()
  const before = await snapshot()
  for (const [shortcut, strike] of [['Control+Shift+X', true], ['Meta+Shift+X', false]]) {
    await bridge().press(shortcut)
    await expect.poll(async () => (await snapshot()).blocks.every(block => block.runs.every(run => run.strike === strike))).toBe(true)
    await expect(bridge()).toBeFocused()
  }
  for (const [shortcut, align] of [['Control+Shift+R', 'right'], ['Meta+Shift+L', 'left'], ['Control+Shift+E', 'center'], ['Meta+Shift+J', 'justify']]) {
    await bridge().press(shortcut)
    await expect.poll(async () => (await snapshot()).blocks.map(block => block.style.align)).toEqual([align, align])
    await expect(bridge()).toBeFocused()
  }
  const aligned = await snapshot()
  await bridge().press('Control+Shift+J')
  assert.deepEqual(await snapshot(), aligned)
  assert.deepEqual(await history(), { undo: 6, redo: 0 })
  for (const [shortcut, type, preset] of [['Control+Shift+7', 'numbered', 'ORDER_LIST'], ['Meta+Shift+7', 'paragraph', undefined], ['Meta+Shift+8', 'bullet', 'BULLET_LIST'], ['Control+Shift+8', 'paragraph', undefined]]) {
    await bridge().press(shortcut)
    await expect.poll(async () => (await snapshot()).blocks.map(block => block.type)).toEqual([type, type])
    const blocks = (await snapshot()).blocks
    if (preset) {
      assert.equal(blocks.every(block => block.list.preset === preset), true)
      assert.equal(new Set(blocks.map(block => block.list.id)).size, 1)
    }
    if (!preset) assert.equal(blocks.every(block => block.list === undefined), true)
    if (type === 'numbered') assert.deepEqual(await page.evaluate(() => window.documentHomeFixture.markers()), ['1.', '2.'])
    await expect(bridge()).toBeFocused()
  }
  assert.deepEqual((await snapshot()).blocks.map(block => block.runs), before.blocks.map(block => block.runs))
  assert.deepEqual(await history(), { undo: 10, redo: 0 })
  const formatted = await snapshot(), previousCommands = await page.evaluate(() => window.documentHomeFixture.apiCommands())
  await bridge().evaluate(element => {
    for (const [code, key] of [['KeyX', 'X'], ['KeyL', 'L'], ['KeyE', 'E'], ['KeyR', 'R'], ['KeyJ', 'J'], ['Digit7', '&'], ['Digit8', '*']]) {
      element.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key, code, ctrlKey: true }))
      element.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key, code, ctrlKey: true, shiftKey: true, altKey: true }))
    }
  })
  await page.getByRole('tab', { name: '开始', exact: true }).focus()
  await bridge().evaluate(element => {
    for (const [code, key] of [['KeyX', 'X'], ['KeyL', 'L'], ['KeyE', 'E'], ['KeyR', 'R'], ['KeyJ', 'J'], ['Digit7', '&'], ['Digit8', '*']]) element.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key, code, metaKey: true, shiftKey: true }))
  })
  assert.deepEqual(await snapshot(), formatted)
  assert.deepEqual(await history(), { undo: 10, redo: 0 })
  assert.deepEqual(await page.evaluate(() => window.documentHomeFixture.apiCommands()), previousCommands)
  await bridge().focus()
  await bridge().press('Control+z')
  await expect.poll(async () => (await snapshot()).blocks.map(block => block.type)).toEqual(['bullet', 'bullet'])
  await bridge().press('Control+y')
  await expect.poll(async () => (await snapshot()).blocks.map(block => block.type)).toEqual(['paragraph', 'paragraph'])
  await bridge().press('Meta+z')
  const saved = await snapshot()
  assert.deepEqual(await history(), { undo: 9, redo: 1 })
  await page.evaluate(value => window.documentHomeFixture.reopen(value), JSON.stringify(saved))
  assert.deepEqual((await snapshot()).blocks, saved.blocks)
  assert.deepEqual((await snapshot()).defaultStyle, before.defaultStyle)
  await page.evaluate(() => window.documentHomeFixture.denyEdit())
  const readonly = await snapshot(), readonlyHistory = await history(), readonlyCommands = await page.evaluate(() => window.documentHomeFixture.apiCommands())
  await bridge().focus()
  await bridge().evaluate(element => {
    for (const [code, key] of [['KeyX', 'X'], ['KeyL', 'L'], ['KeyE', 'E'], ['KeyR', 'R'], ['KeyJ', 'J'], ['Digit7', '&'], ['Digit8', '*']]) element.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key, code, ctrlKey: true, shiftKey: true }))
  })
  assert.deepEqual(await snapshot(), readonly)
  assert.deepEqual(await history(), readonlyHistory)
  assert.deepEqual(await page.evaluate(() => window.documentHomeFixture.apiCommands()), readonlyCommands)
})

test.each([['classic', false], ['grid', true]])('fixed subscript and superscript shortcuts use the body command and typing history in %s, dark %s', async (variant, dark) => {
  await page.goto(baseURL + '?api=1')
  await page.evaluate(({ variant, dark }) => {
    window.documentHomeFixture.clearFormatting()
    window.documentHomeFixture.variant(variant)
    window.documentHomeFixture.appearance(dark, 'zh-CN')
  }, { variant, dark })
  await selectAll()
  const before = await snapshot()
  for (const [shortcut, verticalAlign] of [['Control+Comma', 'subscript'], ['Meta+Period', 'superscript'], ['Control+Period', undefined], ['Meta+Comma', 'subscript'], ['Control+Comma', undefined]]) {
    await bridge().press(shortcut)
    await expect.poll(async () => (await snapshot()).blocks[0].runs.every(run => run.verticalAlign === verticalAlign)).toBe(true)
    await expect(bridge()).toBeFocused()
  }
  assert.deepEqual(await history(), { undo: 5, redo: 0 })
  const formatted = await snapshot(), previousCommands = await page.evaluate(() => window.documentHomeFixture.apiCommands())
  assert.equal(formatted.blocks[0].runs.every(run => run.fontSize === before.blocks[0].runs[0].fontSize && run.italic === true && run.href === 'https://example.org'), true)
  await bridge().evaluate(element => {
    element.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: '<', code: 'Comma', ctrlKey: true, shiftKey: true }))
    element.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: '.', code: 'Period', ctrlKey: true, altKey: true }))
  })
  await page.getByRole('tab', { name: '开始', exact: true }).focus()
  await bridge().evaluate(element => element.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: '.', code: 'Period', ctrlKey: true })))
  assert.deepEqual(await snapshot(), formatted)
  assert.deepEqual(await page.evaluate(() => window.documentHomeFixture.apiCommands()), previousCommands)
  assert.deepEqual(await history(), { undo: 5, redo: 0 })
  await bridge().focus()
  await bridge().press('Control+z')
  await expect.poll(async () => (await snapshot()).blocks[0].runs.every(run => run.verticalAlign === 'subscript')).toBe(true)
  await bridge().press('Control+y')
  await expect.poll(async () => (await snapshot()).blocks[0].runs.every(run => run.verticalAlign === undefined)).toBe(true)
  await bridge().press('ArrowRight')
  const collapsed = await snapshot()
  await bridge().press('Control+Comma')
  assert.deepEqual(await snapshot(), collapsed)
  assert.deepEqual(await history(), { undo: 5, redo: 0 })
  await bridge().press('!')
  await expect.poll(async () => (await snapshot()).blocks[0].runs.map(run => run.text).join('')).toBe('Hello world!')
  assert.equal((await snapshot()).blocks[0].runs.some(run => run.text === '!' && run.verticalAlign === 'subscript'), true)
  assert.deepEqual(await history(), { undo: 6, redo: 0 })
  await page.evaluate(() => window.documentHomeFixture.denyEdit())
  await expect(page.locator('.ribbon-visible-groups, .ribbon-grid-group:not(.is-hidden)').locator('[data-ribbon-item="doc-subscript"]').getByRole('button')).toBeDisabled()
  const readonly = await snapshot(), readonlyHistory = await history(), readonlyCommands = await page.evaluate(() => window.documentHomeFixture.apiCommands())
  await bridge().focus()
  await bridge().press('Control+Period')
  assert.deepEqual(await snapshot(), readonly)
  assert.deepEqual(await history(), readonlyHistory)
  assert.deepEqual(await page.evaluate(() => window.documentHomeFixture.apiCommands()), readonlyCommands)
})

for (const variant of ['classic', 'grid']) test(`${variant} clear formatting uses document defaults, public undo/redo and cleared typing marks`, async () => {
  await page.goto(baseURL + '?api=1')
  await page.evaluate(value => { window.documentHomeFixture.clearFormatting(); window.documentHomeFixture.variant(value) }, variant)
  await bridge().focus(); await bridge().press('Control+End')
  const before = await snapshot()
  await (await homeItem('doc-bold')).getByRole('button').click()
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  const clear = (await homeItem('doc-clear-format')).getByRole('button')
  await expect(clear).toBeEnabled(); await expect(clear).toHaveClass(/ant-btn/)
  const foreground = await clear.evaluate(element => getComputedStyle(element).color)
  await page.evaluate(() => window.documentHomeFixture.appearance(true, 'zh-CN'))
  await expect.poll(() => clear.evaluate(element => getComputedStyle(element).color)).not.toBe(foreground)
  await page.evaluate(() => window.documentHomeFixture.appearance(false, 'zh-CN'))
  await clear.click()
  await expect(bridge()).toBeFocused()
  await expect.poll(async () => (await snapshot()).blocks[0].type).toBe('paragraph')
  const after = await snapshot(), runs = after.blocks[0].runs
  assert.deepEqual(runs.map(run => [run.text, run.bold, run.italic, run.fontFamily, run.fontSize, run.highlight, run.verticalAlign]), [
    ['Hello ', false, true, 'Arial', 26, '#FFFF00', 'superscript'],
    ['world', false, false, 'Georgia', 15, undefined, undefined],
  ])
  assert.deepEqual(after.blocks[0].style, after.defaultStyle.paragraph)
  assert.equal(after.blocks[0].list, undefined); assert.equal(runs[1].href, 'https://example.org')
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  assert.deepEqual((await page.evaluate(() => window.documentHomeFixture.apiCommands())).map(event => event.id), ['doc.command.clear-formatting'])
  await bridge().press('x')
  await expect.poll(async () => (await snapshot()).blocks[0].runs.map(run => run.text).join('')).toBe('Hello x')
  const typedAfterClear = (await snapshot()).blocks[0].runs.find(run => run.text === 'x')
  assert.deepEqual([typedAfterClear.bold, typedAfterClear.italic, typedAfterClear.fontFamily, typedAfterClear.fontSize, typedAfterClear.highlight], [false, false, 'Georgia', 15, undefined])
  await (await homeItem('doc-undo')).getByRole('button').click(); assert.deepEqual(await snapshot(), after)
  await (await homeItem('doc-undo')).getByRole('button').click(); assert.deepEqual(await snapshot(), before)
  await (await homeItem('doc-redo')).getByRole('button').click(); assert.deepEqual(await snapshot(), after)
  await page.evaluate(value => window.documentHomeFixture.reopen(value), JSON.stringify(after))
  assert.deepEqual((await snapshot()).defaultStyle, after.defaultStyle)
  // Reopening resets UI selection, so select the cleared word again by keyboard.
  await bridge().focus(); await bridge().press('Control+End'); await bridge().press('Control+Shift+ArrowLeft')
  await bridge().press('x')
  await expect.poll(async () => (await snapshot()).blocks[0].runs.map(run => run.text).join('')).toBe('Hello x')
  const inserted = (await snapshot()).blocks[0].runs.find(run => run.text === 'x')
  assert.deepEqual([inserted.bold, inserted.italic, inserted.fontFamily, inserted.fontSize, inserted.highlight], [false, false, 'Georgia', 15, undefined])
  await expect(page.locator('.document-surface canvas')).toHaveCount(1)
})

test('a cancelled clear preserves pending typing format and restores the Canvas input focus', async () => {
  await page.goto(baseURL + '?api=1')
  await page.evaluate(() => window.documentHomeFixture.clearFormatting())
  await bridge().focus(); await bridge().press('Control+End')
  const before = await snapshot()
  await (await homeItem('doc-bold')).getByRole('button').click()
  await page.evaluate(() => window.documentHomeFixture.cancelClear(true))
  await (await homeItem('doc-clear-format')).getByRole('button').click()
  await expect(bridge()).toBeFocused()
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), { undo: 0, redo: 0 })
  assert.deepEqual(await page.evaluate(() => window.documentHomeFixture.apiCommands()), [])
  await page.evaluate(() => window.documentHomeFixture.cancelClear(false))
  await bridge().press('x')
  await expect.poll(async () => (await snapshot()).blocks[0].runs.map(run => run.text).join('')).toBe('Hello worldx')
  const inserted = (await snapshot()).blocks[0].runs.find(run => run.text === 'x')
  assert.equal(inserted.bold, true)
})

test('clear formatting disabled state follows actual edit permission without publishing a history entry', async () => {
  await page.goto(baseURL + '?api=1')
  await page.evaluate(() => window.documentHomeFixture.clearFormatting())
  await selectAll()
  const clear = (await homeItem('doc-clear-format')).getByRole('button')
  await expect(clear).toBeEnabled()
  await page.evaluate(() => window.documentHomeFixture.denyEdit())
  await expect(clear).toBeDisabled()
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  assert.deepEqual(await page.evaluate(() => window.documentHomeFixture.apiCommands()), [])
})
