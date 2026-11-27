import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

const fixturePath = '/__document-hyperlinks__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><style>body{margin:0}#app{height:100vh;min-width:0}</style></head><body><div id="app"></div><script type="module">
import { createApp, h, KeepAlive, ref, shallowRef, watch } from 'vue'
import { ConfigProvider, theme } from 'antdv-next'
import DocumentWorkbench from '/packages/office-studio/src/components/DocumentWorkbench.vue'
import { OfficeStudio } from '/packages/office-studio/src/api/office-studio'
import { FDocument } from '/packages/office-studio/src/api/facade'
import { decodeDocumentBody } from '/packages/office-studio/src/api/document-body-codec'
import { createBlankDocument } from '/packages/office-studio/src/domain/model'
import { blocksFromHtml, createRun, normalizeDocument } from '/packages/office-studio/src/features/document/model/document-model'
import { documentHyperlinks } from '/packages/office-studio/src/features/document/services/document-hyperlinks'
import { resolveDocumentHyperlinkOperation } from '/packages/office-studio/src/api/document-public-body'
import { documentSegmentIdForHost } from '/packages/office-studio/src/api/document-segments'
import { documentModelForBlockHost } from '/packages/office-studio/src/features/document/services/document-block-host'
import { EMPTY_SECTION_HEADER_FOOTER } from '/packages/office-studio/src/features/document/services/section-header-footer'
import { documentStoryParagraphs } from '/packages/office-studio/src/features/document/model/document-story'
import { createTextMeasurer, caretRect, layoutDocument } from '/packages/office-studio/src/features/document/layout/document-layout'
import { sceneOffset } from '/packages/office-studio/src/features/document/rendering/document-scene'
const copy = value => JSON.parse(JSON.stringify(value))
const decoded = decodeDocumentBody({
  dataStream: 'Alpha Beta\\rBookmark target\\rHeading target\\r\\n',
  paragraphs: [{ startIndex: 10, paragraphId: 'link-text' }, { startIndex: 26, paragraphId: 'bookmark-target' },
    { startIndex: 41, paragraphId: 'heading-target', paragraphStyle: { headingId: 'heading-alias', namedStyleType: 3 } }],
  sectionBreaks: [{ startIndex: 42, sectionId: 'first' }],
  textRuns: [{ st: 0, ed: 6, ts: { bl: 1, cl: { rgb: '#883311' } } }, { st: 6, ed: 10, ts: { it: 1, ul: { s: 0 }, cl: { rgb: '#114488' } } }],
  customRanges: [{ rangeType: 3, rangeId: 'bookmark-range', startIndex: 11, endIndex: 25, properties: { bookmarkId: 'bookmark-alias' } }],
})
const initial = normalizeDocument({ ...createBlankDocument(), id: 'hyperlink-browser', name: '链接文档',
  blocks: decoded.content.blocks, metadata: decoded.content.metadata, contentControls: decoded.contentControls })
initial.sectionHeaderFooters = { first: { ...EMPTY_SECTION_HEADER_FOOTER, header: 'Header link',
  storyContent: { header: { runs: [createRun('Header ', { bold: true }), createRun('link', { italic: true })],
    paragraphs: documentStoryParagraphs('Header link') } } } }
let studio = new OfficeStudio({ document: initial, currentUserId: 'local-user' })
studio.activateDocument(initial.id)
const documentApi = shallowRef(new FDocument(studio, initial.id)), model = shallowRef(documentApi.value.getModel())
const dark = ref(false), variant = ref('classic'), visible = ref(true), commands = []
let stopWatching
function observe() {
  stopWatching?.()
  stopWatching = watch(() => studio.editor.document.model, value => { model.value = value }, { flush: 'sync' })
  studio.events.addEvent('CommandExecuted', event => { if (event.kind === 'document') commands.push(event.id) })
}
observe()
const selection = (start, end = start, blockId = 'link-text') => ({ anchor: { blockId, offset: start }, focus: { blockId, offset: end } })
function replace(next) {
  studio.editor.document.replace(normalizeDocument(next)); model.value = documentApi.value.getModel()
}
window.documentHyperlinkFixture = {
  snapshot: () => copy(model.value), commands: () => [...commands],
  links: (host) => copy(documentHyperlinks(host ? documentModelForBlockHost(model.value, host) : model.value)),
  history: () => { const unit = studio.editor.history.units[documentApi.value.id]; return { undo: unit.undo.length, redo: unit.redo.length } },
  select: (start, end = start, blockId = 'link-text') => documentApi.value.setInputSelection(selection(start, end, blockId)),
  selectStory: (start, end = start) => documentApi.value.setInputSelection(selection(0), { sectionId: 'first', story: 'header', start, end }),
  input: () => copy(documentApi.value.getInputState()),
  add: (target, start = 0, end = 10) => documentApi.value.executeCommand({ id: 'docs.command.add-hyper-link', payload: target }, selection(start, end)),
  publicLink: host => {
    const segmentId = host ? documentSegmentIdForHost(model.value, host) : ''
    const range = documentApi.value.getBody(segmentId).customRanges.find(range => range.rangeType === 0)
    return { unitId: documentApi.value.id, linkId: range.rangeId, startIndex: range.startIndex, endIndex: range.endIndex,
      ...(segmentId ? { segmentId } : {}) }
  },
  operation: (id, params) => document.querySelector('[data-testid="workbench-ribbon"]').__vueParentComponent.emit('execute', id, params),
  resolveOperation: (id, params) => {
    const resolved = resolveDocumentHyperlinkOperation(model.value, id, params)
    return resolved?.kind === 'link' ? copy({ id: resolved.link.id, label: resolved.link.label, selection: resolved.link.selection, host: resolved.host }) : resolved
  },
  importHtml: html => replace({ ...createBlankDocument(), id: documentApi.value.id, blocks: blocksFromHtml(html) }),
  publicReopen: () => {
    const body = documentApi.value.getBody(), decoded = decodeDocumentBody(body)
    replace({ ...model.value, blocks: decoded.content.blocks, metadata: decoded.content.metadata, contentControls: decoded.contentControls })
    return copy(body)
  },
  appearance: (isDark, layout) => { dark.value = isDark; variant.value = layout },
  rebind: () => { documentApi.value = new FDocument(studio, documentApi.value.id) },
  replaceSame: () => { const next = copy(model.value); next.name += ' changed'; replace(next) },
  replaceId: () => { const next = copy(initial); next.id = 'hyperlink-replacement'; studio.createDocument(next); studio.activateDocument(next.id); documentApi.value = new FDocument(studio, next.id); model.value = documentApi.value.getModel() },
  denyTail: () => { const next = copy(model.value); next.permissionRules = [{ id: 'locked-link-tail', scope: 'paragraph', targetId: 'link-text', action: 'edit', effect: 'deny', userIds: [] }]; replace(next) },
  denyView: () => { const next = copy(model.value); next.permissionRules = [{ id: 'hidden-link-source', scope: 'paragraph', targetId: 'link-text', action: 'view', effect: 'deny', userIds: [] }]; replace(next) },
  show: value => { visible.value = value },
  point: (offset = 1) => {
    const scroll = document.querySelector('.document-scroll'), bounds = scroll.getBoundingClientRect()
    const layout = layoutDocument(model.value, createTextMeasurer())
    const transform = sceneOffset(layout, { width: scroll.clientWidth, height: scroll.clientHeight, scrollLeft: scroll.scrollLeft, scrollTop: scroll.scrollTop, zoom: 100 })
    const caret = caretRect(layout, { blockId: 'link-text', offset })
    return { x: bounds.left + transform.x + caret.x + 1, y: bounds.top + transform.y + caret.y + caret.height / 2 }
  },
  saveReopen: () => {
    const saved = copy(documentApi.value.save()); stopWatching(); studio.dispose()
    studio = new OfficeStudio({ document: saved, currentUserId: 'local-user' }); studio.activateDocument(saved.id)
    documentApi.value = new FDocument(studio, saved.id); model.value = documentApi.value.getModel(); observe()
    return saved
  },
}
const app = createApp({ render: () => h(ConfigProvider, { theme: { algorithm: dark.value ? theme.darkAlgorithm : theme.defaultAlgorithm } }, {
  default: () => h(KeepAlive, { max: 2 }, { default: () => visible.value ? h(DocumentWorkbench, {
    key: 'editor', modelValue: model.value, documentApi: documentApi.value, ribbonVariant: variant.value,
    canUndo: studio.editor.canUndo, canRedo: studio.editor.canRedo,
    'onUpdate:modelValue': value => { model.value = value },
    onUndo: () => documentApi.value.undo(), onRedo: () => documentApi.value.redo(),
  }) : h('div', { key: 'away' }, '工作台暂时隐藏') }),
}) })
app.mount('#app')
window.addEventListener('pagehide', () => { stopWatching(); app.unmount(); studio.dispose() }, { once: true })
</script></body></html>`

let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'document-hyperlinks-fixture',
    configureServer(vite) {
      vite.middlewares.use(fixturePath, async (_request, response, next) => {
        try { response.setHeader('Content-Type', 'text/html; charset=utf-8'); response.end(await vite.transformIndexHtml(fixturePath, fixture)) } catch (error) { next(error) }
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
  errors = []; page = await browser.newPage({ viewport: { width: 1440, height: 900 } })
  page.on('pageerror', error => errors.push(error.message))
  await page.goto(baseURL)
  await expect(bridge()).toBeVisible()
}, { timeout: 60_000 })
afterEach(async () => { try { assert.deepEqual(errors, []) } finally { await page?.close() } })

const bridge = () => page.getByLabel('文档文字输入', { exact: true })
const dialog = () => page.getByRole('dialog', { name: /插入链接|编辑链接/ })
const links = () => page.evaluate(() => window.documentHyperlinkFixture.links())
const snapshot = () => page.evaluate(() => window.documentHyperlinkFixture.snapshot())
const history = () => page.evaluate(() => window.documentHyperlinkFixture.history())
const authoredRuns = model => model.blocks[0].runs.map(({ id, text, bold, italic, underline, color, fontFamily, fontSize }) => ({ id, text, bold, italic, underline, color, fontFamily, fontSize }))
async function linkButton() {
  await page.getByRole('tab', { name: '插入', exact: true }).click()
  const selector = '[data-ribbon-item="doc-link"]'
  const visible = page.locator('.ribbon-visible-groups ' + selector + ', .ribbon-grid-group:not(.is-hidden) ' + selector)
  if (await visible.count()) return visible.getByRole('button')
  await page.locator('.ribbon-more-button:visible, .ribbon-grid-more-button:visible').click()
  return page.locator('.ribbon-overflow-panel:visible ' + selector).getByRole('button')
}
async function openByShortcut() {
  await bridge().focus(); await bridge().press('Control+k')
  await expect(dialog()).toBeVisible()
}
async function submit(target) {
  await dialog().getByRole('textbox', { name: '链接地址', exact: true }).fill(target)
  await dialog().getByRole('button', { name: '应用', exact: true }).click()
}

for (const variant of ['classic', 'grid']) test(`${variant} toolbar and Ctrl+K edit one real hyperlink with history, style preservation and reopen`, async () => {
  await page.evaluate(value => { window.documentHyperlinkFixture.appearance(false, value); window.documentHyperlinkFixture.select(0, 10) }, variant)
  const original = authoredRuns(await snapshot())
  await (await linkButton()).click(); await expect(dialog()).toBeVisible()
  await submit('example.org/first')
  await expect(dialog()).toBeHidden(); await expect(bridge()).toBeFocused()
  assert.deepEqual(authoredRuns(await snapshot()), original)
  const first = (await links())[0]
  assert.equal(first.target, 'https://example.org/first'); assert.equal(first.label, 'Alpha Beta')
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  await page.evaluate(() => window.documentHyperlinkFixture.select(3))
  await openByShortcut()
  await expect(dialog().getByRole('textbox', { name: '显示文字', exact: true })).toHaveValue('Alpha Beta')
  await submit('person@example.org')
  await expect(dialog()).toBeHidden(); await expect(bridge()).toBeFocused()
  assert.equal((await links())[0].id, first.id); assert.equal((await links())[0].target, 'mailto:person@example.org')
  assert.deepEqual(authoredRuns(await snapshot()), original)
  await page.evaluate(() => window.documentHyperlinkFixture.saveReopen())
  assert.equal((await links())[0].id, first.id)
  await page.evaluate(() => window.documentHyperlinkFixture.select(2))
  await openByShortcut(); await dialog().getByRole('button', { name: '移除链接', exact: true }).click()
  await expect(dialog()).toBeHidden(); await expect(bridge()).toBeFocused()
  assert.deepEqual(await links(), []); assert.deepEqual(authoredRuns(await snapshot()), original)
  await page.getByRole('tab', { name: '开始', exact: true }).click()
  await page.getByRole('button', { name: '撤销', exact: true }).click()
  await expect.poll(links).toHaveLength(1)
  await page.getByRole('button', { name: '重做', exact: true }).click()
  await expect.poll(links).toEqual([])
  assert.deepEqual(await page.evaluate(() => window.documentHyperlinkFixture.commands()), [
    'docs.command.add-hyper-link', 'docs.command.update-hyper-link', 'docs.command.delete-hyper-link',
  ])
})

test('an edited label replaces the full entity and keeps a single undo transaction', async () => {
  await page.evaluate(() => { window.documentHyperlinkFixture.add('example.org'); window.documentHyperlinkFixture.select(3, 5) })
  await openByShortcut()
  await dialog().getByRole('textbox', { name: '显示文字', exact: true }).fill('新显示文字😀')
  await submit('new.example.org')
  await expect(dialog()).toBeHidden()
  assert.equal((await links())[0].label, '新显示文字😀')
  assert.equal((await snapshot()).blocks[0].runs.map(run => run.text).join(''), '新显示文字😀')
  assert.deepEqual(await history(), { undo: 2, redo: 0 })
  await page.getByRole('button', { name: '撤销', exact: true }).click()
  assert.equal((await links())[0].label, 'Alpha Beta'); assert.equal((await links())[0].target, 'https://example.org')
})

test('blank and unsafe targets and missing anchors leave the popup and the document unchanged', async () => {
  await page.evaluate(() => window.documentHyperlinkFixture.select(0, 10))
  const before = await snapshot()
  await openByShortcut()
  for (const target of ['', 'javascript:alert(1)', '#missing-target']) {
    await submit(target)
    await expect(dialog()).toBeVisible(); await expect(dialog().getByRole('alert')).toBeVisible()
    assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), { undo: 0, redo: 0 })
  }
  await dialog().getByRole('button', { name: '取消', exact: true }).click()
  await expect(bridge()).toBeFocused()
})

for (const target of ['#bookmark-alias', '#heading-alias']) test(`${target} navigates to the retained public target without creating an external page`, async () => {
  await page.evaluate(value => { window.documentHyperlinkFixture.add(value); window.documentHyperlinkFixture.select(2) }, target)
  const before = await history(), pages = page.context().pages().length
  const point = await page.evaluate(() => window.documentHyperlinkFixture.point())
  await page.keyboard.down('Control'); await page.mouse.click(point.x, point.y); await page.keyboard.up('Control')
  const blockId = target === '#bookmark-alias' ? 'bookmark-target' : 'heading-target'
  await expect.poll(() => page.evaluate(() => window.documentHyperlinkFixture.input().location.selection.focus.blockId)).toBe(blockId)
  assert.deepEqual(await history(), before); assert.equal(page.context().pages().length, pages)
  await expect(bridge()).toBeFocused()
})

test('an external Canvas link opens a real browser page through the shared operation', async () => {
  await page.evaluate(target => window.documentHyperlinkFixture.add(target), baseURL)
  const before = await history(), point = await page.evaluate(() => window.documentHyperlinkFixture.point())
  const opened = page.waitForEvent('popup')
  await page.keyboard.down('Control'); await page.mouse.click(point.x, point.y); await page.keyboard.up('Control')
  const popup = await opened
  await popup.waitForURL(baseURL)
  assert.equal(await popup.evaluate(() => window.opener === null), true)
  await popup.close(); assert.deepEqual(await history(), before)
})

for (const change of ['selection', 'same-document', 'rebind', 'new-document', 'deactivate']) test(`${change} invalidates an open link edit without writing the old range`, async () => {
  await page.evaluate(() => window.documentHyperlinkFixture.select(0, 10))
  await openByShortcut()
  await dialog().getByRole('textbox', { name: '链接地址', exact: true }).fill('example.org')
  await page.evaluate(value => {
    if (value === 'selection') { window.documentHyperlinkFixture.select(2); return }
    if (value === 'same-document') { window.documentHyperlinkFixture.replaceSame(); return }
    if (value === 'rebind') { window.documentHyperlinkFixture.rebind(); return }
    if (value === 'new-document') { window.documentHyperlinkFixture.replaceId(); return }
    window.documentHyperlinkFixture.show(false)
  }, change)
  await expect(dialog()).toBeHidden(); assert.deepEqual(await links(), [])
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
})

test('collapsed insertion and a protected existing link disable both toolbar and shortcut edits', async () => {
  await page.evaluate(() => window.documentHyperlinkFixture.select(0))
  await expect(await linkButton()).toBeDisabled()
  await bridge().focus(); await bridge().press('Control+k'); await expect(dialog()).toBeHidden()
  await page.evaluate(() => { window.documentHyperlinkFixture.add('example.org'); window.documentHyperlinkFixture.select(2); window.documentHyperlinkFixture.denyTail() })
  await expect(await linkButton()).toBeDisabled()
  await bridge().focus(); await bridge().press('Control+k'); await expect(dialog()).toBeHidden()
  assert.equal((await links())[0].target, 'https://example.org')
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
})

test('legacy header links use the existing host and survive a real snapshot reopen', async () => {
  await page.evaluate(() => window.documentHyperlinkFixture.selectStory(0, 11))
  await page.getByLabel('编辑页眉', { exact: true }).focus()
  await page.getByLabel('编辑页眉', { exact: true }).press('Control+k')
  await expect(dialog()).toBeVisible(); await submit('header.example.org')
  await expect(dialog()).toBeHidden()
  const host = { kind: 'story', sectionId: 'first', story: 'header' }
  const headerLinks = () => page.evaluate(value => window.documentHyperlinkFixture.links(value), host)
  assert.equal((await headerLinks())[0].label, 'Header link'); assert.deepEqual(await links(), [])
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  await page.evaluate(() => window.documentHyperlinkFixture.saveReopen())
  assert.equal((await headerLinks())[0].target, 'https://header.example.org')
})

test('link forms follow the theme and synthetic composition blocks confirmation and dismissal', async () => {
  await page.evaluate(() => window.documentHyperlinkFixture.select(0, 10))
  await openByShortcut()
  const light = await dialog().locator('.ant-modal-container').evaluate(element => getComputedStyle(element).backgroundColor)
  await page.evaluate(() => window.documentHyperlinkFixture.appearance(true, 'grid'))
  await expect.poll(() => dialog().locator('.ant-modal-container').evaluate(element => getComputedStyle(element).backgroundColor)).not.toBe(light)
  const input = dialog().getByRole('textbox', { name: '链接地址', exact: true })
  await input.fill('example.org'); await input.dispatchEvent('compositionstart')
  await expect(dialog().getByRole('button', { name: '应用', exact: true })).toBeDisabled()
  await input.press('Escape'); await expect(dialog()).toBeVisible()
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  await input.dispatchEvent('compositionend'); await input.press('Enter')
  await expect(dialog()).toBeHidden(); await expect(bridge()).toBeFocused()
})

for (const variant of [
  { layout: 'classic', dark: false, field: '链接地址' },
  { layout: 'grid', dark: true, field: '显示文字' },
]) test(`${variant.layout} delayed link opens preserve a composing ${variant.field} and its body target`, async () => {
  await page.evaluate(value => {
    window.documentHyperlinkFixture.appearance(value.dark, value.layout)
    window.documentHyperlinkFixture.add('example.org')
    window.documentHyperlinkFixture.selectStory(0, 11)
  }, variant)
  await page.getByLabel('编辑页眉', { exact: true }).focus()
  await page.getByLabel('编辑页眉', { exact: true }).press('Control+k')
  await expect(dialog()).toBeVisible(); await submit('header.example.org')
  await expect(dialog()).toBeHidden()
  const host = { kind: 'story', sectionId: 'first', story: 'header' }
  const requests = await page.evaluate(host => {
    const header = window.documentHyperlinkFixture.publicLink(host)
    window.documentHyperlinkFixture.select(3)
    return { body: window.documentHyperlinkFixture.publicLink(), header }
  }, host)
  await openByShortcut()
  const address = dialog().getByRole('textbox', { name: '链接地址', exact: true })
  const label = dialog().getByRole('textbox', { name: '显示文字', exact: true })
  await address.fill('draft.example.org'); await label.fill('Draft label 😀')
  const input = dialog().getByRole('textbox', { name: variant.field, exact: true })
  await input.focus(); await input.dispatchEvent('compositionstart')
  await expect(dialog().getByRole('button', { name: '应用', exact: true })).toBeDisabled()
  const before = await snapshot(), inputState = await page.evaluate(() => window.documentHyperlinkFixture.input())
  const headerLink = (await page.evaluate(host => window.documentHyperlinkFixture.links(host), host))[0]
  const count = await history(), commands = await page.evaluate(() => window.documentHyperlinkFixture.commands())
  assert.deepEqual(count, { undo: 2, redo: 0 })
  for (const operation of [
    { id: 'doc.operation.show-hyper-link-edit-popup' },
    { id: 'doc.operation.open-link' },
    { id: 'doc.operation.show-hyper-link-edit-popup', params: { link: requests.body } },
    { id: 'doc.operation.show-hyper-link-edit-popup', params: { link: requests.header } },
  ]) {
    await page.evaluate(request => window.documentHyperlinkFixture.operation(request.id, request.params), operation)
    await expect(dialog()).toBeVisible(); await expect(address).toHaveValue('draft.example.org')
    await expect(label).toHaveValue('Draft label 😀'); await expect(input).toBeFocused(); await expect(input).toBeEnabled()
    await expect(dialog().getByRole('button', { name: '应用', exact: true })).toBeDisabled()
    await expect(dialog().getByRole('button', { name: '取消', exact: true })).toBeDisabled()
    assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), count)
    assert.deepEqual(await page.evaluate(() => window.documentHyperlinkFixture.input()), inputState)
    assert.deepEqual(await page.evaluate(() => window.documentHyperlinkFixture.commands()), commands)
  }
  await input.dispatchEvent('compositionend')
  await expect(dialog().getByRole('button', { name: '应用', exact: true })).toBeEnabled()
  await dialog().getByRole('button', { name: '应用', exact: true }).click()
  await expect(dialog()).toBeHidden(); await expect(bridge()).toBeFocused()
  assert.equal((await links())[0].id, requests.body.linkId)
  assert.equal((await links())[0].target, 'https://draft.example.org')
  assert.equal((await links())[0].label, 'Draft label 😀')
  assert.deepEqual((await page.evaluate(host => window.documentHyperlinkFixture.links(host), host))[0], headerLink)
  assert.deepEqual(await history(), { undo: 3, redo: 0 })
  await page.getByRole('tab', { name: '开始', exact: true }).click()
  await page.getByRole('button', { name: '撤销', exact: true }).click()
  await expect.poll(snapshot).toEqual(before)
  await page.getByRole('button', { name: '重做', exact: true }).click()
  await expect.poll(async () => (await links())[0].label).toBe('Draft label 😀')
  assert.deepEqual(await history(), { undo: 3, redo: 0 })
  await page.evaluate(() => window.documentHyperlinkFixture.saveReopen())
  assert.equal((await links())[0].id, requests.body.linkId)
  assert.equal((await links())[0].target, 'https://draft.example.org')
  assert.equal((await links())[0].label, 'Draft label 😀')
  assert.deepEqual((await page.evaluate(host => window.documentHyperlinkFixture.links(host), host))[0], headerLink)
})

test('explicit public edit opens the complete link from an unrelated caret and commits one transaction', async () => {
  const request = await page.evaluate(() => {
    window.documentHyperlinkFixture.add('example.org')
    const request = window.documentHyperlinkFixture.publicLink()
    window.documentHyperlinkFixture.select(2, 2, 'bookmark-target')
    window.documentHyperlinkFixture.operation('doc.operation.show-hyper-link-edit-popup', { link: request })
    return request
  })
  await expect(dialog()).toBeVisible()
  await expect(dialog().getByRole('textbox', { name: '显示文字', exact: true })).toHaveValue('Alpha Beta')
  assert.deepEqual(await page.evaluate(() => window.documentHyperlinkFixture.input().location.selection), {
    anchor: { blockId: 'link-text', offset: 0 }, focus: { blockId: 'link-text', offset: 10 },
  })
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  await submit('updated.example.org')
  await expect(dialog()).toBeHidden(); await expect(bridge()).toBeFocused()
  assert.equal((await links())[0].id, request.linkId)
  assert.equal((await links())[0].target, 'https://updated.example.org')
  assert.deepEqual(await history(), { undo: 2, redo: 0 })
  await page.evaluate(() => window.documentHyperlinkFixture.saveReopen())
  assert.equal((await links())[0].id, request.linkId)
  assert.equal((await links())[0].target, 'https://updated.example.org')
})

test('invalid explicit operation identities and stale inclusive ranges do not fall back to the active link', async () => {
  const request = await page.evaluate(() => {
    window.documentHyperlinkFixture.add('example.org'); window.documentHyperlinkFixture.select(3)
    return window.documentHyperlinkFixture.publicLink()
  })
  const before = await snapshot(), input = await page.evaluate(() => window.documentHyperlinkFixture.input()), count = await history()
  for (const patch of [{ unitId: 'another-document' }, { linkId: 'missing-link' }, { segmentId: 'missing-story' },
    { startIndex: request.startIndex + 1 }, { endIndex: request.endIndex + 1 }]) {
    await page.evaluate(link => window.documentHyperlinkFixture.operation('doc.operation.show-hyper-link-edit-popup', { link }), { ...request, ...patch })
    await expect(dialog()).toBeHidden()
    assert.deepEqual(await snapshot(), before)
    assert.deepEqual(await page.evaluate(() => window.documentHyperlinkFixture.input()), input)
    assert.deepEqual(await history(), count)
  }
})

for (const target of ['#bookmark-alias', '#heading-alias']) test(`public ID click ${target} navigates independently of the active caret and Edit permission`, async () => {
  const request = await page.evaluate(target => {
    window.documentHyperlinkFixture.add(target); window.documentHyperlinkFixture.denyTail()
    window.documentHyperlinkFixture.select(1, 1, 'heading-target')
    return window.documentHyperlinkFixture.publicLink()
  }, target)
  const before = await snapshot(), count = await history(), pages = page.context().pages().length
  await page.evaluate(request => window.documentHyperlinkFixture.operation('doc.operation.click-hyper-link', {
    unitId: request.unitId, linkId: request.linkId,
  }), request)
  await expect.poll(() => page.evaluate(() => window.documentHyperlinkFixture.input().location.selection.focus.blockId))
    .toBe(target === '#bookmark-alias' ? 'bookmark-target' : 'heading-target')
  await expect(bridge()).toBeFocused()
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), count)
  assert.equal(page.context().pages().length, pages)
})

test('public header segment edit activates its real host from a body caret and preserves the link ID', async () => {
  await page.evaluate(() => window.documentHyperlinkFixture.selectStory(0, 11))
  await page.getByLabel('编辑页眉', { exact: true }).focus()
  await page.getByLabel('编辑页眉', { exact: true }).press('Control+k')
  await expect(dialog()).toBeVisible(); await submit('header.example.org'); await expect(dialog()).toBeHidden()
  const host = { kind: 'story', sectionId: 'first', story: 'header' }
  const request = await page.evaluate(host => {
    const request = window.documentHyperlinkFixture.publicLink(host)
    window.documentHyperlinkFixture.select(2)
    window.documentHyperlinkFixture.operation('doc.operation.show-hyper-link-edit-popup', { link: { ...request, segmentPage: 0 } })
    return request
  }, host)
  await expect(dialog()).toBeVisible()
  await expect(dialog().getByRole('textbox', { name: '显示文字', exact: true })).toHaveValue('Header link')
  assert.deepEqual(await page.evaluate(() => window.documentHyperlinkFixture.input().location.host), host)
  await submit('changed-header.example.org'); await expect(dialog()).toBeHidden()
  const headerLinks = await page.evaluate(host => window.documentHyperlinkFixture.links(host), host)
  assert.equal(headerLinks[0].id, request.linkId); assert.equal(headerLinks[0].target, 'https://changed-header.example.org')
  assert.deepEqual(await links(), []); assert.deepEqual(await history(), { undo: 2, redo: 0 })
  await page.evaluate(() => window.documentHyperlinkFixture.saveReopen())
  assert.equal((await page.evaluate(host => window.documentHyperlinkFixture.links(host), host))[0].id, request.linkId)
})

test('Canvas Ctrl click rejects a link whose actual source paragraph denies View', async () => {
  await page.evaluate(() => {
    window.documentHyperlinkFixture.add('#bookmark-alias')
    window.documentHyperlinkFixture.denyView(); window.documentHyperlinkFixture.select(2)
  })
  const before = await snapshot(), input = await page.evaluate(() => window.documentHyperlinkFixture.input()), count = await history()
  const point = await page.evaluate(() => window.documentHyperlinkFixture.point())
  await page.keyboard.down('Control'); await page.mouse.click(point.x, point.y); await page.keyboard.up('Control')
  assert.deepEqual(await page.evaluate(() => window.documentHyperlinkFixture.input()), input)
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), count)
  await expect(dialog()).toBeHidden()
})

test('real HTML styled hyperlink runs resolve one complete public entity and survive public body decode', async () => {
  await page.evaluate(() => window.documentHyperlinkFixture.importHtml('<p><a href="https://html.example"><b>A</b><i>B</i></a></p>'))
  const imported = await snapshot(), original = authoredRuns(imported), blockId = imported.blocks[0].id
  assert.equal(imported.blocks[0].runs.length, 2)
  assert.equal(imported.blocks[0].runs[0].bold, true); assert.equal(imported.blocks[0].runs[1].italic, true)
  const request = await page.evaluate(() => window.documentHyperlinkFixture.publicLink())
  assert.equal(request.linkId, 'hyperlink:' + imported.blocks[0].runs[0].id)
  assert.equal((await links()).length, 1); assert.equal((await links())[0].label, 'AB')
  for (const caret of [0, 1]) {
    const resolved = await page.evaluate(({ request, caret, blockId }) => {
      window.documentHyperlinkFixture.select(caret, caret, blockId)
      return window.documentHyperlinkFixture.resolveOperation('doc.operation.show-hyper-link-edit-popup', { link: request })
    }, { request, caret, blockId })
    assert.equal(resolved.id, request.linkId); assert.equal(resolved.label, 'AB')
    assert.deepEqual(resolved.selection, { anchor: { blockId, offset: 0 }, focus: { blockId, offset: 2 } })
  }
  assert.equal(await page.evaluate(linkId => window.documentHyperlinkFixture.resolveOperation('doc.operation.click-hyper-link', {
    unitId: window.documentHyperlinkFixture.snapshot().id, linkId,
  }), 'hyperlink:' + imported.blocks[0].runs[1].id), undefined)
  const body = await page.evaluate(() => window.documentHyperlinkFixture.publicReopen())
  assert.deepEqual(body.customRanges.filter(range => range.rangeType === 0).map(range => range.rangeId), [request.linkId])
  assert.deepEqual(authoredRuns(await snapshot()), original)
  assert.equal((await links())[0].id, request.linkId); assert.deepEqual(await history(), { undo: 0, redo: 0 })
})
