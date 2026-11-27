import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

// Deferred DOC-12 browser drafts mount the real Workbench, facade and antdv-next.
// They have not been run. Synthetic composition covers guards, not Windows IME acceptance.
const fixturePath = '/__document-note-workbench-test__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>Docs notes</title><style>body{margin:0}#app{height:100vh}</style></head><body><div id="app"></div><script type="module">
import { createApp, h, ref, shallowRef, watch } from 'vue'
import { ConfigProvider, theme } from 'antdv-next'
import { OfficeStudio } from '/packages/office-studio/src/api/office-studio'
import { FDocument } from '/packages/office-studio/src/api/facade'
import DocumentWorkbench from '/packages/office-studio/src/components/DocumentWorkbench.vue'
import { createBlankDocument } from '/packages/office-studio/src/domain/model'
import { createParagraph } from '/packages/office-studio/src/features/document/model/document-model'
import { projectDocumentStoryBlocks } from '/packages/office-studio/src/features/document/model/document-story-blocks'
const copy = value => JSON.parse(JSON.stringify(value))
const paragraph = (id, text = '') => { const block = createParagraph(text); block.id = id; block.runs[0].id = id + '-run'; return block }
const story = blocks => projectDocumentStoryBlocks({ blocks, runs: [], paragraphs: [] }, 'display').content
const image = (id, anchorBlockId) => ({ ...paragraph(id), type: 'image', runs: [], metadata: {
  image: { src: 'data:image/svg+xml,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="180" height="90"><rect width="180" height="90" fill="#ff00ff"/></svg>'), width: 180, height: 90, alt: 'Note picture' },
  anchorBlockId, anchorOffset: 0,
} })
const initial = createBlankDocument(); initial.id = 'note-browser'; initial.name = '注释工作台行为草稿'
initial.page = { width: 620, height: 540, marginTop: 45, marginBottom: 40, marginLeft: 48, marginRight: 48 }
initial.blocks = Array.from({ length: 50 }, (_, index) => paragraph('body-' + index, 'Body paragraph ' + index + ' keeps a real document scroll position.'))
initial.blocks.push(paragraph('outside', 'Outside observer selection'))
const reference = paragraph('note-reference', '1'); reference.runs[0].metadata = { note: { type: 'footnote', noteId: 'linked' } }; reference.runs[0].verticalAlign = 'superscript'
initial.blocks.splice(1, 0, reference)
// These empty paragraphs belong to the fixture's real source stories. Opening
// the Workbench must never fabricate or persist a temporary paragraph for an image.
initial.footnotes = {
  orphan: { id: 'orphan', text: '', referenceBlockId: '', referenceRunId: '', content: story([paragraph('orphan-anchor'), image('orphan-image', 'orphan-anchor')]) },
  linked: { id: 'linked', text: 'SECRET note content', referenceBlockId: reference.id, referenceRunId: reference.runs[0].id, content: story([paragraph('linked-note', 'SECRET note content')]) },
}
const separator = paragraph('separator-anchor'); separator.metadata = { noteSeparator: { version: 1, widthFraction: 1 / 3, color: '#1f2937' } }
const separatorContent = story([separator]); separatorContent.metadata = { univerNoteDecorationSource: { version: 1, kind: 'footnote', type: 'separator', noteId: 'public-separator' } }
const dormant = story([paragraph('dormant-anchor'), image('dormant-image', 'dormant-anchor')]); dormant.metadata = {
  univerNoteDecorationSource: { version: 1, kind: 'footnote', type: 'separator', noteId: 'public-extra-separator' },
}
initial.noteSettings = { footnote: { separator: separatorContent, metadata: { officeDormantNoteDecorations: { version: 1, contents: [dormant] } } } }
let studio = new OfficeStudio({ document: initial, currentUserId: 'editor' }); studio.activateDocument(initial.id)
const api = shallowRef(new FDocument(studio, initial.id)), model = shallowRef(api.value.getModel()), dark = ref(false), variant = ref('classic')
const commands = [], notices = [], observed = []
let subscription, cancellation, reentry
function observe() {
  subscription?.dispose()
  const stop = watch(() => studio.editor.document.model, () => { model.value = api.value.getModel() }, { flush: 'sync' })
  const after = studio.events.addEvent('CommandExecuted', event => { if (event.unitId === api.value.id && event.kind === 'document' && event.id.startsWith('doc.command.')) commands.push(event.id) })
  subscription = { dispose() { stop(); after.dispose() } }
}
observe()
function selectBody() { const point = { blockId: 'body-25', offset: 2 }; return api.value.setInputSelection({ anchor: point, focus: point }) }
function outside() { const point = { blockId: 'outside', offset: 0 }; return api.value.setInputSelection({ anchor: point, focus: point }) }
function armCommitObserver() {
  reentry?.dispose()
  reentry = studio.events.addEvent('CommandExecuted', event => {
    if (event.unitId !== api.value.id || event.id !== 'doc.command.insert-text') return
    reentry.dispose(); reentry = undefined
    const before = copy(api.value.getInputState()), revision = api.value.getModelRevision(); outside()
    observed.push({ before, after: copy(api.value.getInputState()), revision, latestRevision: api.value.getModelRevision() })
  })
}
function armReturnObserver() {
  reentry?.dispose()
  reentry = studio.events.addEvent('DocumentInputStateChanged', event => {
    const location = event.state?.location
    if (event.unitId !== api.value.id || !event.state?.valid || location?.kind !== 'body' || location.host
      || location.selection.focus.blockId !== 'body-25' || location.selection.focus.offset !== 2) return
    reentry.dispose(); reentry = undefined
    const before = copy(api.value.getInputState()), revision = api.value.getModelRevision(); outside()
    observed.push({ before, after: copy(api.value.getInputState()), revision, latestRevision: api.value.getModelRevision() })
  })
}
function replace(source) {
  subscription?.dispose(); cancellation?.dispose(); reentry?.dispose(); cancellation = undefined; reentry = undefined
  studio.dispose(); studio = new OfficeStudio({ document: source, currentUserId: 'editor' }); studio.activateDocument(source.id)
  api.value = new FDocument(studio, source.id); model.value = api.value.getModel(); observe()
}
window.noteWorkbenchFixture = {
  snapshot: () => copy(model.value), input: () => { const value = api.value.getInputState(); return value ? copy(value) : null },
  history: () => ({ undo: studio.editor.history.units[api.value.id]?.undo.length ?? 0, redo: studio.editor.history.units[api.value.id]?.redo.length ?? 0 }),
  commands: () => [...commands], notices: () => [...notices], observed: () => copy(observed), listenerErrors: () => studio.getEventListenerErrors().map(value => String(value.error)),
  selectBody, selectRange: selection => api.value.setInputSelection(selection), outside, armCommitObserver, armReturnObserver, setRules: rules => api.value.setPermissionRules(rules),
  undo: () => api.value.undo(), redo: () => api.value.redo(), appearance: (isDark, layout) => { dark.value = isDark; variant.value = layout },
  cancel: id => { cancellation?.dispose(); cancellation = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === id) event.cancel = true }) },
  clearCancellation: () => { cancellation?.dispose(); cancellation = undefined },
  saveReopen: () => { const saved = copy(api.value.saveModel()); replace(saved); return saved },
  replaceSource: () => { const value = copy(initial); value.id = 'replacement-note-browser'; replace(value) },
}
const app = createApp({ render: () => h(ConfigProvider, { theme: { algorithm: dark.value ? theme.darkAlgorithm : theme.defaultAlgorithm, token: { colorPrimary: '#3064c7' } } }, {
  default: () => h(DocumentWorkbench, { modelValue: model.value, documentApi: api.value, permissionUserId: 'editor',
    canUndo: studio.editor.canUndo, canRedo: studio.editor.canRedo, ribbonVariant: variant.value,
    'onUpdate:modelValue': value => { model.value = value }, onUndo: () => api.value.undo(), onRedo: () => api.value.redo(), onNotify: message => notices.push(message),
  }),
}) })
app.mount('#app'); window.addEventListener('pagehide', () => { subscription?.dispose(); cancellation?.dispose(); reentry?.dispose(); app.unmount(); studio.dispose() }, { once: true })
</script></body></html>`

let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'document-note-workbench-fixture',
    configureServer(vite) { vite.middlewares.use(fixturePath, async (_req, res, next) => { try { res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.end(await vite.transformIndexHtml(fixturePath, fixture)) } catch (error) { next(error) } }) },
  }], server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate an HTTP port')
  baseURL = 'http://127.0.0.1:' + address.port + fixturePath; browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)
afterAll(async () => { await browser?.close(); await server?.close() })
beforeEach(async () => {
  errors = []; page = await browser.newPage({ viewport: { width: 1440, height: 900 }, deviceScaleFactor: 1.25 })
  page.on('pageerror', error => errors.push(error.message)); await page.goto(baseURL)
  await page.waitForFunction(() => !!window.noteWorkbenchFixture); await expect(inputBridge()).toBeAttached()
})
afterEach(async () => {
  const listenerErrors = await page?.evaluate(() => window.noteWorkbenchFixture?.listenerErrors() ?? [])
  await page?.close(); assert.deepEqual(errors, []); assert.deepEqual(listenerErrors ?? [], [])
})

const snapshot = () => page.evaluate(() => window.noteWorkbenchFixture.snapshot())
const history = () => page.evaluate(() => window.noteWorkbenchFixture.history())
const input = () => page.evaluate(() => window.noteWorkbenchFixture.input())
const inputBridge = () => page.getByRole('textbox', { name: '文档文字输入', exact: true })
const inspection = () => page.locator('[data-document-content-inspection]')
const decorationRow = (dialog, key) => dialog.locator('[data-document-note-decoration="' + key + '"]')
const noticeText = model => model.noteSettings?.footnote?.continuationNotice?.blocks.map(block => block.runs.map(run => run.text).join('')).join('\n')
async function references() { await page.getByRole('tab', { name: '引用', exact: true }).click() }
async function openSettings(kind = '脚注') {
  await references(); await page.getByRole('button', { name: kind + '设置', exact: true }).click()
  const dialog = page.getByRole('dialog', { name: kind + '设置', exact: true }); await expect(dialog).toBeVisible(); return dialog
}
async function openNotes() {
  await references(); await page.getByRole('button', { name: '管理脚注和尾注', exact: true }).click()
  const dialog = page.getByRole('dialog', { name: '脚注和尾注', exact: true }); await expect(dialog).toBeVisible(); return dialog
}
async function editDecoration(key) {
  const dialog = await openSettings(); await decorationRow(dialog, key).getByRole('button', { name: '在画布中编辑', exact: true }).click()
  await expect(inspection()).toBeVisible(); await expect(inputBridge()).toBeFocused(); return input()
}
async function returnBody() { await inspection().getByRole('button', { name: '返回正文', exact: true }).click(); await expect(inspection()).toHaveCount(0); await expect(inputBridge()).toBeFocused() }
async function overlaysClosed() { await expect(page.locator('.ant-modal-wrap:visible')).toHaveCount(0) }
async function chooseScope(dialog, label) {
  const item = dialog.locator('.ant-form-item').filter({ hasText: '应用范围' })
  await item.getByRole('combobox').click(); await page.getByRole('option', { name: label, exact: true }).click(); return item
}
async function setZoom(value) {
  await page.getByRole('combobox', { name: '文档缩放', exact: true }).click(); await page.getByRole('option', { name: value + '%', exact: true }).click()
  await expect(page.locator('.zoom-value')).toHaveText(value + '%')
}
async function scroll() { return page.locator('.document-scroll').evaluate(element => ({ left: element.scrollLeft, top: element.scrollTop })) }
async function settleFrames() { await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))) }
async function picturePixels() {
  const raster = await page.locator('.document-surface').screenshot()
  return page.evaluate(async source => {
    const image = new Image(); image.src = source; await image.decode()
    const canvas = document.createElement('canvas'); canvas.width = image.width; canvas.height = image.height
    const context = canvas.getContext('2d'); context.drawImage(image, 0, 0)
    const data = context.getImageData(0, 0, canvas.width, canvas.height).data; let count = 0
    for (let index = 0; index < data.length; index += 4) if (data[index] > 230 && data[index + 1] < 30 && data[index + 2] > 230 && data[index + 3] > 230) count++
    return count
  }, 'data:image/png;base64,' + raster.toString('base64'))
}

for (const [kind, label, reverse, dark, variant] of [
  ['footnote', '脚注', false, false, 'classic'], ['footnote', '脚注', true, true, 'grid'],
  ['endnote', '尾注', false, false, 'classic'], ['endnote', '尾注', true, true, 'grid'],
]) {
  test(kind + ' Ribbon insertion uses the editable logical end of a ' + (reverse ? 'reversed' : 'forward') + ' nonempty selection', async () => {
    const property = kind === 'footnote' ? 'footnotes' : 'endnotes', command = 'doc.command.insert-' + kind
    await page.evaluate(([dark, variant]) => window.noteWorkbenchFixture.appearance(dark, variant), [dark, variant])
    await page.evaluate(() => window.noteWorkbenchFixture.setRules([
      { id: 'deny-start', scope: 'paragraph', targetId: 'body-24', action: 'edit', effect: 'deny', userIds: ['editor'] },
    ]))
    await page.evaluate(reverse => {
      const start = { blockId: 'body-24', offset: 2 }, end = { blockId: 'body-25', offset: 2 }
      window.noteWorkbenchFixture.selectRange(reverse ? { anchor: end, focus: start } : { anchor: start, focus: end })
    }, reverse)
    await references()
    const button = page.getByRole('button', { name: '插入' + label, exact: true }), before = await snapshot(), stack = await history()
    await expect(button).toBeEnabled()
    await page.evaluate(command => window.noteWorkbenchFixture.cancel(command), command); await button.click()
    assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), stack)
    await page.evaluate(() => window.noteWorkbenchFixture.clearCancellation()); await button.click()
    await expect.poll(async () => Object.keys((await snapshot())[property] ?? {}).length).toBe(Object.keys(before[property] ?? {}).length + 1)
    const after = await snapshot(), note = Object.entries(after[property]).find(([id]) => !Object.hasOwn(before[property] ?? {}, id))[1]
    const first = after.blocks.find(block => block.id === 'body-24'), last = after.blocks.find(block => block.id === 'body-25')
    assert.deepEqual(first, before.blocks.find(block => block.id === first.id))
    assert.equal(last.runs.filter(run => !run.metadata?.note).map(run => run.text).join(''), before.blocks.find(block => block.id === last.id).runs.map(run => run.text).join(''))
    assert.equal(note.referenceBlockId, last.id)
    assert.equal(last.runs.slice(0, last.runs.findIndex(run => run.id === note.referenceRunId)).reduce((offset, run) => offset + run.text.length, 0), 2)
    assert.deepEqual((await input()).location.host, { kind: 'note', noteKind: kind, noteId: note.id }); await expect(inputBridge()).toBeFocused()
    assert.equal((await history()).undo, stack.undo + 1)
    assert.deepEqual((await page.evaluate(() => window.noteWorkbenchFixture.commands())).filter(id => id === command), [command])
    await page.evaluate(() => window.noteWorkbenchFixture.undo()); await expect.poll(snapshot).toEqual(before)
    await page.evaluate(() => window.noteWorkbenchFixture.redo()); await expect.poll(snapshot).toEqual(after)
    const saved = await page.evaluate(() => window.noteWorkbenchFixture.saveReopen()); await expect.poll(snapshot).toEqual(saved)
    assert.deepEqual(saved[property][note.id], after[property][note.id]); assert.deepEqual(saved.blocks.find(block => block.id === first.id), first)

    await page.evaluate(() => window.noteWorkbenchFixture.setRules([
      { id: 'deny-end', scope: 'paragraph', targetId: 'body-25', action: 'edit', effect: 'deny', userIds: ['editor'] },
    ]))
    await page.evaluate(reverse => {
      const start = { blockId: 'body-24', offset: 2 }, end = { blockId: 'body-25', offset: 2 }
      window.noteWorkbenchFixture.selectRange(reverse ? { anchor: end, focus: start } : { anchor: start, focus: end })
    }, reverse)
    await references(); await expect(button).toBeDisabled()
  })
}

test('opening an orphan picture story paints its real image without model/history writes and restores body selection, scroll and zoom', async () => {
  await page.evaluate(() => window.noteWorkbenchFixture.selectBody()); await setZoom(125); await settleFrames()
  await page.locator('.document-scroll').evaluate(element => { element.scrollTop = 410; element.scrollLeft = 0 }); await settleFrames()
  const before = await snapshot(), stack = await history(), bodyInput = await input(), offset = await scroll()
  const notes = await openNotes(); await notes.locator('[data-document-note-edit="footnote:orphan"]').click()
  await overlaysClosed(); await expect(inspection()).toHaveAttribute('data-document-inspection-host', 'footnote:orphan')
  assert.deepEqual((await input()).location.host, { kind: 'note', noteKind: 'footnote', noteId: 'orphan' })
  assert.equal((await input()).location.selection.focus.blockId, 'orphan-anchor')
  await expect.poll(picturePixels).toBeGreaterThan(40); assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), stack)
  await expect(inputBridge()).toBeFocused(); await returnBody(); await settleFrames()
  assert.deepEqual((await input()).location, bodyInput.location); assert.deepEqual(await scroll(), offset)
  await expect(page.locator('.zoom-value')).toHaveText('125%'); assert.deepEqual(await history(), stack)
})

test('separator-only and public dormant picture definitions open separate hosts and retain source IDs without ordinary notes', async () => {
  const before = await snapshot(), stack = await history()
  const initialSettings = await openSettings(); await expect(decorationRow(initialSettings, 'separator')).toContainText('公开 ID public-separator')
  await decorationRow(initialSettings, 'separator').getByRole('button', { name: '在画布中编辑', exact: true }).click()
  await expect(inspection()).toBeVisible(); const separator = await input(); assert.equal(separator.location.selection.focus.blockId, 'separator-anchor')
  await expect(inspection()).toHaveAttribute('data-document-inspection-host', 'noteDecoration:footnote:separator:active')
  await returnBody(); const dialog = await openSettings()
  await expect(decorationRow(dialog, 'dormant:0')).toContainText('公开 ID public-extra-separator')
  await decorationRow(dialog, 'dormant:0').getByRole('button', { name: '在画布中编辑', exact: true }).click()
  await expect(inspection()).toHaveAttribute('data-document-inspection-host', 'noteDecoration:footnote:separator:0')
  assert.equal((await input()).location.selection.focus.blockId, 'dormant-anchor'); await expect.poll(picturePixels).toBeGreaterThan(40)
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), stack)
  const content = (await snapshot()).noteSettings.footnote.metadata.officeDormantNoteDecorations.contents[0]
  assert.equal(content.metadata.univerNoteDecorationSource.noteId, 'public-extra-separator'); assert.equal(content.metadata.decorationSource, undefined)
  assert.deepEqual(Object.keys((await snapshot()).footnotes).sort(), ['linked', 'orphan']); await returnBody()
})

test('the default continuation notice registers on its first Canvas edit, supports undo/redo and survives local saveModel reopen', async () => {
  const before = await snapshot(), stack = await history(); const selected = await editDecoration('continuationNotice')
  assert.equal(selected.location.host.decoration, 'continuationNotice'); assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), stack)
  await page.keyboard.insertText('Continue😀'); await expect.poll(async () => noticeText(await snapshot())).toBe('Continue😀')
  assert.equal((await history()).undo, stack.undo + 1); const edited = (await snapshot()).noteSettings.footnote.continuationNotice
  await page.evaluate(() => window.noteWorkbenchFixture.undo()); await expect.poll(async () => noticeText(await snapshot())).toBe(undefined)
  await page.evaluate(() => window.noteWorkbenchFixture.redo()); await expect.poll(async () => noticeText(await snapshot())).toBe('Continue😀')
  assert.deepEqual((await snapshot()).noteSettings.footnote.continuationNotice, edited)
  const saved = await page.evaluate(() => window.noteWorkbenchFixture.saveReopen()); await expect.poll(async () => (await input())?.valid).toBe(true)
  assert.deepEqual((await snapshot()).noteSettings, saved.noteSettings); await editDecoration('continuationNotice')
  await expect(inputBridge()).toBeFocused(); assert.equal(noticeText(await snapshot()), 'Continue😀')
})

test('settings drafts block scope changes and inspection, and custom mark drafts require explicit clearing before inspection', async () => {
  const before = await snapshot(), stack = await history(), dialog = await openSettings()
  const start = dialog.locator('.ant-form-item').filter({ hasText: '起始编号' }).getByRole('spinbutton'); await start.fill('3')
  await expect(decorationRow(dialog, 'separator').getByRole('button', { name: '在画布中编辑', exact: true })).toBeDisabled()
  let scope = await chooseScope(dialog, '当前节'); await expect(scope.locator('.ant-select-selection-item')).toHaveText('整个文档')
  await expect(dialog).toContainText('当前范围有未应用的设置'); assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), stack)
  await dialog.getByRole('button', { name: '还原设置草稿', exact: true }).click(); scope = await chooseScope(dialog, '当前节')
  await expect(scope.locator('.ant-select-selection-item')).toHaveText('当前节'); await chooseScope(dialog, '整个文档')
  await dialog.getByRole('checkbox', { name: '使用自定义标记', exact: true }).check(); await dialog.getByRole('textbox', { name: '自定义注释标记', exact: true }).fill('※')
  await expect(decorationRow(dialog, 'separator').getByRole('button', { name: '在画布中编辑', exact: true })).toBeDisabled()
  await dialog.getByRole('button', { name: '清空标记草稿', exact: true }).click(); await expect(decorationRow(dialog, 'separator').getByRole('button', { name: '在画布中编辑', exact: true })).toBeEnabled()
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), stack)
  await dialog.getByRole('button', { name: '关闭', exact: true }).click(); await overlaysClosed(); await expect(inputBridge()).toBeFocused()
})

test('scoped view revocation clears Notes drafts and navigation, while an edit denial still permits inspection', async () => {
  const notes = await openNotes(), linked = notes.locator('[data-document-note="footnote:linked"]')
  await linked.getByRole('textbox').fill('USER DRAFT SECRET')
  await page.evaluate(() => window.noteWorkbenchFixture.setRules([{ id: 'hide-note', scope: 'paragraph', targetId: 'linked-note', action: 'view', effect: 'deny', userIds: ['editor'] }]))
  await expect(linked.getByRole('textbox')).toHaveValue(''); await expect(linked.getByRole('textbox')).toBeDisabled()
  await expect(linked.locator('[data-document-note-edit]')).toBeDisabled(); await expect(linked.locator('.note-number')).toBeDisabled()
  await expect(notes).not.toContainText('SECRET note content'); await expect(notes).not.toContainText('USER DRAFT SECRET')
  assert.equal((await snapshot()).footnotes.linked.text, 'SECRET note content'); await notes.getByRole('button', { name: '关闭', exact: true }).click(); await overlaysClosed()
  await page.evaluate(() => window.noteWorkbenchFixture.setRules([{ id: 'hide-separator', scope: 'paragraph', targetId: 'separator-anchor', action: 'view', effect: 'deny', userIds: ['editor'] }]))
  let dialog = await openSettings(); await expect(decorationRow(dialog, 'separator').getByRole('button', { name: '在画布中编辑', exact: true })).toBeDisabled()
  await expect(decorationRow(dialog, 'continuationNotice').getByRole('button', { name: '在画布中编辑', exact: true })).toBeEnabled()
  await dialog.getByRole('button', { name: '关闭', exact: true }).click(); await overlaysClosed()
  await page.evaluate(() => window.noteWorkbenchFixture.setRules([{ id: 'lock-separator', scope: 'paragraph', targetId: 'separator-anchor', action: 'edit', effect: 'deny', userIds: ['editor'] }]))
  dialog = await openSettings(); await expect(decorationRow(dialog, 'separator').getByRole('button', { name: '在画布中编辑', exact: true })).toBeEnabled()
  await decorationRow(dialog, 'separator').getByRole('button', { name: '在画布中编辑', exact: true }).click(); await expect(inspection()).toBeVisible()
  const before = await snapshot(), stack = await history(); await inputBridge().focus(); await page.keyboard.insertText('Forbidden')
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), stack)
})

test('a note edit observer keeps its newer body input state instead of restoring the completed note host', async () => {
  await editDecoration('continuationNotice'); const stack = await history()
  await page.evaluate(() => window.noteWorkbenchFixture.armCommitObserver()); await page.keyboard.insertText('Observed note')
  await expect.poll(async () => noticeText(await snapshot())).toBe('Observed note'); await expect(inspection()).toHaveCount(0)
  const records = await page.evaluate(() => window.noteWorkbenchFixture.observed()); assert.equal(records.length, 1)
  assert.equal(records[0].revision, records[0].latestRevision); assert.equal(records[0].after.location.host, undefined)
  assert.equal(records[0].after.location.selection.focus.blockId, 'outside'); await expect.poll(input).toEqual(records[0].after)
  assert.equal((await history()).undo, stack.undo + 1); await expect(inputBridge()).toBeFocused()
})

test('returning from inspection respects a reentrant body selection and does not restore the older scroll offset', async () => {
  await page.evaluate(() => window.noteWorkbenchFixture.selectBody()); await settleFrames()
  await page.locator('.document-scroll').evaluate(element => { element.scrollTop = 120 }); await settleFrames(); const old = await scroll()
  const notes = await openNotes(); await notes.locator('[data-document-note-edit="footnote:orphan"]').click(); await overlaysClosed(); await expect(inspection()).toBeVisible()
  await page.evaluate(() => window.noteWorkbenchFixture.armReturnObserver()); await returnBody(); await settleFrames()
  const records = await page.evaluate(() => window.noteWorkbenchFixture.observed()); assert.equal(records.length, 1)
  await expect.poll(input).toEqual(records[0].after); assert.equal(records[0].after.location.selection.focus.blockId, 'outside')
  await expect.poll(async () => (await scroll()).top).toBeGreaterThan(old.top + 100)
})

test('light/dark themes and classic/grid Ribbon layouts keep real antd overlays and the Canvas bridge interactive', async () => {
  const colors = []
  for (const [dark, variant] of [[false, 'classic'], [true, 'grid'], [false, 'grid'], [true, 'classic']]) {
    await page.evaluate(([dark, variant]) => window.noteWorkbenchFixture.appearance(dark, variant), [dark, variant])
    const dialog = await openSettings(); await expect(decorationRow(dialog, 'continuationNotice').getByRole('button', { name: '在画布中编辑', exact: true })).toBeEnabled()
    colors.push(await dialog.locator('.ant-modal-content').evaluate(element => ({ background: getComputedStyle(element).backgroundColor, color: getComputedStyle(element).color })))
    await dialog.getByRole('button', { name: '关闭', exact: true }).click(); await overlaysClosed(); await expect(inputBridge()).toBeFocused()
    await expect(inputBridge()).toHaveClass(/ant-input/); assert.equal((await history()).undo, 0)
  }
  assert.notDeepEqual(colors[0], colors[1]); assert.deepEqual(colors[0], colors[2]); assert.deepEqual(colors[1], colors[3])
})

test('synthetic composition registers a default notice once and suppresses trailing duplicate input', async () => {
  await editDecoration('continuationNotice'); const before = await snapshot(), stack = await history()
  await inputBridge().evaluate(element => {
    element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }))
    element.dispatchEvent(new CompositionEvent('compositionupdate', { bubbles: true, data: '续页😀' }))
  })
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), stack)
  await inputBridge().evaluate(element => {
    element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '续页😀' }))
    element.value = '续页😀'; element.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertCompositionText', data: '续页😀', isComposing: false }))
  })
  await expect.poll(async () => noticeText(await snapshot())).toBe('续页😀'); assert.equal((await history()).undo, stack.undo + 1)
  assert.deepEqual((await page.evaluate(() => window.noteWorkbenchFixture.commands())).filter(id => id === 'doc.command.insert-text'), ['doc.command.insert-text'])
})

test('canceling a synthetic note composition keeps the detached default and its history unchanged', async () => {
  await editDecoration('continuationNotice'); const before = await snapshot(), stack = await history()
  await page.evaluate(() => window.noteWorkbenchFixture.cancel('doc.command.insert-text'))
  await inputBridge().evaluate(element => {
    element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }))
    element.dispatchEvent(new CompositionEvent('compositionupdate', { bubbles: true, data: '取消续页草稿' }))
    element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '取消续页草稿' }))
  })
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), stack); await expect(inputBridge()).toBeFocused()
})

test('a stale composition input cannot publish an old notice into a replacement document', async () => {
  await editDecoration('continuationNotice')
  await inputBridge().evaluate(element => {
    window.staleNoteInput = element
    element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }))
    element.dispatchEvent(new CompositionEvent('compositionupdate', { bubbles: true, data: '旧文档注释草稿' }))
  })
  await page.evaluate(() => window.noteWorkbenchFixture.replaceSource())
  await page.evaluate(() => window.staleNoteInput.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '旧文档注释草稿' })))
  await expect.poll(async () => (await snapshot()).id).toBe('replacement-note-browser')
  assert.equal(noticeText(await snapshot()), undefined); assert.equal((await history()).undo, 0); await expect(inspection()).toHaveCount(0)
})
