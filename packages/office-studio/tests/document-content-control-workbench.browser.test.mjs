import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

// Deferred DOC-11 browser drafts mount the real Workbench, facade and antdv-next overlays.
// Synthetic composition proves event guards only; real Windows IME, Office samples and both renderers remain unaccepted.
const fixturePath = '/__document-content-control-workbench-test__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>Docs content controls</title><style>body{margin:0}#app{height:100vh}</style></head><body><div id="app"></div><script type="module">
import { createApp, h, ref, shallowRef, watch } from 'vue'
import { ConfigProvider, theme } from 'antdv-next'
import { OfficeStudio } from '/packages/office-studio/src/api/office-studio'
import { FDocument } from '/packages/office-studio/src/api/facade'
import DocumentWorkbench from '/packages/office-studio/src/components/DocumentWorkbench.vue'
import { createBlankDocument } from '/packages/office-studio/src/domain/model'
import { createParagraph } from '/packages/office-studio/src/features/document/model/document-model'
const copy = value => JSON.parse(JSON.stringify(value))
const initial = createBlankDocument(); initial.id = 'sdt-browser'; initial.name = '内容控件行为草稿'
initial.page = { width: 620, height: 540, marginTop: 52, marginBottom: 40, marginLeft: 48, marginRight: 48 }
const paragraph = (id, text) => { const block = createParagraph(text); block.id = id; return block }
initial.blocks = [paragraph('check-p', '☐'), paragraph('date-p', '03 Sep 2026'), paragraph('list-p', 'Alpha'), paragraph('combo-p', 'Free'), paragraph('text-p', 'Name'), paragraph('outside', 'Outside')]
initial.blocks[0].runs[0].fontFamily = 'Segoe UI Symbol'
const list = [{ value: 'A', displayText: 'Alpha' }, { value: 'B', displayText: 'Alpha' }]
initial.contentControls = [
  { id: 'check', type: 'checkbox', blockId: 'check-p', startOffset: 0, endOffset: 1, alias: '确认', checked: false, value: 'false', lock: 'unlocked', checkbox: { checked: false, checkedState: { value: '2612', font: 'Segoe UI Symbol' }, uncheckedState: { value: '2610', font: 'Segoe UI Symbol' } } },
  { id: 'date', type: 'date', blockId: 'date-p', startOffset: 0, endOffset: 11, alias: '截止日期', value: '2026-09-03', dateValue: '2026-09-03', dateFormat: 'dd MMM yyyy', lock: 'unlocked' },
  { id: 'list', type: 'dropDownList', blockId: 'list-p', startOffset: 0, endOffset: 5, alias: '候选项', value: 'A', items: list, lock: 'unlocked' },
  { id: 'combo', type: 'comboBox', blockId: 'combo-p', startOffset: 0, endOffset: 4, alias: '可输入候选', value: 'Free', items: [{ value: 'A', displayText: 'Alpha' }], lock: 'unlocked' },
  { id: 'text', type: 'text', blockId: 'text-p', startOffset: 0, endOffset: 4, alias: '姓名', lock: 'unlocked' },
]
const storyContent = (id, text) => { const block = paragraph(id, text); return { blocks: [block], runs: copy(block.runs), paragraphs: [{ startOffset: 0, endOffset: text.length, style: {} }] } }
initial.sectionHeaderFooters = { first: { header: '☐', footer: '', firstPageHeader: '', firstPageFooter: '', evenPageHeader: '', evenPageFooter: '', differentFirstPage: false, differentOddEven: false, linkToPrevious: false, storyContent: { header: storyContent('header-p', '☐') } } }
initial.footnotes = { note: { id: 'note', text: 'Alpha', referenceBlockId: 'text-p', referenceRunId: initial.blocks[4].runs[0].id, content: storyContent('note-p', 'Alpha') } }
initial.contentControls.push({ id: 'header-check', type: 'checkbox', blockId: 'header-p', startOffset: 0, endOffset: 1, alias: '页眉确认', value: 'false', checked: false, checkbox: copy(initial.contentControls[0].checkbox), lock: 'unlocked' },
  { id: 'note-list', type: 'dropDownList', blockId: 'note-p', startOffset: 0, endOffset: 5, alias: '脚注候选', value: 'A', items: [{ value: 'A', displayText: 'Alpha' }, { value: 'B', displayText: 'Beta' }], lock: 'unlocked' })
let studio = new OfficeStudio({ document: initial, currentUserId: 'editor' }); studio.activateDocument(initial.id)
const api = shallowRef(new FDocument(studio, initial.id)), model = shallowRef(api.value.getModel()), dark = ref(false), variant = ref('classic'), commands = [], notices = []
let unsubscribe, cancellation, reentry
const observerInputs = []
function observe() {
  unsubscribe?.dispose()
  const stopModel = watch(() => studio.editor.document.model, () => { model.value = api.value.getModel() }, { flush: 'sync' })
  const after = studio.events.addEvent('CommandExecuted', event => {
    if (event.unitId === api.value.id && event.kind === 'document' && event.id.startsWith('doc.command.')) commands.push(event.id)
  })
  unsubscribe = { dispose() { stopModel(); after.dispose() } }
}
observe()
function select(id, host) { const control = api.value.getContentControl(id).getInfo(); return api.value.setInputSelection({ anchor: { blockId: control.blockId, offset: control.startOffset }, focus: { blockId: control.blockId, offset: control.endOffset } }, undefined, host) }
function selectRange(id, start, end = start, host) { const control = api.value.getContentControl(id).getInfo(); return api.value.setInputSelection({ anchor: { blockId: control.blockId, offset: start }, focus: { blockId: control.blockId, offset: end } }, undefined, host) }
function outside() { return api.value.setInputSelection({ anchor: { blockId: 'outside', offset: 0 }, focus: { blockId: 'outside', offset: 0 } }) }
function armObserver(mode) {
  reentry?.dispose()
  reentry = studio.events.addEvent('CommandExecuted', event => {
    if (event.unitId !== api.value.id || event.id !== 'doc.command.select-content-control-item' || event.params?.controlId !== 'combo') return
    reentry.dispose(); reentry = undefined
    const before = copy(api.value.getInputState()), revisionBefore = api.value.getModelRevision()
    if (mode === 'selection') outside()
    else if (mode === 'revision') api.value.getContentControl('text').update({ alias: 'observer title' })
    else throw new Error('Unknown observer scenario')
    observerInputs.push({ before, after: copy(api.value.getInputState()), revisionBefore, revisionAfter: api.value.getModelRevision() })
  })
}
window.contentControlFixture = {
  snapshot: () => copy(model.value), history: () => ({ undo: studio.editor.history.units[api.value.id]?.undo.length ?? 0, redo: studio.editor.history.units[api.value.id]?.redo.length ?? 0 }), commands: () => [...commands], notices: () => [...notices],
  input: () => { const state = api.value.getInputState(); return state ? copy(state) : null }, observerInputs: () => copy(observerInputs), armObserver,
  listenerErrors: () => studio.getEventListenerErrors().map(failure => String(failure.error)),
  select: id => select(id), selectRange, outside,
  openHost: area => area === 'header' ? select('header-check', { kind: 'story', sectionId: 'first', story: 'header' }) : select('note-list', { kind: 'note', noteKind: 'footnote', noteId: 'note' }),
  update: (id, patch) => api.value.getContentControl(id).update(patch), setChecked: checked => api.value.getContentControl('check').setChecked(checked),
  setRules: rules => api.value.setPermissionRules(rules), undo: () => api.value.undo(), redo: () => api.value.redo(), appearance: (isDark, layout) => { dark.value = isDark; variant.value = layout },
  cancel: id => { cancellation?.dispose(); cancellation = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === id) event.cancel = true }) },
  saveReopen: () => { const saved = copy(api.value.save()); unsubscribe?.dispose(); cancellation?.dispose(); reentry?.dispose(); studio.dispose(); studio = new OfficeStudio({ document: saved, currentUserId: 'editor' }); studio.activateDocument(saved.id); api.value = new FDocument(studio, saved.id); model.value = api.value.getModel(); observe(); return saved },
  replaceSource: () => { const source = copy(initial); source.id = 'replacement-sdt'; unsubscribe?.dispose(); cancellation?.dispose(); reentry?.dispose(); studio.dispose(); studio = new OfficeStudio({ document: source, currentUserId: 'editor' }); studio.activateDocument(source.id); api.value = new FDocument(studio, source.id); model.value = api.value.getModel(); observe() },
}
const app = createApp({ render: () => h(ConfigProvider, { theme: { algorithm: dark.value ? theme.darkAlgorithm : theme.defaultAlgorithm, token: { colorPrimary: '#3064c7' } } }, { default: () => h(DocumentWorkbench, {
  modelValue: model.value, documentApi: api.value, permissionUserId: 'editor', canUndo: studio.editor.canUndo, canRedo: studio.editor.canRedo, ribbonVariant: variant.value,
  'onUpdate:modelValue': value => { model.value = value }, onUndo: () => api.value.undo(), onRedo: () => api.value.redo(), onNotify: message => notices.push(message),
}) }) })
app.mount('#app'); window.addEventListener('pagehide', () => { unsubscribe?.dispose(); cancellation?.dispose(); reentry?.dispose(); app.unmount(); studio.dispose() }, { once: true })
</script></body></html>`

let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'document-content-control-workbench-fixture',
    configureServer(vite) { vite.middlewares.use(fixturePath, async (_req, res, next) => { try { res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.end(await vite.transformIndexHtml(fixturePath, fixture)) } catch (error) { next(error) } }) },
  }], server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate an HTTP port')
  baseURL = 'http://127.0.0.1:' + address.port + fixturePath; browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)
afterAll(async () => { await browser?.close(); await server?.close() })
beforeEach(async () => { errors = []; page = await browser.newPage({ viewport: { width: 1440, height: 900 }, deviceScaleFactor: 1.25 }); page.on('pageerror', error => errors.push(error.message)); await page.goto(baseURL); await page.waitForFunction(() => !!window.contentControlFixture); await expect(inputBridge()).toBeAttached() })
afterEach(async () => { const listenerErrors = await page?.evaluate(() => window.contentControlFixture?.listenerErrors() ?? []); await page?.close(); assert.deepEqual(errors, []); assert.deepEqual(listenerErrors ?? [], []) })

const snapshot = () => page.evaluate(() => window.contentControlFixture.snapshot())
const history = () => page.evaluate(() => window.contentControlFixture.history())
const control = (model, id) => model.contentControls.find(control => control.id === id)
const blockText = (model, id) => model.blocks.find(block => block.id === id).runs.map(run => run.text).join('')
const inputBridge = () => page.getByRole('textbox', { name: '文档文字输入', exact: true })
async function openEditor(label) { await page.getByRole('button', { name: '编辑' + label, exact: true }).click(); const editor = page.locator('.content-control-editor'); await expect(editor).toBeVisible(); return editor }
async function propertiesFor(id) { await page.evaluate(id => window.contentControlFixture.select(id), id); await page.getByRole('tab', { name: '内容控件', exact: true }).click(); await page.getByRole('button', { name: '属性', exact: true }).click(); const dialog = page.getByRole('dialog', { name: '内容控件属性', exact: true }); await expect(dialog).toBeVisible(); return dialog }

test('Canvas checkbox click and keyboard Space persist symbols and reuse undo/redo', async () => {
  const checkbox = page.getByRole('checkbox', { name: '确认', exact: true })
  await checkbox.check(); await expect.poll(async () => control(await snapshot(), 'check').checked).toBe(true)
  assert.equal(blockText(await snapshot(), 'check-p'), '☒'); assert.equal((await history()).undo, 1)
  await expect(inputBridge()).toBeFocused()
  await page.evaluate(() => window.contentControlFixture.undo()); await expect(checkbox).not.toBeChecked()
  await page.evaluate(() => window.contentControlFixture.redo()); await expect(checkbox).toBeChecked()
  await checkbox.focus(); await page.keyboard.press('Space'); await expect.poll(async () => control(await snapshot(), 'check').checked).toBe(false)
  assert.equal(blockText(await snapshot(), 'check-p'), '☐'); assert.equal((await history()).undo, 2)
})

test('DatePicker uses Dayjs display, stores ISO and updates the persisted formatted text in one command', async () => {
  const editor = await openEditor('截止日期'), picker = editor.locator('.ant-picker-input input')
  await expect(picker).toHaveValue('2026-09-03'); await picker.fill('2026-11-30'); await picker.press('Enter')
  await editor.getByRole('button', { name: '确定', exact: true }).click()
  await expect.poll(async () => control(await snapshot(), 'date').dateValue).toBe('2026-11-30')
  assert.equal(control(await snapshot(), 'date').value, '2026-11-30'); assert.equal(blockText(await snapshot(), 'date-p'), '30 Nov 2026')
  assert.equal((await history()).undo, 1); await expect(inputBridge()).toBeFocused()
})

test('duplicate dropdown display labels select their stored value and combo boxes accept a free value', async () => {
  const list = await openEditor('候选项'); await list.getByRole('combobox').click()
  await page.getByRole('option', { name: 'Alpha', exact: true }).nth(1).click(); await list.getByRole('button', { name: '确定', exact: true }).click()
  await expect.poll(async () => control(await snapshot(), 'list').value).toBe('B')
  assert.equal(blockText(await snapshot(), 'list-p'), 'Alpha')
  const combo = await openEditor('可输入候选'); await combo.getByRole('combobox').fill('自定义内容'); await combo.getByRole('button', { name: '确定', exact: true }).click()
  await expect.poll(async () => control(await snapshot(), 'combo').value).toBe('自定义内容')
  assert.equal(blockText(await snapshot(), 'combo-p'), '自定义内容'); assert.equal((await history()).undo, 2)
})

test('Escape cancels a pending picker value and restores the document input focus', async () => {
  const before = await snapshot(), editor = await openEditor('可输入候选')
  await editor.getByRole('combobox').fill('unsaved'); await page.keyboard.press('Escape')
  await expect(editor).toHaveCount(0); assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0)
  await expect(inputBridge()).toBeFocused()
})

test('external model edits, selection changes and source replacement invalidate pending value editors', async () => {
  await openEditor('可输入候选'); await page.evaluate(() => window.contentControlFixture.update('text', { tag: 'external' }))
  await expect(page.locator('.content-control-editor')).toHaveCount(0)
  await openEditor('可输入候选'); await page.evaluate(() => window.contentControlFixture.outside()); await expect(page.locator('.content-control-editor')).toHaveCount(0)
  const editor = await openEditor('可输入候选'); await editor.getByRole('combobox').fill('old-unit-draft')
  await page.evaluate(() => window.contentControlFixture.replaceSource()); await expect(page.locator('.content-control-editor')).toHaveCount(0)
  assert.equal(control(await snapshot(), 'combo').value, 'Free'); assert.equal((await history()).undo, 0)
})

test('a canceled registered command retains the model and content history', async () => {
  const before = await snapshot(); await page.evaluate(() => window.contentControlFixture.cancel('doc.command.select-content-control-item'))
  const editor = await openEditor('可输入候选'); await editor.getByRole('combobox').fill('canceled'); await editor.getByRole('button', { name: '确定', exact: true }).click()
  assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0)
  assert.deepEqual(await page.evaluate(() => window.contentControlFixture.commands()), [])
})

test('a command observer can move the caret without the completed overlay request restoring its old selection', async () => {
  const editor = await openEditor('可输入候选'); await editor.getByRole('combobox').fill('user value')
  await page.evaluate(() => window.contentControlFixture.armObserver('selection'))
  await editor.getByRole('button', { name: '确定', exact: true }).click(); await expect(editor).toHaveCount(0)
  await expect.poll(async () => control(await snapshot(), 'combo').value).toBe('user value')
  const observations = await page.evaluate(() => window.contentControlFixture.observerInputs()); assert.equal(observations.length, 1)
  const observed = observations[0]
  assert.equal(observed.revisionAfter, observed.revisionBefore); assert.equal(observed.after.generation, observed.before.generation + 1)
  assert.equal(observed.after.valid, true)
  assert.deepEqual(observed.after.location, { kind: 'body', selection: { anchor: { blockId: 'outside', offset: 0 }, focus: { blockId: 'outside', offset: 0 } } })
  await expect.poll(() => page.evaluate(() => window.contentControlFixture.input())).toEqual(observed.after)
  assert.equal((await history()).undo, 1)
  assert.deepEqual(await page.evaluate(() => window.contentControlFixture.commands()), ['doc.command.select-content-control-item'])
  await expect(inputBridge()).toBeFocused(); await page.keyboard.insertText('Z')
  await expect.poll(async () => blockText(await snapshot(), 'outside')).toBe('ZOutside')
  assert.equal(blockText(await snapshot(), 'combo-p'), 'user value')
})

test('an observer content commit leaves its revision and invalid input session intact after overlay confirmation', async () => {
  const editor = await openEditor('可输入候选'); await editor.getByRole('combobox').fill('user value')
  await page.evaluate(() => window.contentControlFixture.armObserver('revision'))
  await editor.getByRole('button', { name: '确定', exact: true }).click(); await expect(editor).toHaveCount(0)
  await expect.poll(async () => control(await snapshot(), 'text').alias).toBe('observer title')
  const observations = await page.evaluate(() => window.contentControlFixture.observerInputs()); assert.equal(observations.length, 1)
  const observed = observations[0]
  assert.equal(observed.revisionAfter, observed.revisionBefore + 1); assert.equal(observed.after.generation, observed.before.generation)
  assert.equal(observed.after.valid, false)
  await expect.poll(() => page.evaluate(() => window.contentControlFixture.input())).toEqual(observed.after)
  assert.equal(blockText(await snapshot(), 'combo-p'), 'user value'); assert.equal((await history()).undo, 2)
  assert.deepEqual(await page.evaluate(() => window.contentControlFixture.commands()), ['doc.command.select-content-control-item', 'doc.command.update-content-control'])
})

test('content locks and permission rules disable their actual overlay controls', async () => {
  await page.evaluate(() => window.contentControlFixture.update('check', { lock: 'contentLocked' }))
  await expect(page.getByRole('checkbox', { name: '确认', exact: true })).toBeDisabled()
  await page.evaluate(() => window.contentControlFixture.setRules([{ id: 'deny-list', scope: 'paragraph', targetId: 'list-p', action: 'edit', effect: 'deny', userIds: ['editor'] }]))
  await expect(page.getByRole('button', { name: '编辑候选项', exact: true })).toBeDisabled()
  await expect(page.getByRole('button', { name: '编辑可输入候选', exact: true })).toBeEnabled()
})

test('list properties support reordering, insertion and validation of duplicate stored values', async () => {
  const dialog = await propertiesFor('list')
  await dialog.getByRole('button', { name: '上移列表项', exact: true }).nth(1).click()
  await dialog.getByLabel('显示文字', { exact: true }).first().fill('Beta')
  await dialog.getByRole('button', { name: '新增选项', exact: true }).click()
  await dialog.getByLabel('显示文字', { exact: true }).nth(2).fill('Gamma'); await dialog.getByLabel('选项值', { exact: true }).nth(2).fill('A')
  await dialog.getByRole('button', { name: '保存', exact: true }).click(); await expect(dialog.getByText('选项值不能重复。').first()).toBeVisible()
  assert.equal((await history()).undo, 0)
  await dialog.getByLabel('选项值', { exact: true }).nth(2).fill('C'); await dialog.getByRole('button', { name: '保存', exact: true }).click()
  await expect(dialog).toBeHidden(); assert.deepEqual(control(await snapshot(), 'list').items.map(item => item.value), ['B', 'A', 'C'])
  assert.equal(control(await snapshot(), 'list').items[0].displayText, 'Beta'); assert.equal((await history()).undo, 1)
})

test('properties save keeps a wrapper lock and can change a title without removing the control', async () => {
  await page.evaluate(() => window.contentControlFixture.update('text', { lock: 'sdtLocked' }))
  const dialog = await propertiesFor('text'); await dialog.getByLabel('标题', { exact: true }).fill('客户姓名'); await dialog.getByRole('button', { name: '保存', exact: true }).click()
  await expect(dialog).toBeHidden(); assert.equal(control(await snapshot(), 'text').alias, '客户姓名'); assert.equal(control(await snapshot(), 'text').lock, 'sdtLocked')
})

test('property permission changes disable fields and commit while keeping cancellation available', async () => {
  const dialog = await propertiesFor('text'); await dialog.getByLabel('标题', { exact: true }).fill('pending title')
  await page.evaluate(() => window.contentControlFixture.setRules([{ id: 'deny-text', scope: 'paragraph', targetId: 'text-p', action: 'edit', effect: 'deny', userIds: ['editor'] }]))
  await expect(dialog.getByLabel('标题', { exact: true })).toBeDisabled(); await expect(dialog.getByRole('button', { name: '保存', exact: true })).toBeDisabled()
  await expect(dialog.getByRole('button', { name: '取消', exact: true })).toBeEnabled(); await dialog.getByRole('button', { name: '取消', exact: true }).click()
  assert.equal(control(await snapshot(), 'text').alias, '姓名'); assert.equal((await history()).undo, 1)
})

test('custom checkbox properties expose symbol input and persist its actual checked glyph', async () => {
  const dialog = await propertiesFor('check')
  await dialog.locator('.ant-select').filter({ hasText: '叉号' }).click(); await page.getByRole('option', { name: '自定义', exact: true }).click()
  const glyph = dialog.getByLabel('选中符号（十六进制）', { exact: true }); await expect(glyph).toBeFocused(); await glyph.fill('1F7E2')
  await dialog.getByLabel('选中符号字体', { exact: true }).fill('Segoe UI Emoji'); await dialog.getByRole('button', { name: '保存', exact: true }).click()
  await expect(dialog).toBeHidden(); await page.getByRole('checkbox', { name: '确认', exact: true }).check()
  await expect.poll(async () => blockText(await snapshot(), 'check-p')).toBe('🟢')
  assert.equal(control(await snapshot(), 'check').endOffset, 2); assert.equal((await history()).undo, 2)
})

test('light/dark themes and classic/grid layouts retain interactive antd overlays', async () => {
  const button = page.getByRole('button', { name: '编辑截止日期', exact: true }), before = await snapshot()
  const light = await button.evaluate(element => getComputedStyle(element).backgroundColor)
  await page.evaluate(() => window.contentControlFixture.appearance(true, 'grid'))
  await expect.poll(() => button.evaluate(element => getComputedStyle(element).backgroundColor)).not.toBe(light)
  const editor = await openEditor('截止日期'); await expect(editor.locator('.ant-picker')).toBeVisible(); await editor.getByRole('button', { name: '取消', exact: true }).click()
  assert.deepEqual(await snapshot(), before)
  await page.evaluate(() => window.contentControlFixture.appearance(false, 'classic')); await expect(button).toBeEnabled()
})

test('synthetic composition disables commit and closing the properties form clears its composition guard', async () => {
  const editor = await openEditor('可输入候选'), confirm = editor.getByRole('button', { name: '确定', exact: true })
  await editor.getByRole('combobox').fill('中文')
  await editor.evaluate(element => element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true })))
  await expect(confirm).toBeDisabled(); assert.equal((await history()).undo, 0)
  await editor.evaluate(element => element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '中文' })))
  await confirm.click(); await expect.poll(async () => control(await snapshot(), 'combo').value).toBe('中文'); assert.equal((await history()).undo, 1)
  let dialog = await propertiesFor('text'); await dialog.locator('form').evaluate(element => element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true })))
  await expect(dialog.getByRole('button', { name: '保存', exact: true })).toBeDisabled(); await dialog.getByRole('button', { name: '取消', exact: true }).click()
  dialog = await propertiesFor('text'); await expect(dialog.getByRole('button', { name: '保存', exact: true })).toBeEnabled()
})

test('snapshot reopen preserves actual values, Unicode text and content-control metadata', async () => {
  await page.getByRole('checkbox', { name: '确认', exact: true }).check()
  const editor = await openEditor('可输入候选'); await editor.getByRole('combobox').fill('saved'); await editor.getByRole('button', { name: '确定', exact: true }).click()
  const saved = await page.evaluate(() => window.contentControlFixture.saveReopen())
  await expect.poll(async () => control(await snapshot(), 'combo').value).toBe('saved')
  assert.deepEqual((await snapshot()).contentControls, saved.contentControls); assert.equal(blockText(await snapshot(), 'check-p'), '☒')
  assert.equal((await history()).undo, 0)
})

test('header and footnote host controls use the same Canvas input overlay and do not edit body blocks', async () => {
  const body = (await snapshot()).blocks
  await page.evaluate(() => window.contentControlFixture.openHost('header'))
  await page.getByRole('checkbox', { name: '页眉确认', exact: true }).check()
  await expect.poll(async () => (await snapshot()).sectionHeaderFooters.first.header).toBe('☒')
  await page.evaluate(() => window.contentControlFixture.openHost('note'))
  const editor = await openEditor('脚注候选'); await editor.getByRole('combobox').click(); await page.getByRole('option', { name: 'Beta', exact: true }).click(); await editor.getByRole('button', { name: '确定', exact: true }).click()
  await expect.poll(async () => (await snapshot()).footnotes.note.text).toBe('Beta')
  assert.deepEqual((await snapshot()).blocks, body)
})

test('the antd temporary property and ordinary input remove only the shell, with undo, redo and snapshot reopen', async () => {
  const dialog = await propertiesFor('text'); await dialog.getByRole('checkbox', { name: '临时控件', exact: true }).check()
  await dialog.getByRole('button', { name: '保存', exact: true }).click(); await expect(dialog).toBeHidden()
  assert.equal(control(await snapshot(), 'text').temporary, true)
  await page.evaluate(() => window.contentControlFixture.selectRange('text', 0, 4)); await inputBridge().focus()
  await page.keyboard.insertText('李😀')
  await expect.poll(async () => blockText(await snapshot(), 'text-p')).toBe('李😀')
  assert.equal(control(await snapshot(), 'text'), undefined); assert.equal((await history()).undo, 2)
  await expect(inputBridge()).toBeFocused()
  await page.evaluate(() => window.contentControlFixture.undo()); assert.equal(control(await snapshot(), 'text').temporary, true)
  assert.equal(blockText(await snapshot(), 'text-p'), 'Name')
  await page.evaluate(() => window.contentControlFixture.redo()); assert.equal(control(await snapshot(), 'text'), undefined)
  await page.evaluate(() => window.contentControlFixture.saveReopen()); assert.equal(blockText(await snapshot(), 'text-p'), '李😀')
  assert.equal(control(await snapshot(), 'text'), undefined); assert.equal((await history()).undo, 0)
})

test('a temporary checkbox restores a valid current caret after its shell disappears and accepts continued typing', async () => {
  await page.evaluate(() => window.contentControlFixture.update('check', { temporary: true }))
  await page.getByRole('checkbox', { name: '确认', exact: true }).check()
  await expect.poll(async () => blockText(await snapshot(), 'check-p')).toBe('☒')
  assert.equal(control(await snapshot(), 'check'), undefined); assert.equal((await history()).undo, 2)
  await expect(inputBridge()).toBeFocused()
  const input = await page.evaluate(() => window.contentControlFixture.input())
  assert.equal(input.valid, true); assert.deepEqual(input.location, { kind: 'body', selection: {
    anchor: { blockId: 'check-p', offset: 1 }, focus: { blockId: 'check-p', offset: 1 },
  } })
  await page.keyboard.insertText('继续')
  await expect.poll(async () => blockText(await snapshot(), 'check-p')).toBe('☒继续')
  assert.equal((await history()).undo, 3)
  await page.evaluate(() => window.contentControlFixture.undo()); assert.equal(blockText(await snapshot(), 'check-p'), '☒')
  await page.evaluate(() => window.contentControlFixture.undo()); assert.equal(control(await snapshot(), 'check').temporary, true)
  assert.equal(blockText(await snapshot(), 'check-p'), '☐')
})

test('a temporary combo value restores its UTF-16 end caret instead of the removed control selection', async () => {
  await page.evaluate(() => window.contentControlFixture.update('combo', { temporary: true }))
  const editor = await openEditor('可输入候选'); await editor.getByRole('combobox').fill('临时😀')
  await editor.getByRole('button', { name: '确定', exact: true }).click(); await expect(editor).toHaveCount(0)
  await expect.poll(async () => blockText(await snapshot(), 'combo-p')).toBe('临时😀')
  assert.equal(control(await snapshot(), 'combo'), undefined); await expect(inputBridge()).toBeFocused()
  const input = await page.evaluate(() => window.contentControlFixture.input())
  assert.equal(input.valid, true); assert.deepEqual(input.location.selection, {
    anchor: { blockId: 'combo-p', offset: 4 }, focus: { blockId: 'combo-p', offset: 4 },
  })
  await page.keyboard.insertText('后续'); await expect.poll(async () => blockText(await snapshot(), 'combo-p')).toBe('临时😀后续')
})

test('confirming an unchanged temporary date retains its shell, revision, history and checked focus session', async () => {
  await page.evaluate(() => window.contentControlFixture.update('date', { temporary: true }))
  const previousHistory = await history(), editor = await openEditor('截止日期')
  const before = await page.evaluate(() => window.contentControlFixture.input())
  await editor.getByRole('button', { name: '确定', exact: true }).click(); await expect(editor).toHaveCount(0)
  assert.equal(control(await snapshot(), 'date').temporary, true); assert.deepEqual(await history(), previousHistory)
  const after = await page.evaluate(() => window.contentControlFixture.input())
  assert.equal(after.modelRevision, before.modelRevision); assert.equal(after.valid, true)
  await expect(inputBridge()).toBeFocused()
})

test('ordinary input synchronizes date, checkbox and actual free combo text, and rejects dropdown editing', async () => {
  for (const sample of [{ id: 'date', length: 11, text: '04 Sep 2026' }, { id: 'check', length: 1, text: '☒' }, { id: 'combo', length: 4, text: 'Alpha' }]) {
    await page.evaluate(sample => window.contentControlFixture.selectRange(sample.id, 0, sample.length), sample)
    await inputBridge().focus(); await page.keyboard.insertText(sample.text)
  }
  await expect.poll(async () => control(await snapshot(), 'date').dateValue).toBe('2026-09-04')
  assert.equal(control(await snapshot(), 'date').value, '2026-09-04'); assert.equal(control(await snapshot(), 'check').checked, true)
  assert.equal(control(await snapshot(), 'combo').value, 'Alpha'); assert.equal((await history()).undo, 3)
  const before = await snapshot(), previousHistory = await history()
  await page.evaluate(() => window.contentControlFixture.selectRange('list', 0, 5)); await inputBridge().focus(); await page.keyboard.insertText('Beta')
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), previousHistory)
  assert.equal(control(await snapshot(), 'list').value, 'A')
})

test('an explicit persistent control selection keeps inside affinity for continued typing at its new end caret', async () => {
  await page.evaluate(() => window.contentControlFixture.select('text'))
  await page.getByRole('button', { name: '选择姓名', exact: true }).click(); await inputBridge().focus()
  await page.keyboard.insertText('输入')
  await expect.poll(async () => control(await snapshot(), 'text').value).toBe('输入')
  await page.keyboard.insertText('连续')
  await expect.poll(async () => blockText(await snapshot(), 'text-p')).toBe('输入连续')
  assert.equal(control(await snapshot(), 'text').value, '输入连续'); assert.equal(control(await snapshot(), 'text').endOffset, 4)
  assert.equal((await history()).undo, 2)
})

test('deleting the full contents of an explicitly selected persistent control retains its empty shell for further input', async () => {
  await page.evaluate(() => window.contentControlFixture.select('text'))
  await page.getByRole('button', { name: '选择姓名', exact: true }).click(); await inputBridge().focus()
  await page.keyboard.press('Backspace')
  await expect.poll(async () => blockText(await snapshot(), 'text-p')).toBe('')
  assert.equal(control(await snapshot(), 'text').value, ''); assert.equal(control(await snapshot(), 'text').startOffset, 0)
  assert.equal(control(await snapshot(), 'text').endOffset, 0); assert.equal((await history()).undo, 1)
  await page.keyboard.insertText('新值'); await expect.poll(async () => control(await snapshot(), 'text').value).toBe('新值')
  assert.equal(blockText(await snapshot(), 'text-p'), '新值'); assert.equal((await history()).undo, 2)
})

test('synthetic native composition replaces a complete placeholder in one commit and suppresses the trailing duplicate input', async () => {
  await page.evaluate(() => window.contentControlFixture.update('text', { placeholder: 'Name', showingPlaceholder: true, temporary: true }))
  const before = await snapshot(), previousHistory = await history()
  await page.evaluate(() => window.contentControlFixture.selectRange('text', 2)); await inputBridge().focus()
  await inputBridge().evaluate(element => {
    element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }))
    element.dispatchEvent(new CompositionEvent('compositionupdate', { bubbles: true, data: '输入😀' }))
  })
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), previousHistory)
  await inputBridge().evaluate(element => {
    element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '输入😀' }))
    element.value = '输入😀'; element.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertCompositionText', data: '输入😀', isComposing: false }))
  })
  await expect.poll(async () => blockText(await snapshot(), 'text-p')).toBe('输入😀')
  assert.equal(control(await snapshot(), 'text'), undefined); assert.equal((await history()).undo, previousHistory.undo + 1)
  assert.deepEqual((await page.evaluate(() => window.contentControlFixture.commands())).filter(id => id === 'doc.command.insert-text'), ['doc.command.insert-text'])
  await expect(inputBridge()).toBeFocused()
})

test('canceling a composition commit retains its temporary placeholder and content history', async () => {
  await page.evaluate(() => window.contentControlFixture.update('text', { placeholder: 'Name', showingPlaceholder: true, temporary: true }))
  const before = await snapshot(), previousHistory = await history()
  await page.evaluate(() => { window.contentControlFixture.selectRange('text', 2); window.contentControlFixture.cancel('doc.command.insert-text') })
  await inputBridge().focus(); await inputBridge().evaluate(element => {
    element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }))
    element.dispatchEvent(new CompositionEvent('compositionupdate', { bubbles: true, data: '取消草稿' }))
    element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '取消草稿' }))
  })
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), previousHistory)
  assert.equal(control(await snapshot(), 'text').showingPlaceholder, true); assert.equal(control(await snapshot(), 'text').temporary, true)
  await expect(inputBridge()).toBeFocused()
})

test('a stale composition input node cannot publish its preview into a replacement document', async () => {
  await page.evaluate(() => window.contentControlFixture.selectRange('text', 1)); await inputBridge().focus()
  await inputBridge().evaluate(element => {
    window.staleContentControlInput = element
    element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }))
    element.dispatchEvent(new CompositionEvent('compositionupdate', { bubbles: true, data: '旧文档草稿' }))
  })
  await page.evaluate(() => window.contentControlFixture.replaceSource())
  await page.evaluate(() => window.staleContentControlInput.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '旧文档草稿' })))
  await expect.poll(async () => (await snapshot()).id).toBe('replacement-sdt')
  assert.equal(blockText(await snapshot(), 'text-p'), 'Name'); assert.equal((await history()).undo, 0)
})

test('ordinary header and footnote input updates their control state while the body remains unchanged', async () => {
  await page.evaluate(() => { window.contentControlFixture.update('header-check', { type: 'text' }); window.contentControlFixture.update('note-list', { type: 'comboBox' }) })
  const body = (await snapshot()).blocks
  await page.evaluate(() => window.contentControlFixture.openHost('header')); await inputBridge().focus(); await page.keyboard.insertText('Header text')
  await expect.poll(async () => control(await snapshot(), 'header-check').value).toBe('Header text')
  assert.equal((await snapshot()).sectionHeaderFooters.first.header, 'Header text')
  await page.evaluate(() => window.contentControlFixture.openHost('note')); await inputBridge().focus(); await page.keyboard.insertText('Typed note')
  await expect.poll(async () => control(await snapshot(), 'note-list').value).toBe('Typed note')
  assert.equal((await snapshot()).footnotes.note.text, 'Typed note'); assert.deepEqual((await snapshot()).blocks, body)
})
