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
import { documentContentControlSelectionForId } from '/packages/office-studio/src/features/document/controllers/content-control-interaction'
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
const noteReference = paragraph('note-reference', '1'); noteReference.runs[0].metadata = { note: { type: 'footnote', noteId: 'note' } }; noteReference.runs[0].verticalAlign = 'superscript'
initial.blocks.push(noteReference)
initial.footnotes = { note: { id: 'note', text: 'Alpha', referenceBlockId: noteReference.id, referenceRunId: noteReference.runs[0].id, content: storyContent('note-p', 'Alpha') } }
initial.contentControls.push({ id: 'header-parent', type: 'group', blockId: 'header-p', startOffset: 0, endOffset: 1, alias: '页眉上层', lock: 'unlocked' },
  { id: 'header-check', type: 'checkbox', blockId: 'header-p', startOffset: 0, endOffset: 1, alias: '页眉确认', value: 'false', checked: false, parentId: 'header-parent', checkbox: copy(initial.contentControls[0].checkbox), lock: 'unlocked' },
  { id: 'note-list', type: 'dropDownList', blockId: 'note-p', startOffset: 0, endOffset: 5, alias: '脚注候选', value: 'A', items: [{ value: 'A', displayText: 'Alpha' }, { value: 'B', displayText: 'Beta' }], lock: 'unlocked' })
let studio = new OfficeStudio({ document: initial, currentUserId: 'editor' }); studio.activateDocument(initial.id)
const api = shallowRef(new FDocument(studio, initial.id)), model = shallowRef(api.value.getModel()), dark = ref(false), variant = ref('classic'), commands = [], notices = []
let unsubscribe, cancellation, reentry
const observerInputs = []
let savedImageFileReader
const pendingImageReads = []
function holdImageReads() {
  const Original = window.FileReader
  if (savedImageFileReader) throw new Error('Image reads are already held')
  savedImageFileReader = Original
  window.FileReader = class {
    result = null; error = null; onload = null; onerror = null
    readAsDataURL(file) {
      const reader = new Original()
      reader.onload = () => { this.result = reader.result; pendingImageReads.push(() => this.onload?.(new Event('load'))) }
      reader.onerror = () => { this.error = reader.error; this.onerror?.(new Event('error')) }
      reader.readAsDataURL(file)
    }
  }
}
function releaseImageReads() {
  if (savedImageFileReader) window.FileReader = savedImageFileReader
  savedImageFileReader = undefined
  for (const release of pendingImageReads.splice(0)) release()
}
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
function selectLogical(id, host) {
  const selection = documentContentControlSelectionForId(api.value.getModel(), id, host)
  if (!selection) throw new Error('The content control has no current logical selection')
  api.value.setInputSelection(selection, undefined, host)
  return copy(selection)
}
function selectRange(id, start, end = start, host) { const control = api.value.getContentControl(id).getInfo(); return api.value.setInputSelection({ anchor: { blockId: control.blockId, offset: start }, focus: { blockId: control.blockId, offset: end } }, undefined, host) }
function outside() { return api.value.setInputSelection({ anchor: { blockId: 'outside', offset: 0 }, focus: { blockId: 'outside', offset: 0 } }) }
function prepareTableBinding() {
  const value = api.value.getContentControl('text').getText()
  const table = api.value.insertTableFromData([[value]], { tableId: 'binding-table', position: 'documentEnd' })
  if (!table) throw new Error('The binding fixture table was not inserted')
  const block = api.value.getModel().blocks.find(block => block.id === table.getId()).metadata.table.rows[0].cells[0].blocks[0]
  const dataBinding = { storeItemID: '{TABLE-BINDING}', xpath: '/value', cachedResult: value, dirty: false }
  api.value.getContentControl('text').update({ dataBinding })
  api.value.addContentControl({ id: 'table-source', type: 'richText', blockId: block.id, startOffset: 0, endOffset: value.length,
    value, temporary: true, dataBinding })
  return table.getId()
}
function prepareRepeatingRows() {
  const table = api.value.insertTableFromData([['Header', 'H'], ['A', 'A2'], ['B', 'B2'], ['Tail', 'T']],
    { tableId: 'repeating-row-table', position: 'documentEnd' })
  if (!table) throw new Error('The repeating row fixture table was not inserted')
  table.selectRows(1, 2)
  return table.getId()
}
function prepareRepeatingBlocks() {
  const first = api.value.appendParagraph('Block A').getInfo().blockId
  const table = api.value.insertTableFromData([['Nested', 'Table']], { tableId: 'repeating-block-table', position: 'documentEnd' })
  if (!table) throw new Error('The repeating block fixture table was not inserted')
  const last = api.value.appendParagraph('Block B').getInfo().blockId
  api.value.setInputSelection({ anchor: { blockId: first, offset: 0 }, focus: { blockId: last, offset: 7 } })
  const blocks = api.value.getModel().blocks, start = blocks.findIndex(block => block.id === first), end = blocks.findIndex(block => block.id === last)
  return blocks.slice(start, end + 1).map(block => block.id)
}
function prepareParentTable() {
  const table = api.value.insertTableFromData([['L😀R']], { tableId: 'parent-selection-table', position: 'documentEnd' })
  if (!table) throw new Error('The parent selection fixture table was not inserted')
  const block = api.value.getModel().blocks.find(block => block.id === table.getId()).metadata.table.rows[0].cells[0].blocks[0]
  api.value.addContentControl({ id: 'table-parent', type: 'group', blockId: block.id, startOffset: 0, endOffset: 4, alias: '表格上层' })
  api.value.addContentControl({ id: 'table-child', type: 'text', blockId: block.id, startOffset: 1, endOffset: 3, value: '😀', alias: '表格子控件', parentId: 'table-parent' })
  const target = { blockId: table.getId(), row: 0, col: 0, cellBlockId: block.id }
  api.value.setInputSelection({ anchor: { ...target, offset: 1 }, focus: { ...target, offset: 3 } })
  return target
}
function preparePicture(inTable = false) {
  let target
  if (inTable) {
    const table = api.value.insertTableFromData([['P']], { tableId: 'picture-selection-table', position: 'documentEnd' })
    if (!table) throw new Error('The picture fixture table was not inserted')
    const block = api.value.getModel().blocks.find(block => block.id === table.getId()).metadata.table.rows[0].cells[0].blocks[0]
    target = { blockId: table.getId(), row: 0, col: 0, cellBlockId: block.id }
  } else target = { blockId: api.value.appendParagraph('P').getInfo().blockId }
  api.value.addContentControl({ id: 'picture-control', type: 'picture', blockId: target.cellBlockId ?? target.blockId,
    startOffset: 0, endOffset: 1, value: 'P', placeholder: 'Choose picture', alias: '图片控件' })
  api.value.setInputSelection({ anchor: { ...target, offset: 0 }, focus: { ...target, offset: 1 } })
  return target
}
function prepareXmlAlias() {
  const blockId = api.value.appendParagraph('A / A').getInfo().blockId, storeItemID = '{XML-ALIAS}'
  api.value.addContentControl({ id: 'xml-peer', type: 'richText', blockId, startOffset: 0, endOffset: 1, value: 'A',
    alias: '同段落关联', placeholder: 'Peer😀', dataBinding: { storeItemID, xpath: '/root/value/text()', cachedResult: 'A', dirty: false } })
  api.value.addContentControl({ id: 'xml-source', type: 'comboBox', blockId, startOffset: 4, endOffset: 5, value: 'A',
    alias: '同段落来源', placeholder: 'Source😀', temporary: true, items: [{ value: 'A', displayText: 'A' }],
    dataBinding: { storeItemID, xpath: '/root/value', cachedResult: 'A', dirty: false } })
  api.value.setCustomXmlParts([{ partName: 'customXml/alias.xml', storeItemID, xml: '<root><value>A</value></root>' }])
  return blockId
}
function armObserver(mode, commandId = 'doc.command.select-content-control-item', controlId = 'combo') {
  reentry?.dispose()
  reentry = studio.events.addEvent('CommandExecuted', event => {
    if (event.unitId !== api.value.id || event.id !== commandId || event.params?.controlId !== controlId) return
    reentry.dispose(); reentry = undefined
    const before = copy(api.value.getInputState()), revisionBefore = api.value.getModelRevision()
    if (mode === 'selection') outside()
    else if (mode === 'revision') api.value.getContentControl('text').update({ alias: 'observer title' })
    else throw new Error('Unknown observer scenario')
    observerInputs.push({ before, after: copy(api.value.getInputState()), revisionBefore, revisionAfter: api.value.getModelRevision() })
  })
}
function armParentSelectionObserver() {
  reentry?.dispose()
  reentry = studio.events.addEvent('DocumentInputStateChanged', event => {
    const location = event.state?.location
    if (event.unitId !== api.value.id || !event.state?.valid || location?.kind !== 'body'
      || location.selection.focus.blockId !== 'parent-selection-table'
      || location.selection.anchor.offset !== 0 || location.selection.focus.offset !== 4) return
    reentry.dispose(); reentry = undefined
    const before = copy(api.value.getInputState()), revisionBefore = api.value.getModelRevision()
    outside()
    observerInputs.push({ before, after: copy(api.value.getInputState()), revisionBefore, revisionAfter: api.value.getModelRevision() })
  })
}
window.contentControlFixture = {
  snapshot: () => copy(model.value), history: () => ({ undo: studio.editor.history.units[api.value.id]?.undo.length ?? 0, redo: studio.editor.history.units[api.value.id]?.redo.length ?? 0 }), commands: () => [...commands], notices: () => [...notices],
  input: () => { const state = api.value.getInputState(); return state ? copy(state) : null }, observerInputs: () => copy(observerInputs), armObserver, armParentSelectionObserver,
  listenerErrors: () => studio.getEventListenerErrors().map(failure => String(failure.error)),
  select: id => select(id), selectLogical, logicalSelection: id => copy(documentContentControlSelectionForId(api.value.getModel(), id)),
  selectRange, outside, prepareTableBinding, prepareParentTable, preparePicture, prepareXmlAlias, prepareRepeatingRows, prepareRepeatingBlocks,
  holdImageReads, releaseImageReads, imageReadCount: () => pendingImageReads.length,
  selectTable: () => api.value.getTable('binding-table').selectCell(0, 0),
  openHost: area => area === 'header' ? select('header-check', { kind: 'story', sectionId: 'first', story: 'header' }) : select('note-list', { kind: 'note', noteKind: 'footnote', noteId: 'note' }),
  update: (id, patch) => api.value.getContentControl(id).update(patch), setChecked: checked => api.value.getContentControl('check').setChecked(checked),
  setRules: rules => api.value.setPermissionRules(rules), undo: () => api.value.undo(), redo: () => api.value.redo(), appearance: (isDark, layout) => { dark.value = isDark; variant.value = layout },
  cancel: id => { cancellation?.dispose(); cancellation = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === id) event.cancel = true }) },
  saveReopen: () => { const saved = copy(api.value.saveModel()); unsubscribe?.dispose(); cancellation?.dispose(); reentry?.dispose(); studio.dispose(); studio = new OfficeStudio({ document: saved, currentUserId: 'editor' }); studio.activateDocument(saved.id); api.value = new FDocument(studio, saved.id); model.value = api.value.getModel(); observe(); return saved },
  replaceSource: () => { const source = copy(initial); source.id = 'replacement-sdt'; unsubscribe?.dispose(); cancellation?.dispose(); reentry?.dispose(); studio.dispose(); studio = new OfficeStudio({ document: source, currentUserId: 'editor' }); studio.activateDocument(source.id); api.value = new FDocument(studio, source.id); model.value = api.value.getModel(); observe() },
}
const app = createApp({ render: () => h(ConfigProvider, { theme: { algorithm: dark.value ? theme.darkAlgorithm : theme.defaultAlgorithm, token: { colorPrimary: '#3064c7' } } }, { default: () => h(DocumentWorkbench, {
  modelValue: model.value, documentApi: api.value, permissionUserId: 'editor', canUndo: studio.editor.canUndo, canRedo: studio.editor.canRedo, ribbonVariant: variant.value,
  'onUpdate:modelValue': value => { model.value = value }, onUndo: () => api.value.undo(), onRedo: () => api.value.redo(), onNotify: message => notices.push(message),
}) }) })
app.mount('#app'); window.addEventListener('pagehide', () => { if (savedImageFileReader) window.FileReader = savedImageFileReader; unsubscribe?.dispose(); cancellation?.dispose(); reentry?.dispose(); app.unmount(); studio.dispose() }, { once: true })
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
async function resetSelected(id) {
  await page.evaluate(id => window.contentControlFixture.select(id), id)
  await page.getByRole('tab', { name: '内容控件', exact: true }).click()
  const button = page.getByRole('button', { name: '恢复占位符', exact: true })
  await expect(button).toBeEnabled(); await button.click()
}
async function openPictureReplacement() {
  await page.getByRole('tab', { name: '内容控件', exact: true }).click()
  const button = page.getByRole('button', { name: '更换图片', exact: true })
  await expect(button).toBeEnabled(); await button.click()
  const dialog = page.getByRole('dialog', { name: '更换图片', exact: true }); await expect(dialog).toBeVisible()
  return dialog
}
const replacementPng = { name: 'replacement.png', mimeType: 'image/png',
  buffer: Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jD3EAAAAASUVORK5CYII=', 'base64') }
const setReplacementImage = dialog => dialog.locator('input[type="file"]').setInputFiles(replacementPng)

async function insertSelectedRepeatingSection() {
  await page.getByRole('tab', { name: '插入', exact: true }).click()
  await page.getByRole('button', { name: '内容控件选项', exact: true }).click()
  await page.getByRole('menuitem', { name: '重复节控件', exact: true }).click()
  await expect.poll(async () => (await snapshot()).contentControls.filter(control => control.type === 'repeatingSectionItem').length).toBe(1)
  const item = (await snapshot()).contentControls.find(control => control.type === 'repeatingSectionItem')
  await page.getByRole('tab', { name: '内容控件', exact: true }).click()
  return item
}
const duplicateItem = () => page.locator('[data-ribbon-item="doc-controls-duplicate-item"]:visible').getByRole('button', { name: '复制', exact: true })
const deleteItem = () => page.locator('[data-ribbon-item="doc-controls-delete-item"]:visible').getByRole('button', { name: '删除项目', exact: true })

test('Ribbon row insert, duplicate, deletion, undo/redo and snapshot reopen retain actual structural input selections', async () => {
  const tableId = await page.evaluate(() => window.contentControlFixture.prepareRepeatingRows()), stack = await history()
  const item = await insertSelectedRepeatingSection(), before = await snapshot()
  assert.equal(item.placement, 'row'); assert.equal(item.tableRowRange.rowIds.length, 2)
  assert.equal((await history()).undo, stack.undo + 1)
  assert.deepEqual((await page.evaluate(() => window.contentControlFixture.input())).location.selection,
    await page.evaluate(id => window.contentControlFixture.logicalSelection(id), item.id))
  await expect(duplicateItem()).toBeEnabled(); await duplicateItem().click()
  await expect.poll(async () => (await snapshot()).blocks.find(block => block.id === tableId).metadata.table.rows.length).toBe(6)
  const duplicated = await snapshot(), copyItem = duplicated.contentControls.find(control => control.type === 'repeatingSectionItem' && control.id !== item.id)
  assert.ok(copyItem && copyItem.parentId === item.parentId)
  assert.deepEqual((await page.evaluate(() => window.contentControlFixture.input())).location.selection,
    await page.evaluate(id => window.contentControlFixture.logicalSelection(id), copyItem.id))
  await expect(inputBridge()).toBeFocused(); await expect(deleteItem()).toBeEnabled(); await deleteItem().click()
  await expect.poll(async () => (await snapshot()).blocks.find(block => block.id === tableId).metadata.table.rows.length).toBe(4)
  assert.equal(control(await snapshot(), copyItem.id), undefined); assert.equal((await history()).undo, stack.undo + 3)
  await page.evaluate(() => window.contentControlFixture.undo()); assert.deepEqual((await snapshot()).contentControls, duplicated.contentControls)
  await page.evaluate(() => window.contentControlFixture.redo()); assert.deepEqual((await snapshot()).contentControls, before.contentControls)
  const saved = await page.evaluate(() => window.contentControlFixture.saveReopen())
  assert.deepEqual((await snapshot()).contentControls, saved.contentControls)
  assert.equal((await snapshot()).blocks.find(block => block.id === tableId).metadata.table.rows.length, 4)
  assert.deepEqual(await page.evaluate(() => window.contentControlFixture.notices()), [])
})

test('Ribbon complete-block insertion and copying select the full copied tree including a nested table', async () => {
  const originalBlocks = await page.evaluate(() => window.contentControlFixture.prepareRepeatingBlocks()), stack = await history()
  const item = await insertSelectedRepeatingSection(); assert.deepEqual(item.blockRange.blockIds, originalBlocks)
  await duplicateItem().click()
  await expect.poll(async () => (await snapshot()).contentControls.filter(control => control.type === 'repeatingSectionItem').length).toBe(2)
  const current = await snapshot(), copied = current.contentControls.find(control => control.type === 'repeatingSectionItem' && control.id !== item.id)
  assert.equal(copied.blockRange.blockIds.length, originalBlocks.length); assert.ok(copied.blockRange.blockIds.every(id => !originalBlocks.includes(id)))
  assert.ok(current.blocks.some(block => copied.blockRange.blockIds.includes(block.id) && block.type === 'table'))
  assert.deepEqual((await page.evaluate(() => window.contentControlFixture.input())).location.selection,
    await page.evaluate(id => window.contentControlFixture.logicalSelection(id), copied.id))
  await expect(inputBridge()).toBeFocused(); assert.equal((await history()).undo, stack.undo + 2)
  await page.evaluate(id => window.contentControlFixture.selectLogical(id), copied.id); await deleteItem().click()
  await expect.poll(async () => control(await snapshot(), copied.id)).toBe(undefined)
  assert.deepEqual(control(await snapshot(), item.id).blockRange.blockIds, originalBlocks)
  const saved = await page.evaluate(() => window.contentControlFixture.saveReopen()); assert.deepEqual((await snapshot()).contentControls, saved.contentControls)
})

test('a canceled Ribbon row duplication retains the full row range, input session and content history', async () => {
  await page.evaluate(() => window.contentControlFixture.prepareRepeatingRows()); const item = await insertSelectedRepeatingSection()
  await page.evaluate(id => window.contentControlFixture.selectLogical(id), item.id)
  const before = await snapshot(), stack = await history(), input = await page.evaluate(() => window.contentControlFixture.input())
  await page.evaluate(() => window.contentControlFixture.cancel('doc.command.duplicate-repeating-section-item'))
  await duplicateItem().click(); await expect(inputBridge()).toBeFocused()
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), stack)
  assert.deepEqual(await page.evaluate(() => window.contentControlFixture.input()), input)
  assert.deepEqual((await page.evaluate(() => window.contentControlFixture.commands())).filter(id => id === 'doc.command.duplicate-repeating-section-item'), [])
})

test('a repeating command observer keeps its newer outside selection instead of the completed copied row selection', async () => {
  await page.evaluate(() => window.contentControlFixture.prepareRepeatingRows()); const item = await insertSelectedRepeatingSection(), stack = await history()
  await page.evaluate(id => window.contentControlFixture.armObserver('selection', 'doc.command.duplicate-repeating-section-item', id), item.id)
  await duplicateItem().click(); await expect.poll(async () => (await snapshot()).contentControls.filter(control => control.type === 'repeatingSectionItem').length).toBe(2)
  const observed = (await page.evaluate(() => window.contentControlFixture.observerInputs()))[0]
  assert.equal(observed.after.valid, true); assert.equal(observed.revisionBefore, observed.revisionAfter)
  assert.deepEqual(observed.after.location.selection, { anchor: { blockId: 'outside', offset: 0 }, focus: { blockId: 'outside', offset: 0 } })
  await expect.poll(() => page.evaluate(() => window.contentControlFixture.input())).toEqual(observed.after)
  assert.equal((await history()).undo, stack.undo + 1); await expect(inputBridge()).toBeFocused()
})

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

test('a structural cell input clears prior text-control affinity and commits its bound peers with one history item', async () => {
  await page.evaluate(() => window.contentControlFixture.selectRange('text', 1))
  await inputBridge().focus(); await page.keyboard.insertText('X')
  await expect.poll(async () => control(await snapshot(), 'text').value).toBe('NXame')
  await page.evaluate(() => window.contentControlFixture.prepareTableBinding())
  const before = await snapshot(), stack = await history()
  await page.evaluate(() => window.contentControlFixture.selectTable())
  await inputBridge().focus(); await page.keyboard.insertText('李😀')
  await expect.poll(async () => control(await snapshot(), 'text').value).toBe('李😀')
  const current = await snapshot(), table = current.blocks.find(block => block.id === 'binding-table')
  assert.equal(table.metadata.table.rows[0].cells[0].blocks[0].runs.map(run => run.text).join(''), '李😀')
  assert.equal(control(current, 'table-source'), undefined); assert.equal(control(current, 'text').dataBinding.cachedResult, '李😀')
  assert.equal((await history()).undo, stack.undo + 1); assert.deepEqual(await page.evaluate(() => window.contentControlFixture.notices()), [])
  assert.equal((await page.evaluate(() => window.contentControlFixture.input())).location.selection.focus.offset, 3)
  await expect(inputBridge()).toBeFocused()
  await page.evaluate(() => window.contentControlFixture.undo())
  await expect.poll(async () => !!control(await snapshot(), 'table-source')).toBe(true)
  assert.deepEqual((await snapshot()).contentControls, before.contentControls)
  await page.evaluate(() => window.contentControlFixture.redo()); await expect.poll(async () => control(await snapshot(), 'text').value).toBe('李😀')
  await page.evaluate(() => window.contentControlFixture.saveReopen()); await expect.poll(async () => control(await snapshot(), 'text').value).toBe('李😀')
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

test('selecting a table parent uses its real row, column and paragraph coordinates without content history', async () => {
  const target = await page.evaluate(() => window.contentControlFixture.prepareParentTable())
  await page.getByRole('tab', { name: '内容控件', exact: true }).click()
  const before = await snapshot(), previousHistory = await history(), inputBefore = await page.evaluate(() => window.contentControlFixture.input())
  const button = page.getByRole('button', { name: '选择上层控件', exact: true })
  await expect(button).toBeEnabled(); await button.click()
  await expect.poll(async () => (await page.evaluate(() => window.contentControlFixture.input())).location.selection).toEqual({
    anchor: { ...target, offset: 0 }, focus: { ...target, offset: 4 },
  })
  const input = await page.evaluate(() => window.contentControlFixture.input())
  assert.equal(input.valid, true); assert.equal(input.modelRevision, inputBefore.modelRevision); assert.equal(input.generation, inputBefore.generation + 1)
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), previousHistory)
  await expect(inputBridge()).toBeFocused(); await expect(button).toBeDisabled()
  assert.deepEqual(await page.evaluate(() => window.contentControlFixture.notices()), [])
})

test('an equal-range header parent preserves its auxiliary host, focus and content revision', async () => {
  await page.evaluate(() => window.contentControlFixture.openHost('header'))
  await page.getByRole('tab', { name: '内容控件', exact: true }).click()
  const before = await snapshot(), previousHistory = await history(), inputBefore = await page.evaluate(() => window.contentControlFixture.input())
  const button = page.getByRole('button', { name: '选择上层控件', exact: true })
  await expect(button).toBeEnabled(); await button.click()
  await expect(button).toBeDisabled(); await expect(inputBridge()).toBeFocused()
  const input = await page.evaluate(() => window.contentControlFixture.input())
  assert.equal(input.valid, true); assert.equal(input.modelRevision, inputBefore.modelRevision); assert.equal(input.generation, inputBefore.generation + 1)
  assert.deepEqual(input.location, { kind: 'body', host: { kind: 'story', sectionId: 'first', story: 'header' },
    selection: { anchor: { blockId: 'header-p', offset: 0 }, focus: { blockId: 'header-p', offset: 1 } } })
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), previousHistory)
})

test('a parent-selection observer retains its newer outside caret without an extra renewal or old parent focus', async () => {
  await page.evaluate(() => window.contentControlFixture.prepareParentTable())
  await page.getByRole('tab', { name: '内容控件', exact: true }).click()
  const before = await snapshot(), previousHistory = await history()
  await page.evaluate(() => window.contentControlFixture.armParentSelectionObserver())
  const button = page.getByRole('button', { name: '选择上层控件', exact: true })
  await expect(button).toBeEnabled(); await button.click()
  const observations = await page.evaluate(() => window.contentControlFixture.observerInputs()); assert.equal(observations.length, 1)
  const observed = observations[0]
  assert.equal(observed.after.valid, true); assert.equal(observed.revisionAfter, observed.revisionBefore)
  assert.equal(observed.after.generation, observed.before.generation + 1)
  assert.deepEqual(observed.after.location.selection, { anchor: { blockId: 'outside', offset: 0 }, focus: { blockId: 'outside', offset: 0 } })
  await expect.poll(() => page.evaluate(() => window.contentControlFixture.input())).toEqual(observed.after)
  await expect(inputBridge()).toBeFocused(); assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), previousHistory)
  assert.deepEqual(await page.evaluate(() => window.contentControlFixture.notices()), [])
})

test('the antd image upload replaces a table picture, keeps a valid anchor and supports no-op, history and snapshot reopen', async () => {
  const target = await page.evaluate(() => window.contentControlFixture.preparePicture(true))
  const before = await snapshot(), previousHistory = await history()
  let dialog = await openPictureReplacement(); await setReplacementImage(dialog)
  await expect.poll(async () => control(await snapshot(), 'picture-control').metadata?.pictureBlockId).toBeTruthy()
  await expect(dialog).toBeHidden(); await expect(inputBridge()).toBeFocused()
  const after = await snapshot(), picture = control(after, 'picture-control'), pictureId = picture.metadata.pictureBlockId
  const cell = after.blocks.find(block => block.id === target.blockId).metadata.table.rows[0].cells[0]
  const images = cell.blocks.filter(block => block.type === 'image')
  assert.equal(images.length, 1); assert.equal(images[0].id, pictureId); assert.equal(images[0].metadata.image.alt, replacementPng.name)
  assert.match(images[0].metadata.image.src, /^data:image\/png;base64,/); assert.ok(images[0].metadata.image.width > 0); assert.ok(images[0].metadata.image.height > 0)
  assert.equal(picture.startOffset, 0); assert.equal(picture.endOffset, 0); assert.equal(picture.showingPlaceholder, false)
  const input = await page.evaluate(() => window.contentControlFixture.input())
  assert.equal(input.valid, true); assert.deepEqual(input.location.selection, { anchor: { ...target, offset: 0 }, focus: { ...target, offset: 0 } })
  assert.equal((await history()).undo, previousHistory.undo + 1)
  dialog = await openPictureReplacement(); await setReplacementImage(dialog); await expect(dialog).toBeHidden(); await expect(inputBridge()).toBeFocused()
  assert.deepEqual(await snapshot(), after); assert.equal((await history()).undo, previousHistory.undo + 1)
  assert.deepEqual(await page.evaluate(() => window.contentControlFixture.input()), input)
  await page.evaluate(() => window.contentControlFixture.undo()); assert.deepEqual((await snapshot()).contentControls, before.contentControls); assert.deepEqual((await snapshot()).blocks, before.blocks)
  await page.evaluate(() => window.contentControlFixture.redo()); assert.deepEqual((await snapshot()).contentControls, after.contentControls); assert.deepEqual((await snapshot()).blocks, after.blocks)
  const saved = await page.evaluate(() => window.contentControlFixture.saveReopen())
  assert.equal(control(await snapshot(), 'picture-control').metadata.pictureBlockId, pictureId); assert.deepEqual((await snapshot()).contentControls, saved.contentControls)
  assert.deepEqual(await page.evaluate(() => window.contentControlFixture.notices()), [])
})

test('a canceled picture command closes the upload dialog and preserves the selected model, history and input session', async () => {
  await page.evaluate(() => window.contentControlFixture.preparePicture())
  const dialog = await openPictureReplacement(), before = await snapshot(), previousHistory = await history(), inputBefore = await page.evaluate(() => window.contentControlFixture.input())
  await page.evaluate(() => window.contentControlFixture.cancel('doc.command.replace-picture-content-control-image'))
  await setReplacementImage(dialog); await expect(dialog).toBeHidden(); await expect(inputBridge()).toBeFocused()
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), previousHistory)
  assert.deepEqual(await page.evaluate(() => window.contentControlFixture.input()), inputBefore)
  assert.deepEqual((await page.evaluate(() => window.contentControlFixture.commands())).filter(id => id === 'doc.command.replace-picture-content-control-image'), [])
})

test('a picture command observer keeps its newer outside caret through modal close without a stale selection renewal', async () => {
  await page.evaluate(() => window.contentControlFixture.preparePicture())
  await page.evaluate(() => window.contentControlFixture.armObserver('selection', 'doc.command.replace-picture-content-control-image', 'picture-control'))
  const previousHistory = await history(), dialog = await openPictureReplacement(); await setReplacementImage(dialog)
  await expect(dialog).toBeHidden(); await expect(inputBridge()).toBeFocused()
  const observations = await page.evaluate(() => window.contentControlFixture.observerInputs()); assert.equal(observations.length, 1)
  const observed = observations[0]
  assert.equal(observed.after.valid, true); assert.equal(observed.revisionBefore, observed.revisionAfter); assert.equal(observed.after.generation, observed.before.generation + 1)
  assert.deepEqual(observed.after.location.selection, { anchor: { blockId: 'outside', offset: 0 }, focus: { blockId: 'outside', offset: 0 } })
  await expect.poll(() => page.evaluate(() => window.contentControlFixture.input())).toEqual(observed.after)
  assert.ok(control(await snapshot(), 'picture-control').metadata.pictureBlockId); assert.equal((await history()).undo, previousHistory.undo + 1)
})

for (const mode of ['selection', 'revision', 'host', 'reopen', 'permission']) {
  test('an image read held across ' + mode + ' changes cannot publish into a newer picture context', async () => {
    const target = await page.evaluate(() => window.contentControlFixture.preparePicture())
    const dialog = await openPictureReplacement()
    await page.evaluate(() => window.contentControlFixture.holdImageReads()); await setReplacementImage(dialog)
    await expect.poll(() => page.evaluate(() => window.contentControlFixture.imageReadCount())).toBe(1)
    await page.evaluate(({ mode, target }) => {
      const fixture = window.contentControlFixture
      if (mode === 'selection') fixture.outside()
      else if (mode === 'revision') fixture.update('text', { alias: 'Changed during upload' })
      else if (mode === 'host') fixture.openHost('header')
      else if (mode === 'reopen') fixture.saveReopen()
      else fixture.setRules([{ id: 'deny-picture-upload', scope: 'paragraph', targetId: target.cellBlockId ?? target.blockId, action: 'edit', effect: 'deny', userIds: ['editor'] }])
    }, { mode, target })
    await expect(dialog).toBeHidden()
    const beforeRelease = await snapshot(), previousHistory = await history(), inputBefore = await page.evaluate(() => window.contentControlFixture.input())
    await page.evaluate(() => window.contentControlFixture.releaseImageReads())
    await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))))
    assert.deepEqual(await snapshot(), beforeRelease); assert.deepEqual(await history(), previousHistory)
    assert.deepEqual(await page.evaluate(() => window.contentControlFixture.input()), inputBefore)
    assert.equal(control(await snapshot(), 'picture-control').metadata?.pictureBlockId, undefined)
    assert.deepEqual((await page.evaluate(() => window.contentControlFixture.commands())).filter(id => id === 'doc.command.replace-picture-content-control-image'), [])
  })
}

test('picture replacement disables for a content lock or edit permission and remains available with a wrapper lock', async () => {
  const target = await page.evaluate(() => window.contentControlFixture.preparePicture())
  await page.getByRole('tab', { name: '内容控件', exact: true }).click()
  const button = page.getByRole('button', { name: '更换图片', exact: true })
  await expect(button).toBeEnabled()
  await page.evaluate(() => window.contentControlFixture.update('picture-control', { lock: 'sdtLocked' })); await expect(button).toBeEnabled()
  await page.evaluate(() => window.contentControlFixture.update('picture-control', { lock: 'contentLocked' })); await expect(button).toBeDisabled()
  await page.evaluate(() => window.contentControlFixture.update('picture-control', { lock: 'unlocked' })); await expect(button).toBeEnabled()
  await page.evaluate(target => window.contentControlFixture.setRules([{ id: 'deny-picture', scope: 'paragraph', targetId: target.cellBlockId ?? target.blockId, action: 'edit', effect: 'deny', userIds: ['editor'] }]), target)
  await expect(button).toBeDisabled(); assert.equal(control(await snapshot(), 'picture-control').metadata?.pictureBlockId, undefined)
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

test('Ribbon reset restores a temporary prompt caret, continued input and the same undo/redo snapshot history', async () => {
  await page.evaluate(() => window.contentControlFixture.update('text', { temporary: true, placeholder: '请填写😀' }))
  const before = await snapshot(), previousHistory = await history()
  await resetSelected('text'); await expect.poll(async () => blockText(await snapshot(), 'text-p')).toBe('请填写😀')
  assert.equal(control(await snapshot(), 'text'), undefined); assert.equal((await history()).undo, previousHistory.undo + 1)
  await expect(inputBridge()).toBeFocused()
  const input = await page.evaluate(() => window.contentControlFixture.input())
  assert.equal(input.valid, true); assert.deepEqual(input.location.selection, {
    anchor: { blockId: 'text-p', offset: '请填写😀'.length }, focus: { blockId: 'text-p', offset: '请填写😀'.length },
  })
  assert.deepEqual((await page.evaluate(() => window.contentControlFixture.commands())).filter(id => id === 'doc.command.reset-content-control-placeholder'), ['doc.command.reset-content-control-placeholder'])
  await page.keyboard.insertText('后续'); await expect.poll(async () => blockText(await snapshot(), 'text-p')).toBe('请填写😀后续')
  await page.evaluate(() => window.contentControlFixture.undo()); assert.equal(blockText(await snapshot(), 'text-p'), '请填写😀')
  await page.evaluate(() => window.contentControlFixture.undo()); assert.deepEqual((await snapshot()).contentControls, before.contentControls)
  assert.equal(blockText(await snapshot(), 'text-p'), 'Name')
  await page.evaluate(() => window.contentControlFixture.redo()); assert.equal(control(await snapshot(), 'text'), undefined)
  const saved = await page.evaluate(() => window.contentControlFixture.saveReopen())
  await expect.poll(async () => blockText(await snapshot(), 'text-p')).toBe('请填写😀')
  assert.equal(control(saved, 'text'), undefined); assert.deepEqual(await history(), { undo: 0, redo: 0 })
})

test('Ribbon reset clears XML aliases in one paragraph and focuses the source prompt after the surviving peer', async () => {
  const blockId = await page.evaluate(() => window.contentControlFixture.prepareXmlAlias()), before = await snapshot(), previousHistory = await history()
  await resetSelected('xml-source')
  await expect.poll(async () => blockText(await snapshot(), blockId)).toBe('Peer😀 / Source😀')
  const after = await snapshot(), peer = control(after, 'xml-peer')
  assert.equal(control(after, 'xml-source'), undefined); assert.equal(peer.showingPlaceholder, true)
  assert.equal(peer.dataBinding.cachedResult, ''); assert.equal(peer.dataBinding.dirty, false)
  assert.deepEqual([peer.startOffset, peer.endOffset], [0, 'Peer😀'.length])
  assert.equal(await page.evaluate(() => new DOMParser().parseFromString(window.contentControlFixture.snapshot().customXmlParts[0].xml, 'application/xml').documentElement.firstElementChild.textContent), '')
  assert.equal((await history()).undo, previousHistory.undo + 1); await expect(inputBridge()).toBeFocused()
  const input = await page.evaluate(() => window.contentControlFixture.input()), offset = 'Peer😀 / Source😀'.length
  assert.equal(input.valid, true); assert.deepEqual(input.location.selection, { anchor: { blockId, offset }, focus: { blockId, offset } })
  await page.evaluate(() => window.contentControlFixture.undo()); assert.equal(blockText(await snapshot(), blockId), 'A / A')
  assert.deepEqual((await snapshot()).customXmlParts, before.customXmlParts)
  await page.evaluate(() => window.contentControlFixture.redo()); assert.deepEqual((await snapshot()).contentControls, after.contentControls)
  assert.deepEqual((await snapshot()).customXmlParts, after.customXmlParts)
})

test('a temporary combo alias edit uses the production overlay and focuses after every same-paragraph replacement', async () => {
  const blockId = await page.evaluate(() => window.contentControlFixture.prepareXmlAlias()), previousHistory = await history()
  await page.evaluate(() => window.contentControlFixture.select('xml-source'))
  const editor = await openEditor('同段落来源'); await editor.getByRole('combobox').fill('中😀')
  await editor.getByRole('button', { name: '确定', exact: true }).click(); await expect(editor).toHaveCount(0)
  await expect.poll(async () => blockText(await snapshot(), blockId)).toBe('中😀 / 中😀')
  assert.equal(control(await snapshot(), 'xml-source'), undefined); assert.equal(control(await snapshot(), 'xml-peer').value, '中😀')
  assert.equal((await history()).undo, previousHistory.undo + 1); await expect(inputBridge()).toBeFocused()
  const input = await page.evaluate(() => window.contentControlFixture.input()), offset = '中😀 / 中😀'.length
  assert.equal(input.valid, true); assert.deepEqual(input.location.selection, { anchor: { blockId, offset }, focus: { blockId, offset } })
  assert.equal(await page.evaluate(() => new DOMParser().parseFromString(window.contentControlFixture.snapshot().customXmlParts[0].xml, 'application/xml').documentElement.firstElementChild.textContent), '中😀')
  await page.keyboard.insertText('后续'); await expect.poll(async () => blockText(await snapshot(), blockId)).toBe('中😀 / 中😀后续')
  assert.equal(control(await snapshot(), 'xml-peer').value, '中😀')
})

test('a canceled Ribbon reset preserves its temporary shell, model, history and valid selected focus', async () => {
  await page.evaluate(() => { window.contentControlFixture.update('text', { temporary: true, placeholder: 'Prompt😀' }); window.contentControlFixture.cancel('doc.command.reset-content-control-placeholder') })
  const before = await snapshot(), previousHistory = await history()
  await resetSelected('text'); assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), previousHistory)
  assert.equal(control(await snapshot(), 'text').temporary, true); assert.equal(control(await snapshot(), 'text').showingPlaceholder, false)
  await expect(inputBridge()).toBeFocused()
  const input = await page.evaluate(() => window.contentControlFixture.input())
  assert.equal(input.valid, true); assert.deepEqual(input.location.selection, { anchor: { blockId: 'text-p', offset: 0 }, focus: { blockId: 'text-p', offset: 4 } })
  assert.deepEqual((await page.evaluate(() => window.contentControlFixture.commands())).filter(id => id === 'doc.command.reset-content-control-placeholder'), [])
})

test('a reset observer selection remains the focused current caret after temporary removal', async () => {
  await page.evaluate(() => { window.contentControlFixture.update('combo', { temporary: true, placeholder: 'Prompt😀' }); window.contentControlFixture.armObserver('selection', 'doc.command.reset-content-control-placeholder', 'combo') })
  const previousHistory = await history()
  await resetSelected('combo'); await expect.poll(async () => blockText(await snapshot(), 'combo-p')).toBe('Prompt😀')
  const observations = await page.evaluate(() => window.contentControlFixture.observerInputs()); assert.equal(observations.length, 1)
  const observed = observations[0]
  assert.equal(observed.after.valid, true); assert.equal(observed.after.generation, observed.before.generation + 1)
  assert.equal(observed.revisionAfter, observed.revisionBefore); assert.equal((await history()).undo, previousHistory.undo + 1)
  await expect.poll(() => page.evaluate(() => window.contentControlFixture.input())).toEqual(observed.after)
  assert.deepEqual(observed.after.location.selection, { anchor: { blockId: 'outside', offset: 0 }, focus: { blockId: 'outside', offset: 0 } })
  await expect(inputBridge()).toBeFocused(); await page.keyboard.insertText('Z')
  await expect.poll(async () => blockText(await snapshot(), 'outside')).toBe('ZOutside')
  assert.equal(blockText(await snapshot(), 'combo-p'), 'Prompt😀')
})

test('a reset observer content commit keeps its newer revision and invalid input session', async () => {
  await page.evaluate(() => { window.contentControlFixture.update('combo', { temporary: true, placeholder: 'Prompt😀' }); window.contentControlFixture.armObserver('revision', 'doc.command.reset-content-control-placeholder', 'combo') })
  const previousHistory = await history()
  await resetSelected('combo'); await expect.poll(async () => control(await snapshot(), 'text').alias).toBe('observer title')
  const observations = await page.evaluate(() => window.contentControlFixture.observerInputs()); assert.equal(observations.length, 1)
  const observed = observations[0]
  assert.equal(observed.revisionAfter, observed.revisionBefore + 1); assert.equal(observed.after.generation, observed.before.generation)
  assert.equal(observed.after.valid, false); assert.equal((await history()).undo, previousHistory.undo + 2)
  await expect.poll(() => page.evaluate(() => window.contentControlFixture.input())).toEqual(observed.after)
  assert.equal(blockText(await snapshot(), 'combo-p'), 'Prompt😀'); assert.equal(control(await snapshot(), 'combo'), undefined)
  await expect(inputBridge()).not.toBeFocused()
})

test('Ribbon reset uses the shared source permission and supported-type capability', async () => {
  await page.evaluate(() => { window.contentControlFixture.update('text', { placeholder: 'Prompt😀' }); window.contentControlFixture.select('text') })
  await page.getByRole('tab', { name: '内容控件', exact: true }).click()
  await expect(page.getByRole('button', { name: '恢复占位符', exact: true })).toBeEnabled()
  await page.evaluate(() => window.contentControlFixture.setRules([{ id: 'deny-reset', scope: 'paragraph', targetId: 'text-p', action: 'edit', effect: 'deny', userIds: ['editor'] }]))
  await expect(page.getByRole('button', { name: '恢复占位符', exact: true })).toBeDisabled()
  await page.evaluate(() => { window.contentControlFixture.update('check', { placeholder: 'Checkbox prompt' }); window.contentControlFixture.select('check') })
  await expect(page.getByRole('button', { name: '恢复占位符', exact: true })).toBeDisabled()
  assert.deepEqual((await page.evaluate(() => window.contentControlFixture.commands())).filter(id => id === 'doc.command.reset-content-control-placeholder'), [])
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
