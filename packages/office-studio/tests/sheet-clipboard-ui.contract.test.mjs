import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { nextTick, shallowRef } from 'vue'

// Draft only. This suite is written for the user's later unified verification pass.
let server, sdk, clipboardUi, clipboardContent
beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[sdk, clipboardUi, clipboardContent] = await Promise.all([
    '/packages/office-studio/src/index.ts',
    '/packages/office-studio/src/components/spreadsheet/use-sheet-clipboard.ts',
    '/packages/office-studio/src/features/sheet-workbench/sheet-clipboard-content.ts',
  ].map(path => server.ssrLoadModule(path)))
})
afterAll(async () => { await server?.close() })
const copy = value => JSON.parse(JSON.stringify(value))
const range = (row, col, endRow = row, endCol = col) => ({ startRow: row, startCol: col, endRow, endCol })
const apiRange = value => ({ startRow: value.startRow, endRow: value.endRow, startColumn: value.startCol, endColumn: value.endCol })
const domainRange = value => ({ startRow: value.startRow, endRow: value.endRow, startCol: value.startColumn, endCol: value.endColumn })
function deferred() { let resolve, reject; const promise = new Promise((yes, no) => { resolve = yes; reject = no }); return { promise, resolve, reject } }
function model() {
  const book = sdk.createBlankWorkbook(); book.id = 'clipboard-ui-book'; book.calculationMode = 'manual'
  book.sheets = ['Source', 'Report'].map((name, index) => ({ ...sdk.createBlankSheet(name), id: 'clipboard-ui-sheet-' + index,
    rowCount: 40, columnCount: 12, defaultRowHeight: 26, defaultColumnWidth: 96 }))
  book.activeSheetId = book.sheets[0].id
  const sheet = book.sheets[0], rich = sdk.createCell(0, 0, 'rich')
  rich.richText = [{ text: 'ri', font: { bold: true }, vendor: { retain: 1 } }, { text: 'ch', font: { italic: true } }]
  rich.hyperlink = 'https://example.test'; rich.metadata = { vendor: { retain: true } }; rich.style.fill = '#ffeeaa'
  sheet.cells['0:0'] = rich
  sheet.cells['0:1'] = { ...sdk.createCell(0, 1, 5), formula: '=A2+$A$2', dataType: 'formula', cachedDataType: 'number', display: '5' }
  sheet.cells['1:0'] = sdk.createCell(1, 0, 5); sheet.columnWidths = { 0: 123, 1: 156 }
  return book
}
function host(t, options = {}) {
  const original = options.model ?? model(), studio = options.standalone ? undefined
    : new sdk.OfficeStudio({ workbook: original, sheetTextMeasure: text => [...text].length * 8 })
  if (studio) t.onTestFinished(() => studio.dispose())
  const owner = shallowRef(studio), view = shallowRef(sdk.toExcelWorkbookView(studio ? studio.getWorkbookModel(original.id) : original))
  if (options.standalone) delete view.value.id
  const sheetId = shallowRef(view.value.activeSheetId), selected = shallowRef(range(0, 0, 0, 1)), active = shallowRef(true), editing = shallowRef(false)
  const currentUserId = shallowRef(options.userId ?? 'local-user')
  const hidePasteOptions = shallowRef(options.hidePasteOptions ?? false), permission = shallowRef(options.permission)
  const exclusiveRange = shallowRef(false), shapeTextEditing = shallowRef(false)
  const calls = { applied: [], changed: [], selected: [], focused: [], notices: [], written: [], reads: 0, commands: [] }
  studio?.events.addEvent('BeforeCommandExecute', event => calls.commands.push({ id: event.id, params: copy(event.params ?? {}) }))
  const context = {
    studio: () => owner.value, workbook: () => view.value, sheetId: () => sheetId.value, range: () => selected.value,
    userId: () => currentUserId.value, region: () => 'en-US', active: () => active.value, editing: () => editing.value,
    hidePasteOptions: () => hidePasteOptions.value, exclusiveRange: () => exclusiveRange.value, shapeTextEditing: () => shapeTextEditing.value,
    get permission() { return permission.value },
    finishEdit: () => !editing.value,
    prepareView: next => options.prepareView ? options.prepareView(next) : next,
    applyView: next => {
      view.value = next; sheetId.value = next.activeSheetId; calls.applied.push(copy(next))
      const state = owner.value?.getSelectionState(next.id, sheetId.value), actual = state?.ranges[state.activeRangeIndex] ?? state?.ranges.at(-1)
      if (actual) selected.value = domainRange(actual)
      options.onApply?.(context)
    },
    selection: ranges => { calls.selected.push(copy(ranges)); selected.value = copy(ranges.at(-1)); },
    changed: value => { if (value.current()) { calls.changed.push(copy(view.value)); options.onChanged?.(context) } },
    focus: () => calls.focused.push(copy(selected.value)), notify: value => calls.notices.push(copy(value)),
    systemClipboard: {
      canRead: () => options.readerAvailable ?? true,
      read: async () => { calls.reads++; return options.read ? options.read() : calls.written.at(-1)
        ? { text: calls.written.at(-1).plain, html: calls.written.at(-1).html } : undefined },
      write: async value => { calls.written.push(copy(value)); return options.write ? options.write(value) : true },
    },
  }
  const select = (value, id = sheetId.value) => {
    if (owner.value) {
      const book = sdk.FOffice.newAPI(owner.value).getWorkbook(view.value.id)
      book.getSheetBySheetId(id).activate(); owner.value.setSelections(view.value.id, id, [apiRange(value)])
    }
    view.value.activeSheetId = id; sheetId.value = id; selected.value = copy(value)
  }
  select(selected.value)
  const ui = clipboardUi.useSheetClipboard(context)
  const current = () => owner.value ? owner.value.getWorkbookModel(view.value.id) : sdk.applyExcelWorkbookView(original, view.value, 'en-US')
  const cell = (row, col, index = 0) => current().sheets[index].cells[row + ':' + col]
  const history = () => {
    const unit = owner.value?.editor.history.units[view.value.id]
    return { undo: unit?.undo.length ?? 0, redo: unit?.redo.length ?? 0 }
  }
  return { ui, context, calls, owner, view, sheetId, selected, active, editing, currentUserId, hidePasteOptions, permission,
    exclusiveRange, shapeTextEditing, original, select, current, cell, history }
}
function clipboardEvent(content = {}) {
  const data = new Map(Object.entries(content))
  return { defaultPrevented: false, stopped: false, target: undefined,
    preventDefault() { this.defaultPrevented = true }, stopPropagation() { this.stopped = true },
    clipboardData: { getData: key => data.get(key) ?? '', setData: (key, value) => data.set(key, value) }, data }
}
function keyEvent(key, extra = {}) { return { key, ctrlKey: true, metaKey: false, shiftKey: false, altKey: false, isComposing: false,
  keyCode: 0, defaultPrevented: false, preventDefault() { this.defaultPrevented = true }, ...extra } }

test('native copy writes plain and HTML flavours while cut retains its source until a paste commits', t => {
  const h = host(t), before = copy(h.current()), event = clipboardEvent()
  h.ui.handleCopy(event)
  assert.equal(event.defaultPrevented, true); assert.equal(event.stopped, true); assert.equal(event.data.get('text/plain'), 'rich\t5')
  assert.ok(clipboardContent.sheetClipboardIdFromHtml(event.data.get('text/html')))
  assert.deepEqual(h.ui.copyRange.value, range(0, 0, 0, 1)); assert.deepEqual(h.current(), before); assert.deepEqual(h.history(), { undo: 0, redo: 0 })
  h.ui.handleCopy(clipboardEvent(), true)
  assert.equal(h.owner.value.sheetClipboard.clipboard.copyType, 'cut'); assert.deepEqual(h.current(), before)
})

test('copy-formulas uses the same snapshot command and leaves nonformula source cells absent', async t => {
  const h = host(t), result = await h.ui.runClipboardCommand('sheet.command.copy-formula-only')
  assert.equal(result.accepted, true); assert.equal(result.clipboard.formulaOnly, true); assert.equal(result.plain, '\t=A2+$A$2')
  assert.deepEqual(Object.keys(result.clipboard.cells), ['0:1']); assert.equal(h.calls.written.length, 1); assert.equal(h.history().undo, 0)
})

test('a matching native HTML token keeps complete rich runs, hyperlinks, metadata and relative formula references', t => {
  const h = host(t), event = clipboardEvent(), source = copy(h.cell(0, 0))
  h.ui.handleCopy(event); h.select(range(4, 4)); h.ui.handlePaste(clipboardEvent(Object.fromEntries(event.data)))
  assert.deepEqual(h.cell(4, 4).richText, source.richText); assert.equal(h.cell(4, 4).hyperlink, source.hyperlink)
  assert.deepEqual(h.cell(4, 4).metadata.vendor, source.metadata.vendor); assert.equal(h.cell(4, 5).formula, '=E6+$A$2')
  assert.equal(h.history().undo, 1); assert.equal(h.calls.changed.length, 1)
})

test('clipboard keydown leaves C X V defaults intact and Shift V routes the following native paste to values', t => {
  const h = host(t), event = clipboardEvent(); h.ui.handleCopy(event)
  for (const key of ['c', 'x', 'v']) { const keyboard = keyEvent(key); assert.equal(h.ui.handleClipboardKeydown(keyboard), true); assert.equal(keyboard.defaultPrevented, false) }
  const keyboard = keyEvent('v', { shiftKey: true }); h.ui.handleClipboardKeydown(keyboard); h.select(range(4, 4))
  h.ui.handlePaste(clipboardEvent(Object.fromEntries(event.data)))
  assert.equal(h.cell(4, 5).formula, null); assert.equal(h.cell(4, 5).value, 5)
  assert.equal(h.cell(4, 4).style.fill, null); assert.deepEqual(h.cell(4, 4).richText, h.cell(0, 0).richText)
})

test('explicit text, clipboard IDs and fixed Univer paste hooks bypass browser reads and share mode validation', async t => {
  const h = host(t); await h.ui.runClipboardCommand('sheet.command.copy'); const id = h.owner.value.sheetClipboard.clipboard.id
  h.select(range(4, 4)); let result = await h.ui.runClipboardCommand('sheet.command.paste', { clipboardId: id, value: 'special-paste-value' })
  assert.equal(result.accepted, true); assert.equal(h.cell(4, 5).value, 5); assert.equal(h.cell(4, 5).formula, null); assert.equal(h.calls.reads, 0)
  h.select(range(8, 4)); result = await h.ui.runClipboardCommand('sheet.command.paste', { text: '12.5%\t\'00123' })
  assert.equal(result.accepted, true); assert.equal(h.cell(8, 4).value, .125); assert.equal(h.cell(8, 5).value, '00123'); assert.equal(h.calls.reads, 0)
  const before = copy(h.current())
  assert.equal(await h.ui.runClipboardCommand('sheet.command.paste-values', { mode: 'formulas', text: 'x' }), undefined)
  assert.equal(await h.ui.runClipboardCommand('sheet.command.paste', { text: 7 }), undefined)
  assert.deepEqual(h.current(), before); assert.ok(h.calls.notices.some(item => item.type === 'error'))
})

test('fixed Univer paste types survive UI routing and reject conflicting modes in production and standalone hosts', async t => {
  for (const standalone of [false, true]) {
    const original = model(); original.sheets[0].cells['4:4'] = sdk.createCell(4, 4, 'old')
    original.sheets[0].cells['4:4'].style.font.bold = true
    const h = host(t, { standalone, model: original }), copied = await h.ui.runClipboardCommand('sheet.command.copy')
    h.select(range(4, 4)); await h.ui.runClipboardCommand('sheet.command.paste')
    const reads = h.calls.reads, result = await h.ui.runClipboardCommand('sheet.command.optional-paste', { type: 'SPECIAL_PASTE_VALUE' })
    assert.equal(result.accepted, true); assert.equal(h.cell(4, 4).style.font.bold, true)
    assert.equal(h.cell(4, 5).value, 5); assert.equal(h.cell(4, 5).formula, null); assert.equal(h.ui.pasteOptions.value.mode, 'values')
    if (!standalone) {
      assert.equal(h.calls.commands.findLast(command => command.id === 'sheet.command.optional-paste').params.type, 'SPECIAL_PASTE_VALUE')
      assert.equal(h.history().undo, 1)
    }
    const before = copy(h.current()), noticeCount = h.calls.notices.length
    for (const params of [{ type: 'SPECIAL_PASTE_FORMULA', mode: 'values' },
      { type: 'SPECIAL_PASTE_FORMULA', value: 'special-paste-value' }, { type: 'unknown-type' }, { type: 7 }]) {
      assert.equal(await h.ui.runClipboardCommand('sheet.command.optional-paste', params), undefined)
    }
    assert.equal(await h.ui.runClipboardCommand('sheet.command.paste-values', { type: 'SPECIAL_PASTE_FORMULA', text: 'x' }), undefined)
    assert.deepEqual(h.current(), before); assert.equal(h.calls.notices.length, noticeCount + 5)
    h.select(range(8, 4))
    assert.equal((await h.ui.runClipboardCommand('sheet.command.paste-formula', {
      clipboardId: copied.clipboard.id, type: 'SPECIAL_PASTE_FORMULA', mode: 'formulas', value: 'special-paste-formula',
    })).accepted, true)
    assert.equal(h.cell(8, 5).formula, '=E10+$A$2'); assert.equal(h.calls.reads, reads)
    if (!standalone) assert.equal(h.calls.commands.findLast(command => command.id === 'sheet.command.paste-formula').params.type, 'SPECIAL_PASTE_FORMULA')
  }
})

test('a genuinely empty or media-only reader result never falls back to an older internal clipboard', async t => {
  for (const content of [{ text: '', html: '' }, {}]) {
    const h = host(t, { read: async () => content }); await h.ui.runClipboardCommand('sheet.command.copy'); h.select(range(4, 4))
    const before = copy(h.current()), result = await h.ui.runClipboardCommand('sheet.command.paste')
    assert.equal(result.accepted, false); assert.equal(result.reason, 'empty-clipboard'); assert.deepEqual(h.current(), before)
    assert.equal(h.history().undo, 0); assert.equal(h.calls.changed.length, 0)
  }
})

test('an unavailable or refused reader can use the captured internal ID without replacing clipboard semantics', async t => {
  for (const read of [async () => undefined, async () => { throw new Error('reader permission denied') }]) {
    const h = host(t, { read }); await h.ui.runClipboardCommand('sheet.command.copy'); h.select(range(4, 4))
    assert.equal((await h.ui.runClipboardCommand('sheet.command.paste')).accepted, true)
    assert.deepEqual(h.cell(4, 4).richText, h.cell(0, 0).richText); assert.equal(h.history().undo, 1)
  }
})

test('a foreign HTML token cannot select an old internal snapshot through a missing cache ID', async t => {
  const h = host(t, { read: async () => ({ text: 'external', html: '<table data-office-clipboard-id="foreign-id"><tr><td>external</td></tr></table>' }) })
  await h.ui.runClipboardCommand('sheet.command.copy'); h.select(range(4, 4))
  assert.equal((await h.ui.runClipboardCommand('sheet.command.paste')).accepted, true)
  assert.equal(h.cell(4, 4).value, 'external'); assert.equal(h.cell(4, 4).richText?.some(run => run.vendor?.retain === 1) ?? false, false)
})

test('an explicit clipboard ID cannot select a conflicting cached snapshot when HTML was also supplied', async t => {
  for (const standalone of [false, true]) {
    const h = host(t, { standalone }), copied = await h.ui.runClipboardCommand('sheet.command.copy')
    h.select(range(4, 4))
    const result = await h.ui.runClipboardCommand('sheet.command.paste', { clipboardId: 'different-id', html: copied.html, text: copied.plain })
    assert.equal(result.accepted, true); assert.equal(h.cell(4, 4).richText?.some(run => run.vendor?.retain === 1) ?? false, false)
    assert.equal(h.calls.reads, 0)
  }
})

test('async paste uses its captured target and preserves a later complete selection and focus', async t => {
  const pending = deferred(), h = host(t, { read: () => pending.promise }); h.select(range(4, 4))
  const result = h.ui.runClipboardCommand('sheet.command.paste'); h.select(range(8, 7)); const after = copy(h.owner.value.getSelectionState())
  pending.resolve({ text: 'fixed target' }); assert.equal((await result).accepted, true); await nextTick()
  assert.equal(h.cell(4, 4).value, 'fixed target'); assert.equal(h.cell(8, 7), undefined)
  assert.deepEqual(h.owner.value.getSelectionState(), after); assert.deepEqual(h.selected.value, range(8, 7)); assert.equal(h.calls.focused.length, 0)
})

test('switching sheets while the reader waits rejects the old paste without any publication', async t => {
  const pending = deferred(), h = host(t, { read: () => pending.promise }), task = h.ui.runClipboardCommand('sheet.command.paste')
  h.select(range(3, 3), h.original.sheets[1].id); pending.resolve({ text: 'late' })
  assert.equal((await task).reason, 'stale-target'); assert.equal(h.calls.applied.length, 0); assert.equal(h.history().undo, 0)
})

test('switching units while the reader waits cannot paste into either the old or new workbook', async t => {
  const pending = deferred(), h = host(t, { read: () => pending.promise }), task = h.ui.runClipboardCommand('sheet.command.paste')
  const replacement = sdk.FOffice.newAPI(h.owner.value).createWorkbook({ ...model(), id: 'clipboard-other-unit' })
  h.view.value = sdk.toExcelWorkbookView(h.owner.value.getWorkbookModel(replacement.getId())); h.sheetId.value = h.view.value.activeSheetId
  pending.resolve({ text: 'late' }); assert.equal((await task).reason, 'stale-target')
  assert.equal(h.calls.changed.length, 0); assert.equal(h.history().undo, 0)
})

test('same-ID model replacement and history reset each invalidate a pending read', async t => {
  for (const reset of [h => h.owner.value.editor.workbook.replace(copy(h.current())), h => h.owner.value.editor.history.clear(h.view.value.id)]) {
    const pending = deferred(), h = host(t, { read: () => pending.promise }), task = h.ui.runClipboardCommand('sheet.command.paste')
    reset(h); pending.resolve({ text: 'late' }); assert.equal((await task).reason, 'stale-target')
    assert.equal(h.calls.changed.length, 0); assert.equal(h.history().undo, 0)
  }
})

test('a new content edit after target capture rejects a late read and preserves that edit', async t => {
  const pending = deferred(), h = host(t, { read: () => pending.promise }), task = h.ui.runClipboardCommand('sheet.command.paste')
  sdk.FOffice.newAPI(h.owner.value).getActiveWorkbook().getActiveSheet().getRange('A4').setValue('new edit')
  const after = copy(h.current()), depth = h.history(); pending.resolve({ text: 'late' })
  assert.equal((await task).reason, 'stale-target'); assert.deepEqual(h.current(), after); assert.deepEqual(h.history(), depth)
  assert.equal(h.calls.changed.length, 0)
})

test('studio rebinding, a replaced view and deactivation independently invalidate a captured host', async t => {
  for (const reset of [h => {
    h.owner.value = new sdk.OfficeStudio({ workbook: model() }); t.onTestFinished(() => h.owner.value.dispose())
  }, h => { h.view.value = copy(h.view.value) }, h => { h.view.value.sheets[0] = copy(h.view.value.sheets[0]) },
  h => { h.active.value = false }, h => { h.editing.value = true }]) {
    const pending = deferred(), h = host(t, { read: () => pending.promise }), task = h.ui.runClipboardCommand('sheet.command.paste')
    reset(h); pending.resolve({ text: 'late' }); assert.equal((await task).reason, 'stale-target'); assert.equal(h.calls.applied.length, 0)
  }
})

test('a newer paste request wins when clipboard readers settle out of order', async t => {
  const first = deferred(), second = deferred(), reads = [first, second], h = host(t, { read: () => reads.shift().promise })
  h.select(range(4, 4)); const old = h.ui.runClipboardCommand('sheet.command.paste'); h.select(range(8, 7)); const newer = h.ui.runClipboardCommand('sheet.command.paste')
  second.resolve({ text: 'new request' }); assert.equal((await newer).accepted, true)
  first.resolve({ text: 'old request' }); assert.equal((await old).reason, 'stale-target')
  assert.equal(h.cell(8, 7).value, 'new request'); assert.equal(h.cell(4, 4), undefined); assert.equal(h.history().undo, 1)
})

test('Escape cancels an awaiting paste and removes marks while retaining ordinary copied content', async t => {
  const pending = deferred(), h = host(t, { read: () => pending.promise }); await h.ui.runClipboardCommand('sheet.command.copy')
  const id = h.owner.value.sheetClipboard.clipboard.id; h.select(range(4, 4)); const task = h.ui.runClipboardCommand('sheet.command.paste')
  const escape = keyEvent('Escape', { ctrlKey: false }); assert.equal(h.ui.handleClipboardKeydown(escape), true); assert.equal(escape.defaultPrevented, true)
  assert.equal(h.owner.value.sheetClipboard.copyMark, undefined); assert.equal(h.owner.value.sheetClipboard.clipboard.id, id)
  pending.resolve({ text: 'late' }); assert.equal((await task).reason, 'stale-target'); assert.equal(h.history().undo, 0)
})

test('Escape on a different sheet also cancels a pending cut and keeps the source untouched', async t => {
  const h = host(t), before = copy(h.current()); await h.ui.runClipboardCommand('sheet.command.cut'); h.select(range(3, 3), h.original.sheets[1].id)
  assert.equal(h.ui.copyRange.value, undefined); assert.equal(h.ui.handleClipboardKeydown(keyEvent('Escape', { ctrlKey: false })), true)
  assert.equal(h.owner.value.sheetClipboard.clipboard, undefined); assert.deepEqual(h.current().sheets[0], before.sheets[0]); assert.equal(h.history().undo, 0)
})

test('a rejected cancel-copy command reports failure while still invalidating an awaiting paste for both UI entry paths', async t => {
  for (const nativeEscape of [false, true]) {
    const pending = deferred(), h = host(t, { read: () => pending.promise })
    await h.ui.runClipboardCommand('sheet.command.cut'); h.select(range(4, 4)); h.calls.focused.length = 0
    const before = copy(h.current()), paste = h.ui.runClipboardCommand('sheet.command.paste')
    h.owner.value.events.addEvent('BeforeCommandExecute', event => { if (event.id === 'sheet.operation.cancel-copy') event.cancel = true })
    if (nativeEscape) {
      const keyboard = keyEvent('Escape', { ctrlKey: false })
      assert.equal(h.ui.handleClipboardKeydown(keyboard), true); assert.equal(keyboard.defaultPrevented, true)
    } else assert.equal(await h.ui.runClipboardCommand('sheet.operation.cancel-copy'), false)
    assert.ok(h.ui.copyRange.value); assert.equal(h.owner.value.sheetClipboard.clipboard.copyType, 'cut')
    assert.equal(h.calls.notices.length, 1); assert.equal(h.calls.notices[0].type, 'error')
    pending.resolve({ text: 'late' }); assert.equal((await paste).reason, 'stale-target'); await nextTick()
    assert.deepEqual(h.current(), before); assert.equal(h.history().undo, 0)
    assert.equal(h.calls.changed.length, 0); assert.equal(h.calls.focused.length, 0)
  }
})

test('a successful cut moves once and leaves one shared content history transaction', async t => {
  const h = host(t), source = copy(h.cell(0, 0)); await h.ui.runClipboardCommand('sheet.command.cut'); h.select(range(4, 4))
  assert.equal((await h.ui.runClipboardCommand('sheet.command.paste')).accepted, true)
  assert.equal(h.cell(0, 0), undefined); assert.deepEqual(h.cell(4, 4).richText, source.richText)
  assert.equal(h.owner.value.sheetClipboard.clipboard, undefined); assert.equal(h.owner.value.sheetClipboard.copyMark, undefined); assert.equal(h.history().undo, 1)
  sdk.FOffice.newAPI(h.owner.value).getActiveWorkbook().undo(); assert.deepEqual(h.cell(0, 0).richText, source.richText); assert.equal(h.cell(4, 4), undefined)
})

test('a changed cut source is refused atomically and remains available for reselection', async t => {
  const h = host(t); await h.ui.runClipboardCommand('sheet.command.cut')
  sdk.FOffice.newAPI(h.owner.value).getActiveWorkbook().getActiveSheet().getRange('A1').setValue('changed source')
  h.view.value = sdk.toExcelWorkbookView(h.current()); h.select(range(4, 4)); const before = copy(h.current()), depth = h.history()
  const result = await h.ui.runClipboardCommand('sheet.command.paste')
  assert.equal(result.accepted, false); assert.equal(result.reason, 'stale-source'); assert.deepEqual(h.current(), before); assert.deepEqual(h.history(), depth)
  assert.equal(h.owner.value.sheetClipboard.clipboard.copyType, 'cut')
})

test('cancellation, protected destinations and stop validation preserve content, cut state and history', async t => {
  for (const setup of [h => h.owner.value.events.addEvent('BeforeCommandExecute', event => { if (event.id === 'sheet.command.paste') event.cancel = true }),
    h => {
      const sheet = h.current().sheets[0]; sheet.rangeProtectionRules = [{ id: 'locked', permissionId: 'locked-permission', ownerUserId: 'remote',
        allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView', description: '', ranges: [range(4, 4, 4, 5)] }]
    }, h => {
      h.current().sheets[0].dataValidations = [{ id: 'stop', type: 'whole', operator: 'between', formula1: '1', formula2: '2',
        allowBlank: false, errorStyle: 'stop', errorTitle: '范围', error: '1 到 2', ranges: [range(4, 5)] }]
    }]) {
    const h = host(t); await h.ui.runClipboardCommand('sheet.command.cut'); setup(h)
    h.view.value = sdk.toExcelWorkbookView(h.current()); h.select(range(4, 4)); const before = copy(h.current()), depth = h.history()
    const result = await h.ui.runClipboardCommand('sheet.command.paste')
    assert.equal(result?.accepted ?? false, false); assert.deepEqual(h.current(), before); assert.deepEqual(h.history(), depth)
    assert.equal(h.owner.value.sheetClipboard.clipboard.copyType, 'cut'); assert.equal(h.calls.changed.length, 0); assert.ok(h.calls.notices.length)
  }
})

test('production optional paste amends the original undo entry and keeps its chosen redo result', async t => {
  const h = host(t); await h.ui.runClipboardCommand('sheet.command.copy'); h.select(range(4, 4))
  await h.ui.runClipboardCommand('sheet.command.paste'); assert.equal(h.ui.pasteOptions.value.mode, 'all')
  assert.deepEqual(h.owner.value.sheetClipboard.recentPaste.target.range, range(4, 4))
  assert.deepEqual(h.ui.pasteOptions.value.range, range(4, 4, 4, 5))
  const result = await h.ui.runClipboardCommand('sheet.command.optional-paste', { value: 'special-paste-value' })
  assert.equal(result.accepted, true); assert.equal(h.cell(4, 5).formula, null); assert.equal(h.cell(4, 5).value, 5)
  assert.equal(h.history().undo, 1); assert.equal(h.ui.pasteOptions.value.mode, 'values')
  const book = sdk.FOffice.newAPI(h.owner.value).getActiveWorkbook(); book.undo(); assert.equal(h.cell(4, 4), undefined)
  book.redo(); assert.equal(h.cell(4, 5).value, 5); assert.equal(h.cell(4, 5).formula, null)
})

test('recent options disappear after an intervening edit and cannot amend its history', async t => {
  const h = host(t); await h.ui.runClipboardCommand('sheet.command.copy'); h.select(range(4, 4)); await h.ui.runClipboardCommand('sheet.command.paste')
  sdk.FOffice.newAPI(h.owner.value).getActiveWorkbook().getActiveSheet().getRange('J12').setValue('later edit')
  h.view.value = sdk.toExcelWorkbookView(h.current()); const before = copy(h.current()), depth = h.history()
  assert.equal(h.ui.pasteOptions.value, undefined)
  assert.equal((await h.ui.runClipboardCommand('sheet.command.optional-paste', { mode: 'values' })).accepted, false)
  assert.deepEqual(h.current(), before); assert.deepEqual(h.history(), depth)
})

test('optional paste refuses a different explicit clipboard ID in production and standalone modes', async t => {
  for (const standalone of [false, true]) {
    const h = host(t, { standalone }); await h.ui.runClipboardCommand('sheet.command.copy'); h.select(range(4, 4))
    await h.ui.runClipboardCommand('sheet.command.paste'); const before = copy(h.current()), depth = h.history()
    const result = await h.ui.runClipboardCommand('sheet.command.optional-paste', { clipboardId: 'different-id', mode: 'values' })
    assert.equal(result.reason, 'stale-paste'); assert.deepEqual(h.current(), before); assert.deepEqual(h.history(), depth)
    assert.equal(h.ui.pasteOptions.value.mode, 'all')
  }
})

test('a post-apply host change prevents old UI notifications, selection writes and focus restoration', async t => {
  const h = host(t, { onApply: context => {
    const other = sdk.FOffice.newAPI(context.studio()).createWorkbook({ ...model(), id: 'post-apply-new-unit' })
    h.view.value = sdk.toExcelWorkbookView(context.studio().getWorkbookModel(other.getId())); h.sheetId.value = h.view.value.activeSheetId
  } })
  await h.ui.runClipboardCommand('sheet.command.copy'); h.select(range(4, 4)); h.calls.focused.length = 0
  const result = await h.ui.runClipboardCommand('sheet.command.paste'); await nextTick()
  assert.equal(result.accepted, true); assert.equal(h.calls.changed.length, 0); assert.equal(h.calls.selected.length, 0); assert.equal(h.calls.focused.length, 0)
})

test('a same-ID view replacement during apply cannot publish that new view as the paste result', async t => {
  const h = host(t, { onApply: () => {
    h.view.value = copy(h.view.value); h.view.value.sheets[0].cells.E5 = { value: 'new UI host' }
  } })
  await h.ui.runClipboardCommand('sheet.command.copy'); h.select(range(4, 4)); h.calls.focused.length = 0
  const result = await h.ui.runClipboardCommand('sheet.command.paste'); await nextTick()
  assert.equal(result.accepted, true); assert.equal(h.cell(4, 4).value, 'rich')
  assert.equal(h.view.value.sheets[0].cells.E5.value, 'new UI host'); assert.equal(h.calls.changed.length, 0); assert.equal(h.calls.focused.length, 0)
})

test('a view change from a change listener keeps the completed transaction and suppresses deferred focus', async t => {
  const h = host(t, { onChanged: () => { h.view.value.sheets[0].cells.J12 = { value: 'new UI input' } } })
  await h.ui.runClipboardCommand('sheet.command.copy'); h.select(range(4, 4)); h.calls.focused.length = 0
  const result = await h.ui.runClipboardCommand('sheet.command.paste'); await nextTick()
  assert.equal(result.accepted, true); assert.equal(h.history().undo, 1); assert.equal(h.calls.changed.length, 1); assert.equal(h.calls.focused.length, 0)
})

test('standalone optional paste restores original styles even when its replacement plan is a no-op on the original', async t => {
  const source = model(), sheet = source.sheets[0]; sheet.cells['0:0'] = sdk.createCell(0, 0, 'same')
  sheet.cells['4:4'] = sdk.createCell(4, 4, 'same'); sheet.cells['4:4'].style.font.bold = true
  const h = host(t, { standalone: true, model: source }); h.select(range(0, 0)); await h.ui.runClipboardCommand('sheet.command.copy')
  h.select(range(4, 4)); assert.equal((await h.ui.runClipboardCommand('sheet.command.paste')).changed, true); assert.equal(h.cell(4, 4).style.font.bold, false)
  const switched = await h.ui.runClipboardCommand('sheet.command.optional-paste', { mode: 'values' })
  assert.equal(switched.accepted, true); assert.equal(switched.changed, true); assert.equal(h.cell(4, 4).style.font.bold, true)
  assert.equal(h.calls.changed.length, 2); assert.equal(h.ui.pasteOptions.value.mode, 'values'); assert.equal(h.context.studio(), undefined)
})

test('optional paste rechecks restored column width permission in production and standalone hosts', async t => {
  for (const standalone of [false, true]) {
    const source = model()
    source.sheets[0].rangeProtectionRules = [{ id: 'width-protection', permissionId: 'width-permission', ownerUserId: 'local-user',
      allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView', description: '', ranges: [range(10, 4)] }]
    const h = host(t, { standalone, model: source }), copied = await h.ui.runClipboardCommand('sheet.command.copy')
    h.select(range(4, 4)); await h.ui.runClipboardCommand('sheet.command.paste-column-widths', { clipboardId: copied.clipboard.id })
    assert.equal(h.ui.pasteOptions.value.mode, 'columnWidths'); assert.equal(h.calls.changed.length, 1)
    if (standalone) h.currentUserId.value = 'remote-user'
    else h.current().sheets[0].rangeProtectionRules[0].ownerUserId = 'remote-user'
    const before = copy(h.current()), depth = h.history()
    assert.equal(h.ui.pasteOptionsCapabilities.value.paste.values.disabled, true)
    assert.equal(h.ui.capabilities.value.paste.values.disabled, false)
    assert.equal(await h.ui.runClipboardCommand('sheet.command.optional-paste', { type: 'SPECIAL_PASTE_VALUE' }), undefined)
    assert.deepEqual(h.current(), before); assert.deepEqual(h.history(), depth); assert.equal(h.ui.pasteOptions.value.mode, 'columnWidths')
    assert.equal(h.calls.changed.length, 1); assert.equal(h.calls.notices.at(-1).type, 'error')
  }
})

test('standalone optional paste keeps the original single-cell target when compact visible rows expand its result', async t => {
  const source = model(); source.sheets[0].cells['5:0'] = sdk.createCell(5, 0, 'hidden by filter')
  source.sheets[0].autoFilter = { range: range(0, 0, 10, 4), columns: { 0: { kind: 'condition', operator: 'doesNotEqual', value: 'hidden by filter' } }, filteredRows: [5] }
  const h = host(t, { standalone: true, model: source }); h.select(range(0, 0, 1, 0))
  await h.ui.runClipboardCommand('sheet.command.copy'); h.select(range(4, 4)); await h.ui.runClipboardCommand('sheet.command.paste')
  assert.equal(h.cell(4, 4).value, 'rich'); assert.equal(h.cell(5, 4), undefined); assert.equal(h.cell(6, 4).value, 5)
  assert.deepEqual(h.ui.pasteOptions.value.range, range(6, 4))
  const result = await h.ui.runClipboardCommand('sheet.command.optional-paste', { mode: 'values' })
  assert.equal(result.accepted, true); assert.equal(h.cell(4, 4).value, 'rich'); assert.equal(h.cell(6, 4).value, 5)
  assert.equal(h.cell(7, 4), undefined); assert.equal(h.cell(8, 4), undefined)
})

test('standalone native paste cannot take focus after the user selects a new range before Vue flush', async t => {
  const h = host(t, { standalone: true }), event = clipboardEvent(); h.ui.handleCopy(event); h.select(range(4, 4))
  h.ui.handlePaste(clipboardEvent(Object.fromEntries(event.data))); h.select(range(8, 7)); await nextTick()
  assert.equal(h.cell(4, 4).value, 'rich'); assert.deepEqual(h.selected.value, range(8, 7)); assert.equal(h.calls.focused.length, 0)
})

test('standalone preparation failure retains its source and clipboard without publishing a candidate', async t => {
  const h = host(t, { standalone: true, prepareView: () => { throw new Error('candidate rejected') } })
  await h.ui.runClipboardCommand('sheet.command.cut'); h.select(range(4, 4)); const before = copy(h.view.value)
  assert.equal(await h.ui.runClipboardCommand('sheet.command.paste'), undefined); assert.deepEqual(h.view.value, before)
  assert.equal(h.calls.changed.length, 0); assert.equal(h.calls.applied.length, 0); assert.ok(h.ui.copyRange.value)
})

test('production and standalone optional paste restore a moved table while a chart data source cut preserves its drawing host', async t => {
  for (const standalone of [false, true]) {
    const source = model(), sheetId = source.sheets[0].id, targetSheetId = source.sheets[1].id, originalRange = range(0, 0, 1, 1)
    source.tables = [{ id: 'table-ui', name: 'SourceTable', sheetId, range: originalRange,
      columns: [{ id: 'column-a', name: 'A' }, { id: 'column-b', name: 'B' }], showHeaderRow: true, showTotalsRow: false, styleName: 'TableStyleMedium2' }]
    source.charts = [{ id: 'chart-ui', sheetId, type: 'bar', source: { range: originalRange, orientation: 'columns' },
      position: { x: 300, y: 200 }, size: { width: 400, height: 250 }, title: 'Reference',
      legend: { visible: true, position: 'bottom' }, categoryField: 0, valueFields: [1],
      xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, revision: 0 }]
    const h = host(t, { standalone, model: source }), initial = copy(h.current()); h.select(originalRange)
    await h.ui.runClipboardCommand('sheet.command.cut'); h.select(range(4, 4), targetSheetId); await h.ui.runClipboardCommand('sheet.command.paste')
    assert.equal(h.current().tables[0].sheetId, targetSheetId); assert.deepEqual(h.current().tables[0].range, range(4, 4, 5, 5))
    assert.deepEqual(h.current().charts[0].source.range, range(4, 4, 5, 5)); assert.equal(h.current().charts[0].source.sheetId, targetSheetId)
    assert.equal(h.current().charts[0].sheetId, sheetId); assert.deepEqual(h.current().charts[0].position, initial.charts[0].position)
    const result = await h.ui.runClipboardCommand('sheet.command.optional-paste', { type: 'SPECIAL_PASTE_VALUE' })
    assert.equal(result.accepted, true); assert.deepEqual(copy(h.current().tables), initial.tables); assert.deepEqual(copy(h.current().charts), initial.charts)
    assert.equal(h.cell(0, 0).value, 'rich'); assert.equal(h.cell(4, 4, 1).value, 'rich')
    if (standalone) assert.equal(h.context.studio(), undefined)
    else assert.equal(h.history().undo, 1)
  }
})

test('standalone in-place changes invalidate an async read through its captured view signature', async t => {
  const pending = deferred(), h = host(t, { standalone: true, read: () => pending.promise }); h.select(range(4, 4))
  const task = h.ui.runClipboardCommand('sheet.command.paste'); h.view.value.sheets[0].cells.J12 = { value: 'new local edit' }
  const before = copy(h.view.value); pending.resolve({ text: 'late' })
  assert.equal((await task).reason, 'stale-target'); assert.deepEqual(h.view.value, before); assert.equal(h.calls.changed.length, 0)
})

test('fixed textContent and htmlContent aliases reach production and standalone paste without system reads', async t => {
  for (const standalone of [false, true]) {
    const h = host(t, { standalone }); h.select(range(4, 4))
    let result = await h.ui.runClipboardCommand('sheet.command.paste-by-short-key', { textContent: '7\t8' })
    assert.equal(result.accepted, true); assert.equal(h.cell(4, 4).value, 7); assert.equal(h.cell(4, 5).value, 8)
    h.select(range(8, 4))
    const html = '<table><tr><td style="font-weight:bold;background-color:#ffeeaa">Styled</td></tr></table>'
    result = await h.ui.runClipboardCommand('sheet.command.paste', { htmlContent: html, html, textContent: 'Styled', text: 'Styled' })
    assert.equal(result.accepted, true); assert.equal(h.cell(8, 4).value, 'Styled'); assert.equal(h.cell(8, 4).style.font.bold, true)
    assert.equal(h.cell(8, 4).style.fill, '#ffeeaa'); assert.equal(h.calls.reads, 0)
    if (!standalone) {
      const command = h.calls.commands.findLast(entry => entry.id === 'sheet.command.paste')
      assert.equal(command.params.html, html); assert.equal(command.params.text, 'Styled')
    }
  }
})

test('conflicting or invalid content aliases retain model, selection, clipboard state and history before reading', async t => {
  for (const standalone of [false, true]) {
    const h = host(t, { standalone }); await h.ui.runClipboardCommand('sheet.command.copy')
    h.select(range(4, 4)); await h.ui.runClipboardCommand('sheet.command.paste'); await nextTick()
    const before = copy(h.current()), selection = copy(h.selected.value), depth = h.history(), options = copy(h.ui.pasteOptions.value)
    const reads = h.calls.reads, notices = h.calls.notices.length, published = h.calls.changed.length
    const invalid = [{ text: '', textContent: 'other' }, { text: 'x', textContent: ' x' },
      { html: '<b>x</b>', htmlContent: '<i>x</i>' }, { textContent: 7 }, { htmlContent: null }]
    for (const params of invalid) assert.equal(await h.ui.runClipboardCommand('sheet.command.paste', params), undefined)
    assert.deepEqual(h.current(), before); assert.deepEqual(h.selected.value, selection); assert.deepEqual(h.history(), depth)
    assert.deepEqual(h.ui.pasteOptions.value, options); assert.equal(h.calls.changed.length, published)
    assert.equal(h.calls.reads, reads); assert.equal(h.calls.notices.length, notices + invalid.length)
    assert.ok(h.calls.notices.slice(notices).every(notice => notice.type === 'error'))
  }
})

test('explicit empty content aliases never read or paste an older cached copy', async t => {
  for (const standalone of [false, true]) for (const params of [{ textContent: '' }, { htmlContent: '' }, { text: '', textContent: '', htmlContent: '' }]) {
    const h = host(t, { standalone }); await h.ui.runClipboardCommand('sheet.command.copy'); h.select(range(4, 4))
    const before = copy(h.current()), mark = copy(h.ui.copyRange.value)
    const result = await h.ui.runClipboardCommand('sheet.command.paste', params)
    assert.equal(result.accepted, false); assert.equal(result.reason, 'empty-clipboard'); assert.equal(h.calls.reads, 0)
    assert.deepEqual(h.current(), before); assert.deepEqual(h.ui.copyRange.value, mark); assert.equal(h.calls.changed.length, 0)
    assert.equal(h.history().undo, 0)
  }
})

test('hidePasteOptions only changes presentation and toggling it retains recent content, selection and history', async t => {
  for (const standalone of [false, true]) {
    const h = host(t, { standalone, hidePasteOptions: true }), copied = await h.ui.runClipboardCommand('sheet.command.copy')
    h.select(range(4, 4)); assert.equal((await h.ui.runClipboardCommand('sheet.command.paste', { clipboardId: copied.clipboard.id })).accepted, true)
    await nextTick()
    const before = copy(h.current()), selected = copy(h.selected.value), depth = h.history(), published = h.calls.changed.length
    assert.equal(h.ui.pasteOptions.value, undefined); assert.ok(h.ui.pasteOptionsCapabilities.value)
    h.hidePasteOptions.value = false
    assert.deepEqual(h.ui.pasteOptions.value, { range: range(4, 4, 4, 5), mode: 'all' })
    h.hidePasteOptions.value = true; assert.equal(h.ui.pasteOptions.value, undefined)
    assert.deepEqual(h.current(), before); assert.deepEqual(h.selected.value, selected); assert.deepEqual(h.history(), depth)
    assert.equal(h.calls.changed.length, published); assert.equal(h.calls.reads, 0)
  }
})

test('hidden recent paste remains callable and later shows its replacement mode without creating a second production history entry', async t => {
  for (const standalone of [false, true]) {
    const original = model(); original.sheets[0].cells['4:4'] = sdk.createCell(4, 4, 'old'); original.sheets[0].cells['4:4'].style.font.bold = true
    const h = host(t, { standalone, hidePasteOptions: true, model: original }), copied = await h.ui.runClipboardCommand('sheet.command.copy')
    h.select(range(4, 4)); await h.ui.runClipboardCommand('sheet.command.paste', { clipboardId: copied.clipboard.id })
    assert.equal(h.cell(4, 4).style.font.bold, false); assert.equal(h.ui.pasteOptions.value, undefined)
    const result = await h.ui.runClipboardCommand('sheet.command.optional-paste', { type: 'SPECIAL_PASTE_VALUE' })
    assert.equal(result.accepted, true); assert.equal(h.cell(4, 4).style.font.bold, true); assert.equal(h.cell(4, 5).formula, null)
    assert.equal(h.ui.pasteOptions.value, undefined); assert.equal(h.calls.reads, 0)
    h.hidePasteOptions.value = false; assert.equal(h.ui.pasteOptions.value.mode, 'values')
    assert.equal(h.calls.changed.length, 2); assert.equal(h.history().undo, standalone ? 0 : 1)
  }
})

test('UI operation permissions distinguish style from values and also guard native and explicit command writes', async t => {
  for (const standalone of [false, true]) {
    const h = host(t, { standalone, readerAvailable: false }), copied = await h.ui.runClipboardCommand('sheet.command.copy')
    h.permission.value = request => request.point !== 'setCellStyle'; h.select(range(4, 4))
    assert.equal(h.ui.capabilities.value.paste.all.disabled, true); assert.equal(h.ui.capabilities.value.paste.values.disabled, false)
    assert.equal(h.ui.capabilities.value.paste.formats.disabled, true); assert.equal(h.ui.capabilities.value.paste.columnWidths.disabled, false)
    const before = copy(h.current()); h.ui.handlePaste(clipboardEvent({ 'text/plain': copied.plain, 'text/html': copied.html }))
    assert.deepEqual(h.current(), before); assert.equal(h.calls.changed.length, 0); assert.equal(h.calls.notices.at(-1).type, 'error')
    const result = await h.ui.runClipboardCommand('sheet.command.paste-values', { clipboardId: copied.clipboard.id })
    assert.equal(result.accepted, true); assert.equal(h.cell(4, 4).value, 'rich'); assert.equal(h.cell(4, 5).formula, null)
    assert.equal(h.calls.reads, 0)
  }
})

test('copy permission and edit permission are separately enforced without erasing the cached ordinary copy', async t => {
  for (const standalone of [false, true]) {
    const h = host(t, { standalone, permission: request => request.point !== 'copy' })
    assert.equal(h.ui.capabilities.value.copy.disabled, true); assert.equal(h.ui.capabilities.value.cut.disabled, true)
    const before = copy(h.current()); assert.equal(await h.ui.runClipboardCommand('sheet.command.copy'), undefined)
    assert.equal(h.calls.written.length, 0); assert.equal(h.ui.copyRange.value, undefined)
    h.permission.value = request => request.point !== 'edit'
    assert.equal(h.ui.capabilities.value.copy.disabled, false); assert.equal(h.ui.capabilities.value.cut.disabled, true)
    const copied = await h.ui.runClipboardCommand('sheet.command.copy'); assert.equal(copied.accepted, true)
    h.ui.handleCopy(clipboardEvent(), true)
    assert.deepEqual(h.current(), before); assert.deepEqual(h.ui.copyRange.value, range(0, 0, 0, 1)); assert.equal(h.history().undo, 0)
    if (!standalone) assert.equal(h.owner.value.sheetClipboard.clipboard.copyType, 'copy')
  }
})

test('replacing optional widths must retain permission to restore the old widths even when the new mode writes only values', async t => {
  for (const standalone of [false, true]) {
    const h = host(t, { standalone, permission: () => true }), copied = await h.ui.runClipboardCommand('sheet.command.copy')
    h.select(range(4, 4)); await h.ui.runClipboardCommand('sheet.command.paste', { clipboardId: copied.clipboard.id })
    await h.ui.runClipboardCommand('sheet.command.optional-paste', { mode: 'columnWidths' })
    h.permission.value = request => request.point !== 'setColumnWidth'
    const before = copy(h.current()), options = copy(h.ui.pasteOptions.value), depth = h.history(), published = h.calls.changed.length
    assert.equal(h.ui.pasteOptionsCapabilities.value.paste.values.disabled, true)
    assert.match(h.ui.pasteOptionsCapabilities.value.paste.values.reason, /列宽/)
    assert.equal(h.ui.capabilities.value.paste.values.disabled, false)
    assert.equal(await h.ui.runClipboardCommand('sheet.command.optional-paste', { mode: 'values' }), undefined)
    assert.deepEqual(h.current(), before); assert.deepEqual(h.ui.pasteOptions.value, options); assert.deepEqual(h.history(), depth)
    assert.equal(h.calls.changed.length, published); assert.equal(h.calls.notices.at(-1).type, 'error')
  }
})

test('an explicitly unavailable reader uses cached content without reading and explicit plugin context does not change content history', async t => {
  for (const standalone of [false, true]) {
    const h = host(t, { standalone, readerAvailable: false })
    assert.equal(h.ui.capabilities.value.source.systemClipboardAvailable, false); assert.equal(h.ui.capabilities.value.paste.all.disabled, true)
    await h.ui.runClipboardCommand('sheet.command.copy'); h.select(range(4, 4))
    assert.equal(h.ui.capabilities.value.paste.all.disabled, false)
    assert.equal((await h.ui.runClipboardCommand('sheet.command.paste')).accepted, true); assert.equal(h.calls.reads, 0)
    const before = copy(h.current()), depth = h.history(), published = h.calls.changed.length
    h.exclusiveRange.value = true; assert.equal(h.ui.capabilities.value.pasteSpecial.hidden, true)
    assert.equal(h.ui.capabilities.value.copy.hidden, false); assert.equal(h.ui.capabilities.value.paste.all.hidden, false)
    h.shapeTextEditing.value = true; assert.equal(h.ui.capabilities.value.ribbonPaste.hidden, true)
    assert.equal(await h.ui.runClipboardCommand('sheet.command.copy'), undefined)
    assert.deepEqual(h.current(), before); assert.deepEqual(h.history(), depth); assert.equal(h.calls.changed.length, published)
  }
})
