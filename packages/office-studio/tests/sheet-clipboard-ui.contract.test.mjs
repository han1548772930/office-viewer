import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { nextTick, shallowRef } from 'vue'
import { sheetImagePng } from './fixtures/sheet-image-fixture.mjs'

// Draft only. This suite is written for the user's later unified verification pass.
let server, sdk, clipboardUi, clipboardContent, clipboardModel, richFonts, clipboardCells, imageSource
beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[sdk, clipboardUi, clipboardContent, clipboardModel, richFonts, clipboardCells, imageSource] = await Promise.all([
    '/packages/office-studio/src/index.ts',
    '/packages/office-studio/src/components/spreadsheet/use-sheet-clipboard.ts',
    '/packages/office-studio/src/features/sheet-workbench/sheet-clipboard-content.ts',
    '/packages/office-studio/src/features/sheet-workbench/sheet-clipboard.ts',
    '/packages/office-studio/src/features/sheet-workbench/sheet-cell-rich-text-style.ts',
    '/packages/office-studio/src/features/sheet-workbench/sheet-clipboard-cell.ts',
    '/packages/office-studio/src/features/sheet-images/sheet-image-source.ts',
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
const formattedRuns = [
  { id: 'target-first', text: 'target ', hyperlink: 'https://target.test/run', font: { language: 'ja-JP' }, vendor: { keep: 1 } },
  { id: 'target-second', text: 'runs', vendor: { keep: 2 } },
]
function formatsModel() {
  const book = model(), source = sdk.createCell(0, 0, 'source'), target = sdk.createCell(4, 4, 'target runs')
  source.style.font = { family: 'Courier New', size: 16, bold: true, italic: false, underline: true, strike: false, color: '#123456' }
  source.richText = [{ id: 'source-run', text: 'source', font: { family: 'Arial', size: 40 }, vendor: { source: true } }]
  source.authoredStyle = copy(source.style)
  target.style = copy(source.style); target.authoredStyle = copy(target.style)
  target.richText = [
    { id: 'target-first', text: 'target ', hyperlink: 'https://target.test/run', vendor: { keep: 1 },
      font: { family: 'Arial', size: 32, bold: false, italic: true, underline: false, strike: true, color: '#ff0000', language: 'ja-JP' } },
    { id: 'target-second', text: 'runs', fontFamily: 'Arial', fontSize: 28, bold: false, italic: true,
      underline: false, strike: true, color: '#ff0000', vendor: { keep: 2 } },
  ]
  target.hyperlink = 'https://target.test/cell'; target.metadata = { vendor: { retain: true } }
  book.sheets[0].cells['0:0'] = source; book.sheets[0].cells['4:4'] = target
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

test('external forced strings keep their literal text and cannot become numbers or formula-only inputs', () => {
  const target = formatsModel().sheets[0].cells['4:4'], before = copy(target)
  for (const externalFormat of ['text', 'html']) for (const type of ['force_string', 'force-string', 'forcestring']) {
    for (const literal of ['00123', '=1+1']) {
      const source = { ...sdk.createCell(0, 0, literal), dataType: type }, original = copy(source)
      for (const mode of ['all', 'values']) {
        const pasted = clipboardCells.createSheetClipboardPastedCell(source, target, { row: 4, col: 4 }, mode,
          { rowDelta: 4, columnDelta: 4, externalFormat, date1904: false, region: 'en-US' })
        assert.equal(pasted.value, literal); assert.equal(pasted.formula, null); assert.equal(pasted.dataType, type)
        if (mode === 'values') assert.deepEqual(pasted.style.font, target.style.font)
      }
      const formulas = clipboardCells.createSheetClipboardPastedCell(source, target, { row: 4, col: 4 }, 'formulas',
        { rowDelta: 4, columnDelta: 4, externalFormat, date1904: false, region: 'en-US' })
      assert.deepEqual(formulas, before); assert.deepEqual(source, original)
    }
  }
  const plain = clipboardCells.createSheetClipboardPastedCell({ ...sdk.createCell(0, 0, '00123'), dataType: 'text' }, target,
    { row: 4, col: 4 }, 'values', { rowDelta: 4, columnDelta: 4, externalFormat: 'text', date1904: false, region: 'en-US' })
  assert.equal(plain.value, 123); assert.equal(plain.dataType, 'number'); assert.deepEqual(target, before)
})

test('values paste retains source rich-run fonts except over a formula while target formatting links recent history and reopen survive', async t => {
  for (const standalone of [false, true]) for (const formulaTarget of [false, true]) {
    const original = formatsModel(), source = original.sheets[0].cells['0:0']
    source.hyperlink = 'https://source.test/cell'
    source.richText = [
      { id: 'source-first', text: 'so', font: { family: 'Arial', size: 24, bold: false }, hyperlink: 'https://source.test/run', vendor: { retain: 1 } },
      { id: 'source-second', text: 'urce', fontFamily: 'Times New Roman', fontSize: 18, italic: true, vendor: { retain: 2 } },
    ]
    if (formulaTarget) Object.assign(original.sheets[0].cells['4:4'], { value: 91, formula: '=90+1', dataType: 'formula', cachedDataType: 'number', display: '91' })
    const h = host(t, { standalone, model: original, permission: request => request.point !== 'setCellStyle' })
    const before = copy(h.cell(4, 4)), values = []
    h.owner.value?.events.addEvent('SheetValueChanged', event => values.push(copy(event)))
    h.select(range(0, 0)); const copied = await h.ui.runClipboardCommand('sheet.command.copy'); h.select(range(4, 4))
    assert.equal(h.ui.capabilities.value.paste.values.disabled, false)
    const result = await h.ui.runClipboardCommand('sheet.command.paste-value', { clipboardId: copied.clipboard.id })
    assert.equal(result.accepted, true); assert.equal(result.changed, true)
    const pasted = copy(h.cell(4, 4)), expectedRuns = formulaTarget ? undefined : source.richText
    assert.equal(pasted.value, 'source'); assert.equal(pasted.formula, null); assert.equal(pasted.cachedDataType, null)
    assert.deepEqual(pasted.richText, expectedRuns)
    assert.deepEqual(pasted.style, before.style); assert.equal(pasted.hyperlink, before.hyperlink); assert.deepEqual(pasted.metadata, before.metadata)
    const view = h.view.value.sheets[0].cells.E5
    assert.deepEqual(view.richText, expectedRuns)
    assert.deepEqual(richFonts.sheetCellFontValues(view, 'family'), formulaTarget ? ['Courier New'] : ['Arial', 'Times New Roman'])
    assert.deepEqual(richFonts.sheetCellFontValues(view, 'size'), formulaTarget ? [16] : [24, 18]); assert.equal(h.calls.changed.length, 1)
    assert.equal(h.history().undo, standalone ? 0 : 1); assert.equal(h.ui.pasteOptions.value.mode, 'values')
    if (!standalone) assert.deepEqual(values.map(event => event.ranges.map(domainRange)), [[range(4, 4)]])
    const repeated = await h.ui.runClipboardCommand('sheet.command.optional-paste', { type: 'SPECIAL_PASTE_VALUE' })
    assert.equal(repeated.accepted, true); assert.equal(repeated.changed, false); assert.equal(h.calls.changed.length, 1)
    assert.deepEqual(h.cell(0, 0).richText, source.richText); assert.deepEqual(h.cell(4, 4), pasted)
    if (standalone) continue
    const api = sdk.FOffice.newAPI(h.owner.value), book = api.getActiveWorkbook()
    book.undo(); assert.deepEqual(h.cell(4, 4), before); book.redo(); assert.deepEqual(h.cell(4, 4), pasted)
    const reopened = api.createWorkbook({ ...book.save(), id: 'clipboard-rich-values-reopened' })
    assert.deepEqual(h.owner.value.getWorkbookModel(reopened.getId()).sheets[0].cells['4:4'].richText, expectedRuns)
  }
})

test('values paste copies a cell image and its public body while a following plain value clears the old image content', () => {
  const source = sdk.createCell(0, 0, null), target = formatsModel().sheets[0].cells['4:4']
  source.image = { id: 'source-cell-image', title: 'Source', altText: 'Source image', vendor: { retain: true },
    resource: imageSource.sheetImageResourceFromDataUrl(sheetImagePng, 'source.png', 'source-cell-resource') }
  source.metadata = { publicBody: sdk.sheetCellImageToPublicBody(source), sheetImagePublicDrawingId: source.image.id }
  const before = copy(source), options = { rowDelta: 4, columnDelta: 4, date1904: false, region: 'en-US' }
  const pasted = clipboardCells.createSheetClipboardPastedCell(source, target, { row: 4, col: 4 }, 'values', options)
  assert.notEqual(pasted.image.id, source.image.id); assert.deepEqual(pasted.image.resource, source.image.resource)
  assert.deepEqual(pasted.image.vendor, source.image.vendor); assert.deepEqual(pasted.style.font, target.style.font)
  assert.deepEqual(pasted.metadata.publicBody, source.metadata.publicBody); assert.equal(pasted.metadata.sheetImagePublicDrawingId, source.image.id)
  const publicBody = sdk.sheetCellImageToPublicBody(pasted)
  assert.deepEqual(publicBody.drawingsOrder, [pasted.image.id]); assert.equal(publicBody.body.customBlocks[0].blockId, pasted.image.id)
  assert.equal(Object.hasOwn(publicBody.drawings, source.image.id), false)
  assert.deepEqual(sdk.toExcelCellView(pasted).image.resource, source.image.resource); assert.deepEqual(source, before)
  const plain = clipboardCells.createSheetClipboardPastedCell(sdk.createCell(0, 0, 'plain'), pasted, { row: 4, col: 4 }, 'values', options)
  assert.equal(plain.value, 'plain'); assert.equal(plain.image, undefined); assert.equal(plain.richText, undefined)
  assert.equal(plain.metadata.publicBody, undefined); assert.equal(plain.metadata.sheetImagePublicDrawingId, undefined)
  assert.equal(plain.hyperlink, target.hyperlink); assert.deepEqual(plain.style.font, target.style.font)
  const formulaTarget = { ...copy(target), value: 91, formula: '=90+1', dataType: 'formula', cachedDataType: 'number', display: '91' }
  const flattened = clipboardCells.createSheetClipboardPastedCell(source, formulaTarget, { row: 4, col: 4 }, 'values', options)
  assert.equal(flattened.formula, null); assert.equal(flattened.image, undefined); assert.equal(flattened.metadata.publicBody, undefined)
  assert.equal(flattened.hyperlink, target.hyperlink); assert.deepEqual(flattened.style, target.style)
})

test('value permission and stop validation reject values paste before content, events, history or recent state change', async t => {
  for (const standalone of [false, true]) for (const blocked of ['permission', 'validation']) {
    const original = formatsModel()
    if (blocked === 'validation') original.sheets[0].dataValidations = [{ id: 'reject-values', type: 'custom', formula1: '=FALSE()',
      allowBlank: false, showErrorMessage: true, errorStyle: 'stop', error: 'Value edits are blocked', ranges: [range(4, 4)] }]
    const h = host(t, { standalone, model: original,
      permission: blocked === 'permission' ? request => request.point !== 'setCellValue' : undefined }), values = []
    h.owner.value?.events.addEvent('SheetValueChanged', event => values.push(copy(event)))
    h.select(range(0, 0)); const copied = await h.ui.runClipboardCommand('sheet.command.copy'); h.select(range(4, 4))
    const before = copy(h.current()), depth = h.history()
    const result = await h.ui.runClipboardCommand('sheet.command.paste-values', { clipboardId: copied.clipboard.id })
    assert.equal(result?.accepted ?? false, false); assert.deepEqual(h.current(), before); assert.deepEqual(h.history(), depth)
    assert.equal(h.calls.changed.length, 0); assert.deepEqual(values, []); assert.equal(h.ui.pasteOptions.value, undefined)
    assert.ok(h.calls.notices.length > 0)
  }
})

test('formats planner replaces target run fonts and retains text, links, identities and extensions in model and view', () => {
  const original = formatsModel(), before = copy(original), sheet = original.sheets[0]
  const captured = clipboardModel.captureSheetClipboard(original, { unitId: original.id, sheetId: sheet.id, range: range(0, 0) }, { userId: 'local-user' })
  const plan = clipboardModel.createSheetClipboardPastePlan(original,
    { unitId: original.id, sheetId: sheet.id, range: range(4, 4) }, captured, 'formats', { userId: 'local-user', region: 'en-US' })
  assert.equal(plan.result.accepted, true); assert.equal(plan.result.changed, true)
  assert.deepEqual(plan.changedRanges[sheet.id], [range(4, 4)])
  const candidate = copy(original); clipboardModel.applySheetClipboardPastePlan(candidate, plan)
  const pasted = candidate.sheets[0].cells['4:4'], target = before.sheets[0].cells['4:4']
  assert.deepEqual(pasted.richText, formattedRuns); assert.equal(pasted.value, target.value)
  assert.equal(pasted.formula, target.formula); assert.equal(pasted.dataType, target.dataType)
  assert.equal(pasted.cachedDataType, target.cachedDataType); assert.equal(pasted.hyperlink, target.hyperlink)
  assert.deepEqual(pasted.metadata, target.metadata); assert.deepEqual(pasted.style.font, sheet.cells['0:0'].style.font)
  const view = sdk.toExcelWorkbookView(candidate).sheets[0].cells.E5
  assert.deepEqual(view.richText, formattedRuns)
  for (const field of Object.keys(pasted.style.font)) {
    assert.deepEqual(richFonts.sheetCellFontValues(view, field), [pasted.style.font[field], pasted.style.font[field]])
  }
  assert.deepEqual(clipboardModel.sheetClipboardChangedValueRanges(sheet, candidate.sheets[0]), [])
  assert.deepEqual(original, before); assert.deepEqual(captured.cells['0:0'].richText, sheet.cells['0:0'].richText)
})

test('run-only formats paste skips stop value validation and keeps one history entry through no-op, undo, redo and reopen', async t => {
  for (const standalone of [false, true]) {
    const original = formatsModel()
    original.sheets[0].dataValidations = [{ id: 'reject-value-edits', type: 'custom', formula1: '=FALSE()', allowBlank: false,
      showErrorMessage: true, errorStyle: 'stop', error: 'Value edits are blocked', ranges: [range(4, 4)] }]
    const h = host(t, { standalone, model: original }), before = copy(h.cell(4, 4)), values = []
    h.owner.value?.events.addEvent('SheetValueChanged', event => values.push(copy(event)))
    h.select(range(0, 0)); const copied = await h.ui.runClipboardCommand('sheet.command.copy')
    h.select(range(4, 4))
    const result = await h.ui.runClipboardCommand('sheet.command.paste-format', { clipboardId: copied.clipboard.id })
    assert.equal(result.accepted, true); assert.equal(result.changed, true); assert.deepEqual(result.issues, [])
    const pasted = copy(h.cell(4, 4))
    assert.deepEqual(pasted.richText, formattedRuns); assert.equal(pasted.value, before.value)
    assert.deepEqual(pasted.style.font, before.style.font); assert.deepEqual(h.view.value.sheets[0].cells.E5.richText, formattedRuns)
    assert.deepEqual(values, []); assert.equal(h.calls.changed.length, 1); assert.equal(h.calls.notices.length, 0)
    assert.deepEqual(h.history(), { undo: standalone ? 0 : 1, redo: 0 })
    assert.equal(h.ui.pasteOptions.value.mode, 'formats'); await nextTick()
    assert.deepEqual(h.calls.focused.at(-1), range(4, 4))
    const redundant = await h.ui.runClipboardCommand('sheet.command.optional-paste', { type: 'SPECIAL_PASTE_FORMAT' })
    assert.equal(redundant.accepted, true); assert.equal(redundant.changed, false)
    assert.deepEqual(h.cell(4, 4), pasted); assert.equal(h.calls.changed.length, 1); assert.deepEqual(values, [])
    assert.equal(h.history().undo, standalone ? 0 : 1)
    if (standalone) continue
    const api = sdk.FOffice.newAPI(h.owner.value), book = api.getActiveWorkbook()
    book.undo(); assert.deepEqual(h.cell(4, 4), before)
    book.redo(); assert.deepEqual(h.cell(4, 4), pasted)
    const saved = book.save(), reopened = api.createWorkbook({ ...saved, id: 'clipboard-formats-reopened' })
    assert.deepEqual(h.owner.value.getWorkbookModel(reopened.getId()).sheets[0].cells['4:4'], pasted)
  }
})

test('recent formats replacement restores real values and formulas while excluding cells changed only by formatting from value events', async t => {
  for (const standalone of [false, true]) {
    const original = formatsModel(), sheet = original.sheets[0]
    sheet.cells['4:5'] = { ...sdk.createCell(4, 5, 4), formula: '=2+2', dataType: 'formula', cachedDataType: 'number' }
    sheet.cells['0:2'] = sdk.createCell(0, 2, 8); sheet.cells['0:2'].style.font.bold = true
    sheet.cells['4:6'] = sdk.createCell(4, 6, 8)
    const h = host(t, { standalone, model: original }), initial = copy(h.current()), values = []
    h.owner.value?.events.addEvent('SheetValueChanged', event => values.push(copy(event)))
    h.select(range(0, 0, 0, 2)); const copied = await h.ui.runClipboardCommand('sheet.command.copy')
    h.select(range(4, 4)); assert.equal((await h.ui.runClipboardCommand('sheet.command.paste', { clipboardId: copied.clipboard.id })).accepted, true)
    assert.equal(h.cell(4, 4).value, 'source'); assert.equal(h.cell(4, 5).formula, '=E6+$A$2')
    assert.equal(h.cell(4, 6).style.font.bold, true)
    if (!standalone) assert.deepEqual(values.at(-1).ranges.map(domainRange), [range(4, 4, 4, 5)])
    values.length = 0
    const result = await h.ui.runClipboardCommand('sheet.command.optional-paste', { type: 'SPECIAL_PASTE_FORMAT' })
    assert.equal(result.accepted, true); assert.equal(result.changed, true)
    assert.equal(h.cell(4, 4).value, 'target runs'); assert.deepEqual(h.cell(4, 4).richText, formattedRuns)
    assert.equal(h.cell(4, 5).formula, '=2+2'); assert.equal(h.cell(4, 5).value, 4); assert.equal(h.cell(4, 6).style.font.bold, true)
    assert.equal(h.ui.pasteOptions.value.mode, 'formats'); assert.equal(h.calls.changed.length, 2)
    assert.equal(h.history().undo, standalone ? 0 : 1)
    if (standalone) continue
    assert.equal(values.length, 1); assert.deepEqual(values[0].ranges.map(domainRange), [range(4, 4, 4, 5)])
    const amended = copy(h.current()), book = sdk.FOffice.newAPI(h.owner.value).getActiveWorkbook()
    assert.equal(h.owner.value.sheetClipboard.recentPaste.mode, 'formats')
    book.undo(); assert.deepEqual(h.current().sheets[0].cells, initial.sheets[0].cells)
    book.redo(); assert.deepEqual(h.current().sheets[0].cells, amended.sheets[0].cells)
  }
})

test('clipboard value ranges ignore modeled run fonts, empty styled cells and geometry but retain sparse content changes', () => {
  const before = formatsModel().sheets[0], after = copy(before)
  after.cells['4:4'].richText = copy(formattedRuns); after.cells['4:4'].style.fill = '#abcdef'
  after.cells['6:8'] = sdk.createCell(6, 8, null); after.cells['6:8'].style.font.bold = true
  after.columnWidths[4] = 200; after.mergedRanges = [range(8, 8, 8, 9)]
  assert.deepEqual(clipboardModel.sheetClipboardChangedValueRanges(before, after), [])
  after.cells['4:4'].richText[0].hyperlink = 'https://changed.test/run'
  after.cells['4:5'] = sdk.createCell(4, 5, 9); after.cells['6:8'].value = 'sparse'; after.cells['6:8'].dataType = 'text'
  assert.deepEqual(clipboardModel.sheetClipboardChangedValueRanges(before, after), [range(4, 4, 4, 5), range(6, 8)])
  assert.equal(before.cells['4:4'].richText[0].hyperlink, 'https://target.test/run')
})

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

test('formula paste writes ordinary values and sparse blanks while copy-formulas and cut retain history, target formatting and translated references', async t => {
  for (const standalone of [false, true]) for (const command of ['sheet.command.copy', 'sheet.command.copy-formula-only', 'sheet.command.cut']) {
    const original = formatsModel(), sheet = original.sheets[0]
    sheet.cells['0:0'] = copy(model().sheets[0].cells['0:0']); sheet.cells['0:2'] = sdk.createCell(0, 2, 8)
    sheet.cells['4:5'] = { ...sdk.createCell(4, 5, 4), formula: '=2+2', dataType: 'formula', cachedDataType: 'number' }
    sheet.cells['4:6'] = sdk.createCell(4, 6, 'old number'); sheet.cells['4:7'] = sdk.createCell(4, 7, 'old blank')
    const h = host(t, { standalone, model: original, permission: request => request.point !== 'setCellStyle' })
    const initial = copy(h.current()), target = copy(h.cell(4, 4)), values = [], dimensions = []
    h.owner.value?.events.addEvent('SheetValueChanged', event => values.push(copy(event)))
    h.owner.value?.events.addEvent('SheetDimensionsChanged', event => dimensions.push(copy(event)))
    h.select(range(0, 0, 0, 4)); const copied = await h.ui.runClipboardCommand(command); h.select(range(4, 4))
    assert.equal(h.ui.capabilities.value.paste.formulas.disabled, false)
    const result = await h.ui.runClipboardCommand(command === 'sheet.command.copy-formula-only' ? 'sheet.command.paste' : 'sheet.command.paste-formula',
      { clipboardId: copied.clipboard.id })
    assert.equal(result.accepted, true); assert.equal(result.changed, true)
    assert.equal(h.cell(4, 4).value, command === 'sheet.command.copy-formula-only' ? null : 'rich')
    assert.equal(h.cell(4, 4).formula, null); assert.equal(h.cell(4, 4).richText, undefined)
    assert.deepEqual(h.cell(4, 4).style, target.style); assert.equal(h.cell(4, 4).hyperlink, target.hyperlink)
    assert.equal(h.cell(4, 5).formula, '=E6+$A$2'); assert.equal(h.cell(4, 5).value, null); assert.equal(h.cell(4, 5).cachedDataType, null)
    assert.equal(h.cell(4, 6).value, command === 'sheet.command.copy-formula-only' ? null : 8)
    assert.equal(h.cell(4, 7).value, null); assert.equal(h.cell(4, 8), undefined)
    for (const key of ['0:0', '0:1', '0:2', '1:0']) assert.deepEqual(h.current().sheets[0].cells[key], initial.sheets[0].cells[key])
    assert.equal(h.ui.pasteOptions.value.mode, 'formulas'); assert.equal(h.history().undo, standalone ? 0 : 1)
    if (!standalone) assert.deepEqual(values.map(event => event.ranges.map(domainRange)), [[range(4, 4, 4, 7)]])
    assert.deepEqual(dimensions, [])
    const pasted = copy(h.current()), repeated = await h.ui.runClipboardCommand('sheet.command.optional-paste', { type: 'SPECIAL_PASTE_FORMULA' })
    assert.equal(repeated.accepted, true); assert.equal(repeated.changed, false)
    assert.deepEqual(h.current(), pasted); assert.equal(h.calls.changed.length, 1); assert.equal(values.length, standalone ? 0 : 1)
    await nextTick(); assert.deepEqual(h.calls.focused.at(-1), range(4, 4, 4, 8))
    if (standalone) continue
    const api = sdk.FOffice.newAPI(h.owner.value), book = api.getActiveWorkbook()
    book.undo(); assert.deepEqual(h.current().sheets[0].cells, initial.sheets[0].cells)
    book.redo(); assert.deepEqual(h.current().sheets[0].cells, pasted.sheets[0].cells)
    const reopened = api.createWorkbook({ ...book.save(), id: 'clipboard-formulas-reopened-' + command })
    assert.deepEqual(h.owner.value.getWorkbookModel(reopened.getId()).sheets[0].cells, pasted.sheets[0].cells)
  }
  for (const standalone of [false, true]) {
    const h = host(t, { standalone, permission: request => request.point !== 'setCellValue' }), values = []
    h.owner.value?.events.addEvent('SheetValueChanged', event => values.push(copy(event)))
    const copied = await h.ui.runClipboardCommand('sheet.command.copy'); h.select(range(4, 4)); const before = copy(h.current())
    const rejected = await h.ui.runClipboardCommand('sheet.command.paste-formula', { clipboardId: copied.clipboard.id })
    assert.equal(rejected?.accepted ?? false, false); assert.deepEqual(h.current(), before)
    assert.deepEqual(values, []); assert.equal(h.calls.changed.length, 0); assert.equal(h.history().undo, 0)
    assert.equal(h.ui.pasteOptions.value, undefined)
  }
})

test('column width paste skips hidden and zero-width columns without value validation and preserves a rejected recent replacement', async t => {
  for (const standalone of [false, true]) {
    const original = model(), source = original.sheets[0], target = original.sheets[1]
    source.columnWidths = { 0: 123, 1: 156, 2: 180, 3: 200, 4: 0, 5: 250 }
    target.columnWidths = { 4: 190, 5: 200, 6: 210, 7: 220, 8: 230, 9: 0 }
    target.hiddenColumns = [5]; target.columnHiddenSpans = [{ start: 6, end: 6, value: true }]
    target.outline = { summaryAbove: false, summaryLeft: false, groups: [{ id: 'collapsed-column', axis: 'column', start: 7, end: 7, collapsed: true }] }
    target.cells['4:4'] = sdk.createCell(4, 4, 'target content')
    target.dataValidations = [{ id: 'reject-values', type: 'custom', formula1: '=FALSE()', allowBlank: false,
      showErrorMessage: true, errorStyle: 'stop', error: 'Value edits are blocked', ranges: [range(4, 4, 4, 9)] }]
    const h = host(t, { standalone, model: original, permission: request => !['setCellValue', 'setCellStyle'].includes(request.point) })
    const before = copy(h.current().sheets[1]), values = [], dimensions = []
    h.owner.value?.events.addEvent('SheetValueChanged', event => values.push(copy(event)))
    h.owner.value?.events.addEvent('SheetDimensionsChanged', event => dimensions.push(copy(event)))
    h.select(range(0, 0, 0, 5)); const copied = await h.ui.runClipboardCommand('sheet.command.copy')
    h.select(range(4, 4), target.id); assert.equal(h.ui.capabilities.value.paste.columnWidths.disabled, false)
    const result = await h.ui.runClipboardCommand('sheet.command.paste-column-widths', { clipboardId: copied.clipboard.id })
    assert.equal(result.accepted, true); assert.equal(result.changed, true); assert.deepEqual(result.issues, [])
    const pasted = copy(h.current().sheets[1])
    assert.deepEqual([4, 5, 6, 7, 8, 9].map(column => sdk.sheetColumnWidth(pasted, column)), [123, 200, 210, 220, 230, 0])
    assert.deepEqual(pasted.cells, before.cells); assert.deepEqual(pasted.dataValidations, before.dataValidations)
    assert.deepEqual(pasted.hiddenColumns, before.hiddenColumns); assert.deepEqual(pasted.columnHiddenSpans, before.columnHiddenSpans)
    assert.deepEqual(pasted.outline, before.outline); assert.deepEqual(values, [])
    if (!standalone) {
      assert.equal(dimensions.length, 1); assert.equal(dimensions[0].axis, 'column'); assert.equal(dimensions[0].automatic, false)
      assert.equal(dimensions[0].sheetId, target.id); assert.deepEqual(dimensions[0].ranges.map(domainRange), [range(0, 4, 39, 4)])
    }
    const repeated = await h.ui.runClipboardCommand('sheet.command.optional-paste', { type: 'SPECIAL_PASTE_COL_WIDTH' })
    assert.equal(repeated.accepted, true); assert.equal(repeated.changed, false); assert.equal(h.calls.changed.length, 1)
    h.permission.value = request => request.point !== 'setCellStyle'
    const current = copy(h.current()), depth = h.history(), options = copy(h.ui.pasteOptions.value)
    const rejected = await h.ui.runClipboardCommand('sheet.command.optional-paste', { type: 'SPECIAL_PASTE_FORMULA' })
    assert.equal(rejected.accepted, false); assert.equal(rejected.reason, 'validation')
    assert.deepEqual(h.current(), current); assert.deepEqual(h.history(), depth); assert.deepEqual(h.ui.pasteOptions.value, options)
    assert.equal(h.calls.changed.length, 1); assert.deepEqual(values, []); assert.equal(dimensions.length, standalone ? 0 : 1)
    if (standalone) continue
    const api = sdk.FOffice.newAPI(h.owner.value), book = api.getActiveWorkbook()
    book.undo(); assert.deepEqual(h.current().sheets[1], before)
    book.redo(); assert.deepEqual(h.current().sheets[1], pasted)
    const reopened = api.createWorkbook({ ...book.save(), id: 'clipboard-hidden-widths-reopened' })
    assert.deepEqual(h.owner.value.getWorkbookModel(reopened.getId()).sheets[1], pasted)
  }
})

test('uniform whole-text runs use ordinary external parsing and target value fonts while forced strings and rich extensions keep their content', () => {
  const target = formatsModel().sheets[0].cells['4:4'], position = { row: 4, col: 4 }
  const options = { rowDelta: 4, columnDelta: 4, date1904: false, region: 'en-US' }
  for (const font of [{ font: { family: 'Arial', size: 24, bold: false, color: '#123456' } },
    { fontFamily: 'Arial', fontSize: 24, bold: false, color: '#123456' }]) {
    const source = sdk.createCell(0, 0, '00123'); source.richText = [{ id: 'uniform', text: '00123', ...font }]
    source.style.fill = '#ffeeaa'; source.authoredStyle = copy(source.style); const before = copy(source)
    const internal = clipboardCells.createSheetClipboardPastedCell(source, target, position, 'values', options)
    assert.equal(internal.value, '00123'); assert.equal(internal.richText, undefined); assert.deepEqual(internal.style, target.style)
    for (const mode of ['all', 'values']) {
      const pasted = clipboardCells.createSheetClipboardPastedCell(source, target, position, mode, { ...options, externalFormat: 'html' })
      assert.equal(pasted.value, 123); assert.equal(pasted.dataType, 'number'); assert.equal(pasted.formula, null); assert.equal(pasted.richText, undefined)
      if (mode === 'values') { assert.deepEqual(pasted.style, target.style); assert.equal(pasted.hyperlink, target.hyperlink); continue }
      assert.equal(pasted.style.font.family, 'Arial'); assert.equal(pasted.style.font.size, 24); assert.equal(pasted.style.font.bold, false)
      assert.equal(pasted.style.font.color, '#123456'); assert.equal(pasted.style.fill, source.style.fill)
      assert.equal(pasted.authoredStyle.font.family, 'Arial'); assert.equal(pasted.authoredStyle.font.size, 24)
    }
    const textTarget = copy(target); textTarget.style.numberFormat = '@'; textTarget.authoredStyle.numberFormat = '@'
    const text = clipboardCells.createSheetClipboardPastedCell(source, textTarget, position, 'values', { ...options, externalFormat: 'html' })
    assert.equal(text.value, '00123'); assert.deepEqual(text.style, textTarget.style); assert.equal(text.richText, undefined)
    for (const type of ['force_string', 'force-string', 'forcestring']) for (const literal of ['00123', '=1+1']) {
      const forced = { ...copy(source), value: literal, dataType: type, richText: [{ id: 'uniform', text: literal, ...font }] }
      for (const mode of ['all', 'values']) {
        const pasted = clipboardCells.createSheetClipboardPastedCell(forced, target, position, mode, { ...options, externalFormat: 'html' })
        assert.equal(pasted.value, literal); assert.equal(pasted.dataType, type); assert.equal(pasted.formula, null)
      }
    }
    assert.deepEqual(source, before)
  }
  for (const runs of [
    [{ text: '00', font: { family: 'Arial', size: 24 } }, { text: '123', font: { family: 'Times New Roman', size: 18 } }],
    [{ text: '00123', font: { family: 'Arial', size: 24, language: 'ja-JP' } }],
    [{ text: '00123', font: { family: 'Arial', size: 24, va: 1 } }],
    [{ text: '00123', font: { family: 'Arial', size: 24 }, vendor: { keep: 1 } }],
    [{ text: '00123', font: { family: 'Arial', size: 24 }, hyperlink: 'https://source.test/run' }],
    [{ text: '00\n123', font: { family: 'Arial', size: 24 } }],
  ]) {
    const source = { ...sdk.createCell(0, 0, runs.map(run => run.text).join('')), richText: runs }
    const pasted = clipboardCells.createSheetClipboardPastedCell(source, target, position, 'values', { ...options, externalFormat: 'html' })
    assert.equal(pasted.value, source.value); assert.deepEqual(pasted.richText, runs); assert.deepEqual(pasted.style, target.style)
  }
  const source = sdk.createCell(0, 0, '00123'); source.richText = [{ text: '00123', font: { family: 'Arial', size: 24 } }]
  source.image = { id: 'uniform-image', resource: imageSource.sheetImageResourceFromDataUrl(sheetImagePng, 'uniform.png', 'uniform-resource') }
  source.metadata = { publicBody: sdk.sheetCellImageToPublicBody(source), sheetImagePublicDrawingId: source.image.id }
  const image = clipboardCells.createSheetClipboardPastedCell(source, target, position, 'values', { ...options, externalFormat: 'html' })
  assert.equal(image.value, '00123'); assert.deepEqual(image.richText, source.richText); assert.notEqual(image.image.id, source.image.id)
  assert.deepEqual(image.metadata.publicBody, source.metadata.publicBody)
  delete source.image; delete source.metadata.sheetImagePublicDrawingId
  const body = clipboardCells.createSheetClipboardPastedCell(source, target, position, 'values', { ...options, externalFormat: 'html' })
  assert.equal(body.value, '00123'); assert.deepEqual(body.richText, source.richText)
})

test('uniform-run values preserve target fonts and formula cleanup through recent history, undo, redo and internal reopen', async t => {
  for (const standalone of [false, true]) for (const formulaTarget of [false, true]) {
    const original = formatsModel(), source = original.sheets[0].cells['0:0']
    source.richText = [{ id: 'source-uniform', text: 'source', font: { family: 'Arial', size: 24, bold: false } }]
    if (formulaTarget) Object.assign(original.sheets[0].cells['4:4'], { value: 91, formula: '=90+1', dataType: 'formula', cachedDataType: 'number', display: '91' })
    const h = host(t, { standalone, model: original, permission: request => request.point !== 'setCellStyle' }), before = copy(h.cell(4, 4)), values = []
    h.owner.value?.events.addEvent('SheetValueChanged', event => values.push(copy(event)))
    h.select(range(0, 0)); const copied = await h.ui.runClipboardCommand('sheet.command.copy'); h.select(range(4, 4))
    assert.equal((await h.ui.runClipboardCommand('sheet.command.paste-values', { clipboardId: copied.clipboard.id })).accepted, true)
    assert.equal(h.cell(4, 4).value, 'source'); assert.equal(h.cell(4, 4).formula, null); assert.equal(h.cell(4, 4).richText, undefined)
    assert.deepEqual(h.cell(4, 4).style, before.style); assert.equal(h.cell(4, 4).hyperlink, before.hyperlink)
    assert.deepEqual(richFonts.sheetCellFontValues(h.view.value.sheets[0].cells.E5, 'family'), ['Courier New'])
    assert.deepEqual(richFonts.sheetCellFontValues(h.view.value.sheets[0].cells.E5, 'size'), [16])
    h.permission.value = undefined
    assert.equal((await h.ui.runClipboardCommand('sheet.command.optional-paste', { mode: 'all' })).accepted, true)
    assert.deepEqual(h.cell(4, 4).richText, source.richText)
    assert.equal((await h.ui.runClipboardCommand('sheet.command.optional-paste', { mode: 'values' })).accepted, true)
    const pasted = copy(h.cell(4, 4)), repeated = await h.ui.runClipboardCommand('sheet.command.optional-paste', { mode: 'values' })
    assert.equal(repeated.accepted, true); assert.equal(repeated.changed, false); assert.equal(pasted.richText, undefined)
    assert.deepEqual(pasted.style, before.style); assert.equal(h.calls.changed.length, 3); assert.equal(h.history().undo, standalone ? 0 : 1)
    assert.deepEqual(h.cell(0, 0).richText, source.richText); assert.equal(h.ui.pasteOptions.value.mode, 'values')
    if (!standalone) assert.deepEqual(values.map(event => event.ranges.map(domainRange)), [[range(4, 4)], [range(4, 4)], [range(4, 4)]])
    await nextTick(); assert.deepEqual(h.calls.focused.at(-1), range(4, 4))
    if (standalone) continue
    const api = sdk.FOffice.newAPI(h.owner.value), book = api.getActiveWorkbook()
    book.undo(); assert.deepEqual(h.cell(4, 4), before); book.redo(); assert.deepEqual(h.cell(4, 4), pasted)
    const reopened = api.createWorkbook({ ...book.save(), id: 'clipboard-uniform-run-reopened' })
    assert.deepEqual(h.owner.value.getWorkbookModel(reopened.getId()).sheets[0].cells['4:4'], pasted)
  }
})
