import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { applyExcelWorkbookView, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { SHEET_COLOR_PRESETS, SHEET_STANDARD_COLORS, rememberSheetColor, sheetBorderColor, sheetColorDraftValue, sheetColorHex, sheetOpaqueColorHex } from '../src/features/sheet-workbench/sheet-colors.ts'
import { applySheetFontFormatCommand, applySheetViewFontFormatCommand, assertSheetFontFormatRanges, fontFormatActionForCommand, sheetFontFormatCommandForShortcut, sheetFontFormatInputRanges } from '../src/features/sheet-workbench/sheet-font-format-command.ts'
import { sheetToolbarColorLocales } from '../src/i18n/sheets/locales/toolbar-colors.ts'
import { supportedOfficeLocales } from '../src/i18n/regional-settings.ts'

const range = (startRow, endRow = startRow, startCol = 0, endCol = startCol) => ({ startRow, endRow, startCol, endCol })
const content = sheet => Object.entries(sheet.cells).map(([key, { value, formula, display, dataType, cachedDataType }]) => [key, { value, formula, display, dataType, cachedDataType }])
const sample = () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  workbook.id = 'font-color-contract'; workbook.calculationMode = 'manual'
  sheet.rowCount = 100; sheet.columnCount = 26
  sheet.cells['0:0'] = { ...createCell(0, 0, '00123'), dataType: 'text' }
  sheet.cells['1:0'] = { ...createCell(1, 0, .25), formula: '=1/4', dataType: 'formula', cachedDataType: 'number' }
  sheet.cells['2:0'] = createCell(2, 0, 7)
  sheet.cells['2:0'].style.font.bold = true
  sheet.cells['4:4'] = createCell(4, 4, 'outside')
  return workbook
}

test('the complete opaque palette has five ordered rows of nine unique colors', () => {
  assert.deepEqual(SHEET_COLOR_PRESETS.map(row => row.length), [9, 9, 9, 9, 9])
  assert.equal(SHEET_STANDARD_COLORS.length, 45)
  assert.equal(new Set(SHEET_STANDARD_COLORS).size, 45)
  assert.deepEqual(SHEET_COLOR_PRESETS[0], ['#ffffff', '#e1effe', '#fde8e8', '#feecdc', '#fff4b9', '#def7ec', '#d5f5f6', '#edebfe', '#fce8f3'])
  assert.equal(SHEET_COLOR_PRESETS[2][1], '#3f83f8')
  assert.equal(SHEET_COLOR_PRESETS[4][8], '#751a3d')
  assert.ok(SHEET_STANDARD_COLORS.every(value => /^#[0-9a-f]{6}$/.test(value)))
})

test('picker hex normalization accepts three or six digits without treating imported CSS colors as picker values', () => {
  assert.equal(sheetColorHex(' #AbC '), '#aabbcc')
  assert.equal(sheetColorHex('#0A1B2C'), '#0a1b2c')
  for (const value of ['abc', '#ab', '#abcd', '#12345678', '#zzzzzz', 'red', '', null, 12]) assert.equal(sheetColorHex(value), undefined)
})

test('theme defaults use opaque RGB values from light and dark ConfigProvider text tokens', () => {
  assert.equal(sheetOpaqueColorHex('rgba(0, 0, 0, 0.88)'), '#000000')
  assert.equal(sheetOpaqueColorHex('rgba(255, 255, 255, 0.85)'), '#ffffff')
  assert.equal(sheetOpaqueColorHex('rgb(48, 100, 199)'), '#3064c7')
  assert.equal(sheetOpaqueColorHex('#abc'), '#aabbcc')
  for (const value of ['rgb(256,0,0)', 'rgba(0,0,0,2)', 'rgb(-1,0,0)', 'transparent']) assert.equal(sheetOpaqueColorHex(value), undefined)
})

test('incomplete HEX drafts cannot confirm the previous valid picker scalar', () => {
  assert.equal(sheetColorDraftValue('hex', ['0A1B2C']), '#0a1b2c')
  assert.equal(sheetColorDraftValue('hex', ['#0a1b2c']), '#0a1b2c')
  for (const fields of [[''], ['12'], ['abc'], ['#12345'], ['12345678'], ['12345g'], ['123456', '00'], []]) assert.equal(sheetColorDraftValue('hex', fields), undefined)
})

test('RGB drafts require three complete integer channels within their visible limits', () => {
  assert.equal(sheetColorDraftValue('rgb', ['10', '20', '30']), '#0a141e')
  assert.equal(sheetColorDraftValue('rgb', ['0', '255', '128']), '#00ff80')
  for (const fields of [['', '20', '30'], ['256', '0', '0'], ['-1', '0', '0'], ['1.5', '0', '0'], ['10%', '0', '0'], ['1e2', '0', '0'], ['10', '20']]) assert.equal(sheetColorDraftValue('rgb', fields), undefined)
})

test('HSB drafts convert primary hues, percentages, zero saturation and the 360-degree boundary', () => {
  for (const [hue, expected] of [[0, '#ff0000'], [60, '#ffff00'], [120, '#00ff00'], [240, '#0000ff'], [360, '#ff0000']]) assert.equal(sheetColorDraftValue('hsb', [String(hue), '100%', '100%']), expected)
  assert.equal(sheetColorDraftValue('hsb', ['180', '100', '50']), '#008080')
  assert.equal(sheetColorDraftValue('hsb', ['30', '0%', '50%']), '#808080')
  for (const fields of [['361', '100', '100'], ['10%', '100', '100'], ['0', '100.1', '100'], ['0', '100', ''], ['0', '100', '-1']]) assert.equal(sheetColorDraftValue('hsb', fields), undefined)
})

test('text and fill memories are independent, and reset, invalid or identical choices keep the existing memory', () => {
  const original = { color: '#123456' }
  const next = rememberSheetColor(original, 'background', '#AbC')
  assert.deepEqual(next, { color: '#123456', background: '#aabbcc' })
  assert.deepEqual(original, { color: '#123456' })
  for (const value of [null, undefined, '', 'red', '#1234', '#aabbcc']) assert.equal(rememberSheetColor(next, 'background', value), next)
})

test('border selection distinguishes absent, uniform and mixed edge colors including diagonals', () => {
  assert.equal(sheetBorderColor(undefined), null)
  assert.equal(sheetBorderColor({ bottom: { style: 'thin' } }), null)
  assert.equal(sheetBorderColor({ top: { color: '#AbC' }, diagonalDown: { color: '#aabbcc' } }), '#aabbcc')
  assert.equal(sheetBorderColor({ top: { color: '#123456' }, bottom: { color: '#654321' } }), undefined)
  assert.equal(sheetBorderColor({ left: { color: 'theme-color' }, right: { color: 'theme-color' }, metadata: 'kept' }), 'theme-color')
})

test('all supported language dictionaries contain the same nonempty Sheets color labels', () => {
  const keys = Object.keys(sheetToolbarColorLocales['en-US']).sort()
  assert.equal(keys.length, 13)
  for (const locale of supportedOfficeLocales) {
    const values = sheetToolbarColorLocales[locale]
    assert.deepEqual(Object.keys(values).sort(), keys)
    assert.ok(Object.values(values).every(value => typeof value === 'string' && value.length > 0), locale)
    assert.ok(values['sheet.color.options'].includes('{title}'), locale)
  }
  assert.equal(sheetToolbarColorLocales['en-US']['sheet.color.no-fill'], 'No Fill')
  assert.equal(sheetToolbarColorLocales['zh-CN']['sheet.color.invalid'], '请输入有效的颜色值')
})

test('public range aliases resolve to the same mark and color actions as toolbar commands', () => {
  for (const [field, ids] of [
    ['bold', ['sheet.command.toggle-bold', 'sheet.command.set-range-bold', 'sheet.command.set-bold']],
    ['italic', ['sheet.command.toggle-italic', 'sheet.command.set-range-italic', 'sheet.command.set-italic']],
    ['underline', ['sheet.command.toggle-underline', 'sheet.command.set-range-underline', 'sheet.command.set-underline']],
    ['strike', ['sheet.command.toggle-strike', 'sheet.command.set-range-stroke', 'sheet.command.set-stroke']],
  ]) for (const id of ids) {
    assert.deepEqual(fontFormatActionForCommand(id), { kind: 'mark', field })
    assert.deepEqual(fontFormatActionForCommand(id, { value: 0 }), { kind: 'mark', field, value: false })
    assert.deepEqual(fontFormatActionForCommand(id, true), { kind: 'mark', field, value: true })
  }
  for (const id of ['sheet.command.set-color', 'sheet.command.set-font-color', 'sheet.command.set-text-color', 'sheet.command.set-range-text-color']) assert.deepEqual(fontFormatActionForCommand(id, { value: '#AbC' }), { kind: 'color', field: 'color', value: '#aabbcc' })
  for (const id of ['sheet.command.set-background', 'sheet.command.set-background-color']) assert.deepEqual(fontFormatActionForCommand(id, '#abcdef'), { kind: 'color', field: 'background', value: '#abcdef' })
  assert.deepEqual(fontFormatActionForCommand('sheet.command.reset-range-text-color'), { kind: 'color', field: 'color', value: null })
  assert.equal(fontFormatActionForCommand('unregistered-command'), undefined)
})

test('font enum setters keep their contracts while invalid action parameters reject before changes', () => {
  assert.deepEqual(fontFormatActionForCommand('sheet.command.set-font-weight', { value: 'bold' }), { kind: 'mark', field: 'bold', value: true })
  assert.deepEqual(fontFormatActionForCommand('sheet.command.set-font-weight', null), { kind: 'mark', field: 'bold', value: null })
  assert.deepEqual(fontFormatActionForCommand('sheet.command.set-font-style', 'italic'), { kind: 'mark', field: 'italic', value: true })
  assert.deepEqual(fontFormatActionForCommand('sheet.command.set-color', 'red'), { kind: 'color', field: 'color', value: 'red' })
  assert.equal(fontFormatActionForCommand('sheet.command.set-color', ''), undefined)
  assert.throws(() => fontFormatActionForCommand('sheet.command.set-font-weight', true), TypeError)
  assert.throws(() => fontFormatActionForCommand('sheet.command.set-font-style', 'oblique'), TypeError)
  assert.throws(() => fontFormatActionForCommand('sheet.command.toggle-bold', 'true'), TypeError)
  assert.throws(() => fontFormatActionForCommand('sheet.command.set-color', 12), TypeError)
})

test('font shortcuts match desktop bindings and protect IME, Alt and already-handled keys', () => {
  const base = { ctrlKey: true, metaKey: false, shiftKey: false }
  for (const [key, command] of [['B', 'toggle-bold'], ['i', 'toggle-italic'], ['u', 'toggle-underline']]) assert.equal(sheetFontFormatCommandForShortcut({ ...base, key }), 'sheet.command.' + command)
  assert.equal(sheetFontFormatCommandForShortcut({ ...base, key: 'x', shiftKey: true }), 'sheet.command.toggle-strike')
  assert.equal(sheetFontFormatCommandForShortcut({ ...base, key: 'b', ctrlKey: false, metaKey: true }), 'sheet.command.toggle-bold')
  for (const extra of [{ ctrlKey: false }, { shiftKey: true }, { altKey: true }, { isComposing: true }, { keyCode: 229 }, { defaultPrevented: true }]) assert.equal(sheetFontFormatCommandForShortcut({ ...base, key: 'b', ...extra }), undefined)
  assert.equal(sheetFontFormatCommandForShortcut({ ...base, key: 'x' }), undefined)
})

test('single range coordinate aliases normalize without changing explicit multi-range input', () => {
  const ranges = [range(0), range(2)]
  assert.equal(sheetFontFormatInputRanges({ ranges }), ranges)
  assert.deepEqual(sheetFontFormatInputRanges({ range: { startRow: 0, endRow: 1, startColumn: 2, endColumn: 3 } }), [range(0, 1, 2, 3)])
  assert.deepEqual(sheetFontFormatInputRanges({ range: range(3, 4, 1, 2) }), [range(3, 4, 1, 2)])
  assert.equal(sheetFontFormatInputRanges(), undefined)
  for (const invalid of [range(-1), range(1, 0), range(0, 0, 2, 1), range(0, 1_048_576), range(0, 0, 0, 16_384), range(.5)]) assert.throws(() => assertSheetFontFormatRanges([invalid]), RangeError)
})

test('mixed-selection toggles set all targets and overlapping ranges never toggle a cell twice', () => {
  const sheet = sample().sheets[0], before = content(sheet)
  const ranges = [range(0, 2), range(1, 2)]
  assert.equal(applySheetFontFormatCommand(sheet, ranges, { kind: 'mark', field: 'bold' }), true)
  assert.deepEqual([0, 1, 2].map(row => sheet.cells[row + ':0'].style.font.bold), [true, true, true])
  assert.equal(applySheetFontFormatCommand(sheet, ranges, { kind: 'mark', field: 'bold' }), true)
  assert.deepEqual([0, 1, 2].map(row => sheet.cells[row + ':0'].style.font.bold), [false, false, false])
  assert.deepEqual(content(sheet), before)
  assert.equal(sheet.cells['4:4'].style.font.bold, false)
})

test('format commands skip filter-hidden rows and retain access to manually hidden rows', () => {
  const workbook = sample(), sheet = workbook.sheets[0]
  sheet.autoFilter = { range: range(0, 2), columns: {}, filteredRows: [1] }
  sheet.hiddenRows = [2]
  const view = toExcelWorkbookView(workbook).sheets[0]
  view.filteredRows = [1]
  const action = { kind: 'color', field: 'background', value: '#f05252' }
  assert.equal(applySheetFontFormatCommand(sheet, [range(0, 2)], action), true)
  assert.equal(applySheetViewFontFormatCommand(view, [range(0, 2)], action), true)
  assert.deepEqual([0, 1, 2].map(row => sheet.cells[row + ':0'].style.fill), ['#f05252', null, '#f05252'])
  assert.deepEqual(['A1', 'A2', 'A3'].map(key => view.cells[key].background), ['#f05252', null, '#f05252'])
})

test('sparse full-row resets touch existing cells only in both model and standalone paths', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  for (const row of [0, 2, 1_048_575]) {
    const cell = createCell(row, 0, 'kept')
    cell.style.font.color = '#123456'; cell.style.fill = '#abcdef'; cell.style.font.italic = true
    sheet.cells[row + ':0'] = cell
  }
  const view = toExcelWorkbookView(workbook).sheets[0], before = content(sheet)
  const ranges = [range(0, 0, 0, 16_383), range(1_048_575, 1_048_575, 0, 16_383)]
  for (const field of ['color', 'background']) {
    assert.equal(applySheetFontFormatCommand(sheet, ranges, { kind: 'color', field, value: null }), true)
    assert.equal(applySheetViewFontFormatCommand(view, ranges, { kind: 'color', field, value: null }), true)
    assert.equal(applySheetFontFormatCommand(sheet, ranges, { kind: 'color', field, value: null }), false)
    assert.equal(applySheetViewFontFormatCommand(view, ranges, { kind: 'color', field, value: null }), false)
  }
  assert.equal(Object.keys(sheet.cells).length, 3)
  assert.equal(Object.keys(view.cells).length, 3)
  assert.equal(sheet.cells['2:0'].style.font.color, '#123456')
  assert.equal(sheet.cells['2:0'].style.fill, '#abcdef')
  assert.deepEqual(content(sheet), before)
  assert.ok(Object.values(sheet.cells).every(cell => cell.style.font.italic))
})

test('model and standalone formatting produce the same fields without altering values or formula caches', () => {
  const workbook = sample(), sheet = workbook.sheets[0], view = toExcelWorkbookView(workbook).sheets[0]
  const before = content(sheet), ranges = [range(0, 2)]
  for (const action of [
    { kind: 'mark', field: 'italic' }, { kind: 'mark', field: 'underline', value: true },
    { kind: 'mark', field: 'strike', value: true }, { kind: 'color', field: 'color', value: '#0a141e' },
    { kind: 'color', field: 'background', value: '#ff8000' },
  ]) {
    assert.equal(applySheetFontFormatCommand(sheet, ranges, action), true)
    assert.equal(applySheetViewFontFormatCommand(view, ranges, action), true)
  }
  for (const row of [0, 1, 2]) {
    const modelCell = sheet.cells[row + ':0'], viewCell = view.cells['A' + (row + 1)]
    for (const field of ['bold', 'italic', 'underline', 'strike', 'color']) assert.equal(viewCell[field], modelCell.style.font[field])
    assert.equal(viewCell.background, modelCell.style.fill)
  }
  assert.deepEqual(content(sheet), before)
  const roundtrip = applyExcelWorkbookView(workbook, { ...toExcelWorkbookView(workbook), sheets: [view] })
  assert.equal(roundtrip.sheets[0].cells['0:0'].value, '00123')
  assert.equal(roundtrip.sheets[0].cells['1:0'].formula, '=1/4')
  assert.equal(roundtrip.sheets[0].cells['1:0'].style.font.color, '#0a141e')
  assert.equal(roundtrip.sheets[0].cells['1:0'].style.fill, '#ff8000')
})

test('no-op clears and explicit false marks do not create blank cells or report content changes', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0], view = toExcelWorkbookView(workbook).sheets[0]
  for (const action of [{ kind: 'mark', field: 'bold', value: false }, { kind: 'color', field: 'background', value: null }]) {
    assert.equal(applySheetFontFormatCommand(sheet, [range(0, 1)], action), false)
    assert.equal(applySheetViewFontFormatCommand(view, [range(0, 1)], action), false)
  }
  assert.deepEqual(sheet.cells, {})
  assert.deepEqual(view.cells, {})
})

test('a protected target rejects the complete model and view operation before any cell is modified', () => {
  const workbook = sample(), sheet = workbook.sheets[0]
  sheet.rangeProtectionRules = [{ id: 'locked', permissionId: 'locked-permission', ranges: [range(1)], ownerUserId: 'owner', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView', description: '' }]
  const view = toExcelWorkbookView(workbook).sheets[0], beforeModel = structuredClone(sheet), beforeView = structuredClone(view)
  const action = { kind: 'color', field: 'color', value: '#f05252' }
  assert.throws(() => applySheetFontFormatCommand(sheet, [range(0, 2)], action, 'guest'), /保护/)
  assert.throws(() => applySheetViewFontFormatCommand(view, [range(0, 2)], action, 'guest'), /保护/)
  assert.deepEqual(sheet, beforeModel)
  assert.deepEqual(view, beforeView)
})

let vite, OfficeStudio, FOffice, createSheetRibbon
beforeAll(async () => {
  vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio, FOffice } = await vite.ssrLoadModule('/packages/office-studio/src/index.ts'))
  ;({ createSheetRibbon } = await vite.ssrLoadModule('/packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts'))
}, { timeout: 60_000 })
afterAll(async () => { await vite?.close() })
function withStudio(run) {
  const studio = new OfficeStudio({ workbook: sample(), currentUserId: 'local-user' })
  try { return run(studio, FOffice.newAPI(studio)) } finally { studio.dispose() }
}

test('font ribbon retains eleven classic roots and the exact seven second-row positions', () => {
  const font = createSheetRibbon({ hasSheet: true, fontEditable: true, color: '#f05252', background: null, borderColor: null, lastTextColor: '#3f83f8', defaultBackgroundColor: '#3064c7' })[0].groups.find(group => group.id === 'font')
  assert.equal(font.items.length, 11)
  assert.deepEqual(font.items.slice(4).map(item => [item.id, item.gridLayout.row, item.gridLayout.column]), [
    ['sheet.command.toggle-bold', 2, 1], ['sheet.command.toggle-italic', 2, 2], ['sheet.command.toggle-underline', 2, 3], ['sheet.command.toggle-strike', 2, 4],
    ['sheet.menu.text-color', 2, 7], ['sheet.menu.fill', 2, 6], ['sheet.menu.borders', 2, 5],
  ])
  const text = font.items.find(item => item.id === 'sheet.menu.text-color')
  assert.equal(text.params, '#3f83f8')
  assert.equal(text.popupProps.selectedColor, '#f05252')
  assert.equal(text.children.length, 46)
  const disabled = createSheetRibbon({ hasSheet: true, fontEditable: false })[0].groups.find(group => group.id === 'font')
  assert.ok(disabled.items.every(item => item.disabled))
  assert.ok(disabled.items.flatMap(item => item.children ?? []).every(item => item.disabled))
})

test('Facade marks and colors use registered commands, preserve content and share workbook history', () => withStudio((studio, api) => {
  const workbook = api.getActiveWorkbook(), sheet = workbook.getActiveSheet(), before = content(studio.getWorkbookModel(workbook.getId()).sheets[0]), executed = []
  studio.events.addEvent('CommandExecuted', event => executed.push(event.id))
  sheet.getRange('A1:A3').setFontWeight('bold').setFontStyle('italic').setFontColor('#AbC').setBackground('#ff8000')
  assert.deepEqual(executed, ['sheet.command.set-font-weight', 'sheet.command.set-font-style', 'sheet.command.set-font-color', 'sheet.command.set-background'])
  let model = studio.getWorkbookModel(workbook.getId())
  assert.deepEqual(content(model.sheets[0]), before)
  assert.deepEqual([0, 1, 2].map(row => model.sheets[0].cells[row + ':0'].style.font.color), ['#aabbcc', '#aabbcc', '#aabbcc'])
  assert.equal(studio.editor.history.units[workbook.getId()].undo.length, 4)
  workbook.undo()
  assert.equal(studio.getWorkbookModel(workbook.getId()).sheets[0].cells['0:0'].style.fill, null)
  workbook.redo()
  sheet.getRange('A1').setFontColor(null)
  assert.equal(studio.getWorkbookModel(workbook.getId()).sheets[0].cells['0:0'].style.font.color, null)
  workbook.undo()
  assert.equal(studio.getWorkbookModel(workbook.getId()).sheets[0].cells['0:0'].style.font.color, '#aabbcc')
  const saved = workbook.save(), reopened = api.createWorkbook({ ...saved, id: 'font-color-reopened' })
  model = studio.getWorkbookModel(reopened.getId())
  assert.deepEqual(content(model.sheets[0]), before)
  assert.equal(model.sheets[0].cells['0:0'].style.font.italic, true)
  assert.equal(model.sheets[0].cells['0:0'].style.fill, '#ff8000')
}))

test('alias commands default to the current selection and support explicit coordinate-style ranges', () => withStudio((studio, api) => {
  const book = api.getActiveWorkbook(), sheet = book.getActiveSheet()
  sheet.getRange('A1:A2').activate()
  assert.equal(studio.executeCommandSync('sheet.command.set-range-underline', undefined), true)
  assert.equal(studio.getWorkbookModel(book.getId()).sheets[0].cells['0:0'].style.font.underline, true)
  assert.equal(studio.getWorkbookModel(book.getId()).sheets[0].cells['2:0'].style.font.underline, false)
  assert.equal(studio.executeCommandSync('sheet.command.set-range-text-color', { unitId: book.getId(), subUnitId: sheet.getSheetId(), range: { startRow: 2, endRow: 2, startColumn: 0, endColumn: 0 }, value: '#F05252' }), true)
  assert.equal(studio.getWorkbookModel(book.getId()).sheets[0].cells['2:0'].style.font.color, '#f05252')
  assert.equal(studio.executeCommandSync('sheet.command.reset-range-text-color', { unitId: book.getId(), sheetId: sheet.getSheetId(), ranges: [range(2)] }), true)
  assert.equal(studio.getWorkbookModel(book.getId()).sheets[0].cells['2:0'].style.font.color, null)
}))

test('canceled, invalid and protected API commands leave model, command events and history unchanged', () => withStudio((studio, api) => {
  const book = api.getActiveWorkbook(), sheet = book.getActiveSheet(), events = [], before = structuredClone(studio.getWorkbookModel(book.getId()))
  studio.events.addEvent('CommandExecuted', event => events.push(event.id))
  const cancel = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === 'sheet.command.set-font-color') event.cancel = true })
  assert.throws(() => sheet.getRange('A1').setFontColor('#f05252'), /canceled/)
  cancel.dispose()
  assert.throws(() => sheet.getRange('A1').setFontStyle('oblique'), TypeError)
  assert.deepEqual(studio.getWorkbookModel(book.getId()), before)
  assert.deepEqual(events, [])
  assert.equal(studio.editor.history.units[book.getId()]?.undo.length ?? 0, 0)
  const locked = structuredClone(before)
  locked.sheets[0].rangeProtectionRules = [{ id: 'locked', permissionId: 'permission', ranges: [range(1)], ownerUserId: 'owner', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView', description: '' }]
  studio.editor.workbook.replace(locked)
  assert.throws(() => sheet.getRange('A1:A3').setBackground('#abcdef'), /保护|permission/i)
  assert.deepEqual(studio.getWorkbookModel(book.getId()), locked)
  assert.deepEqual(events, [])
}))

test('font history is isolated when a second workbook is active and unchanged color creates no undo entry', () => withStudio((studio, api) => {
  const first = api.getActiveWorkbook(), firstSheet = first.getActiveSheet()
  firstSheet.getRange('A1').setFontColor('#f05252')
  firstSheet.getRange('A1').setFontColor('#f05252')
  assert.equal(studio.editor.history.units[first.getId()].undo.length, 1)
  const second = api.createWorkbook({ ...sample(), id: 'font-color-second' })
  second.getActiveSheet().getRange('A1').setBackground('#3f83f8')
  first.undo()
  assert.equal(studio.getWorkbookModel(first.getId()).sheets[0].cells['0:0'].style.font.color, null)
  assert.equal(studio.getWorkbookModel(second.getId()).sheets[0].cells['0:0'].style.fill, '#3f83f8')
  assert.equal(studio.activeWorkbookId, second.getId())
  first.redo()
  assert.equal(studio.getWorkbookModel(first.getId()).sheets[0].cells['0:0'].style.font.color, '#f05252')
}))
