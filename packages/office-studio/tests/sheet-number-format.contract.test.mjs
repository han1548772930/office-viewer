import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { sheetNumberFormatMenu, SHEET_NUMBER_FORMATS } from '../src/features/sheet-workbench/sheet-number-format.ts'
import { applySheetNumberFormatCommand, applySheetViewNumberFormatCommand, numberFormatActionForCommand } from '../src/features/sheet-workbench/sheet-number-format-command.ts'
import { presentSheetCellNumber } from '../src/features/sheet-workbench/sheet-number-presentation.ts'
import { positionSheetNumberFormatRuns, renderSheetNumberFormat } from '../src/features/sheet-workbench/number-format-render.ts'
import { validateCustomNumberFormat } from '../src/features/sheet-workbench/custom-number-format.ts'
import { numberFormatCurrency, sheetCurrencyPattern } from '../src/features/sheet-workbench/sheet-currency.ts'
import { createSheetNumberFormatDraft, readSheetNumberFormatPreferences, rememberSheetNumberFormat,
  SHEET_NUMBER_FORMAT_PANEL_CURRENCIES, sheetNumberFormatDraftPattern } from '../src/features/sheet-workbench/number-format-panel.ts'

const range = (startCol, endCol = startCol) => ({ startRow: 0, endRow: 0, startCol, endCol })

test('format menu retains the complete ordered catalogue, separate percent button and regional currency', () => {
  const entries = sheetNumberFormatMenu('en-US')
  assert.deepEqual(entries.filter(entry => entry.type === 'format').map(entry => entry.id), [
    'general', 'text', 'number', 'percent', 'scientific', 'accounting', 'financial', 'currency', 'roundingCurrency', 'date', 'time', 'datetime', 'duration',
  ])
  assert.equal(entries.filter(entry => entry.type === 'divider').length, 4)
  assert.equal(entries.at(-1).type, 'more')
  assert.equal(entries.find(entry => entry.id === 'number').pattern, '0')
  assert.equal(entries.find(entry => entry.id === 'percent').pattern, '0.00%')
  assert.equal(SHEET_NUMBER_FORMATS.percent, '0%')
  assert.equal(entries.find(entry => entry.id === 'accounting').pattern, '"$" #,##0.00_);[Red]("$"#,##0.00)')
  assert.equal(entries.find(entry => entry.id === 'date').pattern, 'yyyy-mm-dd;@')
  assert.equal(entries.find(entry => entry.id === 'time').pattern, 'am/pm h":"mm":"ss')
  assert.equal(entries.find(entry => entry.id === 'datetime').pattern, 'yyyy-m-d am/pm h:mm')
  assert.deepEqual(numberFormatActionForCommand('sheet.command.set-percent-format'), { kind: 'percent' })
  assert.deepEqual(numberFormatActionForCommand('sheet.command.set-number-format', { pattern: '0.00%' }), { kind: 'set', pattern: '0.00%' })
})

test('regional currency templates resolve all formats and retain full currency tokens', () => {
  for (const [region, symbol] of [
    ['ar-SA', '﷼'], ['ca-ES', '€'], ['de-DE', '€'], ['en-US', '$'], ['es-ES', '€'],
    ['fa-IR', '﷼'], ['fr-FR', '€'], ['id-ID', 'Rp'], ['it-IT', '€'], ['ja-JP', '¥'],
    ['ko-KR', '₩'], ['pl-PL', 'zł'], ['pt-BR', 'R$'], ['ru-RU', '₽'], ['sk-SK', '€'],
    ['vi-VN', '₫'], ['zh-CN', '¥'], ['zh-HK', 'HK$'], ['zh-TW', 'NT$'],
  ]) {
    const menu = sheetNumberFormatMenu(region)
    const currency = `"${symbol}"#,##0.00_);[Red]("${symbol}"#,##0.00)`
    assert.equal(sheetCurrencyPattern(region), currency)
    assert.equal(menu.find(entry => entry.id === 'currency').pattern, currency)
    assert.equal(menu.find(entry => entry.id === 'accounting').pattern, `"${symbol}" #,##0.00_);[Red]("${symbol}"#,##0.00)`)
    assert.equal(menu.find(entry => entry.id === 'roundingCurrency').pattern, `"${symbol}"#,##0;[Red]"${symbol}"#,##0`)
    assert.equal(numberFormatCurrency(currency), symbol)
    assert.ok(SHEET_NUMBER_FORMAT_PANEL_CURRENCIES.includes(symbol))
  }
  assert.equal(numberFormatCurrency('[$CAD-en-CA]#,##0.00'), 'CAD')
  assert.equal(numberFormatCurrency('#,##0.00Rp'), 'Rp')
  assert.equal(numberFormatCurrency('0*€'), undefined)
  assert.equal(numberFormatCurrency('yyyy-mm-dd;@ "€"'), undefined)
})

test('currency and accounting drafts keep new region, explicit symbol and original sections independent', () => {
  const newCurrency = createSheetNumberFormatDraft('General', 'zh-CN', 'de-DE', [], 'currency')
  assert.equal(newCurrency.currency, '€')
  assert.equal(newCurrency.decimals, 2)
  assert.equal(createSheetNumberFormatDraft('General', 'zh-CN', 'de-DE', ['Rp'], 'currency').currency, 'Rp')
  const accounting = '"HK$" #,##0.00_);[Red]("HK$"#,##0.00);"-";@'
  const draft = createSheetNumberFormatDraft(accounting, 'zh-CN', 'de-DE', ['Rp'])
  assert.equal(draft.category, 'accounting')
  assert.equal(draft.currency, 'HK$')
  assert.equal(createSheetNumberFormatDraft('_("HK$"#,##0.00)', 'zh-CN', 'de-DE', []).category, 'accounting')
  assert.equal(sheetNumberFormatDraftPattern(draft), accounting)
  assert.equal(sheetNumberFormatDraftPattern({ ...draft, decimals: 3 }), '"HK$" #,##0.000_);[Red]("HK$"#,##0.000);"-";@')
  const imported = '#,##0.00[$Rp-421];[Red](#,##0.00[$Rp-421])'
  const importedDraft = createSheetNumberFormatDraft(imported, 'zh-CN', 'de-DE', [])
  assert.equal(importedDraft.negative, -1)
  assert.equal(sheetNumberFormatDraftPattern({ ...importedDraft, decimals: 1 }), '#,##0.0[$Rp-421];[Red](#,##0.0[$Rp-421])')
  assert.equal(sheetNumberFormatDraftPattern({ ...importedDraft, pattern: '', negative: 2 }), '"Rp"#,##0.00_);("Rp"#,##0.00)')
  const preferences = rememberSheetNumberFormat({ history: [], currencies: [] }, accounting, false)
  assert.deepEqual(readSheetNumberFormatPreferences({ getItem: () => JSON.stringify(preferences) }), { history: [], currencies: ['HK$'] })
})

test('overlapping range formatting changes only presentation, respects unknown caches and detects no-ops', () => {
  const sheet = createBlankSheet()
  sheet.cells['0:0'] = createCell(0, 0, 12.25)
  sheet.cells['0:1'] = { ...createCell(0, 1, '00123'), dataType: 'text' }
  sheet.cells['0:2'] = { ...createCell(0, 2, null), formula: '=UNSUPPORTED()', dataType: 'formula', cachedDataType: null, display: 'preserved source' }
  const content = Object.values(sheet.cells).map(cell => ({ value: cell.value, formula: cell.formula, dataType: cell.dataType, cachedDataType: cell.cachedDataType }))
  assert.equal(applySheetNumberFormatCommand(sheet, [range(0, 2), range(0, 1)], { kind: 'set', pattern: '0.0' }), true)
  assert.deepEqual(Object.values(sheet.cells).map(cell => ({ value: cell.value, formula: cell.formula, dataType: cell.dataType, cachedDataType: cell.cachedDataType })), content)
  assert.deepEqual(Object.values(sheet.cells).map(cell => cell.display), ['12.3', '00123', 'preserved source'])
  assert.equal(applySheetNumberFormatCommand(sheet, [range(0, 2)], { kind: 'set', pattern: '0.0' }), false)
})

test('domain and standalone view commands share decimal plans and keep text types', () => {
  const workbook = createBlankWorkbook()
  const sheet = workbook.sheets[0]
  sheet.cells['0:0'] = createCell(0, 0, 1.125)
  sheet.cells['0:1'] = { ...createCell(0, 1, '00123'), dataType: 'text' }
  const view = toExcelWorkbookView(workbook).sheets[0]
  delete view.cells.B1.rawValue
  const action = { kind: 'decimals', delta: 1 }
  assert.equal(applySheetNumberFormatCommand(sheet, [range(0, 1)], action), true)
  assert.equal(applySheetViewNumberFormatCommand(view, [range(0, 1)], action), true)
  assert.equal(sheet.cells['0:0'].style.numberFormat, '0.0000')
  assert.equal(view.cells.A1.rawNumberFormat, sheet.cells['0:0'].style.numberFormat)
  assert.equal(view.cells.B1.rawNumberFormat, sheet.cells['0:1'].style.numberFormat)
  assert.equal(view.cells.B1.displayValue, '00123')
})

test('a malformed pattern or protected target rejects the whole format plan before mutation', () => {
  const sheet = createBlankSheet()
  sheet.cells['0:0'] = createCell(0, 0, 12)
  const before = structuredClone(sheet)
  assert.throws(() => applySheetNumberFormatCommand(sheet, [], { kind: 'cells', values: [
    { row: 0, col: 0, pattern: '0.0' }, { row: 0, col: 1, pattern: '0"unfinished' },
  ] }))
  assert.deepEqual(sheet, before)
  assert.throws(() => applySheetNumberFormatCommand(sheet, [range(0, 16_384)], { kind: 'set', pattern: '0.0' }), RangeError)
  assert.deepEqual(sheet, before)
  sheet.rangeProtectionRules = [{ id: 'locked', permissionId: 'locked-permission', ranges: [range(1)],
    ownerUserId: 'owner', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView', description: '' }]
  const locked = structuredClone(sheet)
  assert.throws(() => applySheetNumberFormatCommand(sheet, [range(0, 1)], { kind: 'set', pattern: '0.0' }, 'zh-CN', false, 'guest'), /保护/)
  assert.deepEqual(sheet, locked)
})

test('preview formats numeric values while leaving typed text, formula caches and source fields intact', () => {
  const cells = [
    { value: '12.25', rawValue: 12.25, rawNumberFormat: 'General' },
    { value: '00123', dataType: 's', rawNumberFormat: 'General' },
    { value: '', formula: '=UNSUPPORTED()', rawValue: null, cachedValue: '25%', displayValue: '25%' },
    { value: '', formula: '=A1/100', rawValue: .1225, cachedValue: '12.25%', rawNumberFormat: '0.00%' },
  ]
  const before = structuredClone(cells)
  assert.deepEqual(cells.map(cell => presentSheetCellNumber(cell, false, 'zh-CN', '0.0').text), ['12.3', '00123', '25%', '0.1'])
  assert.deepEqual(cells, before)
})

test('Canvas format runs reserve underscore widths and repeat fill characters without changing the scalar', () => {
  const presentation = renderSheetNumberFormat(12, '*x0_)')
  assert.equal(presentation.text, '12 ')
  assert.deepEqual(presentation.runs, [{ kind: 'fill', character: 'x' }, { kind: 'text', text: '12' }, { kind: 'space', character: ')' }])
  const measure = text => text === '12' ? 10 : text === ')' ? 3 : text.length * 2
  assert.deepEqual(positionSheetNumberFormatRuns(presentation.runs, 20, measure), {
    runs: [{ text: 'xxx', offset: 0 }, { text: '12', offset: 7 }], width: 20,
  })
  assert.deepEqual(positionSheetNumberFormatRuns(presentation.runs, 5, measure), { runs: [{ text: '12', offset: 0 }], width: 13 })
  assert.deepEqual(renderSheetNumberFormat(12, '*\\x0_\\)').runs, presentation.runs)
  const negative = renderSheetNumberFormat(-12, '0_);[Red](0)')
  assert.equal(negative.color, '#ff0000')
  assert.equal(negative.text, '(12)')
  assert.equal(negative.runs, undefined)
})

test('custom format validation rejects controls inside quoted or escaped text as well as at the edges', () => {
  for (const character of ['\u0000', '\t', '\r', '\n', '\u007f', '\u0085']) {
    for (const pattern of ['0"a' + character + 'b"', '0\\' + character, character + '0', '0' + character]) {
      assert.throws(() => validateCustomNumberFormat(pattern), /控制字符/)
    }
  }
  assert.equal(validateCustomNumberFormat('  0.00" units"  ').pattern, '0.00" units"')
  assert.equal(validateCustomNumberFormat('0.00" units"').preview, '1234.57 units')
  assert.throws(() => validateCustomNumberFormat('0"unfinished'), /引号/)
  assert.throws(() => validateCustomNumberFormat('0;0;0;@;@'), /四个分段/)
})
