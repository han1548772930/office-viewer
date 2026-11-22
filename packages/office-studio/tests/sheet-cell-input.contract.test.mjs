import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { applyExcelWorkbookView, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { createSheetCellEditorStart, editSheetCellText, sheetCellInputText } from '../src/features/sheet-workbench/sheet-cell-edit.ts'
import { parseSheetNumberEdit, sheetNumberEditText } from '../src/features/sheet-workbench/sheet-number-format-edit.ts'
import { DateSystem } from '../src/features/sheet-workbench/number-format-engine/date-system.ts'
import { excelDateTimePartsToSerial } from '../src/features/sheet-workbench/number-format-engine/serial-date.ts'

const close = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-10, `${actual} != ${expected}`)
const numeric = (value, pattern) => ({ value: String(value), rawValue: value, dataType: 'number', rawNumberFormat: pattern })
const serial = (year, month, day, hours = 0, minutes = 0, seconds = 0, fractionalSecond = 0, dateSystem = DateSystem.Date1900) => {
  const value = excelDateTimePartsToSerial({ year, month, day, hours, minutes, seconds, fractionalSecond }, { dateSystem })
  assert.notEqual(value, null)
  return value
}

test('percent editing retains the full stored decimal intent and a locale-consistent separator', () => {
  for (const [value, expected] of [[1.001234567, '100.1234567%'], [.001234567893, '0.1234567893%'], [-.125, '-12.5%'], [0, '0%']]) {
    const cell = numeric(value, '0.00%')
    assert.equal(sheetCellInputText(cell), expected)
    assert.equal(editSheetCellText(cell, expected).rawValue, value)
    assert.equal(cell.rawValue, value)
    assert.equal(editSheetCellText(cell, expected).rawNumberFormat, '0.00%')
  }
  assert.equal(sheetCellInputText(numeric(.12345, '0%'), false, 'de-DE'), '12,345%')
  assert.equal(editSheetCellText(numeric(.1, '0.00%'), '12,345%', false, 'de-DE').rawValue, .12345)
  assert.equal(sheetCellInputText(numeric(.12345, '0.00%'), false, 'en-US'), '12.345%')
})

test('percent editor starts retain the suffix only for numeric cells and numeric overwrite input', () => {
  const percent = numeric(.125, '0.00%')
  assert.deepEqual(createSheetCellEditorStart(percent), { value: '12.5%', selection: { start: 4, end: 4 } })
  assert.deepEqual(createSheetCellEditorStart(percent, '5'), { value: '5%', selection: { start: 1, end: 1 } })
  assert.deepEqual(createSheetCellEditorStart(percent, '='), { value: '=', selection: { start: 1, end: 1 } })
  assert.deepEqual(createSheetCellEditorStart(percent, 'x'), { value: 'x', selection: { start: 1, end: 1 } })
  assert.deepEqual(createSheetCellEditorStart(percent, '５'), { value: '５%', selection: { start: 1, end: 1 } })
  assert.equal(createSheetCellEditorStart({ ...percent, formula: '=A1%' }, '5').value, '5')
  assert.equal(createSheetCellEditorStart({ value: '001%', rawValue: '001%', dataType: 'text', rawNumberFormat: '0%' }, '5').value, '5')
})

test('new number inputs use strict regional grouping, currencies, accounting signs and localized booleans', () => {
  for (const [source, region, expected] of [
    ['1.234,5', 'de-DE', 1234.5], ['€1.234,5', 'de-DE', 1234.5], ['(€1.234,5)', 'de-DE', -1234.5],
    ['1,234.5', 'en-US', 1234.5], ['1\u202f234,5', 'fr-FR', 1234.5], ['.125', 'en-US', .125],
    ['１２３４.５', 'zh-CN', 1234.5], ['TRUE', 'de-DE', true], ['FALSE', 'zh-CN', false],
  ]) {
    const cell = editSheetCellText(undefined, source, false, region)
    assert.equal(cell.rawValue, expected, `${source} (${region})`)
    assert.equal(cell.dataType, typeof expected === 'boolean' ? 'boolean' : 'number')
  }
})

test('malformed and lossy numeric inputs remain complete explicit text instead of approximate scalars', () => {
  for (const source of ['1000,1.00', '1,00,000', '1,000,', '1.2.3', '1e999', '123456789123456789', '1212121212121212.2345']) {
    const cell = editSheetCellText(undefined, source, false, 'en-US')
    assert.equal(cell.rawValue, source)
    assert.equal(cell.value, source)
    assert.equal(cell.dataType, 'text')
    assert.equal(sheetCellInputText(cell, false, 'en-US'), source)
  }
  for (const source of ['123456789123456789', '1212121212121212.2345']) {
    for (const dataType of [undefined, 'number']) {
      const legacy = { value: source, dataType, rawNumberFormat: '0.00%' }
      assert.equal(sheetCellInputText(legacy), source)
      assert.equal(createSheetCellEditorStart(legacy, '5').value, '5')
      assert.equal(editSheetCellText(legacy, source).rawValue, source)
    }
  }
})

test('text formats and apostrophes outrank number, calendar and formula detection', () => {
  for (const pattern of ['@', '@@@']) for (const source of ['1.234,5', '12.5%', '2026/11/22', '=1+1', "'00123"]) {
    const cell = editSheetCellText({ value: '', rawNumberFormat: pattern }, source, false, 'de-DE')
    assert.equal(cell.rawValue, source)
    assert.equal(cell.dataType, 'text')
    assert.equal(cell.formula, undefined)
    assert.equal(cell.rawNumberFormat, pattern)
  }
  const forced = editSheetCellText(undefined, "'12,345%", false, 'de-DE')
  assert.equal(forced.rawValue, '12,345%')
  assert.equal(forced.dataType, 'force_string')
  assert.equal(sheetCellInputText(forced, false, 'de-DE'), "'12,345%")
  assert.equal(editSheetCellText(undefined, '=1%').formula, '=1%')
})

test('detected format changes share the value edit while formats of the same type retain their settings', () => {
  const percent = editSheetCellText(undefined, '12.345%', false, 'en-US')
  assert.equal(percent.rawValue, .12345)
  assert.equal(percent.numberFormat, 'percent')
  assert.equal(percent.rawNumberFormat, '0.00%')
  const currency = numeric(1, '"¥"#,##0.000;[Red]("¥"#,##0.000)')
  const replacement = editSheetCellText(currency, '€2,50', false, 'de-DE')
  assert.equal(replacement.rawValue, 2.5)
  assert.equal(replacement.rawNumberFormat, currency.rawNumberFormat)
  assert.equal(editSheetCellText(numeric(1, '0.0000%'), '2,5%', false, 'de-DE').rawNumberFormat, '0.0000%')
  for (const [original, pattern, input, expected] of [
    [1, '[<0]0;0.0000%', '2,5%', .025],
    [1, '[<0]yyyy/mm/dd;0.0000%', '2,5%', .025],
    [-1, '0;0.0000%', '-2,5%', -.025],
    [-1, '0%;"€"0.000', '€2,50', 2.5],
  ]) {
    const edited = editSheetCellText(numeric(original, pattern), input, false, 'de-DE')
    assert.equal(edited.rawValue, expected)
    assert.equal(edited.rawNumberFormat, pattern)
  }
  const workbook = createBlankWorkbook()
  workbook.sheets[0].cells['0:0'] = createCell(0, 0, 1)
  const view = toExcelWorkbookView(workbook)
  view.sheets[0].cells.A1 = percent
  const committed = applyExcelWorkbookView(workbook, view, 'en-US')
  assert.equal(committed.sheets[0].cells['0:0'].value, .12345)
  assert.equal(committed.sheets[0].cells['0:0'].style.numberFormat, '0.00%')
  assert.equal(committed.sheets[0].cells['0:0'].dataType, 'number')
  const reopened = toExcelWorkbookView(JSON.parse(JSON.stringify(committed)))
  assert.equal(sheetCellInputText(reopened.sheets[0].cells.A1, false, 'en-US'), '12.345%')
  assert.equal(workbook.sheets[0].cells['0:0'].value, 1)
})

test('partial calendar formats expose complete dates and date edits retain a hidden time fraction', () => {
  const original = serial(2026, 11, 22, 14, 30, 15, .125)
  const cell = numeric(original, 'm/d')
  assert.equal(sheetCellInputText(cell), '2026/11/22')
  const next = editSheetCellText(cell, '2026/11/23')
  close(next.rawValue, original + 1)
  assert.equal(next.rawNumberFormat, 'm/d')
  assert.equal(cell.rawValue, original)
  const withClock = editSheetCellText(cell, '2026/11/23 01:02:03.25')
  assert.equal(withClock.rawValue, original + 1)
  assert.equal(withClock.rawNumberFormat, 'm/d')
  const datetime = editSheetCellText(numeric(original, 'yyyy/mm/dd h:mm'), '2026/11/23 01:02:03.25')
  assert.equal(datetime.rawValue, serial(2026, 11, 23, 1, 2, 3, .25))
  assert.equal(datetime.rawNumberFormat, 'yyyy/mm/dd h:mm')
  const tinyFraction = numeric(serial(2026, 11, 22, 0, 0, 0, .0000005), 'yyyy-mm-dd')
  assert.equal(editSheetCellText(tinyFraction, '2026-11-23').rawValue, tinyFraction.rawValue + 1)
  const invalid = editSheetCellText(cell, '2026/02/30')
  assert.equal(invalid.rawValue, '2026/02/30')
  assert.equal(invalid.dataType, 'text')
})

test('clock edits retain hidden dates, retain fractional seconds and accept an explicit next midnight', () => {
  const original = serial(2026, 11, 22, 12, 30)
  const cell = numeric(original, 'h:mm')
  assert.equal(sheetCellInputText(cell, false, 'de-DE'), '12:30:00')
  const next = editSheetCellText(cell, '13:14:15,25', false, 'de-DE')
  close(next.rawValue, serial(2026, 11, 22, 13, 14, 15, .25))
  assert.equal(next.rawNumberFormat, 'h:mm')
  assert.equal(editSheetCellText(cell, '24:00:00').rawValue, Math.floor(original) + 1)
  const detailed = numeric(serial(2026, 11, 22, 12, 30, 15, .125), 'yyyy/mm/dd h:mm')
  const text = sheetCellInputText(detailed)
  assert.match(text, /^2026\/11\/22 12:30:15\./)
  assert.equal(editSheetCellText(detailed, text).rawValue, detailed.rawValue)
  const prefixed = numeric(serial(2026, 11, 22, 13, 14, 15), 'AM/PM h:mm')
  assert.equal(sheetCellInputText(prefixed), '下午 1:14:15')
  assert.equal(editSheetCellText(prefixed, '下午 2:15:16.25').rawValue, serial(2026, 11, 22, 14, 15, 16, .25))
  assert.equal(editSheetCellText(prefixed, '下午 1:14:15').rawValue, prefixed.rawValue)
})

test('elapsed-time inputs retain their hours, minutes or seconds unit and permit negative values', () => {
  for (const [value, pattern, expected] of [[2, '[h]:mm:ss', '48:00:00'], [2, '[m]:ss', '2880:00'], [2, '[s]', '172800'], [-.5, '[h]:mm:ss', '-12:00:00']]) {
    const cell = numeric(value, pattern)
    assert.equal(sheetCellInputText(cell, false, 'de-DE'), expected)
    assert.equal(editSheetCellText(cell, expected, false, 'de-DE').rawValue, value)
    assert.equal(editSheetCellText(cell, expected, false, 'de-DE').rawNumberFormat, pattern)
  }
  close(editSheetCellText(numeric(0, '[m]:ss'), '90:30,25', false, 'de-DE').rawValue, (90 * 60 + 30.25) / 86400)
  close(editSheetCellText(numeric(0, '[s].000'), '90,25', false, 'de-DE').rawValue, 90.25 / 86400)
  const long = editSheetCellText(undefined, '49:02:03.25', false, 'en-US')
  close(long.rawValue, (49 * 3600 + 2 * 60 + 3.25) / 86400)
  assert.equal(long.numberFormat, 'duration')
  assert.equal(long.dataType, 'number')
  assert.equal(editSheetCellText(undefined, '-25:00:00').rawValue, -25 / 24)
  assert.equal(editSheetCellText(numeric(0, '[h]:mm:ss'), '49:61:00').rawValue, '49:61:00')
})

test('calendar editing retains serial zero, the compatibility leap day and the selected date system', () => {
  assert.equal(sheetCellInputText(numeric(0, 'yyyy-mm-dd')), '1900-01-00')
  assert.equal(editSheetCellText(numeric(0, 'yyyy-mm-dd'), '1900-01-01').rawValue, 1)
  assert.equal(sheetCellInputText(numeric(60.5, 'yyyy-mm-dd')), '1900-02-29')
  assert.equal(editSheetCellText(numeric(60.5, 'yyyy-mm-dd'), '1900-03-01').rawValue, 61.5)
  assert.equal(sheetCellInputText(numeric(0, 'yyyy-mm-dd'), true), '1904-01-01')
  assert.equal(editSheetCellText(numeric(.5, 'yyyy-mm-dd'), '1904-01-02', true).rawValue, 1.5)
  assert.equal(editSheetCellText(undefined, '1900-02-29', true).rawValue, '1900-02-29')
})

test('partial date input can use an injected current year and plain numbers remain ordinary scalars', () => {
  const result = parseSheetNumberEdit('11/22', undefined, 'General', { region: 'en-US', currentYear: 2026 })
  assert.equal(result.value, serial(2026, 11, 22))
  assert.ok(result.numberFormat)
  assert.equal(editSheetCellText(undefined, '12', false, 'en-US').rawValue, 12)
  assert.equal(sheetNumberEditText(1234.5, '#,##0.00', { region: 'de-DE' }), '1234,5')
})
