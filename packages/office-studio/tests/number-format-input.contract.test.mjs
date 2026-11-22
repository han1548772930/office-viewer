import assert from 'node:assert/strict'
import { test } from 'vitest'
import { parseSheetCellText } from '../src/domain/sheet-cell-value.ts'
import { createDateTimeEditPattern } from '../src/features/sheet-workbench/number-format-engine/date-time-edit-pattern.ts'
import { DateSystem } from '../src/features/sheet-workbench/number-format-engine/date-system.ts'
import { dateInfo } from '../src/features/sheet-workbench/number-format-engine/format-info.ts'
import { formatValue } from '../src/features/sheet-workbench/number-format-engine/format-number.ts'
import { decimalInputKey, shiftDecimalInput } from '../src/features/sheet-workbench/number-format-engine/number-input.ts'
import { defaultOptions } from '../src/features/sheet-workbench/number-format-engine/options.ts'
import { parsePattern } from '../src/features/sheet-workbench/number-format-engine/parse-pattern.ts'
import { isValidDate, parseBool, parseDate, parseDuration, parseNumber, parseTime, parseValue } from '../src/features/sheet-workbench/number-format-engine/parse-value.ts'
import { excelDateTimePartsToSerial, excelSerialToDateTimeParts } from '../src/features/sheet-workbench/number-format-engine/serial-date.ts'
import { parseExactDecimalInput } from '../src/shared/decimal-input.ts'

const options = { ...defaultOptions, locale: 'en' }
const render = (value, pattern, overrides = {}) => formatValue(value, parsePattern(pattern), { ...options, ...overrides })

test('localized numeric input validates grouping, accounting signs and currency placement', () => {
  for (const [input, locale, expected, kind] of [
    ['$1,234.50', 'en', 1234.5, 'currency'],
    ['($1,234.50)', 'en', -1234.5, 'currency'],
    ['1.234,50€', 'de', 1234.5, 'currency'],
    ['1\u202f234,50 €', 'fr', 1234.5, 'currency'],
    ['1,234,567', 'en', 1234567, 'grouped'],
    ['１２３４.５', 'en', 1234.5, 'number'],
    ['١٢٣٤.٥', 'en', 1234.5, 'number'],
    ['1.25E-3', 'en', .00125, 'scientific'],
  ]) {
    const parsed = parseNumber(input, { locale })
    assert.equal(parsed?.v, expected, input)
    assert.equal(parsed?.kind, kind, input)
  }
  for (const input of ['1,00', '1000,1.00', '12,3456', '1,234,56', '1,234.', '1,234.5,0', '$$12', '(12', '12)', '--12', '12%%', '€12%']) {
    assert.equal(parseNumber(input, { locale: 'en' }), null, input)
  }
})

test('numeric conversion preserves decimal intent and percent editing keeps its complete precision', () => {
  for (const input of ['12345678901234567', '1.0000000000000001', '1e309', '1e-400']) {
    assert.equal(parseNumber(input, { locale: 'en' }), null, input)
  }
  assert.equal(parseNumber('000123', { locale: 'en' })?.v, 123)
  assert.equal(parseNumber('100.1234567%', { locale: 'en' })?.v, 1.001234567)
  assert.equal(parseNumber('100,1234567%', { locale: 'de' })?.v, 1.001234567)
  assert.equal(shiftDecimalInput(String(1.001234567), 2), '100.1234567')
  assert.equal(shiftDecimalInput(String(.29), 2), '29')
  assert.equal(decimalInputKey('1.2500e3'), decimalInputKey('1250'))
  assert.equal(decimalInputKey('0.10'), decimalInputKey('.1'))
  assert.equal(parseNumber('0.0000000000001', { locale: 'en' })?.v, 1e-13)
})

test('clock parsing supports localized markers, seconds, decimal separators and complete midnight', () => {
  for (const [input, locale, expected, pattern] of [
    ['13:02:03.125', 'en', (13 * 3600 + 2 * 60 + 3.125) / 86400, 'hh:mm:ss.000'],
    ['1:02 PM', 'en', (13 * 3600 + 2 * 60) / 86400, 'h:mm AM/PM'],
    ['PM 1:02:03', 'en', (13 * 3600 + 2 * 60 + 3) / 86400, 'AM/PM h:mm:ss'],
    ['下午1:02:03', 'zh', (13 * 3600 + 2 * 60 + 3) / 86400, 'AM/PM h:mm:ss'],
    ['12:00 AM', 'en', 0, 'hh:mm AM/PM'],
    ['12:00 PM', 'en', .5, 'hh:mm AM/PM'],
    ['00:00:01,25', 'de', 1.25 / 86400, 'hh:mm:ss.00'],
    ['０１：０２：０３', 'en', 3723 / 86400, 'hh:mm:ss'],
    ['24:00:00', 'en', 1, 'hh:mm:ss'],
    ['12:', 'en', .5, 'hh:mm:ss'],
    ['12:34:', 'en', (12 * 3600 + 34 * 60) / 86400, 'hh:mm:ss'],
  ]) {
    const parsed = parseTime(input, { locale })
    assert.equal(parsed?.v, expected, input)
    assert.equal(parsed?.z, pattern, input)
  }
  for (const input of ['24:00:01', '24:01', '24:00:00.001', '13:00 PM', '0:00 AM', '-01:00', '12:60', '12:00:60', '5A', 'AM5', 'AM 5 PM', '12:00:00:00', '00:00:00.12345678901']) {
    assert.equal(parseTime(input, { locale: 'en' }), null, input)
  }
})

test('elapsed parsing keeps the selected unit and rejects lost seconds or malformed input', () => {
  assert.deepEqual(parseDuration('36:15:30.25'), { v: (36 * 3600 + 15 * 60 + 30.25) / 86400, z: '[h]:mm:ss.00' })
  assert.deepEqual(parseDuration('-36:00'), { v: -1.5, z: '[h]:mm' })
  assert.deepEqual(parseDuration('-2160:30,5', { locale: 'de', unit: 'minutes' }), { v: -(2160 * 60 + 30.5) / 86400, z: '[m]:ss.0' })
  assert.deepEqual(parseDuration('129600.25', { unit: 'seconds' }), { v: 129600.25 / 86400, z: '[s].00' })
  assert.equal(parseDuration('12'), null)
  assert.equal(parseDuration('12', { unit: 'hours' })?.v, .5)
  assert.equal(parseDuration('12', { unit: 'minutes' })?.v, 12 / 1440)
  assert.equal(parseDuration('12', { unit: 'seconds' })?.v, 12 / 86400)
  assert.equal(parseDuration('00:00:00.0000000001')?.v, 1e-10 / 86400)
  for (const input of ['1:60', '1:00:60', '1:30.5', '--36:00', '36:00 PM', '36:00:00:00', '123456789:00:00.0000000001']) {
    assert.equal(parseDuration(input), null, input)
  }
})

test('date parsing is deterministic with injected years and respects calendar and serial boundaries', () => {
  assert.equal(parseDate('1900-01-00')?.v, 0)
  assert.equal(parseDate('1900-02-29')?.v, 60)
  assert.equal(parseDate('1904-01-01', { dateSystem: DateSystem.Date1904 })?.v, 0)
  assert.equal(parseDate('1900-02-29', { dateSystem: DateSystem.Date1904 }), null)
  assert.equal(parseDate('29/02/2024', { locale: 'de' })?.v, 45351)
  assert.equal(parseDate('2/29/2024', { locale: 'en' })?.v, 45351)
  assert.equal(parseDate('２０２４年２月２９日', { locale: 'zh' })?.v, 45351)
  assert.equal(parseDate(' ２０２４年２月２９日 ', { locale: 'zh' })?.v, 45351)
  assert.equal(parseDate('2/29', { locale: 'en', currentYear: 2024 })?.v, 45351)
  assert.equal(parseDate('2/29', { locale: 'en', currentYear: 2025 }), null)
  assert.equal(parseDate('2024 / 02 / 29')?.v, 45351)
  assert.equal(parseDate('2024 / 02 - 29'), null)
  assert.equal(parseDate('2024-02-29T12:00:00')?.v, 45351.5)
  assert.equal(parseDate('1 23', { locale: 'en', currentYear: 2024 }), null)
  assert.equal(parseDate('9999-12-31 24:00:00'), null)
  assert.equal(isValidDate(2024.5, 2, 29), false)
  assert.equal(isValidDate(2024, 2, 29.5), false)
  assert.equal(isValidDate(2023, 2, 29), false)
})

test('ordinary scalar parsing remains separate from opt-in elapsed-unit interpretation', () => {
  assert.equal(parseValue('12')?.v, 12)
  assert.equal(parseValue('1.25')?.v, 1.25)
  assert.equal(parseBool('TRUE')?.v, true)
  assert.equal(parseBool('false')?.v, false)
  assert.equal(parseValue('invalid 12'), null)
})

test('legacy numeric text and normalized API inputs share the same precision-preserving boundary', () => {
  for (const type of [undefined, 'number', 'n', 'date', 2]) {
    for (const value of ['12345678901234567', '1.0000000000000001', '1e-400']) {
      assert.equal(parseSheetCellText(value, type), value)
    }
    assert.equal(parseSheetCellText('1.25e2', type), 125)
    assert.equal(parseSheetCellText('0.0000000000001', type), 1e-13)
  }
  assert.equal(parseSheetCellText('00123', 'text'), '00123')
  assert.equal(parseSheetCellText('00123', 4), '00123')
  assert.equal(parseSheetCellText('1', 'boolean'), true)
  assert.equal(parseExactDecimalInput('1.0000000000000001'), null)
  assert.equal(parseExactDecimalInput('0.29', 2), 29)
})

test('elapsed date metadata and editor patterns preserve units, calendar fields and second precision', () => {
  for (const [pattern, unit] of [['[h]:mm:ss', 'hours'], ['[m]:ss', 'minutes'], ['[s].000', 'seconds']]) {
    const parsed = parsePattern(pattern)
    assert.equal(dateInfo(parsed.partitions).isDuration, true)
    assert.equal(dateInfo(parsed.partitions).elapsedUnit, unit)
    assert.equal(createDateTimeEditPattern(parsed.partitions, 1.5), null)
  }
  assert.equal(createDateTimeEditPattern(parsePattern('mm-dd').partitions, 45351), 'yyyy-mm-dd')
  assert.equal(createDateTimeEditPattern(parsePattern('ddd').partitions, 45351), 'yyyy/mm/dd')
  assert.equal(createDateTimeEditPattern(parsePattern('hh:mm:ss.0000').partitions, .5), 'hh:mm:ss.0000')
  assert.equal(createDateTimeEditPattern(parsePattern('hh:mm').partitions, .5, 8), 'hh:mm:ss.00000000')
  assert.equal(createDateTimeEditPattern(parsePattern('hh:mm').partitions, .5, -1), 'hh:mm:ss')
  assert.equal(createDateTimeEditPattern(parsePattern('hh:mm').partitions, .5, Number.NaN), 'hh:mm:ss')
  assert.equal(createDateTimeEditPattern(parsePattern('[<1]hh:mm;yyyy-mm-dd').partitions, .5), 'hh:mm:ss')
  assert.equal(createDateTimeEditPattern(parsePattern('[<1]hh:mm;yyyy-mm-dd').partitions, 45351), 'yyyy-mm-dd')
})

test('clock and elapsed rendering keep subsecond precision and complete accumulated negative time', () => {
  assert.equal(render(-1.5, '[h]:mm:ss'), '-36:00:00')
  assert.equal(render(-1.5, '[m]:ss'), '-2160:00')
  assert.equal(render(-1.5, '[s].000'), '-129600.000')
  assert.equal(render(-.5, '[h]:mm'), '-12:00')
  assert.equal(render(1.5, '[h]:mm:ss'), '36:00:00')
  assert.equal(render(1.125 / 86400, 'hh:mm:ss.0000000000'), '00:00:01.1250000000')
  assert.equal(render(1e-8 / 86400, 'hh:mm:ss.0000000000'), '00:00:00.0000000100')
  assert.equal(render(1.25 / 86400, 'hh:mm:ss.0000', { locale: 'de' }), '00:00:01,2500')
  assert.throws(() => parsePattern('hh:mm:ss.00000000000'), /fractional second/)
})

test('time rounding carries calendar dates and rejects a rounded value beyond the final date', () => {
  const midnight = excelDateTimePartsToSerial({ year: 2024, month: 2, day: 28, hours: 23, minutes: 59, seconds: 59, fractionalSecond: .9996 }, { dateSystem: DateSystem.Date1900 })
  assert.equal(render(midnight, 'yyyy-mm-dd hh:mm:ss.000'), '2024-02-29 00:00:00.000')
  assert.equal(render(59 + 86399.9996 / 86400, 'yyyy-mm-dd hh:mm:ss.000'), '1900-02-29 00:00:00.000')
  assert.equal(render(2958465 + 86399.9996 / 86400, 'yyyy-mm-dd hh:mm:ss.000'), '######')
})

test('calendar rendering decodes the zero day in both date systems and normalized weekday indices', () => {
  assert.equal(render(0, 'yyyy-mm-dd'), '1900-01-00')
  assert.equal(render(.5, 'yyyy-mm-dd hh:mm:ss'), '1900-01-00 12:00:00')
  assert.equal(render(0, 'yyyy-mm-dd ddd', { dateSystem: DateSystem.Date1904 }), '1904-01-01 Fri')
  assert.equal(render(.5, 'yyyy-mm-dd hh:mm:ss', { dateSystem: DateSystem.Date1904 }), '1904-01-01 12:00:00')
  assert.equal(render(-7, 'ddd', { dateSpanLarge: true }), 'Sat')
})

test('serial decoding repairs canonical whole seconds without erasing real small fractions', () => {
  for (const wholeDays of [0, 60, 45000, 2958465]) {
    for (const totalSeconds of [1, 59, 60, 3600, 86399]) {
      const parts = excelSerialToDateTimeParts(wholeDays + totalSeconds / 86400, { dateSystem: DateSystem.Date1900 })
      assert.ok(parts)
      assert.equal(parts.hours * 3600 + parts.minutes * 60 + parts.seconds, totalSeconds)
      assert.equal(parts.fractionalSecond, 0)
    }
  }
  const tiny = excelSerialToDateTimeParts(1e-7 / 86400, { dateSystem: DateSystem.Date1900 })
  assert.ok(tiny)
  assert.ok(Math.abs(tiny.fractionalSecond - 1e-7) < 1e-20)
  assert.equal(tiny.seconds, 0)
})
