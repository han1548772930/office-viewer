import assert from 'node:assert/strict'
import { test } from 'vitest'
import { OfficeRegionalSettings } from '../src/i18n/regional-settings.ts'
import { displayCell } from '../src/rendering/spreadsheet/scene.ts'
import { previewCustomNumberFormat } from '../src/features/sheet-workbench/custom-number-format.ts'

test('format region follows language until explicitly set and is isolated per instance', () => {
  const first = new OfficeRegionalSettings('zhCN'), other = new OfficeRegionalSettings('enUS')
  first.setLocale('jaJP')
  assert.equal(first.region.value, 'ja-JP')
  first.setRegion('deDE')
  first.setLocale('fr-FR')
  assert.equal(first.locale.value, 'fr-FR')
  assert.equal(first.region.value, 'de-DE')
  assert.equal(other.region.value, 'en-US')
  assert.throws(() => first.setRegion('bogus'), RangeError)
  assert.equal(first.region.value, 'de-DE')
  first.followLanguage()
  assert.equal(first.region.value, 'fr-FR')
})

test('regional display and preview agree without changing imported values or the date system', () => {
  const cell = Object.freeze({ rawNumberFormat: '#,##0.00', rawValue: 1234.5, value: '1234.5' })
  for (const [region, expected] of [['en-US', '1,234.50'], ['de-DE', '1.234,50']]) {
    assert.equal(displayCell(cell, true, region), expected)
    assert.equal(previewCustomNumberFormat(cell.rawNumberFormat, cell.rawValue, region, true), expected)
  }
  assert.equal(previewCustomNumberFormat('yyyy-mm-dd', 0, 'de-DE', true), '1904-01-01')
})
