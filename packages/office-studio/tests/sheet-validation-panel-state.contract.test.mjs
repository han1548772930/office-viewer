import assert from 'node:assert/strict'
import { test } from 'vitest'
import {
  changeValidationType,
  formatValidationRange,
  validationPickedAddress,
  parseValidationRanges,
  validationPanelSettings,
  validationRuleFromPanel,
} from '../src/features/sheet-validation/validation-panel-state.ts'

test('validation ranges accept multiple addresses, current-sheet qualifiers, rows and columns', () => {
  assert.deepEqual(parseValidationRanges("'Sales, Q1'!$B$2:$D$4,C6", 'Sales, Q1'), [
    { startRow: 1, startCol: 1, endRow: 3, endCol: 3 },
    { startRow: 5, startCol: 2, endRow: 5, endCol: 2 },
  ])
  assert.deepEqual(parseValidationRanges('B:B;3:5', 'Sheet1'), [
    { startRow: 0, startCol: 1, endRow: 1_048_575, endCol: 1 },
    { startRow: 2, startCol: 0, endRow: 4, endCol: 16_383 },
  ])
})

test('validation range parsing rejects another sheet and addresses beyond Excel bounds', () => {
  assert.throws(() => parseValidationRanges("'Other Sheet'!A1", 'Sheet1'), RangeError)
  assert.throws(() => parseValidationRanges('XFE1', 'Sheet1'), TypeError)
  assert.throws(() => parseValidationRanges('A1048577', 'Sheet1'), TypeError)
})

test('picking keeps separate rule areas and accepts a single operand reference', () => {
  const ranges = parseValidationRanges('B2:C3,F7')
  assert.equal(validationPickedAddress('range', ranges), 'B2:C3, F7')
  for (const target of ['list', 'formula1', 'formula2']) {
    assert.equal(validationPickedAddress(target, ranges), undefined)
    assert.equal(validationPickedAddress(target, ranges.slice(1)), '=F7')
  }
  assert.equal(validationPickedAddress('range', []), undefined)
})

test('formatted sheet-qualified ranges can be parsed back without changing their coordinates', () => {
  const source = { startRow: 8, startCol: 26, endRow: 12, endCol: 27 }
  const formatted = formatValidationRange(source, "O'Brien Sheet")
  assert.equal(formatted, "'O''Brien Sheet'!AA9:AB13")
  assert.deepEqual(parseValidationRanges(formatted, "O'Brien Sheet"), [source])
})

test('list and checkbox rules preserve their specialized values and common validation options', () => {
  const listSettings = validationPanelSettings({
    id: 'list-rule', type: 'list', ranges: [{ startRow: 0, startCol: 0, endRow: 4, endCol: 0 }],
    formula1: '"Open,Closed"', formula2: '#1677ff,#ff4d4f', allowBlank: false,
    showDropDown: false, showErrorMessage: false, errorStyle: 'warning', error: 'Choose a listed status',
    showInputMessage: true, promptTitle: 'Status', prompt: 'Select a status', renderMode: 'custom',
  })
  const listRule = validationRuleFromPanel(listSettings, listSettings.type === 'list' ? [{ startRow: 0, startCol: 0, endRow: 4, endCol: 0 }] : [], {})
  assert.equal(listRule.formula1, '"Open,Closed"')
  assert.equal(listRule.formula2, '#1677ff,#ff4d4f')
  assert.equal(listRule.allowBlank, false)
  assert.equal(listRule.showDropDown, false)
  assert.equal(listRule.errorStyle, 'warning')
  assert.equal(listRule.renderMode, 'custom')

  const checkboxSettings = validationPanelSettings({
    id: 'checkbox-rule', type: 'checkbox', ranges: [{ startRow: 0, startCol: 1, endRow: 0, endCol: 1 }],
    formula1: 'DONE', formula2: 'TODO',
  })
  const checkboxRule = validationRuleFromPanel(checkboxSettings, [{ startRow: 0, startCol: 1, endRow: 0, endCol: 1 }], {})
  assert.equal(checkboxRule.formula1, 'DONE')
  assert.equal(checkboxRule.formula2, 'TODO')
})

test('switching between list modes retains items, colors and reference configuration', () => {
  const settings = validationPanelSettings({
    id: 'colored-list', type: 'list', ranges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }],
    formula1: '"Open,Closed"', formula2: '#1677ff,#ff4d4f',
  })
  const multi = changeValidationType(settings, 'listMultiple')
  assert.equal(multi.listMode, 'custom')
  assert.deepEqual(multi.listItems.map(item => [item.label, item.color]), [
    ['Open', '#1677ff'], ['Closed', '#ff4d4f'],
  ])
  const single = changeValidationType(multi, 'list')
  assert.equal(single.type, 'list')
  assert.equal(single.listMode, multi.listMode)
  assert.deepEqual(single.listItems, multi.listItems)
  assert.equal(single.listReference, multi.listReference)
})

test('validation panel uses antdv-next Form items for editable fields', async () => {
  const { readFile } = await import('node:fs/promises')
  const panel = await readFile(new URL('../src/components/spreadsheet/SheetDataValidationPanel.vue', import.meta.url), 'utf8')
  assert.match(panel, /import \{[^}]*\bForm,\s*FormItem\b[^}]*\} from 'antdv-next'/)
  assert.match(panel, /<Form v-else class="validation-body validation-form"/)
  assert.match(panel, /<FormItem :label=/)
  assert.doesNotMatch(panel, /<input\b|<select\b|<textarea\b/)
})
