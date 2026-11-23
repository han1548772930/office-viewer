import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { normalizeWorkbook } from '../src/parser/workbookNormalizer.ts'
import { normalizeWorkbookDataValidations } from '../src/parser/workbookDataValidation.ts'
import { createSheetDataValidationRule, validateSheetDataValidationChanges, validationPromptForCell } from '../src/features/sheet-validation/validation-engine.ts'
import { sheetValidationUsesRelativeAnchor } from '../src/features/sheet-validation/validation-formulas.ts'
import { validationPanelSettings, validationRuleFromPanel } from '../src/features/sheet-validation/validation-panel-state.ts'

const range = { startRow: 1, startCol: 1, endRow: 1, endCol: 1 }
function source(overrides = {}) {
  return { id: 'dv', type: 'custom', ranges: [range, { startRow: 3, startCol: 3, endRow: 3, endCol: 3 }],
    formulaAnchor: { row: 1, col: 1 }, formula1: '=A1>0', allowBlank: false,
    showInputMessage: true, promptTitle: 'Input', prompt: 'Use the referenced value', renderMode: 'text',
    metadata: { colors: ['#123456'], future: { retained: true } }, vendor: { nested: ['keep'] }, ...overrides }
}

test('direct parser rules and explicit wrappers both keep ranges origins and extension data', () => {
  const direct = source()
  const wrapped = { rule: source({ id: undefined, uid: 'wrapped', type: 'CUSTOM', formulaAnchor: { row: 2, col: 2 } }) }
  const normalized = normalizeWorkbookDataValidations([direct, wrapped])
  assert.equal(normalized.length, 2)
  assert.equal(normalized[0].id, 'dv')
  assert.equal(normalized[1].id, 'wrapped')
  assert.deepEqual(normalized[1].formulaAnchor, { row: 2, col: 2 })
  assert.deepEqual(normalized[0].ranges, direct.ranges)
  assert.deepEqual(normalized[0].metadata, direct.metadata)
  assert.deepEqual(normalized[0].vendor, direct.vendor)
  direct.metadata.colors.push('#abcdef')
  direct.vendor.nested.push('changed')
  assert.deepEqual(normalized[0].metadata.colors, ['#123456'])
  assert.deepEqual(normalized[0].vendor.nested, ['keep'])
})

test('the production workbook normalizer preserves a disjoint rule predicate at both target cells', () => {
  const workbook = normalizeWorkbook({ schemaVersion: 1, parserVersion: 'draft', format: 'xlsx', fileName: 'validation.xlsx', size: 10,
    model: 'spreadsheet', payload: { modelType: 'spreadsheet', data: { sheets: [{ name: 'Data', cells: [
      { row: 0, col: 0, value: 5 }, { row: 1, col: 1, value: 'input' },
      { row: 2, col: 2, value: -5 }, { row: 3, col: 3, value: 'input' },
    ], dataValidations: [source()] }] } }, diagnostics: { warnings: [], errors: [], stats: {}, capabilities: [] } })
  const sheet = workbook.sheets[0]
  assert.equal(sheet.dataValidations.length, 1)
  assert.deepEqual(sheet.dataValidations[0].formulaAnchor, { row: 1, col: 1 })
  assert.equal(validateSheetDataValidationChanges(workbook, sheet.id, [{ row: 1, column: 1 }]).valid, true)
  assert.equal(validateSheetDataValidationChanges(workbook, sheet.id, [{ row: 3, column: 3 }]).blocked, true)
  assert.deepEqual(validationPromptForCell(sheet, 3, 3), { title: 'Input', message: 'Use the referenced value' })
})

test('invalid origins coverage typed settings and duplicate IDs reject before model installation', () => {
  for (const anchor of [{ row: '1', col: 1 }, { row: 0.5, col: 0 }, { row: -1, col: 0 }, { row: 1_048_576, col: 0 }, { row: 0, col: 16_384 }]) {
    assert.throws(() => normalizeWorkbookDataValidations([source({ formulaAnchor: anchor })]), /公式起点/)
  }
  for (const coverage of [[], [{ ...range, endRow: 0 }], [{ ...range, endCol: 16_384 }], [{ ...range, startRow: 1.5 }]]) {
    assert.throws(() => normalizeWorkbookDataValidations([source({ ranges: coverage })]))
  }
  assert.throws(() => normalizeWorkbookDataValidations([source(), source()]), /重复/)
  assert.throws(() => normalizeWorkbookDataValidations([source({ showDropDown: 'false' })]), /布尔值/)
  assert.throws(() => normalizeWorkbookDataValidations([source({ renderMode: 'invalid' })]), /显示方式/)
  assert.throws(() => normalizeWorkbookDataValidations([{ rule: 'invalid' }]), /包装/)
  assert.throws(() => normalizeWorkbookDataValidations({ rule: source() }), /列表/)
})

test('ordinary Office none rules normalize to any value and preserve guidance', () => {
  const [rule] = normalizeWorkbookDataValidations([source({ type: 'none', formula1: undefined, formulaAnchor: undefined })])
  assert.equal(rule.type, 'any')
  assert.equal(sheetValidationUsesRelativeAnchor(rule), false)
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.dataValidations = [rule]
  for (const value of [null, '', 'text', 123, true]) {
    sheet.cells['1:1'] = createCell(1, 1, value)
    assert.equal(validateSheetDataValidationChanges(workbook, sheet.id, [{ row: 1, column: 1 }]).valid, true)
  }
  sheet.cells['1:1'] = { ...createCell(1, 1, null), formula: '=#REF!' }
  assert.equal(validateSheetDataValidationChanges(workbook, sheet.id, [{ row: 1, column: 1 }]).valid, true)
  assert.equal(validationPromptForCell(sheet, 1, 1).message, rule.prompt)
})

test('panel any-value changes produce a real prompt rule without stale comparison operands', () => {
  const settings = { ...validationPanelSettings(null, 'any'), operator: 'between', formula1: '1', formula2: '10', showInputMessage: true,
    promptTitle: 'Guidance', prompt: 'Any value is accepted' }
  const rule = validationRuleFromPanel(settings, [range], {})
  assert.equal(rule.type, 'any')
  assert.equal(rule.operator, undefined)
  assert.equal(rule.formula1, undefined)
  assert.equal(rule.formula2, undefined)
  assert.equal(rule.prompt, 'Any value is accepted')
  assert.doesNotThrow(() => createSheetDataValidationRule({ type: 'any', operator: 'between' }, [range]))
})

test('checkbox custom values and multi-selection palettes retain their distinct editable contracts', () => {
  const [checkbox, multiple] = normalizeWorkbookDataValidations([
    source({ id: 'check', type: 'checkbox', formula1: 'yes', formula2: 'no' }),
    source({ id: 'multi', type: 'listMultiple', formula1: '"Red,Green"', formula2: '#ff0000,#00ff00', renderMode: 'custom' }),
  ])
  assert.equal(checkbox.formula2, 'no')
  assert.equal(multiple.formula2, '#ff0000,#00ff00')
  assert.equal(multiple.renderMode, 'custom')
  assert.equal(sheetValidationUsesRelativeAnchor(checkbox), false)
  assert.equal(sheetValidationUsesRelativeAnchor(multiple), false)
})

test('any-value API drafts still require valid coverage and typed guidance settings', () => {
  for (const ranges of [[{ ...range, endRow: 1_048_576 }], [{ ...range, endCol: 16_384 }], [{ ...range, startRow: 1.5 }]]) {
    assert.throws(() => createSheetDataValidationRule({ type: 'any' }, ranges), /有效/)
  }
  for (const settings of [{ showInputMessage: 'true' }, { prompt: 3 }, { errorStyle: 'invalid' }, { renderMode: 'invalid' }, { metadata: [] }]) {
    assert.throws(() => createSheetDataValidationRule({ type: 'any', ...settings }, [range]))
  }
  const draft = { type: 'any', showInputMessage: true, prompt: 'Guidance', metadata: { nested: ['keep'] }, vendor: { retained: true } }
  const rule = createSheetDataValidationRule(draft, [range])
  draft.metadata.nested.push('changed')
  draft.vendor.retained = false
  assert.deepEqual(rule.metadata.nested, ['keep'])
  assert.equal(rule.vendor.retained, true)
})
