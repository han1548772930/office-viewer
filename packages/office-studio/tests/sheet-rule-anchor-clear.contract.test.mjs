import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { conditionalFormatFormulaAt } from '../src/features/conditional-formatting/conditional-format-formulas.ts'
import { sheetValidationFormulaAt } from '../src/features/sheet-validation/validation-formulas.ts'
import { validateSheetDataValidationChanges, validationPromptForCell } from '../src/features/sheet-validation/validation-engine.ts'
import { applySheetClearCommand, applySheetViewClearCommand } from '../src/features/sheet-workbench/sheet-clear-command.ts'
import { toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'

const area = (startRow, startCol, endRow = startRow, endCol = startCol) => ({ startRow, startCol, endRow, endCol })
function sample() {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.cells['0:0'] = createCell(0, 0, 5)
  sheet.cells['2:2'] = createCell(2, 2, 'input')
  sheet.dataValidations = [{ id: 'dv', type: 'custom', formula1: '=A1>0', formulaAnchor: { row: 2, col: 2 },
    ranges: [area(0, 0), area(2, 2)], showInputMessage: true, prompt: 'Guidance', metadata: { nested: ['keep'] } }]
  sheet.conditionalFormats = [{ id: 'cf', type: 'expression', operator: '', formula1: '=A1>0',
    formulaAnchor: { row: 2, col: 2 }, ranges: [area(0, 0), area(2, 2)], priority: 1, stopIfTrue: false,
    style: { fill: '#123456' }, config: { min: { type: 'formula', value: '=SUM(A:A)' } } }]
  return workbook
}

test('clear-all trims the first range using the retained original predicate for validation and conditional rules', () => {
  const workbook = sample(), sheet = workbook.sheets[0]
  const beforeValidation = sheetValidationFormulaAt(sheet.dataValidations[0].formula1, sheet.dataValidations[0], 2, 2)
  const beforeConditional = conditionalFormatFormulaAt(sheet.conditionalFormats[0].formula1, sheet.conditionalFormats[0], 2, 2)
  applySheetClearCommand(sheet, [area(0, 0)], 'all')
  assert.equal(sheetValidationFormulaAt(sheet.dataValidations[0].formula1, sheet.dataValidations[0], 2, 2), beforeValidation)
  assert.equal(conditionalFormatFormulaAt(sheet.conditionalFormats[0].formula1, sheet.conditionalFormats[0], 2, 2), beforeConditional)
  assert.equal(sheet.dataValidations[0].formula1, '=A1>0')
  assert.equal(sheet.conditionalFormats[0].config.min.value, '=SUM(A:A)')
  assert.equal(validateSheetDataValidationChanges(workbook, sheet.id, [{ row: 2, column: 2 }]).blocked, true)
})

test('a partly invalid predicate remains recoverable when clear-all removes a later disjoint range', () => {
  const workbook = sample(), sheet = workbook.sheets[0]
  applySheetClearCommand(sheet, [area(2, 2)], 'all')
  assert.deepEqual(sheet.dataValidations[0].formulaAnchor, { row: 2, col: 2 })
  assert.equal(sheet.dataValidations[0].formula1, '=A1>0')
  assert.deepEqual(sheet.conditionalFormats[0].formulaAnchor, { row: 2, col: 2 })
  assert.equal(sheet.conditionalFormats[0].formula1, '=A1>0')
  assert.equal(sheetValidationFormulaAt(sheet.dataValidations[0].formula1, sheet.dataValidations[0], 0, 0), '=#REF!>0')
})

test('Canvas view and domain clear paths preserve the same source anchor and threshold formulas', () => {
  const workbook = sample(), modelSheet = workbook.sheets[0]
  const viewSheet = toExcelWorkbookView(workbook).sheets[0]
  applySheetClearCommand(modelSheet, [area(2, 2)], 'all')
  applySheetViewClearCommand(viewSheet, [area(2, 2)], 'all')
  assert.deepEqual(viewSheet.dataValidations, modelSheet.dataValidations)
  assert.deepEqual(viewSheet.conditionalFormats, modelSheet.conditionalFormats)
})

test('any-value guidance and literal list labels survive range trimming without formula translation', () => {
  const workbook = sample(), sheet = workbook.sheets[0]
  sheet.dataValidations = [
    { id: 'any', type: 'any', ranges: [area(0, 0), area(2, 2)], formula1: '=A1', showInputMessage: true, prompt: 'Any input' },
    { id: 'list', type: 'list', ranges: [area(0, 1), area(2, 3)], formula1: '"A1,B2"' },
  ]
  applySheetClearCommand(sheet, [area(0, 0, 0, 1)], 'all')
  assert.equal(sheet.dataValidations[0].formula1, '=A1')
  assert.equal(sheet.dataValidations[1].formula1, '"A1,B2"')
  assert.deepEqual(validationPromptForCell(sheet, 2, 2), { title: '数据验证', message: 'Any input' })
})
