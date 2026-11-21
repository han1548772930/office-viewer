import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import {
  applySheetDataValidationRule,
  clearSheetDataValidationRanges,
  createSheetDataValidationRule,
  dataValidationsAt,
  sheetDataValidationOptions,
  validateSheetDataValidationChanges,
} from '../src/features/sheet-validation/validation-engine.ts'

function setCell(sheet, row, column, value, formula = null) {
  sheet.cells[`${row}:${column}`] = { ...createCell(row, column, value), formula }
}

function rule(type, range, extra = {}) {
  return createSheetDataValidationRule({ type, ...extra }, [range])
}

test('integer and decimal rules apply every comparison and blank policy', () => {
  const workbook = createBlankWorkbook()
  const sheet = workbook.sheets[0]
  const range = { startRow: 0, startCol: 0, endRow: 0, endCol: 0 }

  sheet.dataValidations = [rule('whole', range, { operator: 'between', formula1: '1', formula2: '5', allowBlank: false })]
  setCell(sheet, 0, 0, 3)
  assert.equal(validateSheetDataValidationChanges(workbook, sheet.id, [{ row: 0, column: 0 }]).valid, true)
  setCell(sheet, 0, 0, 3.5)
  assert.equal(validateSheetDataValidationChanges(workbook, sheet.id, [{ row: 0, column: 0 }]).blocked, true)
  setCell(sheet, 0, 0, null)
  assert.equal(validateSheetDataValidationChanges(workbook, sheet.id, [{ row: 0, column: 0 }]).blocked, true)

  sheet.dataValidations = [rule('decimal', range, { operator: 'notBetween', formula1: '1', formula2: '5' })]
  setCell(sheet, 0, 0, 6.25)
  assert.equal(validateSheetDataValidationChanges(workbook, sheet.id, [{ row: 0, column: 0 }]).valid, true)
  setCell(sheet, 0, 0, 4)
  assert.equal(validateSheetDataValidationChanges(workbook, sheet.id, [{ row: 0, column: 0 }]).blocked, true)
})

test('text length, calendar date, time and checkbox rules use workbook-aware values', () => {
  const workbook = createBlankWorkbook()
  const sheet = workbook.sheets[0]
  workbook.date1904 = true
  const range = { startRow: 0, startCol: 0, endRow: 0, endCol: 0 }

  sheet.dataValidations = [rule('textLength', range, { operator: 'lessThanOrEqual', formula1: '3' })]
  setCell(sheet, 0, 0, 'abc')
  assert.equal(validateSheetDataValidationChanges(workbook, sheet.id, [{ row: 0, column: 0 }]).valid, true)
  setCell(sheet, 0, 0, 'four')
  assert.equal(validateSheetDataValidationChanges(workbook, sheet.id, [{ row: 0, column: 0 }]).blocked, true)

  sheet.dataValidations = [rule('date', range, { operator: 'greaterThan', formula1: '1' })]
  setCell(sheet, 0, 0, '1904-01-03')
  assert.equal(validateSheetDataValidationChanges(workbook, sheet.id, [{ row: 0, column: 0 }]).valid, true)
  setCell(sheet, 0, 0, '1904-01-01')
  assert.equal(validateSheetDataValidationChanges(workbook, sheet.id, [{ row: 0, column: 0 }]).blocked, true)

  sheet.dataValidations = [rule('time', range, { operator: 'lessThan', formula1: '0.5' })]
  setCell(sheet, 0, 0, '11:59:59')
  assert.equal(validateSheetDataValidationChanges(workbook, sheet.id, [{ row: 0, column: 0 }]).valid, true)
  setCell(sheet, 0, 0, '12:00:00')
  assert.equal(validateSheetDataValidationChanges(workbook, sheet.id, [{ row: 0, column: 0 }]).blocked, true)

  sheet.dataValidations = [rule('checkbox', range)]
  setCell(sheet, 0, 0, true)
  assert.equal(validateSheetDataValidationChanges(workbook, sheet.id, [{ row: 0, column: 0 }]).valid, true)
  setCell(sheet, 0, 0, 'yes')
  assert.equal(validateSheetDataValidationChanges(workbook, sheet.id, [{ row: 0, column: 0 }]).blocked, true)
})

test('list rules validate inline and range values, including each multi-select token', () => {
  const workbook = createBlankWorkbook()
  const sheet = workbook.sheets[0]
  setCell(sheet, 0, 0, 'Red')
  setCell(sheet, 1, 0, 'Green')
  setCell(sheet, 2, 0, 'Blue')

  const inline = rule('list', { startRow: 0, startCol: 1, endRow: 0, endCol: 1 }, { formula1: '"Red,Green,Blue"' })
  sheet.dataValidations = [inline]
  assert.deepEqual(sheetDataValidationOptions(inline, workbook, sheet.id), ['Red', 'Green', 'Blue'])
  setCell(sheet, 0, 1, 'Green')
  assert.equal(validateSheetDataValidationChanges(workbook, sheet.id, [{ row: 0, column: 1 }]).valid, true)
  setCell(sheet, 0, 1, 'Purple')
  assert.equal(validateSheetDataValidationChanges(workbook, sheet.id, [{ row: 0, column: 1 }]).blocked, true)

  const fromRange = rule('listMultiple', { startRow: 0, startCol: 2, endRow: 0, endCol: 2 }, { formula1: '=$A$1:$A$3' })
  sheet.dataValidations = [fromRange]
  assert.deepEqual(sheetDataValidationOptions(fromRange, workbook, sheet.id), ['Red', 'Green', 'Blue'])
  setCell(sheet, 0, 2, 'Red,Blue')
  assert.equal(validateSheetDataValidationChanges(workbook, sheet.id, [{ row: 0, column: 2 }]).valid, true)
  setCell(sheet, 0, 2, 'Red,Purple')
  assert.equal(validateSheetDataValidationChanges(workbook, sheet.id, [{ row: 0, column: 2 }]).blocked, true)
})

test('custom formula is translated from rule anchor for each cell', () => {
  const workbook = createBlankWorkbook()
  const sheet = workbook.sheets[0]
  setCell(sheet, 0, 0, 2)
  setCell(sheet, 1, 0, 3)
  setCell(sheet, 2, 0, 4)
  setCell(sheet, 0, 1, 'accepted')
  setCell(sheet, 1, 1, 'accepted')
  const range = { startRow: 0, startCol: 1, endRow: 2, endCol: 1 }
  sheet.dataValidations = [rule('custom', range, { formula1: '=MOD(A1,2)=0' })]

  assert.equal(validateSheetDataValidationChanges(workbook, sheet.id, [{ row: 0, column: 1 }]).valid, true)
  assert.equal(validateSheetDataValidationChanges(workbook, sheet.id, [{ row: 1, column: 1 }]).blocked, true)
  assert.equal(validateSheetDataValidationChanges(workbook, sheet.id, [{ row: 2, column: 1 }]).valid, true)
})

test('adding and clearing rules replace only intersecting range coverage', () => {
  const workbook = createBlankWorkbook()
  const sheet = workbook.sheets[0]
  const original = rule('whole', { startRow: 0, startCol: 0, endRow: 4, endCol: 4 }, { operator: 'greaterThan', formula1: '0' })
  applySheetDataValidationRule(sheet, original)
  const replacement = rule('decimal', { startRow: 1, startCol: 1, endRow: 3, endCol: 3 }, { operator: 'greaterThan', formula1: '1' })
  applySheetDataValidationRule(sheet, replacement)

  assert.deepEqual(dataValidationsAt(sheet, 2, 2).map(item => item.id), [replacement.id])
  assert.equal(dataValidationsAt(sheet, 0, 0)[0]?.id, original.id)
  assert.equal(sheet.dataValidations?.find(item => item.id === original.id)?.ranges.length, 4)

  clearSheetDataValidationRanges(sheet, [{ startRow: 2, startCol: 2, endRow: 2, endCol: 2 }])
  assert.equal(dataValidationsAt(sheet, 2, 2).length, 0)
  assert.equal(dataValidationsAt(sheet, 1, 2)[0]?.id, replacement.id)
})

test('warning and information validation failures are reported without blocking writes', () => {
  const workbook = createBlankWorkbook()
  const sheet = workbook.sheets[0]
  const range = { startRow: 0, startCol: 0, endRow: 0, endCol: 0 }
  sheet.dataValidations = [rule('whole', range, { operator: 'greaterThan', formula1: '10', errorStyle: 'warning', error: 'Must be greater than 10' })]
  setCell(sheet, 0, 0, 2)
  const result = validateSheetDataValidationChanges(workbook, sheet.id, [{ row: 0, column: 0 }])
  assert.equal(result.valid, false)
  assert.equal(result.blocked, false)
  assert.equal(result.issues[0]?.message, 'Must be greater than 10')
})

test('updating one range of a rule keeps that rule on its other cells', () => {
  const workbook = createBlankWorkbook()
  const sheet = workbook.sheets[0]
  const id = 'shared-rule'
  applySheetDataValidationRule(sheet, createSheetDataValidationRule({ id, type: 'whole', operator: 'greaterThan', formula1: '0' }, [
    { startRow: 0, startCol: 0, endRow: 4, endCol: 0 },
  ]))
  applySheetDataValidationRule(sheet, createSheetDataValidationRule({ id, type: 'whole', operator: 'greaterThan', formula1: '10' }, [
    { startRow: 2, startCol: 0, endRow: 2, endCol: 0 },
  ]))

  assert.equal(dataValidationsAt(sheet, 2, 0)[0]?.formula1, '10')
  assert.equal(dataValidationsAt(sheet, 0, 0)[0]?.formula1, '0')
  assert.equal(sheet.dataValidations?.length, 2)
})
