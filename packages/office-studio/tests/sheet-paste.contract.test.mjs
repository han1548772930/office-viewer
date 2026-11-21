import assert from 'node:assert/strict'
import { test } from 'node:test'
import { createSheetPastePlan } from '../src/features/sheet-workbench/sheet-paste.ts'

const sourceCell = {
  value: '', formula: '=A1+$B$2', cachedValue: '17', rawValue: 17, displayValue: '17',
  bold: true, numberFormat: 'currency', rawNumberFormat: '$#,##0.00',
  border: { top: { style: 'thin', color: '#111111' } },
}

test('normal paste translates relative formulas and clones source data', () => {
  const plan = createSheetPastePlan({
    mode: 'all', sourceCells: [[sourceCell]], targetCells: [[{ value: 'old' }]],
    sourceStartRow: 0, sourceStartColumn: 0, targetStartRow: 1, targetStartColumn: 1,
  })
  const pasted = plan.cellPatches[0].cell
  assert.equal(pasted.formula, '=B2+$B$2')
  assert.equal(pasted.value, '')
  assert.equal(pasted.cachedValue, undefined)
  assert.equal(pasted.bold, true)
  assert.notEqual(pasted.border, sourceCell.border)
  assert.notEqual(pasted.border.top, sourceCell.border.top)
  assert.deepEqual(plan.columnWidthPatches, [])
})

test('values paste calculated/raw values while retaining destination formatting', () => {
  const plan = createSheetPastePlan({
    mode: 'values', sourceCells: [[sourceCell, { value: 'text', rawValue: 'text' }, undefined]],
    targetCells: [[{ value: '', formula: '=1+1', bold: true, numberFormat: 'percent' }, { value: 'old', italic: true }, { value: 'old', underline: true }]],
    sourceStartRow: 0, sourceStartColumn: 0, targetStartRow: 0, targetStartColumn: 3,
  })
  assert.deepEqual(plan.cellPatches.map(patch => patch.cell), [
    { value: '17', rawValue: 17, displayValue: undefined, bold: true, numberFormat: 'percent' },
    { value: 'text', rawValue: 'text', displayValue: undefined, italic: true },
    { value: '', rawValue: undefined, displayValue: undefined, underline: true },
  ])
})

test('formats paste resets absent source formats but preserves content and formula cache', () => {
  const plan = createSheetPastePlan({
    mode: 'formats', sourceCells: [[{ value: 'format', bold: true }]],
    targetCells: [[{ value: '', formula: '=A1', cachedValue: '5', rawValue: 5, displayValue: '5', italic: true }]],
    sourceStartRow: 0, sourceStartColumn: 0, targetStartRow: 0, targetStartColumn: 0,
  })
  const pasted = plan.cellPatches[0].cell
  assert.equal(pasted.bold, true)
  assert.equal(pasted.italic, false)
  assert.equal(pasted.numberFormat, 'general')
  assert.equal(pasted.formula, '=A1')
  assert.equal(pasted.cachedValue, '5')
  assert.equal(pasted.rawValue, 5)
})

test('paste except borders uses source content and formatting with target borders', () => {
  const targetBorder = { left: { style: 'double', color: '#222222' } }
  const plan = createSheetPastePlan({
    mode: 'exceptBorders', sourceCells: [[sourceCell]], targetCells: [[{ value: 'old', border: targetBorder }]],
    sourceStartRow: 0, sourceStartColumn: 0, targetStartRow: 0, targetStartColumn: 1,
  })
  const pasted = plan.cellPatches[0].cell
  assert.equal(pasted.formula, '=B1+$B$2')
  assert.deepEqual(pasted.border, targetBorder)
  assert.notEqual(pasted.border, targetBorder)
})

test('column-width paste maps source absolute columns to destination columns', () => {
  const plan = createSheetPastePlan({
    mode: 'columnWidths', sourceCells: [[{}, {}, {}]], targetCells: [],
    sourceStartRow: 0, sourceStartColumn: 2, targetStartRow: 0, targetStartColumn: 5,
    sourceColumnWidths: { 2: 144, 4: 72 }, sourceDefaultColumnWidth: 96,
  })
  assert.deepEqual(plan.cellPatches, [])
  assert.deepEqual(plan.columnWidthPatches, [
    { column: 5, width: 144 }, { column: 6, width: 96 }, { column: 7, width: 72 },
  ])
})
