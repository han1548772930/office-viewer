import assert from 'node:assert/strict'
import { test } from 'vitest'
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
    { value: '17', rawValue: 17, dataType: 'number', cachedDataType: null, displayValue: undefined, bold: true, numberFormat: 'percent' },
    { value: 'text', rawValue: 'text', dataType: 'text', cachedDataType: null, displayValue: undefined, italic: true },
    { value: '', rawValue: null, dataType: 'blank', cachedDataType: null, displayValue: undefined, underline: true },
  ])
})

test('formula-only paste shifts references, preserves destination formatting and skips source constants', () => {
  const plan = createSheetPastePlan({
    mode: 'formulas',
    sourceCells: [[sourceCell, { value: 'constant', rawValue: 23 }], [undefined, { value: '', formula: '=C3+$A4+B$2+$D$8', cachedValue: '9', rawValue: 9 }]],
    targetCells: [[{ value: 'keep', bold: true, background: '#fafafa' }, { value: 'leave me' }], [{ value: 'also leave me' }, { value: 'keep format', italic: true, cachedValue: 'old' }]],
    sourceStartRow: 0, sourceStartColumn: 0, targetStartRow: 2, targetStartColumn: 1,
  })

  assert.deepEqual(plan.cellPatches.map(({ rowOffset, columnOffset, cell }) => ({ rowOffset, columnOffset, cell })), [
    { rowOffset: 0, columnOffset: 0, cell: { value: '', formula: '=B3+$B$2', bold: true, background: '#fafafa', cachedValue: undefined, rawValue: null, dataType: 'formula', cachedDataType: null, displayValue: undefined } },
    { rowOffset: 1, columnOffset: 1, cell: { value: '', formula: '=D5+$A6+C$2+$D$8', italic: true, cachedValue: undefined, rawValue: null, dataType: 'formula', cachedDataType: null, displayValue: undefined } },
  ])
})

test('formula-only paste leaves destinations untouched when the source range has no formulas', () => {
  const plan = createSheetPastePlan({
    mode: 'formulas', sourceCells: [[{ value: 'constant' }]], targetCells: [[{ value: 'keep', bold: true }]],
    sourceStartRow: 0, sourceStartColumn: 0, targetStartRow: 1, targetStartColumn: 1,
  })
  assert.deepEqual(plan.cellPatches, [])
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

test('values paste preserves raw numeric caches, typed numeric text and errors instead of destination types', () => {
  const plan = createSheetPastePlan({
    mode: 'values', sourceCells: [[
      { value: '', formula: '=1/4', rawValue: .25, cachedValue: '25%', cachedDataType: 'number' },
      { value: '00123', dataType: 's' },
      { value: '', formula: '=1/0', cachedValue: '#DIV/0!', cachedDataType: 'e' },
    ]],
    targetCells: [[{ value: '#VALUE!', dataType: 'error' }, { value: '0', dataType: 'number' }, { value: 'old', dataType: 'text' }]],
    sourceStartRow: 0, sourceStartColumn: 0, targetStartRow: 2, targetStartColumn: 0,
  })
  assert.deepEqual(plan.cellPatches.map(({ cell }) => [cell.value, cell.rawValue, cell.dataType, cell.cachedDataType]), [
    ['0.25', .25, 'number', null], ['00123', '00123', 'text', null], ['#DIV/0!', '#DIV/0!', 'error', null],
  ])
})

test('format-only paste keeps destination content and cache types', () => {
  const target = { value: '', formula: '=FALSE()', rawValue: false, cachedValue: 'FALSE', dataType: 'formula', cachedDataType: 'b' }
  const plan = createSheetPastePlan({ mode: 'formats', sourceCells: [[{ value: 'source', dataType: 'text', bold: true }]], targetCells: [[target]],
    sourceStartRow: 0, sourceStartColumn: 0, targetStartRow: 0, targetStartColumn: 0 })
  assert.equal(plan.cellPatches[0].cell.dataType, 'formula')
  assert.equal(plan.cellPatches[0].cell.cachedDataType, 'b')
  assert.equal(plan.cellPatches[0].cell.rawValue, false)
})
