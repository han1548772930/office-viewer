import assert from 'node:assert/strict'
import { test } from 'vitest'
import { moveSheetCellRange, structureEdit, transformWorkbookReferencesForSheetEdit } from '../src/features/sheet-workbench/sheet-cell-range-move.ts'

function sheet(overrides = {}) {
  return {
    id: 'data', name: 'Data', cells: {}, rowCount: 4, columnCount: 4,
    ...overrides,
  }
}

function range(startRow, startColumn, endRow, endColumn) { return { startRow, startColumn, endRow, endColumn } }

test('insert-right shifts selected and trailing cells while retaining source formatting in the inserted cells', () => {
  const source = sheet({ cells: {
    B1: { value: 'selected', background: '#ffeeaa' }, C1: { value: 'selected-2' }, D1: { value: 'trailing' },
  } })
  const result = moveSheetCellRange(source, range(0, 1, 0, 2), 'right')
  assert.equal(result.status, 'applied')
  assert.equal(result.sheet.cells.B1.value, '')
  assert.equal(result.sheet.cells.B1.background, '#ffeeaa')
  assert.equal(result.sheet.cells.C1, undefined)
  assert.equal(result.sheet.cells.D1.value, 'selected')
  assert.equal(result.sheet.cells.E1.value, 'selected-2')
  assert.equal(result.sheet.cells.F1.value, 'trailing')
  assert.equal(result.sheet.columnCount, 6)
})

test('insert-down expands rows and moves only cells in the selected columns', () => {
  const source = sheet({ cells: { B1: { value: 'selected' }, C1: { value: 'other-column' }, B2: { value: 'trailing' } } })
  const result = moveSheetCellRange(source, range(0, 1, 0, 1), 'down')
  assert.equal(result.status, 'applied')
  assert.equal(result.sheet.cells.B1.value, '')
  assert.equal(result.sheet.cells.B2.value, 'selected')
  assert.equal(result.sheet.cells.B3.value, 'trailing')
  assert.equal(result.sheet.cells.C1.value, 'other-column')
  assert.equal(result.sheet.rowCount, 5)
})

test('delete-left and delete-up remove the selected cells and shift trailing cells into the selection', () => {
  const horizontal = moveSheetCellRange(sheet({ cells: {
    A1: { value: 'left' }, B1: { value: 'delete-1' }, C1: { value: 'delete-2' }, D1: { value: 'right' },
  } }), range(0, 1, 0, 2), 'left')
  assert.equal(horizontal.status, 'applied')
  assert.equal(horizontal.sheet.cells.A1.value, 'left')
  assert.equal(horizontal.sheet.cells.B1.value, 'right')
  assert.equal(horizontal.sheet.cells.C1, undefined)

  const vertical = moveSheetCellRange(sheet({ cells: {
    B1: { value: 'top' }, B2: { value: 'delete-1' }, B3: { value: 'delete-2' }, B4: { value: 'bottom' },
  } }), range(1, 1, 2, 1), 'up')
  assert.equal(vertical.status, 'applied')
  assert.equal(vertical.sheet.cells.B1.value, 'top')
  assert.equal(vertical.sheet.cells.B2.value, 'bottom')
  assert.equal(vertical.sheet.cells.B3, undefined)
})

test('partial merged ranges require confirmation and are removed only after confirmation', () => {
  const source = sheet({ merges: ['A1:C1', 'D2:E2'], cells: { A1: { value: 'merged' }, B1: { value: 'inside' } } })
  const selection = range(0, 1, 0, 1)
  const prompt = moveSheetCellRange(source, selection, 'right')
  assert.deepEqual(prompt, { status: 'confirm-merge', mergeRanges: ['A1:C1'] })
  const applied = moveSheetCellRange(source, selection, 'right', true)
  assert.equal(applied.status, 'applied')
  assert.deepEqual(applied.sheet.merges, ['D2:E2'])
  assert.deepEqual(source.merges, ['A1:C1', 'D2:E2'])
})

test('insert at sheet limits is rejected and formula and defined names follow the structural edit', () => {
  const tooWide = sheet({ columnCount: 16_384 })
  assert.equal(moveSheetCellRange(tooWide, range(0, 16_383, 0, 16_383), 'right').status, 'invalid')

  const workbook = {
    sheets: [sheet({ cells: { A2: { value: 'formula', formula: '=B1+Other!B1' } } }), sheet({ id: 'other', name: 'Other', cells: {} })],
    metadata: {
      definedNames: [{ name: 'Total', refersTo: 'Data!$B$1' }],
      localDefinedNames: [{ sheet: 'Data', name: 'Local', refersTo: 'B1' }],
    },
  }
  const edit = structureEdit(range(0, 1, 0, 1), 'right')
  const transformed = transformWorkbookReferencesForSheetEdit(workbook, 'data', edit)
  assert.equal(transformed.sheets[0].cells.A2.formula, '=C1+Other!B1')
  assert.equal(transformed.metadata.definedNames[0].refersTo, 'Data!$C$1')
  assert.equal(transformed.metadata.localDefinedNames[0].refersTo, 'C1')
  assert.equal(workbook.sheets[0].cells.A2.formula, '=B1+Other!B1')
})
