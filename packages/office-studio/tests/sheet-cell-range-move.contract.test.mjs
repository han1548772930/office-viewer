import assert from 'node:assert/strict'
import { test } from 'vitest'
import { moveSheetCellRange, structureEdit, transformWorkbookReferencesForSheetEdit } from '../src/features/sheet-workbench/sheet-cell-range-move.ts'
import { transformFormulaReferencesForStructureEdit, translateFormulaReferences } from '../src/features/formula/formula-references.ts'

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

test('insert-down keeps spare canvas capacity and moves only cells in the selected columns', () => {
  const source = sheet({ cells: { B1: { value: 'selected' }, C1: { value: 'other-column' }, B2: { value: 'trailing' } } })
  const result = moveSheetCellRange(source, range(0, 1, 0, 1), 'down')
  assert.equal(result.status, 'applied')
  assert.equal(result.sheet.cells.B1, undefined)
  assert.equal(result.sheet.cells.B2.value, 'selected')
  assert.equal(result.sheet.cells.B3.value, 'trailing')
  assert.equal(result.sheet.cells.C1.value, 'other-column')
  assert.equal(result.sheet.rowCount, 4)
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

test('real content overflow at sheet limits is rejected and formula and defined names follow the structural edit', () => {
  const tooWide = sheet({ columnCount: 16_384, cells: { XFD1: { value: 'edge' } } })
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

test('references crossing the perpendicular move band retain the upstream rectangular range semantics', () => {
  const right = structureEdit(range(1, 2, 2, 3), 'right')
  assert.equal(transformFormulaReferencesForStructureEdit('=SUM(B1:E4)', right, 'Data', 'Data'), '=SUM(B1:E4)')
  assert.equal(transformFormulaReferencesForStructureEdit('=SUM(B2:E3)', right, 'Data', 'Data'), '=SUM(B2:G3)')
  assert.equal(transformFormulaReferencesForStructureEdit('=SUM(C2:E3)', right, 'Data', 'Data'), '=SUM(E2:G3)')
  const left = structureEdit(range(1, 2, 2, 3), 'left')
  assert.equal(transformFormulaReferencesForStructureEdit('=SUM(B1:F4)', left, 'Data', 'Data'), '=SUM(B1:F4)')
  assert.equal(transformFormulaReferencesForStructureEdit('=SUM(D2:F3)', left, 'Data', 'Data'), '=SUM(C2:D3)')
  assert.equal(transformFormulaReferencesForStructureEdit('=SUM(C2:D3)', left, 'Data', 'Data'), '=SUM(#REF!)')
  const up = structureEdit(range(1, 1, 2, 2), 'up')
  assert.equal(transformFormulaReferencesForStructureEdit('=SUM(A1:D4)', up, 'Data', 'Data'), '=SUM(A1:D4)')
  assert.equal(transformFormulaReferencesForStructureEdit('=SUM(B1:C4)', up, 'Data', 'Data'), '=SUM(B1:C2)')
})

test('qualified ranges and cross-sheet local names preserve their sheet context', () => {
  const edit = structureEdit(range(0, 1, 0, 1), 'right')
  const source = { sheets: [sheet(), sheet({ id: 'other', name: 'Other' })], metadata: { localDefinedNames: [
    { sheet: 'Other', name: 'Local', refersTo: 'B1' },
    { sheet: 'Other', name: 'Source', refersTo: 'Data!B1' },
  ] } }
  const transformed = transformWorkbookReferencesForSheetEdit(source, 'data', edit)
  assert.equal(transformed.metadata.localDefinedNames[0].refersTo, 'B1')
  assert.equal(transformed.metadata.localDefinedNames[1].refersTo, 'Data!C1')
  assert.equal(transformFormulaReferencesForStructureEdit('=SUM(Other!A1:B1)', edit, 'Data', 'Data'), '=SUM(Other!A1:B1)')
  assert.equal(transformFormulaReferencesForStructureEdit('=SUM(Data!A1:Data!B1)', edit, 'Other', 'Data'), '=SUM(Data!A1:Data!C1)')
  assert.equal(transformFormulaReferencesForStructureEdit('=数据!B1+Other!B1+B1', edit, '数据', '数据'), '=数据!C1+Other!B1+C1')
  assert.equal(translateFormulaReferences('=数据!A1+名称A1+Data!A1', 1, 1), '=数据!B2+名称A1+Data!B2')
})

test('references moved beyond the worksheet bounds become reference errors', () => {
  const edit = structureEdit(range(0, 0, 0, 0), 'right')
  assert.equal(transformFormulaReferencesForStructureEdit('=XFD1', edit, 'Data', 'Data'), '=#REF!')
  assert.equal(transformFormulaReferencesForStructureEdit('=SUM(XFC1:XFD1)', edit, 'Data', 'Data'), '=SUM(#REF!)')
})

test('protected ranges anywhere in the trailing move band reject each direction before changing data', () => {
  const protectedSheet = sheet({ rangeProtectionRules: [{ ranges: [{ startRow: 0, endRow: 0, startCol: 3, endCol: 3 }] }], cells: { B1: { value: 'keep' } } })
  for (const direction of ['right', 'left']) assert.equal(moveSheetCellRange(protectedSheet, range(0, 1, 0, 1), direction).status, 'invalid')
  const vertical = sheet({ rangeProtectionRules: [{ ranges: [{ startRow: 3, endRow: 3, startCol: 1, endCol: 1 }] }] })
  for (const direction of ['down', 'up']) assert.equal(moveSheetCellRange(vertical, range(0, 1, 0, 1), direction).status, 'invalid')
  const beyondCanvas = sheet({ rangeProtectionRules: [{ ranges: [{ startRow: 0, endRow: 0, startCol: 100, endCol: 100 }] }] })
  assert.equal(moveSheetCellRange(beyondCanvas, range(0, 1, 0, 1), 'right').status, 'invalid')
  assert.equal(moveSheetCellRange(protectedSheet, range(1, 1, 1, 1), 'right').status, 'applied')
  assert.equal(protectedSheet.cells.B1.value, 'keep')
})

test('inserted blanks retain formatting while moving cell-bound sparklines only with their source data', () => {
  const source = sheet({ cells: { B1: { value: '12', rawValue: 12, formula: '=A1', cachedValue: '12', displayValue: '12', bold: true, sparkline: { type: 'line', sourceRange: 'A1:A4' } } } })
  const result = moveSheetCellRange(source, range(0, 1, 0, 1), 'down')
  assert.equal(result.status, 'applied')
  assert.equal(result.sheet.cells.B1.value, '')
  assert.equal(result.sheet.cells.B1.rawValue, null)
  assert.equal(result.sheet.cells.B1.formula, undefined)
  assert.equal(result.sheet.cells.B1.cachedValue, undefined)
  assert.equal(result.sheet.cells.B1.sparkline, undefined)
  assert.equal(result.sheet.cells.B1.bold, true)
  assert.notEqual(result.sheet.cells.B2, source.cells.B1)
  assert.equal(result.sheet.cells.B2.formula, '=A1')
  assert.equal(result.sheet.cells.B2.rawValue, 12)
  assert.deepEqual(result.sheet.cells.B2.sparkline, source.cells.B1.sparkline)
})
