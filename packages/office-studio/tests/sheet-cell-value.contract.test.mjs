import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { normalizeSheetCellValueType, parseSheetCellText } from '../src/domain/sheet-cell-value.ts'
import { applyCellValue, cellFromApi } from '../src/api/model-utils.ts'
import { applyExcelWorkbookView, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { editSheetCellText, sheetCellInputText } from '../src/features/sheet-workbench/sheet-cell-edit.ts'
import { sheetCellNumberValue, presentSheetCellNumber } from '../src/features/sheet-workbench/sheet-number-presentation.ts'
import { sheetCellTextAlignment } from '../src/features/sheet-workbench/sheet-cell-alignment.ts'
import { calculateSelectionStatistics } from '../src/features/sheet-workbench/sheet-workbench.ts'
import { recalculateSheetViewFormulas } from '../src/features/sheet-workbench/sheet-view-formulas.ts'
import { createSheetFillPlan } from '../src/features/sheet-workbench/sheet-fill.ts'
import { createSpreadsheetScene } from '../src/rendering/spreadsheet/scene.ts'
import { SpreadsheetGeometry } from '../src/rendering/spreadsheet/layout.ts'
import { evaluateFormula, recalculateWorkbookFormulas } from '../src/features/formula/formula-engine.ts'

test('Office aliases and numeric API types retain text, boolean and error semantics', () => {
  for (const type of ['text', 'string', 's', 'str', 'inlineStr', 'force_string', 1, 4]) {
    assert.equal(normalizeSheetCellValueType(type), 'text')
    assert.equal(parseSheetCellText('00123', type), '00123')
  }
  for (const type of ['number', 'n', 'numeric', 'date', 2]) assert.equal(parseSheetCellText('1.25e2', type), 125)
  for (const type of ['boolean', 'bool', 'b', 3]) {
    assert.equal(parseSheetCellText('1', type), true)
    assert.equal(parseSheetCellText('0', type), false)
    assert.equal(parseSheetCellText('false', type), false)
  }
  assert.equal(parseSheetCellText('00123', 'e'), '00123')
  assert.equal(parseSheetCellText('1e999'), '1e999')
  assert.equal(parseSheetCellText('  '), '  ')
})

test('legacy text, typed raw values and unknown formula caches remain distinct', () => {
  const text = { value: '00123', dataType: 'str', rawNumberFormat: '0.00' }
  assert.equal(sheetCellNumberValue(text), '00123')
  assert.equal(presentSheetCellNumber(text, false, 'zh-CN', '0.0%').text, '00123')
  assert.equal(sheetCellNumberValue({ value: '00123', rawValue: '00123' }), '00123')
  assert.equal(sheetCellNumberValue({ value: '00123' }), 123)
  assert.equal(sheetCellNumberValue({ value: '00123', rawNumberFormat: '@' }), '00123')
  assert.equal(sheetCellNumberValue({ value: '', formula: '=TEXT(A1,"00000")', cachedValue: '00123', cachedDataType: 's' }), '00123')
  assert.equal(sheetCellNumberValue({ value: '', formula: '=FALSE()', cachedValue: '0', cachedDataType: 'b' }), false)
  const unknown = { value: '', formula: '=UNSUPPORTED()', rawValue: null, cachedValue: '25%', displayValue: '25%' }
  assert.equal(sheetCellNumberValue(unknown), undefined)
  assert.deepEqual(presentSheetCellNumber(unknown, false, 'zh-CN', '0.00'), { text: '25%' })
})

test('General alignment uses the current scalar or cached type and retains explicit alignment', () => {
  const cases = [
    [{ value: '00123', dataType: 'text' }, 'left'],
    [{ value: '12', dataType: 'n' }, 'right'],
    [{ value: 'TRUE', rawValue: true }, 'center'],
    [{ value: '0', dataType: 'b' }, 'center'],
    [{ value: '#DIV/0!', dataType: 'e' }, 'center'],
    [{ value: '', formula: '=1/0', dataType: 'formula', cachedDataType: 'error', rawValue: '#DIV/0!' }, 'center'],
    [{ value: '', formula: '="00123"', dataType: 'formula', cachedDataType: 'str', cachedValue: '00123' }, 'left'],
    [{ value: '', formula: '=2', rawValue: 2 }, 'right'],
    [{ value: '12', rawValue: 12, align: 'left' }, 'left'],
    [{ value: 'text', align: 'right' }, 'right'],
  ]
  for (const [cell, expected] of cases) assert.equal(sheetCellTextAlignment(cell), expected)
})

test('API imports and replacements preserve declared types and clear stale formula errors', () => {
  const text = cellFromApi(0, 0, { v: '00123', t: 1 })
  assert.equal(text.value, '00123')
  assert.equal(text.dataType, 'text')
  assert.equal(cellFromApi(0, 0, { v: '00123', t: 4 }).dataType, 'force_string')
  assert.equal(cellFromApi(0, 0, { v: '12', t: 2 }).value, 12)
  const cached = cellFromApi(0, 0, { f: '=FALSE()', v: '0', t: 3 })
  assert.equal(cached.value, false)
  assert.equal(cached.dataType, 'formula')
  assert.equal(cached.cachedDataType, 'boolean')
  const error = { ...cached, value: '#DIV/0!', cachedDataType: 'error', metadata: { formulaError: '#DIV/0!', retained: true } }
  const replacement = applyCellValue(error, 0, 0, 25)
  assert.equal(replacement.formula, null)
  assert.equal(replacement.value, 25)
  assert.equal(replacement.dataType, 'number')
  assert.equal(replacement.cachedDataType, null)
  assert.deepEqual(replacement.metadata, { retained: true })
  assert.equal(error.metadata.formulaError, '#DIV/0!')
  const unknown = applyCellValue(replacement, 0, 0, '=UNSUPPORTED()')
  assert.equal(unknown.value, null)
  assert.equal(unknown.cachedDataType, null)
  assert.equal(unknown.display, '=UNSUPPORTED()')
  for (const value of [Infinity, NaN, [], { v: {} }, { f: 5 }, { v: 1, t: 5 }]) {
    assert.throws(() => applyCellValue(replacement, 0, 0, value), TypeError)
  }
})

test('editor text format priority, apostrophes, empty values and cache replacement survive view roundtrips', () => {
  const workbook = createBlankWorkbook()
  workbook.sheets[0].cells['0:0'] = { ...createCell(0, 0, '#DIV/0!'), formula: '=1/0', dataType: 'formula', cachedDataType: 'error', metadata: { formulaError: '#DIV/0!' } }
  const view = toExcelWorkbookView(workbook)
  view.sheets[0].cells.A1 = editSheetCellText(view.sheets[0].cells.A1, '12')
  const result = applyExcelWorkbookView(workbook, view).sheets[0].cells['0:0']
  assert.equal(result.formula, null)
  assert.equal(result.value, 12)
  assert.equal(result.dataType, 'number')
  assert.equal(result.cachedDataType, null)
  assert.equal(result.metadata.formulaError, undefined)
  const forced = editSheetCellText(undefined, "'00123")
  assert.equal(forced.rawValue, '00123')
  assert.equal(sheetCellInputText(forced), "'00123")
  const emptyText = editSheetCellText(undefined, "'")
  assert.equal(emptyText.rawValue, '')
  assert.equal(editSheetCellText(forced, '').rawValue, null)
  for (const pattern of ['@', '@@@']) {
    const cell = editSheetCellText({ value: '', rawNumberFormat: pattern }, '=1+1')
    assert.equal(cell.formula, undefined)
    assert.equal(cell.rawValue, '=1+1')
    assert.equal(presentSheetCellNumber(cell).text, '=1+1')
    assert.equal(editSheetCellText({ value: '', rawNumberFormat: pattern }, "'00123").rawValue, "'00123")
  }
  const noCache = applyExcelWorkbookView(workbook, { ...view, sheets: [{ ...view.sheets[0], cells: { A1: { value: '', formula: '=UNSUPPORTED()', rawValue: null, displayValue: 'source cache' } } }] })
  assert.equal(noCache.sheets[0].cells['0:0'].value, null)
  assert.equal(noCache.sheets[0].cells['0:0'].display, 'source cache')
})

test('formula recalculation projects boolean and error cache types to the standalone view', () => {
  const workbook = { name: 'Types', sheets: [{ id: 's', name: 'Types', rowCount: 20, columnCount: 5, cells: {
    A1: { value: '', formula: '=FALSE()', rawValue: null, dataType: 'formula', cachedDataType: null },
    B1: { value: '', formula: '=1/0', rawValue: null, dataType: 'formula', cachedDataType: null },
  } }] }
  const result = recalculateSheetViewFormulas(workbook).workbook.sheets[0].cells
  assert.equal(result.A1.rawValue, false)
  assert.equal(result.A1.cachedDataType, 'boolean')
  assert.equal(result.B1.cachedDataType, 'error')
  assert.equal(sheetCellTextAlignment(result.B1), 'center')
})

test('an unchanged legacy cell does not gain inferred types just from a view roundtrip', () => {
  const workbook = createBlankWorkbook()
  const original = createCell(0, 0, 12)
  workbook.sheets[0].cells['0:0'] = original
  const result = applyExcelWorkbookView(workbook, toExcelWorkbookView(workbook)).sheets[0].cells['0:0']
  assert.equal(result.dataType, original.dataType)
  assert.equal(result.cachedDataType, original.cachedDataType)
  assert.equal(result.value, original.value)
})

test('selection statistics and numeric series exclude explicitly typed numeric text', () => {
  const sheet = { id: 's', name: 'Types', rowCount: 20, columnCount: 5, cells: {
    A1: { value: '00123', dataType: 'text' }, B1: { value: '1', dataType: 'b' },
    C1: { value: '12', dataType: 'n' }, D1: { value: '3', rawValue: 3 },
  } }
  const statistics = calculateSelectionStatistics(sheet, [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 3 }])
  assert.equal(statistics.sum, 15)
  assert.equal(statistics.count, 2)
  assert.equal(statistics.countA, 4)
  const plan = createSheetFillPlan({ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, { row: 1, column: 0 }, [[sheet.cells.A1]])
  assert.equal(plan.cellPatches[0].cell.value, '00123')
  assert.equal(plan.cellPatches[0].cell.dataType, 'text')
})

test('Canvas draws declared text, numeric values and error caches with consistent alignment at each zoom', () => {
  const geometry = new SpreadsheetGeometry({ id: 's', name: 'Types', rowCount: 20, columnCount: 5, cells: {
    A1: { value: '00123', dataType: 's' }, B1: { value: '12', rawValue: 12 },
    C1: { value: '#DIV/0!', dataType: 'e' },
  } })
  for (const zoom of [0.5, 1, 1.5, 2]) {
    const viewport = { width: 900, height: 500, scrollLeft: 0, scrollTop: 0, zoom }
    const scene = createSpreadsheetScene(geometry, viewport, { activeCell: 'E5', selectedCells: [] })
    for (const [column, text, align] of [[0, '00123', 'left'], [1, '12', 'right'], [2, '#DIV/0!', 'center']]) {
      const drawn = scene.commands.find(command => command.kind === 'text' && command.text === text)
      const rect = geometry.cellRect(0, column, viewport)
      assert.equal(drawn.align, align)
      assert.equal(drawn.x, align === 'center' ? rect.x + rect.width / 2 : align === 'right' ? rect.x + rect.width - 6 * zoom : rect.x + 6 * zoom)
    }
  }
})

test('quoted error labels and declared text remain text in references, predicates and aggregates', () => {
  const workbook = createBlankWorkbook()
  const sheet = workbook.sheets[0]
  const evaluate = expression => evaluateFormula(expression, workbook, sheet.id)
  for (const type of ['text', 's', 'str', 'string', 'inlineStr', 'force_string']) {
    sheet.cells['0:0'] = { ...createCell(0, 0, '#DIV/0!'), dataType: type,
      cachedDataType: 'error', metadata: { formulaError: '#DIV/0!' } }
    assert.equal(evaluate('=A1'), '#DIV/0!', type)
    assert.equal(evaluate('=ISTEXT(A1)'), true, type)
    assert.equal(evaluate('=ISERROR(A1)'), false, type)
    assert.equal(evaluate('=SUM(A1:A1)'), 0, type)
    assert.equal(evaluate('=IFERROR(A1,"wrong")'), '#DIV/0!', type)
  }
  sheet.cells['0:0'] = { ...createCell(0, 0, 7), cachedDataType: 'error' }
  assert.equal(evaluate('=A1'), 7)
  assert.equal(evaluate('=SUM(A1:A1)'), 7)
  for (const code of ['#DIV/0!', '#N/A', '#NAME?', '#VALUE!']) {
    assert.equal(evaluate(`="${code}"`), code)
    assert.equal(evaluate(`=ISTEXT("${code}")`), true)
    assert.equal(evaluate(`=ISERROR("${code}")`), false)
    assert.equal(evaluate(`=IFERROR("${code}","wrong")`), code)
    assert.deepEqual(evaluate(`=SUM("${code}")`), { kind: 'error', code: '#VALUE!' })
  }
})

test('declared and legacy errors still propagate while stale scalar cache tags are ignored', () => {
  const workbook = createBlankWorkbook()
  const sheet = workbook.sheets[0]
  for (const type of ['error', 'e', undefined]) {
    sheet.cells['0:0'] = { ...createCell(0, 0, '#DIV/0!'), dataType: type, cachedDataType: 'text' }
    for (const expression of ['=A1', '=SUM(A1)', '=SUM(A1:A1)']) {
      assert.deepEqual(evaluateFormula(expression, workbook, sheet.id), { kind: 'error', code: '#DIV/0!' })
    }
    assert.equal(evaluateFormula('=ISTEXT(A1)', workbook, sheet.id), false)
    assert.equal(evaluateFormula('=ISERROR(A1)', workbook, sheet.id), true)
    assert.equal(evaluateFormula('=IFERROR(A1,"caught")', workbook, sheet.id), 'caught')
  }
  assert.deepEqual(evaluateFormula('=#DIV/0!', workbook, sheet.id), { kind: 'error', code: '#DIV/0!' })
})

test('unsupported imported caches preserve text, errors, empty strings, zero and false during dependency recalculation', () => {
  const workbook = createBlankWorkbook()
  const sheet = workbook.sheets[0]
  const caches = [['#DIV/0!', 'e'], ['#DIV/0!', 'str'], ['', 'text'], [0, 'number'], [false, 'boolean'], ['#NAME?', 'error']]
  for (const [row, [value, cachedDataType]] of caches.entries()) {
    sheet.cells[`${row}:0`] = { ...createCell(row, 0, value), formula: '=UNSUPPORTED()', dataType: 'formula', cachedDataType }
    if (row === 3 || row === 4) sheet.cells[`${row}:0`].display = ''
  }
  for (const [row, expression] of ['=A1', '=A2', '=A3', '=SUM(A2:A5)', '=ISTEXT(A2)', '=ISERROR(A1)'].entries()) {
    sheet.cells[`${row}:1`] = { ...createCell(row, 1, null), formula: expression, dataType: 'formula', cachedDataType: null }
  }
  sheet.cells['6:0'] = { ...createCell(6, 0, null), formula: '=UNSUPPORTED()', dataType: 'formula', cachedDataType: null, display: 'source-only display' }
  const before = JSON.parse(JSON.stringify(workbook))
  const result = recalculateWorkbookFormulas(workbook, before)
  const cells = result.workbook.sheets[0].cells
  assert.deepEqual(caches.map((_cache, row) => cells[`${row}:0`].value), caches.map(([value]) => value))
  assert.deepEqual(caches.map((_cache, row) => cells[`${row}:0`].cachedDataType), ['error', 'string', 'string', 'number', 'boolean', 'error'])
  assert.deepEqual(caches.map((_cache, row) => cells[`${row}:1`].value), ['#DIV/0!', '#DIV/0!', '', 0, true, true])
  assert.equal(cells['0:1'].cachedDataType, 'error')
  assert.equal(cells['1:1'].cachedDataType, 'string')
  assert.equal(cells['1:1'].metadata?.formulaError, undefined)
  assert.equal(cells['6:0'].cachedDataType, 'error', 'a source display without a typed cache is not a computed value')
  assert.equal(result.diagnostics.filter(item => item.code === 'FORMULA_UNSUPPORTED_CACHE_PRESERVED').length, caches.length)
  assert.deepEqual(JSON.parse(JSON.stringify(workbook)), before, 'recalculation does not mutate its input model')
})

test('view edits preserve unchanged regional displays and format changed values with the supplied region', () => {
  const workbook = createBlankWorkbook()
  const cell = createCell(0, 0, 1234.5)
  cell.style.numberFormat = '#,##0.00'; cell.display = '1.234,50'
  workbook.sheets[0].cells['0:0'] = cell
  const cached = { ...createCell(0, 1, 12.345), formula: '=A1/100', dataType: 'formula', cachedDataType: 'number', display: '12,35' }
  cached.style.numberFormat = '0.00'
  workbook.sheets[0].cells['0:1'] = cached
  const view = toExcelWorkbookView(workbook)
  view.sheets[0].cells.A1.bold = true
  const styled = applyExcelWorkbookView(workbook, view)
  assert.equal(styled.sheets[0].cells['0:0'].display, '1.234,50')
  assert.equal(styled.sheets[0].cells['0:1'].display, '12,35')
  view.sheets[0].cells.A1.value = '25.5'; view.sheets[0].cells.A1.rawValue = 25.5
  const edited = applyExcelWorkbookView(workbook, view, 'de-DE')
  assert.equal(edited.sheets[0].cells['0:0'].value, 25.5)
  assert.equal(edited.sheets[0].cells['0:0'].display, '25,50')
  assert.equal(edited.sheets[0].cells['0:1'].display, '12,35')
  const formatView = toExcelWorkbookView(workbook)
  formatView.sheets[0].cells.A1.rawNumberFormat = '0.0'
  assert.equal(applyExcelWorkbookView(workbook, formatView, 'de-DE').sheets[0].cells['0:0'].display, '1234,5')
  const dateCell = createCell(1, 0, 1)
  dateCell.style.numberFormat = 'yyyy-mm-dd'; dateCell.display = '1900-01-01'
  workbook.sheets[0].cells['1:0'] = dateCell
  const dateView = toExcelWorkbookView(workbook)
  dateView.date1904 = true
  assert.equal(applyExcelWorkbookView(workbook, dateView, 'de-DE').sheets[0].cells['1:0'].display, '1904-01-02')
})

test('standalone formula recalculation uses the host region for numeric result displays', () => {
  const view = { name: 'Regional formulas', sheets: [{ id: 'regional', name: 'Regional formulas', rowCount: 20, columnCount: 5, cells: {
    A1: { value: '1001.25', rawValue: 1001.25, dataType: 'number' },
    B1: { value: '', formula: '=A1/100', rawValue: null, dataType: 'formula', cachedDataType: null, rawNumberFormat: '0.00' },
  } }] }
  const result = recalculateSheetViewFormulas(view, undefined, { numberFormatLocale: 'de-DE' }).workbook.sheets[0].cells
  assert.equal(result.B1.rawValue, 10.0125)
  assert.equal(result.B1.cachedValue, '10,01')
  assert.equal(result.B1.displayValue, '10,01')
  assert.deepEqual(result.A1, view.sheets[0].cells.A1)
})
