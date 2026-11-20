import assert from 'node:assert/strict'
import { test } from 'node:test'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { applyExcelWorkbookView, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { applySheetBorders, applySheetNumberFormat, adjustSheetDecimals, clearSheetCellFormat, nextSheetFontSize } from '../src/features/sheet-workbench/sheet-toolbar-format.ts'
import { changeSheetDecimalPlaces, formatSheetValue, SHEET_NUMBER_FORMATS } from '../src/features/sheet-workbench/sheet-number-format.ts'
import { createSheetRibbon } from '../src/components/spreadsheet/sheet-ribbon.ts'
import { ribbonInvocation } from '../src/platform/ribbon/ribbon-schema.ts'
import { createSpreadsheetScene, displayCell } from '../src/rendering/spreadsheet/scene.ts'
import { SpreadsheetGeometry } from '../src/rendering/spreadsheet/layout.ts'
import { buildSheetFunctionFormula, SHEET_FUNCTION_CATALOG } from '../src/features/sheet-workbench/sheet-function-catalog.ts'
import { recalculateWorkbookFormulas } from '../src/features/formula/formula-engine.ts'

test('numeric presentation uses raw types, fixed decimals, currency/accounting and preserved custom display', () => {
  assert.equal(displayCell({ value: '1234.5', rawValue: 1234.5, numberFormat: 'currency' }), '¥1,234.50')
  assert.equal(displayCell({ value: '1234.5', rawValue: '1234.5', numberFormat: 'currency' }), '1234.5')
  assert.equal(displayCell({ value: '', formula: '=1/4', rawValue: 0.25, cachedValue: '25%', numberFormat: 'percent' }), '25%')
  assert.equal(displayCell({ value: '-1234.5', rawValue: -1234.5, numberFormat: 'accounting' }), '¥ (1,234.50)')
  assert.equal(formatSheetValue(0, SHEET_NUMBER_FORMATS.accounting), '¥ —')
  assert.equal(displayCell({ value: '1234.5', rawValue: 1234.5, numberFormat: 'number' }), '1,234.50')
  assert.equal(displayCell({ value: '1234.5', rawValue: 1234.5, rawNumberFormat: '0.0" units"', displayValue: '1234.5 units' }), '1234.5 units')
  assert.equal(formatSheetValue(1234.5, '#,##0.00', false, 'de-DE'), '1.234,50')
})

test('date and time rendering respects workbook date systems, compatibility serials and fractions', () => {
  assert.equal(displayCell({ value: '1', rawValue: 1, numberFormat: 'date' }), '1900-01-01')
  assert.equal(displayCell({ value: '60', rawValue: 60, numberFormat: 'date' }), '1900-02-29')
  assert.equal(displayCell({ value: '0', rawValue: 0, numberFormat: 'date' }, true), '1904-01-01')
  assert.equal(displayCell({ value: '61.5', rawValue: 61.5, numberFormat: 'datetime' }), '1900-03-01 12:00:00')
  assert.equal(displayCell({ value: '0.5', rawValue: 0.5, numberFormat: 'time' }), '12:00:00')
  assert.equal(displayCell({ value: '61', rawValue: 61, rawNumberFormat: 'm/d/yyyy', displayValue: '3/1/1900' }), '3/1/1900')
})

test('decimal adjustment changes placeholders without corrupting accounting or percent sections', () => {
  assert.equal(changeSheetDecimalPlaces('General', 1), '0.0')
  assert.equal(changeSheetDecimalPlaces('0%', 1), '0.0%')
  assert.equal(changeSheetDecimalPlaces('0.000%', -1), '0.00%')
  assert.equal(changeSheetDecimalPlaces('0', -1), '0')
  assert.equal(changeSheetDecimalPlaces('yyyy-mm-dd', 1), 'yyyy-mm-dd')
  assert.equal(changeSheetDecimalPlaces('@', 1), '@')
  assert.equal(changeSheetDecimalPlaces('¥* #,##0.00;¥* (#,##0.00);¥* "-"', 1), '¥* #,##0.000;¥* (#,##0.000);¥* "-"')
  assert.equal(changeSheetDecimalPlaces('0.00".000"', -1), '0.0".000"')
  const adjusted = adjustSheetDecimals(applySheetNumberFormat({ value: '.25', rawValue: .25 }, 'percent'), 1)
  assert.equal(displayCell(adjusted), '25.0%')
})

test('range borders distinguish outside/inside/edges and preserve unrelated source fields', () => {
  const cells = { A1: { value: 'keep', rawValue: 'keep', color: '#445566' } }
  const ranges = [{ startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }]
  applySheetBorders(cells, ranges, 'outside')
  assert.deepEqual(Object.keys(cells.A1.border).sort(), ['left', 'top'])
  assert.deepEqual(Object.keys(cells.B2.border).sort(), ['bottom', 'right'])
  assert.equal(cells.A1.color, '#445566')
  assert.equal(cells.A1.rawValue, 'keep')
  applySheetBorders(cells, ranges, 'inside')
  assert.deepEqual(Object.keys(cells.A1.border).sort(), ['bottom', 'left', 'right', 'top'])
  applySheetBorders(cells, ranges, 'none')
  assert.deepEqual(cells.A1.border, {})
})

test('clear format and font stepping preserve values, formulas and typed caches', () => {
  const cell = { value: '', formula: '=1/4', cachedValue: '25%', rawValue: .25, bold: true, strike: true, vertical: 'top', border: { top: 'thin' }, rawNumberFormat: '0%', rotation: 45, indent: 2, shrinkToFit: true }
  const reset = clearSheetCellFormat(cell)
  assert.equal(reset.formula, '=1/4')
  assert.equal(reset.rawValue, .25)
  assert.equal(reset.strike, false)
  assert.equal(reset.rawNumberFormat, 'General')
  assert.equal(reset.rotation, 0)
  assert.equal(reset.shrinkToFit, false)
  assert.equal(displayCell(reset), '0.25')
  assert.equal(nextSheetFontSize(11, 1), 12)
  assert.equal(nextSheetFontSize(13, -1), 12)
  assert.equal(nextSheetFontSize(1, -1), 1)
})

test('view roundtrip persists exact number format/style and raw formula values without stale display', () => {
  const workbook = createBlankWorkbook()
  workbook.date1904 = true
  const cell = createCell(0, 0, 1234.5)
  workbook.sheets[0].cells['0:0'] = cell
  const view = toExcelWorkbookView(workbook)
  view.sheets[0].cells.A1 = { ...applySheetNumberFormat(view.sheets[0].cells.A1, 'currency'), strike: true, vertical: 'bottom' }
  let result = applyExcelWorkbookView(workbook, view)
  assert.equal(result.date1904, true)
  assert.equal(result.sheets[0].cells['0:0'].value, 1234.5)
  assert.equal(result.sheets[0].cells['0:0'].display, '¥1,234.50')
  assert.equal(result.sheets[0].cells['0:0'].style.numberFormat, SHEET_NUMBER_FORMATS.currency)
  assert.equal(result.sheets[0].cells['0:0'].style.font.strike, true)
  assert.equal(result.sheets[0].cells['0:0'].style.vertical, 'bottom')
  view.sheets[0].cells.A1 = adjustSheetDecimals(view.sheets[0].cells.A1, 1)
  result = applyExcelWorkbookView(result, view)
  assert.equal(result.sheets[0].cells['0:0'].display, '¥1,234.500')
  view.sheets[0].cells.A1 = clearSheetCellFormat(view.sheets[0].cells.A1)
  result = applyExcelWorkbookView(result, view)
  assert.equal(result.sheets[0].cells['0:0'].style.numberFormat, 'General')
  assert.equal(result.sheets[0].cells['0:0'].display, '1234.5')
  const formulaView = toExcelWorkbookView(result)
  formulaView.sheets[0].cells.A1 = { ...formulaView.sheets[0].cells.A1, formula: '=1/4', rawValue: .25, cachedValue: '25%', rawNumberFormat: '0%' }
  result = applyExcelWorkbookView(result, formulaView)
  assert.equal(result.sheets[0].cells['0:0'].value, .25)
  assert.equal(result.sheets[0].cells['0:0'].display, '25%')
})

test('scene draws vertical alignment, strike and real borders at current zoom', () => {
  const geometry = new SpreadsheetGeometry({ id: 's', name: 'Styles', cells: { A1: { value: 'Text', strike: true, vertical: 'top', border: { bottom: { style: 'medium', color: '#123456' } } } }, rowCount: 10, columnCount: 5, rowHeights: { 0: 60 } })
  const viewport = { width: 600, height: 300, scrollLeft: 0, scrollTop: 0, zoom: 1.5 }
  const scene = createSpreadsheetScene(geometry, viewport, { activeCell: 'E5', selectedCells: [] })
  const text = scene.commands.find(command => command.kind === 'text' && command.text === 'Text')
  const rect = geometry.cellRect(0, 0, viewport)
  assert.ok(text.y < rect.y + rect.height / 2)
  assert.ok(scene.commands.some(command => command.kind === 'line' && command.y1 === text.y && command.y2 === text.y))
  assert.ok(scene.commands.some(command => command.kind === 'line' && command.color === '#123456' && command.width === 3))
})

test('expanded Ribbon offers executable style/data intents and Pro insert aliases', () => {
  const state = { canUndo: true, canRedo: true, hasSheet: true, canMerge: true, canUnmerge: false, canDeleteRow: true, canDeleteColumn: true, chartAvailable: true, canSort: true,
    fontFamily: 'Aptos', fontSize: 11, numberFormat: 'currency', strike: true, vertical: 'bottom', gridlines: true, frozen: false, zoom: 100 }
  const tabs = createSheetRibbon(state)
  const items = tabs.flatMap(tab => tab.groups.flatMap(group => group.items.flatMap(item => [item, ...(item.children ?? [])])))
  assert.equal(new Set(items.map(item => item.id)).size, items.length)
  assert.equal(items.find(item => item.title === '删除线').activated, true)
  assert.equal(items.find(item => item.title === '底端对齐').activated, true)
  assert.ok(items.find(item => item.title === '数字格式').options.some(option => option.value === 'accounting'))
  assert.deepEqual(ribbonInvocation(items.find(item => item.title === '筛选')), { commandId: 'sheet.operation.open-data-panel', params: 'filter' })
  assert.deepEqual(ribbonInvocation(items.find(item => item.title === '自定义排序')), { commandId: 'sheet.operation.open-data-panel', params: 'sort' })
  assert.deepEqual(ribbonInvocation(items.find(item => item.title === '所有边框')), { commandId: 'sheet.command.set-borders', params: 'all' })
  assert.ok(tabs.find(tab => tab.id === 'insert').groups.flatMap(group => group.items).some(item => item.commandId === 'sheet.operation.open-slicer'))
})

test('the function picker lists only formulas that the real engine can evaluate', () => {
  for (const entry of SHEET_FUNCTION_CATALOG) {
    const workbook = createBlankWorkbook()
    workbook.sheets[0].cells['0:0'] = createCell(0, 0, 2)
    workbook.sheets[0].cells['0:1'] = createCell(0, 1, 4)
    const cell = createCell(0, 2, null)
    cell.formula = buildSheetFunctionFormula(entry.name, entry.example)
    workbook.sheets[0].cells['0:2'] = cell
    const result = recalculateWorkbookFormulas(workbook)
    assert.ok(!String(result.workbook.sheets[0].cells['0:2'].value).startsWith('#'), `${entry.name} should evaluate`)
  }
  assert.throws(() => buildSheetFunctionFormula('XLOOKUP', 'A1,B1:B2,C1:C2'), /当前函数目录/)
  assert.throws(() => buildSheetFunctionFormula('SUM', ''), /请输入/)
  assert.throws(() => buildSheetFunctionFormula('SUM', 'A1:A2)'), /语法无效/)
})
