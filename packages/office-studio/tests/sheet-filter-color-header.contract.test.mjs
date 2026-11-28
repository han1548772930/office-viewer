import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankSheet, createCell } from '../src/domain/model.ts'
import { toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { SpreadsheetGeometry } from '../src/rendering/spreadsheet/layout.ts'
import { canonicalSheetFilterColor, createSheetFilterColorReader } from '../src/features/sheet-sort-filter/sheet-filter-colors.ts'
import { sheetFilterHeaderControls } from '../src/features/sheet-sort-filter/sheet-filter-header.ts'
import { getFilteredRows, setAutoFilter, setFilterCriteria } from '../src/features/sheet-sort-filter/sort-filter-engine.ts'
import { readDataPanelColors, readDataPanelValues } from '../src/features/sheet-sort-filter/sheet-data-panel-state.ts'
import { isSheetQuickSumShortcut } from '../src/components/spreadsheet/sheet-function-shortcuts.ts'
import { sheetFilterPermissionRanges } from '../src/features/sheet-sort-filter/sheet-filter-column.ts'
import { canEditProtectedRanges } from '../src/features/sheet-protection/range-protection.ts'

// Implementation-phase drafts. No runtime verification has been performed.
const range = { startRow: 0, endRow: 4, startCol: 0, endCol: 1 }
function setup() {
  const sheet = createBlankSheet('Colors'); sheet.rowCount = 30; sheet.columnCount = 10
  for (let row = 0; row <= 4; row++) {
    sheet.cells[`${row}:0`] = createCell(row, 0, row === 0 ? 'Color' : row)
    sheet.cells[`${row}:1`] = createCell(row, 1, row % 2 === 0 ? 'keep' : 'drop')
  }
  const workbook = { kind: 'workbook', id: 'colors', name: 'Colors', sheets: [sheet], activeSheetId: sheet.id, date1904: false, modified: false }
  return { sheet, workbook }
}

test('color equality canonicalizes RGB/short hex and derives inherited, explicit and missing-cell colors', () => {
  const { sheet, workbook } = setup()
  for (const cell of Object.values(sheet.cells)) cell.authoredStyle = {}
  workbook.defaultStyle = { fill: '#f00' }
  sheet.rowStyles = { 2: { fill: '#00ff00' } }
  sheet.cells['3:0'].authoredStyle.fill = '#0000ff'
  const read = createSheetFilterColorReader(sheet, workbook)
  assert.deepEqual(read(1, 0), { fill: '#ff0000', text: '#000000' })
  assert.equal(read(2, 0).fill, '#00ff00')
  assert.equal(read(3, 0).fill, '#0000ff')
  assert.equal(read(8, 0).fill, '#ff0000')
  assert.equal(canonicalSheetFilterColor('rgb(255, 0, 0)'), '#ff0000')
  setAutoFilter(sheet, range, workbook)
  setFilterCriteria(sheet, 0, { kind: 'color', cellFillColors: ['rgb(255,0,0)'] }, workbook)
  assert.deepEqual(getFilteredRows(sheet, workbook), [2, 3])
  assert.equal(sheet.cells['3:0'].authoredStyle.fill, '#0000ff')
})

test('fill criteria take precedence over text criteria, with null fill and empty API arrays retained', () => {
  const { sheet, workbook } = setup()
  sheet.cells['1:0'].style.fill = '#ff0000'
  sheet.cells['2:0'].style.font.color = '#ff0000'
  setAutoFilter(sheet, range, workbook)
  setFilterCriteria(sheet, 0, { kind: 'color', cellFillColors: [null], cellTextColors: ['#ff0000'] }, workbook)
  assert.deepEqual(getFilteredRows(sheet, workbook), [1])
  setFilterCriteria(sheet, 0, { kind: 'color', cellTextColors: ['rgb(0,0,0)'] }, workbook)
  assert.deepEqual(getFilteredRows(sheet, workbook), [2])
  setFilterCriteria(sheet, 0, { kind: 'color', cellFillColors: [] }, workbook)
  assert.deepEqual(getFilteredRows(sheet, workbook), [1, 2, 3, 4])
  const before = structuredClone(sheet.autoFilter)
  for (const criterion of [{ kind: 'color' }, { kind: 'color', cellTextColors: [null] }, { kind: 'color', cellFillColors: [''] }]) {
    assert.throws(() => setFilterCriteria(sheet, 0, criterion, workbook), TypeError)
    assert.deepEqual(sheet.autoFilter, before)
  }
})

test('conditional color projection uses the existing rule evaluator and formula results without authoring cell styles', () => {
  const { sheet, workbook } = setup(), original = structuredClone(sheet.cells)
  sheet.conditionalFormats = [{ id: 'above-two', type: 'expression', operator: 'equals', formula1: '=A1>2', ranges: [range],
    style: { fill: '#00ff00', fontColor: '#ff0000' }, priority: 1, stopIfTrue: false }]
  setAutoFilter(sheet, range, workbook)
  setFilterCriteria(sheet, 0, { kind: 'color', cellFillColors: ['#0f0'] }, workbook)
  assert.deepEqual(getFilteredRows(sheet, workbook), [1, 2])
  assert.deepEqual(sheet.cells, original)
  sheet.cells['1:0'].value = 5
  assert.deepEqual(getFilteredRows(sheet, workbook), [2])
})

test('value and color options exclude other column failures and ignore the current column condition', () => {
  const { sheet, workbook } = setup()
  sheet.cells['1:0'].style.fill = '#ff0000'; sheet.cells['2:0'].style.fill = '#00ff00'
  sheet.cells['3:0'].style.fill = '#0000ff'
  setAutoFilter(sheet, range, workbook)
  setFilterCriteria(sheet, 1, { kind: 'values', selectedItems: ['keep'] }, workbook)
  setFilterCriteria(sheet, 0, { kind: 'color', cellFillColors: ['#ff0000'] }, workbook)
  assert.deepEqual(readDataPanelValues(sheet, range, 0, workbook).map(item => item.value), [2, 4])
  assert.deepEqual(readDataPanelColors(workbook, sheet, range, 0, 'fill').map(item => item.value), ['#00ff00', null])
  assert.deepEqual(readDataPanelColors(workbook, sheet, range, 0, 'text').map(item => item.value), ['#000000'])
})

test('vertically merged groups use their anchor value and failing criteria hide the complete group without altering manual hide reasons', () => {
  const { sheet, workbook } = setup()
  sheet.mergedRanges = [{ startRow: 1, endRow: 2, startCol: 0, endCol: 0 }]
  sheet.cells['2:0'].value = 'interior garbage'; sheet.hiddenRows = [4]
  setAutoFilter(sheet, range, workbook)
  setFilterCriteria(sheet, 0, { kind: 'values', selectedItems: [1] }, workbook)
  assert.deepEqual(getFilteredRows(sheet, workbook), [3, 4])
  setFilterCriteria(sheet, 1, { kind: 'values', selectedItems: ['keep'] }, workbook)
  assert.deepEqual(getFilteredRows(sheet, workbook), [1, 3, 4])
  setFilterCriteria(sheet, 0, { kind: 'values', selectedItems: [4] }, workbook)
  assert.deepEqual(getFilteredRows(sheet, workbook), [1, 2, 3])
  assert.deepEqual(sheet.hiddenRows, [4]); assert.equal(sheet.cells['2:0'].value, 'interior garbage')
})

test('horizontal merge interiors filter through the left anchor while color choices skip a non-top-left target column', () => {
  const { sheet, workbook } = setup()
  sheet.mergedRanges = [{ startRow: 3, endRow: 3, startCol: 0, endCol: 1 }]
  sheet.cells['3:0'].style.fill = '#ff0000'; sheet.cells['3:1'].style.fill = '#00ff00'
  setAutoFilter(sheet, range, workbook)
  setFilterCriteria(sheet, 1, { kind: 'values', selectedItems: [3] }, workbook)
  assert.deepEqual(getFilteredRows(sheet, workbook), [1, 2, 4])
  assert.deepEqual(readDataPanelValues(sheet, range, 1, workbook).map(item => item.value), ['drop', 'keep', 3])
  assert.deepEqual(readDataPanelColors(workbook, sheet, range, 1, 'fill').map(item => item.value), [null])
})

test('a merge beginning above the first data row is skipped while a merge crossing the filter end retains its full rowSpan', () => {
  const { sheet, workbook } = setup()
  sheet.mergedRanges = [{ startRow: 1, endRow: 5, startCol: 0, endCol: 0 }]
  const short = { startRow: 0, endRow: 2, startCol: 0, endCol: 0 }
  setAutoFilter(sheet, short, workbook)
  setFilterCriteria(sheet, 0, { kind: 'values', selectedItems: [999] }, workbook)
  assert.deepEqual(getFilteredRows(sheet, workbook), [1, 2, 3, 4, 5])
  setAutoFilter(sheet, { startRow: 2, endRow: 7, startCol: 0, endCol: 0 }, workbook)
  setFilterCriteria(sheet, 0, { kind: 'values', selectedItems: [] }, workbook)
  assert.deepEqual(getFilteredRows(sheet, workbook), [6, 7])
})

test('merged value candidates retain an anchor if any constituent row survives another column while colors follow the top-row candidate rule', () => {
  const { sheet, workbook } = setup()
  sheet.mergedRanges = [{ startRow: 1, endRow: 2, startCol: 0, endCol: 0 }]
  sheet.cells['1:0'].style.fill = '#ff0000'
  setAutoFilter(sheet, range, workbook)
  setFilterCriteria(sheet, 1, { kind: 'values', selectedItems: ['keep'] }, workbook)
  assert.deepEqual(readDataPanelValues(sheet, range, 0, workbook).map(item => item.value), [1, 4])
  assert.deepEqual(readDataPanelColors(workbook, sheet, range, 0, 'fill').map(item => item.value), [null])
})

test('filter permission scopes include a merged tail beyond ref while leaving unrelated protected cells outside the footprint', () => {
  const { sheet } = setup(), source = { startRow: 0, endRow: 2, startCol: 0, endCol: 1 }
  const merged = { startRow: 1, endRow: 5, startCol: 0, endCol: 0 }
  sheet.mergedRanges = [merged, { startRow: 7, endRow: 9, startCol: 5, endCol: 7 }]
  sheet.rangeProtectionRules = [{ id: 'tail', permissionId: 'tail-permission', ownerUserId: 'owner', editState: 'onlyMe',
    viewState: 'othersCanView', allowedUserIds: [], ranges: [{ startRow: 5, endRow: 5, startCol: 0, endCol: 0 }] }]
  const scopes = sheetFilterPermissionRanges(sheet, source)
  assert.deepEqual(scopes, [source, merged]); assert.deepEqual(sheetFilterPermissionRanges(sheet, null), [])
  assert.equal(canEditProtectedRanges(sheet, [source], 'reader'), true)
  assert.equal(canEditProtectedRanges(sheet, scopes, 'reader'), false)
  sheet.rangeProtectionRules[0].ranges = [{ startRow: 8, endRow: 8, startCol: 6, endCol: 6 }]
  assert.equal(canEditProtectedRanges(sheet, scopes, 'reader'), true)
})

test('header controls use the actual filter heading row, visible panes, disabled state and zoomed 16px geometry', () => {
  const { sheet, workbook } = setup()
  sheet.autoFilter = { range: { startRow: 2, endRow: 8, startCol: 1, endCol: 3 }, columns: { 2: { kind: 'values', selectedItems: [1] } }, filteredRows: [] }
  sheet.frozenRows = 3; sheet.frozenCols = 2
  const geometry = new SpreadsheetGeometry(toExcelWorkbookView(workbook).sheets[0])
  const viewport = { width: 600, height: 400, scrollLeft: 50, scrollTop: 30, zoom: 1.5 }
  const controls = sheetFilterHeaderControls(geometry, viewport, sheet.autoFilter, column => column !== 3)
  assert.equal(controls.length, 3)
  assert.deepEqual(controls.map(control => control.column), [1, 2, 3])
  for (const control of controls) {
    assert.equal(control.row, 2); assert.equal(control.rect.width, 24); assert.equal(control.rect.height, 24)
    assert.equal(control.active, control.column === 2); assert.equal(control.disabled, control.column === 3)
    assert.match(control.clipPath, /^inset\(/)
  }
  assert.deepEqual(sheetFilterHeaderControls(geometry, viewport, undefined, () => true), [])
})

test('quick sum matches the fixed platform modifiers and rejects composition, Shift and prehandled keys', () => {
  const event = { key: '=', code: 'Equal', altKey: true, ctrlKey: false, metaKey: false, shiftKey: false,
    isComposing: false, keyCode: 187, defaultPrevented: false }
  assert.equal(isSheetQuickSumShortcut(event), true)
  assert.equal(isSheetQuickSumShortcut(event, true), false)
  assert.equal(isSheetQuickSumShortcut({ ...event, metaKey: true }, true), true)
  for (const change of [{ ctrlKey: true }, { metaKey: true }, { shiftKey: true }, { isComposing: true },
    { keyCode: 229 }, { defaultPrevented: true }, { key: '-', code: 'Minus' }, { altKey: false }]) {
    assert.equal(isSheetQuickSumShortcut({ ...event, ...change }), false)
  }
})
