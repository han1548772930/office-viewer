import assert from 'node:assert/strict'
import { after, before, test } from 'node:test'
import { createServer } from 'vite'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { toExcelWorkbookView, applyExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { setAutoFilter, setFilterCriteria } from '../src/features/sheet-sort-filter/sort-filter-engine.ts'
import { SpreadsheetGeometry, GRID_HEADER_HEIGHT } from '../src/rendering/spreadsheet/layout.ts'
import { createSpreadsheetScene } from '../src/rendering/spreadsheet/scene.ts'
import { calculateSelectionStatistics, normalizeSheetViewState } from '../src/features/sheet-workbench/sheet-workbench.ts'
import { findVisibleSheetRow, hiddenSheetRows } from '../src/features/sheet-workbench/sheet-row-visibility.ts'
import { applySheetNumberFormat } from '../src/features/sheet-workbench/sheet-toolbar-format.ts'

let server, api, adapter
before(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
  adapter = await server.ssrLoadModule('/packages/office-studio/src/services/excelViewAdapter.ts')
})
after(async () => { await server?.close() })

function filteredWorkbook() {
  const workbook = createBlankWorkbook()
  const sheet = workbook.sheets[0]
  sheet.defaultRowHeight = 20
  sheet.hiddenRows = [3]
  ;[['Region', 'Amount'], ['East', 10], ['West', 20], ['East', 30], ['West', 40]].forEach((row, rowIndex) => row.forEach((value, column) => { sheet.cells[`${rowIndex}:${column}`] = createCell(rowIndex, column, value) }))
  setAutoFilter(sheet, { startRow: 0, endRow: 4, startCol: 0, endCol: 1 })
  setFilterCriteria(sheet, 0, { kind: 'values', selectedItems: ['East'] })
  return workbook
}

test('filter view keeps separate hidden reasons and shares visible axes, scene, hit-testing and statistics', () => {
  const workbook = filteredWorkbook()
  const view = toExcelWorkbookView(workbook), sheet = view.sheets[0]
  assert.deepEqual(sheet.hiddenRows, [3])
  assert.deepEqual(sheet.filteredRows, [2, 4])
  const geometry = new SpreadsheetGeometry(sheet)
  assert.deepEqual(geometry.rows.visible(0, 100).slice(0, 4), [0, 1, 5, 6])
  assert.equal(geometry.rows.offset(5), 40)
  const viewport = { width: 600, height: 300, scrollLeft: 0, scrollTop: 0, zoom: 1 }
  assert.equal(geometry.hitTest(80, GRID_HEADER_HEIGHT + 45, viewport).row, 5)
  const scene = createSpreadsheetScene(geometry, viewport, { activeCell: 'A2', selectedCells: [] })
  assert.ok(scene.commands.some(command => command.kind === 'text' && command.text === 'East'))
  assert.ok(!scene.commands.some(command => command.kind === 'text' && command.text === 'West'))
  const statistics = calculateSelectionStatistics(sheet, [{ startRow: 1, endRow: 4, startColumn: 1, endColumn: 1 }])
  assert.equal(statistics.sum, 10)
  assert.equal(statistics.count, 1)
  assert.equal(findVisibleSheetRow(sheet, 2, 1), 5)
  assert.equal(findVisibleSheetRow(sheet, 4, -1), 1)
  assert.equal(normalizeSheetViewState(sheet, { key: 'A3', anchor: 'A3' }).key, 'A6')
  assert.deepEqual([...hiddenSheetRows(sheet)].sort(), [2, 3, 4])
})

test('view format edits preserve filter configuration and never convert derived rows into manual hiding', () => {
  const workbook = filteredWorkbook(), source = JSON.parse(JSON.stringify(workbook))
  const view = toExcelWorkbookView(workbook)
  view.sheets[0].cells.B2 = applySheetNumberFormat(view.sheets[0].cells.B2, 'currency')
  view.sheets[0].filteredRows = [1, 2, 3, 4] // Untrusted/stale view cannot rewrite the authoritative filter.
  const next = applyExcelWorkbookView(workbook, view)
  assert.deepEqual(next.sheets[0].hiddenRows, [3])
  assert.deepEqual(next.sheets[0].autoFilter, source.sheets[0].autoFilter)
  assert.equal(next.sheets[0].cells['1:1'].display, '¥10.00')
  assert.deepEqual(workbook, source, 'projection must not mutate the previous domain snapshot')
  view.sheets[0].cells.A3 = { ...view.sheets[0].cells.A3, value: 'East', rawValue: undefined }
  const edited = applyExcelWorkbookView(workbook, view)
  assert.deepEqual(edited.sheets[0].autoFilter.filteredRows, [4])
  assert.deepEqual(edited.sheets[0].hiddenRows, [3])
})

test('filter → toolbar format → undo/redo → clear preserves visibility and manual hiding', () => {
  const model = filteredWorkbook()
  // Start before filtering so the public commands establish a real history transaction.
  model.sheets[0].autoFilter = undefined
  const studio = new api.OfficeStudio({ workbook: model })
  try {
    const workbook = api.FUniver.newAPI(studio).getActiveWorkbook(), sheet = workbook.getActiveSheet()
    sheet.setAutoFilter('A1:B5').setFilterCriteria(0, { kind: 'values', selectedItems: ['East'] })
    const view = adapter.toExcelWorkbookView(studio.getWorkbookModel(workbook.getId()))
    view.sheets[0].cells.B2 = { ...view.sheets[0].cells.B2, strike: true, numberFormat: 'currency', rawNumberFormat: '¥#,##0.00;[Red]-¥#,##0.00' }
    studio.editor.replaceWorkbookFromView(view)
    assert.equal(studio.getWorkbookModel(workbook.getId()).sheets[0].cells['1:1'].style.font.strike, true)
    assert.deepEqual(sheet.getFilteredRows(), [2, 4])
    workbook.undo()
    assert.equal(studio.getWorkbookModel(workbook.getId()).sheets[0].cells['1:1'].style.font.strike, false)
    assert.deepEqual(sheet.getFilteredRows(), [2, 4])
    workbook.redo()
    assert.equal(studio.getWorkbookModel(workbook.getId()).sheets[0].cells['1:1'].style.font.strike, true)
    sheet.clearFilter()
    const restored = adapter.toExcelWorkbookView(studio.getWorkbookModel(workbook.getId())).sheets[0]
    assert.deepEqual(restored.filteredRows, [])
    assert.deepEqual(restored.hiddenRows, [3])
    assert.equal(new SpreadsheetGeometry(restored).rows.offset(5), 80)
    workbook.undo()
    assert.deepEqual(sheet.getFilteredRows(), [2, 4])
    assert.deepEqual(studio.getWorkbookModel(workbook.getId()).sheets[0].hiddenRows, [3])
  } finally { studio.dispose() }
})
