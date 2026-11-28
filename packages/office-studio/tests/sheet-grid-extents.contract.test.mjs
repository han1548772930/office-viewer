import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankSheet, createBlankWorkbook, createCell, DEFAULT_SHEET_ROW_COUNT, DEFAULT_SHEET_COLUMN_COUNT } from '../src/domain/model.ts'
import { workbookFromInput } from '../src/api/model-utils.ts'
import { normalizeWorkbook } from '../src/parser/workbookNormalizer.ts'
import { toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { createSheetAutoFillPlan } from '../src/features/sheet-workbench/sheet-auto-fill.ts'
import { SpreadsheetGeometry } from '../src/rendering/spreadsheet/layout.ts'
import { createSpreadsheetScene } from '../src/rendering/spreadsheet/scene.ts'

// Deferred regression for the screenshot: filling C5:C17 must not turn the
// whole worksheet into a 17-row, 3-column drawing surface.
test('blank worksheets materialize the fixed Univer grid dimensions', () => {
  const sheet = createBlankSheet()
  assert.equal(DEFAULT_SHEET_ROW_COUNT, 1000); assert.equal(DEFAULT_SHEET_COLUMN_COUNT, 20)
  assert.equal(sheet.rowCount, 1000); assert.equal(sheet.columnCount, 20)
})

test.each(['public', 'domain', 'parser'])('%s input with omitted counts preserves a real default grid', kind => {
  const domain = createBlankWorkbook(); delete domain.sheets[0].rowCount; delete domain.sheets[0].columnCount
  const workbook = kind === 'public' ? workbookFromInput({ id: 'public-grid', sheetOrder: ['data'], sheets: { data: { name: 'Data', cellData: {} } } })
    : kind === 'domain' ? workbookFromInput(domain)
      : normalizeWorkbook({ format: 'xlsx', fileName: 'grid.xlsx', parserVersion: 'fixture', payload: { data: { sheets: [{ name: 'Data', cells: [] }] } } })
  assert.equal(workbook.sheets[0].rowCount, 1000); assert.equal(workbook.sheets[0].columnCount, 20)
  const view = toExcelWorkbookView(workbook).sheets[0]
  assert.equal(view.rowCount, 1000); assert.equal(view.columnCount, 20)
})

test('filling a legacy worksheet without counts leaves empty grid lines to the right and below', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  delete sheet.rowCount; delete sheet.columnCount
  sheet.cells['4:2'] = createCell(4, 2, 12313)
  const before = toExcelWorkbookView(workbook).sheets[0]
  const plan = createSheetAutoFillPlan(workbook, sheet.id, { startRow: 4, endRow: 4, startCol: 2, endCol: 2 },
    { startRow: 4, endRow: 16, startCol: 2, endCol: 2 }, { applyType: 'COPY' })
  assert.equal(plan.result.accepted, true)
  const after = toExcelWorkbookView({ ...workbook, sheets: [plan.sheet] }).sheets[0]
  assert.equal(after.rowCount, before.rowCount); assert.equal(after.columnCount, before.columnCount)
  assert.equal(after.cells.C17.rawValue, 12313); assert.equal(after.cells.D18, undefined)
  const geometry = new SpreadsheetGeometry(after), viewport = { width: 1000, height: 800, scrollLeft: 0, scrollTop: 0, zoom: 1 }
  const endpoint = geometry.cellRect(16, 2, viewport, false)
  const scene = createSpreadsheetScene(geometry, viewport, { activeCell: 'C5', selectedCells: [],
    selectedRanges: [{ startRow: 4, endRow: 16, startColumn: 2, endColumn: 2 }] })
  const grid = scene.commands.filter(command => command.kind === 'rect' && command.stroke === '#e5e9f0')
  assert.ok(grid.some(command => command.x >= endpoint.x + endpoint.width), 'blank columns beyond C stay drawn')
  assert.ok(grid.some(command => command.y >= endpoint.y + endpoint.height), 'blank rows beyond 17 stay drawn')
  assert.equal(new SpreadsheetGeometry(before).rows.count, geometry.rows.count)
  assert.equal(new SpreadsheetGeometry(before).columns.count, geometry.columns.count)
})

test('explicit worksheet counts survive normalization and filling without being expanded to defaults', () => {
  const workbook = workbookFromInput({ sheetOrder: ['small'], sheets: { small: { name: 'Small', rowCount: 30, columnCount: 12,
    cellData: { 4: { 2: { v: 12313 } } } } } }), sheet = workbook.sheets[0]
  const plan = createSheetAutoFillPlan(workbook, sheet.id, { startRow: 4, endRow: 4, startCol: 2, endCol: 2 },
    { startRow: 4, endRow: 16, startCol: 2, endCol: 2 }, { applyType: 'COPY' })
  assert.equal(plan.sheet.rowCount, 30); assert.equal(plan.sheet.columnCount, 12)
  const reopened = workbookFromInput(JSON.parse(JSON.stringify({ ...workbook, sheets: [plan.sheet] })))
  assert.equal(reopened.sheets[0].rowCount, 30); assert.equal(reopened.sheets[0].columnCount, 12)
  assert.equal(reopened.sheets[0].cells['16:2'].value, 12313)
})
