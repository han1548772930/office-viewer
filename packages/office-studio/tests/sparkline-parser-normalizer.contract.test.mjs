import assert from 'node:assert/strict'
import { test } from 'vitest'
import { normalizeWorkbook } from '../src/parser/workbookNormalizer.ts'

test('parser sparkline payload becomes a persisted cell-level model entry', () => {
  const workbook = normalizeWorkbook({
    schemaVersion: 1, parserVersion: 'test', format: 'xlsx', fileName: 'sparklines.xlsx', size: 10, model: 'spreadsheet',
    payload: { modelType: 'spreadsheet', data: { date1904: false, activeSheet: 0, sheets: [{ name: 'Data', cells: [{
      row: 0, col: 5, value: null, dataType: 'blank', display: '', sparkline: {
        type: 'column', sourceRange: 'B1:E1', color: '#167A55', negativeColor: '#C43B3B',
        highPointColor: '#F0A000', lowPointColor: '#2B579A', emptyCells: 'zero', axis: 'custom', min: -5, max: 20,
      },
    }] }] } },
    diagnostics: { warnings: [], errors: [], stats: {}, capabilities: ['sparklines'] },
  })
  assert.deepEqual(workbook.sheets[0].cells['0:5'].sparkline, {
    type: 'column', sourceRange: 'B1:E1', color: '#167A55', negativeColor: '#C43B3B',
    highPointColor: '#F0A000', lowPointColor: '#2B579A', emptyCells: 'zero', axis: 'custom', min: -5, max: 20,
  })
})

test('parser pivot cache and pivotSource become editable workbook pivot and chart models', () => {
  const workbook = normalizeWorkbook({
    schemaVersion: 1, parserVersion: 'test', format: 'xlsx', fileName: 'pivot.xlsx', size: 10, model: 'spreadsheet',
    payload: { modelType: 'spreadsheet', data: {
      activeSheet: 1, date1904: false,
      sheets: [{ name: 'Data', cells: [] }, { name: 'Pivot View', cells: [{ row: 0, col: 0, value: 'Region', display: 'Region' }], charts: [{ index: 0, kind: 'bar', pivotTableId: 'pivot-1', legend: true, drawing: { fromOffsetEmu: [95250, 190500], absoluteSizeEmu: [3810000, 2286000] } }] }],
      pivotTables: [{ id: 'pivot-1', source: { sheetName: 'Data', subUnitId: 'Data', range: { startRow: 0, startCol: 0, endRow: 2, endCol: 1 } }, anchor: { subUnitId: 'Pivot View', row: 0, col: 0 }, outputRange: { startRow: 0, startCol: 0, endRow: 3, endCol: 2 }, fields: [{ id: '0', sourceColumn: 0, name: 'Region', area: 'row', position: 0 }, { id: '1', sourceColumn: 1, name: 'Sales', area: 'value', position: 0, aggregation: 'sum' }], options: { showRowGrandTotal: true, showColGrandTotal: true, showRowSubtotals: true, showColSubtotals: true, valuePosition: 'column' }, revision: 0 }],
    } },
    diagnostics: { warnings: [], errors: [], stats: {}, capabilities: ['pivotTables', 'pivotCharts'] },
  })
  assert.equal(workbook.pivotTables[0].source.subUnitId, workbook.sheets[0].id)
  assert.equal(workbook.pivotTables[0].anchor.subUnitId, workbook.sheets[1].id)
  assert.equal(workbook.charts[0].pivotTableId, workbook.pivotTables[0].id)
  assert.equal(workbook.charts[0].sheetId, workbook.sheets[1].id)
  assert.deepEqual(workbook.charts[0].source.range, workbook.pivotTables[0].outputRange)
})
