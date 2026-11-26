import { createBlankSheet, createCell } from '../../src/domain/model.ts'
import { createPivotTable } from '../../src/features/pivot/services/pivot-engine.ts'
import { createPivotChartReport, applyPivotChartEdit } from '../../src/features/pivot/services/pivot-chart-engine.ts'

export const PIVOT_SHOW_DATA_AS_RECORDS = [
  ['North', 'A', 2024, 'Q1', 10], ['North', 'A', 2024, 'Q2', 20],
  ['North', 'B', 2024, 'Q1', 30], ['North', 'B', 2024, 'Q2', 40],
  ['South', 'C', 2024, 'Q1', 20], ['South', 'C', 2024, 'Q2', 10],
  ['South', 'D', 2024, 'Q1', 10], ['South', 'D', 2024, 'Q2', 60],
]

export function pivotShowDataAsFields() {
  return [
    { id: 'region-label', sourceFieldId: 'region-data', sourceColumn: 2, name: 'Region', area: 'row', position: 0, sort: 'ascending' },
    { id: 'product-label', sourceFieldId: 'product-data', sourceColumn: 3, name: 'Product', area: 'row', position: 1, sort: 'ascending' },
    { id: 'year-label', sourceFieldId: 'year-data', sourceColumn: 4, name: 'Year', area: 'column', position: 0, sort: 'ascending' },
    { id: 'quarter-label', sourceFieldId: 'quarter-data', sourceColumn: 5, name: 'Quarter', area: 'column', position: 1, sort: 'ascending' },
    { id: 'amount-value', sourceFieldId: 'amount-data', sourceColumn: 6, name: 'Amount', area: 'value', position: 0, aggregation: 'sum' },
  ]
}

/** Two real hierarchy axes; grand=200, Q1=70, North/Q1=40, North/A=30. */
export function pivotShowDataAsWorkbook({ ownership = 'owned', records = PIVOT_SHOW_DATA_AS_RECORDS, totals = true } = {}) {
  const data = createBlankSheet('Data'), host = createBlankSheet('Report'), second = createBlankSheet('Dashboard')
  data.id = 'show-data'; host.id = 'show-host'; second.id = 'show-dashboard'
  for (const sheet of [data, host, second]) { sheet.rowCount = 300; sheet.columnCount = 60 }
  const unitId = 'show-data-as-unit', rows = [['Region', 'Product', 'Year', 'Quarter', 'Amount'], ...records]
  rows.forEach((row, rowIndex) => row.forEach((value, columnIndex) => {
    const r = rowIndex + 1, c = columnIndex + 2; data.cells[`${r}:${c}`] = createCell(r, c, value)
  }))
  const source = { unitId, subUnitId: data.id, sheetName: data.name, range: { startRow: 1, endRow: rows.length, startCol: 2, endCol: 6 } }
  const input = { ...source, pivotCacheId: 'show-cache', range: { startRow: 1, endRow: rows.length, startColumn: 2, endColumn: 6 } }
  const workbook = { id: unitId, kind: 'workbook', name: 'Show Values As', sheets: [data, host, second], activeSheetId: host.id,
    pivotTables: [], charts: [], pivotCaches: [{ id: 'show-cache', source, revision: 0,
      fields: ['region-data', 'product-data', 'year-data', 'quarter-data', 'amount-data'].map((id, index) => ({ id, sourceColumn: index + 2, name: rows[0][index] })), nextFieldId: 0 }],
    slicers: [], slicerCaches: [], date1904: false, modified: false }
  const report = ownership === 'owned' ? createPivotChartReport(workbook, input, host.id, 'show-owned')
    : createPivotTable(workbook, input, { unitId, subUnitId: host.id, row: 12, col: 1 }, 'show-table')
  const chart = { id: 'show-chart-a', sheetId: host.id, type: 'column',
    source: { sheetId: data.id, range: structuredClone(source.range), orientation: 'columns' },
    position: { x: 360, y: 60 }, size: { width: 320, height: 240 }, categoryField: 0, valueFields: [2],
    legend: { visible: true, position: 'bottom' }, xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, revision: 0,
    ...(ownership === 'owned' ? { ownedPivot: report } : { pivotTableId: report.id }) }
  workbook.charts.push(chart)
  applyPivotChartEdit(workbook, chart, { type: 'setFieldsConfig', config: { fields: pivotShowDataAsFields(), options: {
    repeatRowLabels: true, showRowGrandTotal: totals, showColGrandTotal: totals,
    showRowSubtotals: totals, showColSubtotals: totals, valuePosition: 'column',
  } } })
  if (ownership !== 'owned') workbook.charts.push({ ...structuredClone(chart), id: 'show-chart-b', sheetId: second.id, title: 'Second visual' })
  return workbook
}
