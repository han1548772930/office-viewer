import { createBlankSheet, createCell } from '../../src/domain/model.ts'
import { createPivotTable } from '../../src/features/pivot/services/pivot-engine.ts'
import { applyPivotChartEdit, createPivotChartReport } from '../../src/features/pivot/services/pivot-chart-engine.ts'

const leaves = [['North', 'A', 'a1', 10], ['North', 'A', 'a2', 20], ['North', 'B', 'b1', 30], ['North', 'B', 'b2', 40],
  ['South', 'C', 'c1', 20], ['South', 'C', 'c2', 10], ['South', 'D', 'd1', 10], ['South', 'D', 'd2', 60]]
const columns = [[2024, 'Q1', 'Retail', 1], [2024, 'Q1', 'Web', 0.5], [2024, 'Q2', 'Retail', 2], [2024, 'Q2', 'Web', 1.5]]
export const PIVOT_AXIS_RECORDS = leaves.flatMap(([region, product, item, amount]) =>
  columns.map(([year, quarter, channel, factor]) => [region, product, item, year, quarter, channel, amount * factor]))

export function pivotAxisFields({ aggregation = 'sum', multipleValues = false } = {}) {
  const fields = [
    { id: 'region-label', sourceFieldId: 'region-data', sourceColumn: 2, name: 'Region', area: 'row', position: 0, sort: 'ascending' },
    { id: 'product-label', sourceFieldId: 'product-data', sourceColumn: 3, name: 'Product', area: 'row', position: 1, sort: 'ascending' },
    { id: 'item-label', sourceFieldId: 'item-data', sourceColumn: 4, name: 'Item', area: 'row', position: 2, sort: 'ascending' },
    { id: 'year-label', sourceFieldId: 'year-data', sourceColumn: 5, name: 'Year', area: 'column', position: 0, sort: 'ascending' },
    { id: 'quarter-label', sourceFieldId: 'quarter-data', sourceColumn: 6, name: 'Quarter', area: 'column', position: 1, sort: 'ascending' },
    { id: 'channel-label', sourceFieldId: 'channel-data', sourceColumn: 7, name: 'Channel', area: 'column', position: 2, sort: 'ascending' },
    { id: 'amount-value', sourceFieldId: 'amount-data', sourceColumn: 8, name: 'Amount', area: 'value', position: 0, aggregation },
  ]
  if (multipleValues) fields.push({ ...structuredClone(fields[6]), id: 'amount-average', position: 1, aggregation: 'average', displayName: 'Average' })
  return fields
}

/** Three real levels on each axis; North/Q1/Retail=100, grand=1000, source starts at column 2. */
export function pivotAxisWorkbook({ ownership = 'owned', records = PIVOT_AXIS_RECORDS, totals = false,
  valuePosition = 'column', aggregation = 'sum', multipleValues = false } = {}) {
  const data = createBlankSheet('Data'), host = createBlankSheet('Report'), dashboard = createBlankSheet('Dashboard')
  data.id = 'axis-data'; host.id = 'axis-host'; dashboard.id = 'axis-dashboard'
  for (const sheet of [data, host, dashboard]) { sheet.rowCount = 400; sheet.columnCount = 80 }
  const headers = ['Region', 'Product', 'Item', 'Year', 'Quarter', 'Channel', 'Amount'], rows = [headers, ...records]
  rows.forEach((row, index) => row.forEach((value, column) => { const r = index + 1, c = column + 2; data.cells[`${r}:${c}`] = createCell(r, c, value) }))
  const id = 'pivot-axis-unit', source = { unitId: id, subUnitId: data.id, sheetName: data.name,
    range: { startRow: 1, endRow: rows.length, startCol: 2, endCol: 8 } }
  const input = { ...source, pivotCacheId: 'axis-cache', range: { startRow: 1, endRow: rows.length, startColumn: 2, endColumn: 8 } }
  const workbook = { id, kind: 'workbook', name: 'Pivot axis state', sheets: [data, host, dashboard], activeSheetId: host.id,
    pivotTables: [], charts: [], pivotCaches: [{ id: 'axis-cache', source, revision: 0, nextFieldId: 0,
      fields: ['region-data', 'product-data', 'item-data', 'year-data', 'quarter-data', 'channel-data', 'amount-data']
        .map((id, index) => ({ id, sourceColumn: index + 2, name: headers[index] })) }], slicers: [], slicerCaches: [], modified: false, date1904: false }
  const report = ownership === 'owned' ? createPivotChartReport(workbook, input, host.id, 'axis-owned')
    : createPivotTable(workbook, input, { unitId: id, subUnitId: host.id, row: 20, col: 2 }, 'axis-table')
  const chart = { id: 'axis-chart-a', sheetId: host.id, type: 'column',
    source: { sheetId: data.id, range: structuredClone(source.range), orientation: 'columns' },
    position: { x: 420, y: 60 }, size: { width: 360, height: 260 }, categoryField: 0, valueFields: [3],
    legend: { visible: true, position: 'bottom' }, xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, revision: 0,
    ...(ownership === 'owned' ? { ownedPivot: report } : { pivotTableId: report.id }) }
  workbook.charts.push(chart)
  applyPivotChartEdit(workbook, chart, { type: 'setFieldsConfig', config: { fields: pivotAxisFields({ aggregation, multipleValues }), options: {
    repeatRowLabels: true, showRowGrandTotal: totals, showColGrandTotal: totals,
    showRowSubtotals: totals, showColSubtotals: totals, valuePosition,
  } } })
  if (ownership === 'linked') workbook.charts.push({ ...structuredClone(chart), id: 'axis-chart-b', sheetId: dashboard.id, title: 'Other visual' })
  return workbook
}
