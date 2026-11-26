import { createBlankSheet, createCell } from '../../src/domain/model.ts'
import { createPivotTable } from '../../src/features/pivot/services/pivot-engine.ts'
import { applyPivotChartEdit, createPivotChartReport } from '../../src/features/pivot/services/pivot-chart-engine.ts'

export const pivotChartSourceInput = {
  unitId: 'pivot-chart-engine-unit', subUnitId: 'data', sheetName: 'Data', pivotCacheId: 'engine-cache',
  range: { startRow: 2, endRow: 6, startColumn: 3, endColumn: 7 },
}

export const pivotChartReportOptions = {
  repeatRowLabels: false, showRowGrandTotal: false, showColGrandTotal: false,
  showRowSubtotals: false, showColSubtotals: false, valuePosition: 'column',
}

export function pivotChartFields() {
  return [
    { id: 'region-local', sourceFieldId: 'region-source', sourceColumn: 3, name: 'Region', area: 'row', position: 0 },
    { id: 'amount-local', sourceFieldId: 'amount-source', sourceColumn: 4, name: 'Amount', area: 'value', position: 0, aggregation: 'sum' },
  ]
}

/** Real sparse source cells at a nonzero offset, one retained source cache, and actual report ownership. */
export function pivotChartEngineWorkbook({ ownership = 'owned', configured = true, connected = false, staleSibling = false } = {}) {
  const data = createBlankSheet('Data'), reports = createBlankSheet('Reports'), dashboard = createBlankSheet('Dashboard')
  data.id = 'data'; reports.id = 'reports'; dashboard.id = 'dashboard'
  for (const sheet of [data, reports, dashboard]) { sheet.rowCount = 200; sheet.columnCount = 40 }
  const rows = [['Region', 'Amount', 'Date', 'Segment', 'Score'],
    ['North', 10, 45292, 'A', 5], ['South', 20, 45323, 'A', 15], ['North', 30, 45352, 'B', 25], ['West', 40, 45383, 'B', 35]]
  rows.forEach((row, index) => row.forEach((value, column) => {
    const r = index + 2, c = column + 3; data.cells[`${r}:${c}`] = createCell(r, c, value)
  }))
  const source = { unitId: pivotChartSourceInput.unitId, subUnitId: data.id, sheetName: data.name,
    range: { startRow: 2, endRow: 6, startCol: 3, endCol: 7 } }
  const workbook = { id: pivotChartSourceInput.unitId, kind: 'workbook', name: 'Pivot chart engine', sheets: [data, reports, dashboard],
    activeSheetId: reports.id, charts: [], pivotTables: [], pivotCaches: [{ id: 'engine-cache', source,
      fields: ['region-source', 'amount-source', 'date-source', 'segment-source', 'score-source'].map((id, index) => ({ id, sourceColumn: index + 3, name: rows[0][index] })),
      nextFieldId: 0, revision: 0 }], slicers: [], slicerCaches: [], date1904: false, modified: false }
  if (ownership === 'none') return workbook
  const report = ownership === 'owned'
    ? createPivotChartReport(workbook, structuredClone(pivotChartSourceInput), reports.id, 'owned-report')
    : createPivotTable(workbook, structuredClone(pivotChartSourceInput), { unitId: workbook.id, subUnitId: reports.id, row: 8, col: 1 }, 'linked-report')
  const chart = { id: 'chart-a', sheetId: reports.id, type: 'column',
    source: { sheetId: data.id, range: structuredClone(source.range), orientation: 'columns' },
    position: { x: 300, y: 40 }, size: { width: 240, height: 160 },
    categoryField: 0, valueFields: [1], legend: { visible: true, position: 'bottom' },
    xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, revision: 0,
    ...(ownership === 'owned' ? { ownedPivot: report } : { pivotTableId: report.id }) }
  workbook.charts.push(chart)
  if (configured) applyPivotChartEdit(workbook, chart, { type: 'setFieldsConfig', config: { fields: pivotChartFields(), options: pivotChartReportOptions } })
  if (ownership === 'linked') workbook.charts.push({ ...structuredClone(chart), id: 'chart-b', sheetId: dashboard.id })
  if (connected) {
    const connections = ownership === 'owned' ? { pivotTables: [], ownedPivotCharts: [{ subUnitId: reports.id, pivotChartId: chart.id }] }
      : { pivotTables: [{ subUnitId: reports.id, pivotTableId: report.id }] }
    if (staleSibling) connections.pivotTables.push({ subUnitId: dashboard.id, pivotTableId: 'missing-sibling' })
    workbook.slicerCaches.push({ id: 'engine-control-cache', unitId: workbook.id, type: 'slicer',
      source: { kind: 'pivot', pivotCacheId: 'engine-cache', sourceFieldId: 'region-source' }, connections, revision: 0 })
    workbook.slicers.push({ id: 'engine-slicer', unitId: workbook.id, hostSheetId: dashboard.id, type: 'slicer', cacheId: 'engine-control-cache',
      sourceFieldId: 'region-source', ...(ownership === 'owned' ? { ownedPivotChartId: chart.id } : { pivotTableId: report.id }),
      caption: 'Region', position: { row: 10, column: 3 }, size: { rows: 5, columns: 3 },
      layout: { left: 200, top: 240, width: 160, height: 120 }, style: 'SlicerStyleLight2', revision: 0 })
  }
  return workbook
}
