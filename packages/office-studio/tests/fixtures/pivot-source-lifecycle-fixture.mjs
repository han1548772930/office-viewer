import { createBlankSheet, createCell } from '../../src/domain/model.ts'
import { renderPivotTable } from '../../src/features/pivot/services/pivot-engine.ts'

export const lifecycleRange = { startRow: 2, endRow: 6, startCol: 3, endCol: 7 }
export const lifecycleFieldIds = ['region-source', 'segment-source', 'amount-source', 'other-source', 'date-source']
export const lifecycleOptions = { repeatRowLabels: false, showRowGrandTotal: false, showColGrandTotal: false,
  showRowSubtotals: false, showColSubtotals: false, valuePosition: 'column' }
export const allLifecycleReports = workbook => [...workbook.pivotTables, ...workbook.charts.flatMap(chart => chart.ownedPivot ? [chart.ownedPivot] : [])]

/** Real sparse source cells, duplicate value instances, two linked reports and one owned report share one cache. */
export function pivotSourceLifecycleWorkbook({ selfContained = false, disconnected = false } = {}) {
  const unitId = 'pivot-source-lifecycle-unit', sheets = ['Data', 'Reports', 'Dashboard', 'Paste'].map(name => {
    const sheet = createBlankSheet(name); sheet.id = name.toLowerCase(); sheet.rowCount = 120; sheet.columnCount = 40
    sheet.defaultRowHeight = 20; sheet.defaultColumnWidth = 80; sheet.defaultRowAutoHeight = false
    return sheet
  })
  const data = sheets[0], host = selfContained ? data : sheets[1], dashboard = selfContained ? data : sheets[2]
  const rows = [['Region', 'Segment', 'Amount', 'Other', 'Date'],
    ['North', 'Retail', 10, 100, 45292], ['South', 'Retail', 20, 200, 45323],
    ['North', 'Wholesale', 30, 300, 45352], ['West', 'Wholesale', 40, 400, 45383]]
  rows.forEach((row, index) => row.forEach((value, column) => {
    const r = index + lifecycleRange.startRow, c = column + lifecycleRange.startCol
    data.cells[`${r}:${c}`] = createCell(r, c, value)
  }))
  const source = { unitId, subUnitId: data.id, sheetName: data.name, range: structuredClone(lifecycleRange), pivotCacheId: 'source-cache' }
  const fields = () => [
    { id: 'region-local', sourceFieldId: 'region-source', sourceColumn: 3, name: 'Region', area: 'row', position: 0 },
    { id: 'segment-local', sourceFieldId: 'segment-source', sourceColumn: 4, name: 'Segment', area: 'row', position: 1 },
    { id: 'amount-local', sourceFieldId: 'amount-source', sourceColumn: 5, name: 'Amount', area: 'value', position: 0, aggregation: 'sum' },
    { id: 'amount-average-local', sourceFieldId: 'amount-source', sourceColumn: 5, name: 'Amount', displayName: 'Average Amount', area: 'value', position: 1, aggregation: 'average' },
  ]
  const report = (id, row, sheetId = host.id) => ({ id, source: structuredClone(source), pivotCacheId: 'source-cache',
    anchor: { unitId, subUnitId: sheetId, row, col: 1 }, fields: fields(), options: structuredClone(lifecycleOptions),
    collapseInfo: { 'region-local': { North: true } }, generatedCellKeys: [], drillDownMap: {}, revision: 0 })
  const pivotTables = [report('report-a', 12), report('report-b', 28)]
  const owned = report('owned-report', 0, dashboard.id)
  delete owned.generatedCellKeys; delete owned.drillDownMap
  const chart = (id, sheetId, extra) => ({ id, sheetId, type: 'column',
    source: { sheetId: data.id, range: structuredClone(lifecycleRange), orientation: 'columns' },
    position: { x: 900, y: 400 }, size: { width: 160, height: 100 }, categoryField: 0, valueFields: [2, 3],
    legend: { visible: true, position: 'bottom' }, xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' },
    revision: 0, ...extra })
  const charts = [chart('linked-a', host.id, { pivotTableId: 'report-a' }),
    chart('linked-b', dashboard.id, { pivotTableId: 'report-a' }), chart('owned-chart', dashboard.id, { ownedPivot: owned })]
  const workbook = { id: unitId, kind: 'workbook', name: 'Pivot source lifecycle', sheets, activeSheetId: host.id,
    date1904: false, calculationMode: 'manual', pivotTables, charts, tables: [], modified: false,
    pivotCaches: [{ id: 'source-cache', source: structuredClone(source),
      fields: lifecycleFieldIds.map((id, index) => ({ id, sourceColumn: index + 3, name: rows[0][index] })), nextFieldId: 0, revision: 0 }],
    slicerCaches: [{ id: 'control-cache', unitId, type: 'slicer', source: { kind: 'pivot', pivotCacheId: 'source-cache', sourceFieldId: 'region-source' },
      connections: { pivotTables: pivotTables.map(report => ({ subUnitId: host.id, pivotTableId: report.id })),
        ownedPivotCharts: [{ subUnitId: dashboard.id, pivotChartId: 'owned-chart' }] }, revision: 0 }],
    slicers: [{ id: 'region-view', unitId, hostSheetId: dashboard.id, cacheId: 'control-cache', pivotTableId: 'report-a',
      sourceFieldId: 'region-source', caption: 'Region', type: 'slicer', position: { row: 70, column: 1 }, size: { rows: 5, columns: 2 },
      style: 'SlicerStyleLight2', revision: 0 }],
  }
  if (disconnected) workbook.pivotCaches.push({ id: 'disconnected-cache', source: { ...structuredClone(source), pivotCacheId: 'disconnected-cache' },
    fields: lifecycleFieldIds.map((id, index) => ({ id: `disconnected-${id}`, sourceColumn: index + 3, name: rows[0][index] })), nextFieldId: 4, revision: 0 })
  for (const pivot of pivotTables) renderPivotTable(workbook, pivot)
  for (const linked of charts.filter(chart => chart.pivotTableId)) {
    const pivot = pivotTables.find(pivot => pivot.id === linked.pivotTableId)
    linked.source = { ...linked.source, sheetId: pivot.anchor.subUnitId, range: structuredClone(pivot.outputRange) }
  }
  return workbook
}
