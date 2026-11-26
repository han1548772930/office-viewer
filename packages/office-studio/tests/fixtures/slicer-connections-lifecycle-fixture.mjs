import { createBlankSheet, createCell } from '../../src/domain/model.ts'

/** Real sparse source/report models shared by deferred connection, copy and lifecycle contracts. */
export function slicerConnectionsWorkbook({ singleSheet = false, owned = true, stableFields = false } = {}) {
  const unitId = 'connections-unit', data = createBlankSheet('Data'), reports = createBlankSheet('Reports'), paste = createBlankSheet('Paste')
  data.id = 'data'; reports.id = 'reports'; paste.id = 'paste'
  for (const sheet of [data, reports, paste]) {
    sheet.rowCount = 120; sheet.columnCount = 40; sheet.defaultRowHeight = 20; sheet.defaultColumnWidth = 80; sheet.defaultRowAutoHeight = false
  }
  const host = singleSheet ? data : reports
  const rows = [['Region', 'Amount', 'Date'], ['North', 10, 45292], ['South', 20, 45323], ['North', 30, 45352], ['West', 40, 45383]]
  rows.forEach((row, index) => row.forEach((value, col) => { data.cells[`${index}:${col}`] = createCell(index, col, value) }))
  const source = { unitId, subUnitId: data.id, sheetName: data.name, range: { startRow: 0, endRow: 4, startCol: 0, endCol: 2 } }
  const fieldId = stableFields ? 'region-source' : '0'
  const makeReport = (id, row) => ({ id, pivotCacheId: 'source-cache', source: structuredClone(source),
    anchor: { unitId, subUnitId: host.id, row, col: 1 },
    fields: [{ id: stableFields ? `${id}-region-instance` : '0', sourceFieldId: fieldId, sourceColumn: 0, name: 'Region', area: 'row', position: 0, selectedItems: ['North'] },
      { id: '1', sourceFieldId: '1', sourceColumn: 1, name: 'Amount', area: 'value', position: 0, aggregation: 'sum' }],
    options: { repeatRowLabels: false, showRowGrandTotal: true, showColGrandTotal: true, showRowSubtotals: true, showColSubtotals: true, valuePosition: 'column' },
    outputRange: { startRow: row, endRow: row + 3, startCol: 1, endCol: 2 }, generatedCellKeys: [], drillDownMap: {}, revision: 0 })
  const pivotTables = [makeReport('report-a', 8), makeReport('report-b', 15)]
  const chart = { id: 'owned-chart', sheetId: host.id, type: 'lineStacked',
    source: { sheetId: host.id, range: structuredClone(pivotTables[0].outputRange), orientation: 'columns' },
    position: { x: 360, y: 40 }, size: { width: 160, height: 100 }, legend: { visible: true, position: 'bottom' },
    categoryField: 0, valueFields: [1], xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, revision: 0 }
  if (owned) {
    chart.ownedPivot = makeReport('owned-report', 0)
    delete chart.ownedPivot.outputRange; delete chart.ownedPivot.generatedCellKeys; delete chart.ownedPivot.drillDownMap
  } else chart.pivotTableId = 'report-a'
  const connections = { pivotTables: pivotTables.map(pivot => ({ subUnitId: host.id, pivotTableId: pivot.id })),
    ...(owned ? { ownedPivotCharts: [{ subUnitId: host.id, pivotChartId: chart.id }] } : {}) }
  const slicers = [0, 1].map(index => ({ id: `view-${index + 1}`, unitId, hostSheetId: host.id, cacheId: 'control-cache',
    pivotTableId: 'report-a', sourceFieldId: fieldId, type: 'slicer', caption: 'Region', position: { row: 21, column: index ? 4 : 1 },
    size: { rows: 5, columns: 2 }, layout: { left: 44 + (index ? 320 : 80), top: 446, width: 120, height: 120 },
    drawingGeometry: { width: 120, height: 120, matrix: { a: 1, b: 0, c: 0, d: 1, e: index ? 320 : 80, f: 420 } },
    selectedItems: ['North'], style: 'SlicerStyleLight2', revision: 0 }))
  return { id: unitId, kind: 'workbook', name: 'Shared reports', sheets: [data, reports, paste], activeSheetId: host.id,
    pivotTables, charts: [chart], slicers, pivotCaches: [{ id: 'source-cache', source: structuredClone(source),
      fields: [{ id: fieldId, sourceColumn: 0, name: 'Region' }, { id: '1', sourceColumn: 1, name: 'Amount' }, { id: '2', sourceColumn: 2, name: 'Date' }], nextFieldId: 0, revision: 0 }],
    slicerCaches: [{ id: 'control-cache', unitId, type: 'slicer', source: { kind: 'pivot', pivotCacheId: 'source-cache', sourceFieldId: fieldId },
      connections, selectedItems: ['North'], revision: 0 }], date1904: false, modified: false }
}
