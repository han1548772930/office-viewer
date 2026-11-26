import { createBlankWorkbook, createBlankSheet, createCell } from '../../src/domain/model.ts'
import { createSheetTable } from '../../src/features/sheet-tables/sheet-tables.ts'
import { createPivotTable, setPivotField } from '../../src/features/pivot/services/pivot-engine.ts'
import { createChartDefinition } from '../../src/features/charts/chart-engine.ts'
import { createSlicer, setSlicerItems, setSlicerTimelineRange } from '../../src/features/slicer/slicer-engine.ts'
import { slicerWithDrawingGeometry } from '../../src/features/slicer/slicer-drawing.ts'
import { initializeWorkbookDrawingStates } from '../../src/features/sheet-drawings/sheet-drawing-model-lifecycle.ts'
import { GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH } from '../../src/rendering/spreadsheet/layout.ts'

// Real source, report and control models. This fixture has not been executed.
export const connectionFixtureIds = Object.freeze({
  unit: 'slicer-connections-book', source: 'source', host: 'reports', other: 'other',
  cache: 'shared-pivot-cache', primary: 'primary-report', secondary: 'secondary-report',
  incompatible: 'other-cache-report', owned: 'owned-chart', linked: 'linked-chart',
  slicer: 'slicer-a', shared: 'slicer-b', timeline: 'timeline-a', table: 'table-slicer',
})

export function slicerConnectionsWorkbook({ id = connectionFixtureIds.unit, filtered = false, affine = true, readonly = false } = {}) {
  const ids = connectionFixtureIds
  let workbook = createBlankWorkbook()
  workbook.id = id; workbook.name = 'Slicer report connections'; workbook.calculationMode = 'manual'
  const source = workbook.sheets[0], reports = createBlankSheet('Reports'), other = createBlankSheet('Other reports')
  source.id = ids.source; source.name = 'Source'; reports.id = ids.host; other.id = ids.other
  workbook.sheets = [source, reports, other]; workbook.activeSheetId = reports.id
  for (const sheet of workbook.sheets) {
    sheet.rowCount = 120; sheet.columnCount = 30; sheet.defaultRowHeight = 20; sheet.defaultColumnWidth = 80; sheet.defaultRowAutoHeight = false
  }
  const rows = [
    ['Region', 'Date', 'Amount'], ['North', '2026-01-15', 10], ['South', '2026-02-02', 20],
    ['North', '2026-02-20', 30], ['South', '2026-03-01', 40],
  ]
  rows.forEach((values, row) => values.forEach((value, column) => {
    source.cells[`${row}:${column}`] = createCell(row, column, value)
    reports.cells[`${row}:${column + 8}`] = createCell(row, column + 8, value)
  }))
  workbook = createSheetTable(workbook, { id: 'connection-table', sheetId: ids.host, name: 'LocalTable',
    range: { startRow: 0, endRow: 4, startCol: 8, endCol: 10 } })
  const sourceInput = { unitId: id, subUnitId: ids.source, sheetName: 'Source', pivotCacheId: ids.cache,
    range: { startRow: 0, endRow: 4, startColumn: 0, endColumn: 2 } }
  function report(reportId, hostSheetId, row, pivotCacheId = ids.cache) {
    const pivot = createPivotTable(workbook, { ...sourceInput, pivotCacheId }, { unitId: id, subUnitId: hostSheetId, row, col: 0 }, reportId)
    setPivotField(workbook, pivot, '0', 'row'); setPivotField(workbook, pivot, '1', 'filter'); setPivotField(workbook, pivot, '2', 'value')
    return pivot
  }
  const primary = report(ids.primary, ids.host, 24)
  report(ids.secondary, ids.other, 14)
  report(ids.incompatible, ids.host, 45, 'independent-pivot-cache')
  const ownedReport = structuredClone(primary)
  ownedReport.id = 'owned-report'; ownedReport.anchor = { unitId: id, subUnitId: ids.host, row: 0, col: 14 }
  delete ownedReport.outputRange; delete ownedReport.generatedCellKeys; delete ownedReport.drillDownMap
  // Report-local IDs differ while canonical source identities stay 0 / 1 / 2.
  ownedReport.fields = ownedReport.fields.map(field => ({ ...field, sourceFieldId: field.sourceFieldId ?? String(field.sourceColumn), id: `owned-field-${field.sourceColumn}` }))
  const chartSource = { sheetId: ids.host, range: structuredClone(primary.outputRange), orientation: 'columns' }
  workbook.charts = [
    createChartDefinition(workbook, ids.host, { type: 'line', source: chartSource, ownedPivot: ownedReport, title: 'Owned chart',
      position: { x: GRID_HEADER_WIDTH + 740, y: GRID_HEADER_HEIGHT + 80 }, size: { width: 240, height: 150 }, valueFields: [1] }, ids.owned),
    createChartDefinition(workbook, ids.host, { type: 'line', source: chartSource, pivotTableId: primary.id, title: 'Linked chart',
      position: { x: GRID_HEADER_WIDTH + 740, y: GRID_HEADER_HEIGHT + 260 }, size: { width: 240, height: 150 }, valueFields: [1] }, ids.linked),
  ]
  const layouts = [
    [ids.slicer, 'slicer', 80, 70, 180, 180], [ids.shared, 'slicer', 320, 70, 180, 180],
    [ids.timeline, 'timeline', 80, 310, 360, 130], [ids.table, 'slicer', 510, 310, 180, 170],
  ]
  for (const [controlId, type, left, top, width, height] of layouts) {
    const table = workbook.tables.find(table => table.id === 'connection-table')
    const control = createSlicer(workbook, { id: controlId, unitId: id, hostSheetId: ids.host, type,
      ...(controlId === ids.table ? { tableId: table.id, sourceFieldId: table.columns[0].id } : { pivotTableId: primary.id, sourceFieldId: type === 'timeline' ? '1' : '0' }),
      layout: { left: GRID_HEADER_WIDTH + left, top: GRID_HEADER_HEIGHT + top, width, height } })
    if (affine) Object.assign(control, slicerWithDrawingGeometry(workbook, control,
      { width, height, matrix: { a: 1, b: .2, c: .15, d: 1, e: left, f: top } }))
  }
  if (filtered) {
    setSlicerItems(workbook, workbook.slicers.find(control => control.id === ids.slicer), ['North'])
    setSlicerTimelineRange(workbook, workbook.slicers.find(control => control.id === ids.timeline), { start: '2026-02-01', end: '2026-02-28', level: 'month' })
  }
  if (readonly) workbook.sheets.find(sheet => sheet.id === ids.host).worksheetProtection = {
    id: 'readonly-host', ownerUserId: 'owner', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView',
  }
  initializeWorkbookDrawingStates(workbook)
  return workbook
}

export function connectionReportSelection(workbook, reportId, fieldId = '0') {
  const report = workbook.pivotTables?.find(pivot => pivot.id === reportId)
    ?? workbook.charts?.find(chart => chart.id === reportId)?.ownedPivot
  const field = report?.fields.find(field => (field.sourceFieldId ?? String(field.sourceColumn - report.source.range.startCol)) === fieldId)
  return field?.selectedItems ? [...field.selectedItems] : undefined
}
