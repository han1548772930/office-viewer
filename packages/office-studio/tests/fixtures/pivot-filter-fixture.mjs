import { createBlankSheet, createCell } from '../../src/domain/model.ts'
import { renderPivotTable } from '../../src/features/pivot/services/pivot-engine.ts'
import { timelineCalendarToSerial } from '../../src/features/slicer/timeline-date.ts'
import { createSlicer, setSlicerConnections } from '../../src/features/slicer/slicer-engine.ts'
import { createSheetTable } from '../../src/features/sheet-tables/sheet-tables.ts'

export const pivotFilterOptions = { repeatRowLabels: false, showRowGrandTotal: false, showColGrandTotal: false,
  showRowSubtotals: false, showColSubtotals: false, valuePosition: 'column' }
export const filterReports = workbook => [...workbook.pivotTables, ...workbook.charts.flatMap(chart => chart.ownedPivot ? [chart.ownedPivot] : [])]

/** Real sparse records with formatted serial dates, invalid dates and colliding scalar display keys. */
export function pivotFilterWorkbook({ date1904 = false } = {}) {
  const unitId = 'pivot-filter-unit', sheets = ['Data', 'Output', 'Dashboard'].map(name => {
    const sheet = createBlankSheet(name); sheet.id = name.toLowerCase(); sheet.rowCount = 180; sheet.columnCount = 30; return sheet
  })
  const first = timelineCalendarToSerial('2026-01-01', date1904), second = timelineCalendarToSerial('2026-01-02', date1904)
  const february = timelineCalendarToSerial('2026-02-01', date1904)
  const rows = [['Region', 'Date', 'Amount', 'Mixed', 'Date count'],
    ['North', first + 0.25, 10, 1, 1], ['North', first + 0.75, 20, '1', 2],
    ['north', second, 30, true, 3], ['South', second + 0.5, 40, 'true', 4],
    ['West', february, 50, null, 5], ['North', february + 0.5, 60, '', 6],
    ['South', 'not a date', 70, 'beta', 7]]
  rows.forEach((row, index) => row.forEach((value, column) => {
    const r = index + 2, c = column + 3, cell = createCell(r, c, value)
    if (index > 0 && column === 1) cell.style.numberFormat = 'yyyy-mm-dd hh:mm:ss'
    sheets[0].cells[`${r}:${c}`] = cell
  }))
  const source = { unitId, subUnitId: 'data', sheetName: 'Data', pivotCacheId: 'filter-cache',
    range: { startRow: 2, endRow: 9, startCol: 3, endCol: 7 } }
  const fields = () => [
    { id: 'region-local', sourceFieldId: 'region-source', sourceColumn: 3, name: 'Region', area: 'row', position: 0 },
    { id: 'date-local', sourceFieldId: 'date-source', sourceColumn: 4, name: 'Date', area: 'row', position: 1 },
    { id: 'amount-local', sourceFieldId: 'amount-source', sourceColumn: 5, name: 'Amount', area: 'value', position: 0, aggregation: 'sum' },
    { id: 'mixed-local', sourceFieldId: 'mixed-source', sourceColumn: 6, name: 'Mixed', area: 'filter', position: 0 },
    { id: 'number-local', sourceFieldId: 'number-source', sourceColumn: 7, name: 'Date count', area: 'filter', position: 1 },
  ]
  const report = (id, host, row) => ({ id, pivotCacheId: 'filter-cache', source: structuredClone(source),
    anchor: { unitId, subUnitId: host, row, col: 1 }, fields: fields(), options: structuredClone(pivotFilterOptions),
    generatedCellKeys: [], drillDownMap: {}, revision: 0 })
  const ordinary = report('ordinary-report', 'output', 12), owned = report('owned-report', 'dashboard', 0)
  delete owned.generatedCellKeys; delete owned.drillDownMap
  const chart = (id, extra) => ({ id, sheetId: 'dashboard', type: 'column',
    source: { sheetId: 'data', range: structuredClone(source.range), orientation: 'columns' },
    position: { x: 900, y: 100 }, size: { width: 240, height: 180 }, categoryField: 0, valueFields: [2],
    legend: { visible: true, position: 'bottom' }, xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, revision: 0, ...extra })
  const workbook = { id: unitId, kind: 'workbook', name: 'Pivot filters', sheets, activeSheetId: 'output', date1904,
    modified: false, pivotTables: [ordinary], charts: [chart('linked-chart', { pivotTableId: ordinary.id }), chart('owned-chart', { ownedPivot: owned })],
    pivotCaches: [{ id: 'filter-cache', source: structuredClone(source), revision: 0, fields: fields().map(field => ({
      id: field.sourceFieldId, sourceColumn: field.sourceColumn, name: field.name })), nextFieldId: 0 }],
    slicers: [], slicerCaches: [], tables: [] }
  renderPivotTable(workbook, ordinary)
  workbook.charts[0].source = { ...workbook.charts[0].source, sheetId: 'output', range: structuredClone(ordinary.outputRange) }
  return workbook
}

/** Shared raw-cache fixture includes distinct local IDs, two ordinary reports, linked and owned charts, and two views. */
export function slicerRecordSelectionWorkbook() {
  const workbook = pivotFilterWorkbook(), ordinary = workbook.pivotTables[0], secondary = structuredClone(ordinary)
  secondary.id = 'secondary-report'; secondary.anchor.row = 55
  secondary.fields = secondary.fields.map(field => ({ ...field, id: `secondary-${field.id}` }))
  secondary.generatedCellKeys = []; delete secondary.outputRange; delete secondary.drillDownMap; delete secondary.axisOutput
  workbook.pivotTables.push(secondary)
  renderPivotTable(workbook, secondary)
  const owned = workbook.charts.find(chart => chart.ownedPivot).ownedPivot
  owned.fields = owned.fields.map(field => ({ ...field, id: `owned-${field.id}` }))
  const input = { unitId: workbook.id, hostSheetId: 'dashboard', pivotTableId: ordinary.id,
    sourceFieldId: 'mixed-source', type: 'slicer' }
  const control = createSlicer(workbook, { ...input, id: 'typed-slicer' })
  createSlicer(workbook, { ...input, id: 'typed-slicer-copy', cacheId: control.cacheId })
  setSlicerConnections(workbook, control, { pivotTables: [{ subUnitId: 'output', pivotTableId: secondary.id }],
    ownedPivotCharts: [{ subUnitId: 'dashboard', pivotChartId: 'linked-chart' }, { subUnitId: 'dashboard', pivotChartId: 'owned-chart' }] })
  return workbook
}

/** Native Table filtering uses the same seven real raw values and two cache views. */
export function slicerTableRecordSelectionWorkbook() {
  const workbook = createSheetTable(pivotFilterWorkbook(), { id: 'typed-table', name: 'TypedTable', sheetId: 'data',
    range: { startRow: 2, endRow: 9, startCol: 3, endCol: 7 } })
  const table = workbook.tables.find(table => table.id === 'typed-table'), input = { unitId: workbook.id, hostSheetId: 'dashboard',
    tableId: table.id, sourceFieldId: table.columns[3].id, type: 'slicer' }
  const control = createSlicer(workbook, { ...input, id: 'typed-table-slicer' })
  createSlicer(workbook, { ...input, id: 'typed-table-slicer-copy', cacheId: control.cacheId })
  return workbook
}
