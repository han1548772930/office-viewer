import { createCell } from '../../src/domain/model.ts'
import { renderPivotTable, calculatePivotReport } from '../../src/features/pivot/services/pivot-engine.ts'
import { createSlicer, setSlicerConnections } from '../../src/features/slicer/slicer-engine.ts'
import { getSlicerCache } from '../../src/features/slicer/slicer-connections.ts'
import { timelineCalendarToSerial } from '../../src/features/slicer/timeline-date.ts'
import { pivotFilterWorkbook, filterReports } from './pivot-filter-fixture.mjs'
import { customTimelineStyle } from './slicer-style-fixture.mjs'

/** Eleven source records, including colliding raw numbers/text/errors and a cached typed Date. Never executed. */
export function timelineFilterWorkbook({ date1904 = false, grouped = false } = {}) {
  const workbook = pivotFilterWorkbook({ date1904 }), source = workbook.sheets.find(sheet => sheet.id === 'data')
  const first = timelineCalendarToSerial('2026-01-01', date1904), date = first + .5
  const extra = [['North', date, 1000, 1, 8], ['North', '2026-01-01T12:00:00', 2000, 1, 9],
    ['North', date, 80, 1, 10], ['North', '#N/A', 3000, 1, 11]]
  extra.forEach((values, index) => values.forEach((value, offset) => {
    const row = index + 10, column = offset + 3, cell = createCell(row, column, value)
    if (row === 12 && column === 4) { cell.formula = '=DATE(2026,1,1)+0.5'; cell.cachedDataType = 'date' }
    if (row === 13 && column === 4) { cell.formula = '=NA()'; cell.cachedDataType = 'error'; cell.metadata = { error: '#N/A' } }
    source.cells[`${row}:${column}`] = cell
  }))
  for (const report of filterReports(workbook)) report.source.range.endRow = 13
  for (const cache of workbook.pivotCaches) cache.source.range.endRow = 13
  const primary = workbook.pivotTables[0], secondary = structuredClone(primary)
  secondary.id = 'timeline-secondary'; secondary.anchor.row = 70
  secondary.fields = secondary.fields.map(field => ({ ...field, id: `secondary-${field.id}` }))
  secondary.generatedCellKeys = []; delete secondary.outputRange; delete secondary.drillDownMap; delete secondary.axisOutput
  workbook.pivotTables.push(secondary)
  const owned = workbook.charts.find(chart => chart.ownedPivot).ownedPivot
  owned.fields = owned.fields.map(field => ({ ...field, id: `owned-${field.id}` }))
  if (grouped) for (const report of filterReports(workbook)) timelineReportField(report).dateGroupType = 'month'
  for (const report of workbook.pivotTables) renderPivotTable(workbook, report)
  const input = { unitId: workbook.id, hostSheetId: 'dashboard', pivotTableId: primary.id, sourceFieldId: 'date-source', type: 'timeline' }
  const timeline = createSlicer(workbook, { ...input, id: 'timeline-primary', name: 'Timeline A', level: 'year', selectionLevel: 'day', scrollPosition: first + .25 })
  createSlicer(workbook, { ...input, id: 'timeline-copy', cacheId: timeline.cacheId, name: 'Timeline B', level: 'quarter', selectionLevel: 'month', scrollPosition: first - .75 })
  createSlicer(workbook, { id: 'timeline-mixed-slicer', unitId: workbook.id, hostSheetId: 'dashboard', pivotTableId: primary.id,
    sourceFieldId: 'mixed-source', type: 'slicer', name: 'Mixed controls' })
  setSlicerConnections(workbook, timeline, timelineAllConnections(workbook))
  workbook.timelineStyles = { defaultTimelineStyle: 'Calendar bands', styles: [customTimelineStyle()] }
  return workbook
}

export function timelineAllConnections(workbook) {
  return { pivotTables: workbook.pivotTables.map(report => ({ subUnitId: report.anchor.subUnitId, pivotTableId: report.id })),
    ownedPivotCharts: [{ subUnitId: 'dashboard', pivotChartId: 'owned-chart' }] }
}
export function timelineViews(workbook) { return workbook.slicers.filter(view => view.type === 'timeline') }
export function timelineReportField(report) { return report.fields.find(field => field.area !== 'value' && field.sourceFieldId === 'date-source') }
export function timelineCache(workbook) { return getSlicerCache(workbook, timelineViews(workbook)[0]) }
export function timelineFirst(workbook) { return timelineCalendarToSerial('2026-01-01', workbook.date1904) }
export function timelineReportTotals(workbook) {
  return filterReports(workbook).map(report => {
    const result = calculatePivotReport(workbook, report)
    return result.rowAxis.filter(entry => entry.kind === 'detail' || entry.kind === 'collapsed')
      .reduce((sum, entry) => sum + Number(result.values[entry.index].at(-1) ?? 0), 0)
  })
}
export { filterReports }
