import { createCell } from '../../src/domain/model.ts'
import { createPivotTable, setPivotField, setPivotOptions, setPivotRecordSelection } from '../../src/features/pivot/services/pivot-engine.ts'
import { parseTimelineDate } from '../../src/features/slicer/timeline-date.ts'
import { applyPivotChartEdit, getPivotChartReport } from '../../src/features/pivot/services/pivot-chart-engine.ts'
import { createChartDefinition } from '../../src/features/charts/chart-engine.ts'
import { initializeWorkbookDrawingStates } from '../../src/features/sheet-drawings/sheet-drawing-model-lifecycle.ts'
import { sheetDrawingRefAffectedRanges } from '../../src/features/sheet-drawings/sheet-drawing-host.ts'
import { pivotChartUIIds, pivotChartUIRows, pivotChartUIOptions, pivotChartUIWorkbook } from './pivot-chart-ui-fixture.mjs'

// Deferred real sparse source/report/chart fixture. It has not been executed.
export const pivotChartFieldButtonsIds = pivotChartUIIds
export function pivotChartFieldButtonsWorkbook({ overflow = false, grouped = false, manyMembers = false,
  partialDate = false, typedMembers = false, date1904 = false, dateSerials = false, compatibilityDate = false,
  readonly = false, hiddenSource = false, hiddenOutput = false, protectedOutput = false, protectedChart = false } = {}) {
  const ids = pivotChartUIIds, workbook = pivotChartUIWorkbook({ existing: false, readonly, hiddenSource })
  const source = workbook.sheets.find(sheet => sheet.id === ids.source), host = workbook.sheets.find(sheet => sheet.id === ids.host)
  const rows = (manyMembers ? [pivotChartUIRows[0], ...Array.from({ length: 130 }, (_, index) =>
    [`Region ${String(index).padStart(3, '0')}`, index + 1, '2026-01-15', index % 2 ? 'B' : 'A', 1])] : pivotChartUIRows).map(values => [...values])
  workbook.date1904 = date1904
  if (typedMembers) rows.slice(1).forEach((values, index) => { values[0] = [1, '1', 2, '2'][index] })
  if (dateSerials || compatibilityDate) rows.slice(1).forEach((values, index) => {
    values[2] = compatibilityDate ? [60, 61, 59.5, 60.25][index] : parseTimelineDate(values[2], date1904).serial
  })
  source.cells = {}
  for (const sheet of workbook.sheets) sheet.rowCount = 220
  rows.forEach((values, index) => values.forEach((value, offset) => {
    const row = index + 2, col = offset + 3; source.cells[`${row}:${col}`] = createCell(row, col, value)
    if (index && offset === 2 && (dateSerials || compatibilityDate)) source.cells[`${row}:${col}`].style.numberFormat = 'yyyy-mm-dd hh:mm:ss'
  }))
  const range = { startRow: 2, endRow: rows.length + 1, startColumn: 3, endColumn: 7 }
  const input = { unitId: ids.unit, subUnitId: ids.source, sheetName: 'Data', pivotCacheId: ids.cache, range }
  function createReport(id, sheetId) {
    const report = createPivotTable(workbook, input, { unitId: ids.unit, subUnitId: sheetId, row: 24, col: 0 }, id)
    for (const [fieldId, area] of [['0', 'row'], ['3', 'column'], ['2', 'filter'], ['1', 'value']]) setPivotField(workbook, report, fieldId, area)
    setPivotOptions(workbook, report, { ...pivotChartUIOptions }); return report
  }
  const primary = createReport(ids.primary, ids.host); createReport(ids.secondary, ids.other)
  const owned = structuredClone(primary)
  owned.id = 'ui-owned-report'; owned.anchor = { unitId: ids.unit, subUnitId: ids.host, row: 0, col: 0 }
  delete owned.outputRange; delete owned.generatedCellKeys; delete owned.drillDownMap; delete owned.axisOutput
  owned.fields = owned.fields.map(field => ({ ...field, sourceFieldId: field.sourceFieldId ?? String(field.sourceColumn - 3), id: `ui-owned-${field.sourceColumn}` }))
  const author = { type: 'column', categoryField: 0, valueFields: [1], size: { width: 480, height: 320 },
    options: { threeD: { view: { rotY: 42 } }, series: { 1: { fillOpacity: .35 } } } }
  const sourceRange = { startRow: range.startRow, endRow: range.endRow, startCol: range.startColumn, endCol: range.endColumn }
  workbook.charts = [
    createChartDefinition(workbook, ids.host, { ...structuredClone(author), title: 'Owned field controls', ownedPivot: owned,
      source: { sheetId: ids.source, range: sourceRange, orientation: 'columns' }, position: { x: 90, y: 40 } }, ids.owned),
    createChartDefinition(workbook, ids.host, { ...structuredClone(author), title: 'Linked field controls', pivotTableId: primary.id,
      source: { sheetId: ids.host, range: structuredClone(primary.outputRange), orientation: 'columns' }, position: { x: 600, y: 60 } }, ids.linked),
    createChartDefinition(workbook, ids.other, { ...structuredClone(author), title: 'Other host field controls', pivotTableId: primary.id,
      source: { sheetId: ids.host, range: structuredClone(primary.outputRange), orientation: 'columns' }, position: { x: 90, y: 40 } }, ids.foreign),
  ]
  for (const chartId of [ids.owned, ids.linked]) {
    const chart = workbook.charts.find(chart => chart.id === chartId), report = getPivotChartReport(workbook, chart)
    if (overflow) {
      const value = report.fields.find(field => field.area === 'value'), fields = structuredClone(report.fields)
      fields.push({ ...structuredClone(value), id: `${report.id}:value-average`, displayName: 'Amount average', aggregation: 'avg', position: 1 },
        { ...structuredClone(value), id: `${report.id}:value-count`, displayName: 'Amount count', aggregation: 'count', position: 2 })
      applyPivotChartEdit(workbook, chart, { type: 'setFieldsConfig', config: { fields, options: structuredClone(report.options) } })
    }
    if (grouped || partialDate) applyPivotChartEdit(workbook, chart, { type: 'setDateGroup', fieldId: report.fields.find(field => field.name === 'Date').id, dateType: 'year' })
    if (partialDate || typedMembers) {
      const field = report.fields.find(field => field.name === (partialDate ? 'Date' : 'Region'))
      setPivotRecordSelection(workbook, report, field.id, { mode: 'include', values: [rows[1][partialDate ? 2 : 0]] }, chart.ownedPivot ? 'data-only' : 'worksheet')
    }
  }
  const protect = (name, ranges, hidden = false) => {
    host.rangeProtectionRules = [...(host.rangeProtectionRules ?? []), { id: name, permissionId: name, description: '', ownerUserId: 'owner',
      allowedUserIds: [], editState: 'onlyMe', viewState: hidden ? 'noOneElseCanView' : 'othersCanView', ranges }]
  }
  if (hiddenOutput || protectedOutput) protect('field-ui-output', [structuredClone(primary.outputRange)], hiddenOutput)
  initializeWorkbookDrawingStates(workbook)
  if (protectedChart) protect('field-ui-chart', sheetDrawingRefAffectedRanges(workbook, ids.host, [{ kind: 'chart', id: ids.owned }]))
  return workbook
}
export function pivotChartFieldButtonsReport(workbook, chartId = pivotChartUIIds.owned) {
  return getPivotChartReport(workbook, workbook.charts.find(chart => chart.id === chartId))
}
export function pivotChartFieldButtonField(workbook, chartId, name, area) {
  const field = pivotChartFieldButtonsReport(workbook, chartId).fields.find(field => field.name === name && (!area || field.area === area))
  if (!field) throw new Error(`Missing fixture field ${name}/${area ?? '*'}`)
  return field
}
