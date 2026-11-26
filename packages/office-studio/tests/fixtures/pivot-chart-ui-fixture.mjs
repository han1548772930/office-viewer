import { createBlankWorkbook, createBlankSheet, createCell } from '../../src/domain/model.ts'
import { createPivotTable, setPivotField, setPivotOptions } from '../../src/features/pivot/services/pivot-engine.ts'
import { createChartDefinition } from '../../src/features/charts/chart-engine.ts'
import { initializeWorkbookDrawingStates } from '../../src/features/sheet-drawings/sheet-drawing-model-lifecycle.ts'
import { sheetDrawingRefAffectedRanges } from '../../src/features/sheet-drawings/sheet-drawing-host.ts'

// Real sparse cells and real report/chart models. This deferred fixture has not been executed.
export const pivotChartUIIds = Object.freeze({ unit: 'pivot-chart-ui-book', source: 'data', host: 'reports', other: 'dashboard',
  cache: 'pivot-chart-ui-cache', primary: 'ui-primary-report', secondary: 'ui-secondary-report',
  owned: 'ui-owned-chart', linked: 'ui-linked-chart', foreign: 'ui-foreign-chart' })
export const pivotChartUIRange = Object.freeze({ startRow: 2, endRow: 6, startCol: 3, endCol: 7 })
export const pivotChartUIOptions = Object.freeze({ repeatRowLabels: false, showRowGrandTotal: false, showColGrandTotal: false,
  showRowSubtotals: false, showColSubtotals: false, valuePosition: 'column' })
export const pivotChartUIRows = Object.freeze([
  ['Region', 'Amount', 'Date', 'Segment', 'Score'], ['North', 10, '2026-01-15', 'A', 5],
  ['South', 20, '2026-02-02', 'A', 15], ['North', 30, '2026-02-20', 'B', 25], ['South', 40, '2026-03-01', 'B', 35],
])

export function pivotChartUIWorkbook({ id = pivotChartUIIds.unit, existing = true, readonly = false,
  hiddenSource = false, hiddenOutput = false, protectedOutput = false, protectedChart = false } = {}) {
  const ids = pivotChartUIIds, workbook = createBlankWorkbook(), source = workbook.sheets[0]
  const reports = createBlankSheet('Reports'), dashboard = createBlankSheet('Dashboard')
  source.id = ids.source; source.name = 'Data'; reports.id = ids.host; dashboard.id = ids.other
  workbook.id = id; workbook.name = 'Pivot chart UI'; workbook.calculationMode = 'manual'
  workbook.sheets = [source, reports, dashboard]; workbook.activeSheetId = ids.host
  workbook.charts = []; workbook.pivotTables = []; workbook.pivotCaches = []; workbook.slicers = []; workbook.slicerCaches = []
  for (const sheet of workbook.sheets) {
    sheet.rowCount = 120; sheet.columnCount = 30; sheet.defaultColumnWidth = 80; sheet.defaultRowHeight = 20; sheet.defaultRowAutoHeight = false
  }
  pivotChartUIRows.forEach((values, index) => values.forEach((value, column) => {
    const row = index + pivotChartUIRange.startRow, col = column + pivotChartUIRange.startCol
    source.cells[`${row}:${col}`] = createCell(row, col, value)
  }))
  const sourceInput = { unitId: id, subUnitId: source.id, sheetName: source.name, pivotCacheId: ids.cache,
    range: { startRow: 2, endRow: 6, startColumn: 3, endColumn: 7 } }
  let primary
  if (existing) {
    function report(reportId, hostSheetId, row) {
      const value = createPivotTable(workbook, sourceInput, { unitId: id, subUnitId: hostSheetId, row, col: 0 }, reportId)
      setPivotField(workbook, value, '0', 'row'); setPivotField(workbook, value, '2', 'filter'); setPivotField(workbook, value, '1', 'value')
      setPivotOptions(workbook, value, { ...pivotChartUIOptions }); return value
    }
    primary = report(ids.primary, ids.host, 24); report(ids.secondary, ids.other, 24)
    const owned = structuredClone(primary)
    owned.id = 'ui-owned-report'; owned.anchor = { unitId: id, subUnitId: ids.host, row: 0, col: 0 }
    delete owned.outputRange; delete owned.generatedCellKeys; delete owned.drillDownMap
    owned.fields = owned.fields.map(field => ({ ...field, sourceFieldId: field.sourceFieldId ?? String(field.sourceColumn - 3), id: `ui-owned-${field.sourceColumn}` }))
    const author = { type: 'column', categoryField: 0, valueFields: [1], position: { x: 500.5, y: 40.25 }, size: { width: 360, height: 240 },
      options: { threeD: { view: { rotY: 42 } }, series: { 1: { fillOpacity: .35 } } } }
    workbook.charts = [
      createChartDefinition(workbook, ids.host, { ...structuredClone(author), title: 'Owned chart', ownedPivot: owned,
        source: { sheetId: ids.source, range: { ...pivotChartUIRange }, orientation: 'columns' } }, ids.owned),
      createChartDefinition(workbook, ids.host, { ...structuredClone(author), title: 'Linked chart', pivotTableId: primary.id,
        source: { sheetId: ids.host, range: structuredClone(primary.outputRange), orientation: 'columns' }, position: { x: 880, y: 60 } }, ids.linked),
      createChartDefinition(workbook, ids.other, { ...structuredClone(author), title: 'Other sheet chart', pivotTableId: primary.id,
        source: { sheetId: ids.host, range: structuredClone(primary.outputRange), orientation: 'columns' } }, ids.foreign),
    ]
  }
  if (readonly) reports.worksheetProtection = { id: 'ui-readonly', permissionId: 'ui-readonly', unitType: 'worksheet', ownerUserId: 'owner',
    allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView' }
  const protect = (sheet, name, ranges, hidden) => {
    sheet.rangeProtectionRules = [...(sheet.rangeProtectionRules ?? []), { id: name, permissionId: name, description: '', ownerUserId: 'owner',
      allowedUserIds: [], editState: 'onlyMe', viewState: hidden ? 'noOneElseCanView' : 'othersCanView', ranges }]
  }
  if (hiddenSource) protect(source, 'ui-hidden-source', [{ ...pivotChartUIRange }], true)
  if (primary && (hiddenOutput || protectedOutput)) protect(reports, 'ui-output-rule', [structuredClone(primary.outputRange)], hiddenOutput)
  initializeWorkbookDrawingStates(workbook)
  if (existing && protectedChart) protect(reports, 'ui-chart-rule', sheetDrawingRefAffectedRanges(workbook, ids.host, [{ kind: 'chart', id: ids.owned }]), false)
  return workbook
}

export function pivotChartUIReport(workbook, id = pivotChartUIIds.owned) {
  const chart = workbook.charts?.find(value => value.id === id)
  return chart?.ownedPivot ?? workbook.pivotTables?.find(value => value.id === chart?.pivotTableId)
}
