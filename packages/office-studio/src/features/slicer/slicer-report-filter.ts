import type { PivotRecordSelection, PivotTableModel, WorkbookModel } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { calculatePivotReport, getPivotSourceRecordSelection, PivotOutputConflictError, renderPivotTable,
  setPivotRecordSelection, setPivotSourceMemberFilter } from '../pivot/services/pivot-engine.ts'
import type { SlicerConnectionReport } from './slicer-connections.ts'
import { SlicerConnectionError } from './slicer-connections.ts'
import { normalizeSlicerRecordSelection, sameSlicerRecordSelection } from './slicer-record-selection.ts'
import { getPivotFieldFilterCategories, pivotFilterSourceKey } from '../pivot/services/pivot-filter-predicate.ts'

export function sameSlicerSelection(first: readonly string[] | undefined, second: readonly string[] | undefined): boolean {
  if (first === undefined || second === undefined) return first === second
  return first.length === second.length && first.every(value => second.includes(value))
}

/** Stage all output cells and owned reports before changing any supplied model.
 * Source cells are borrowed. Only changed reports and their output maps are copied. */
export function applySlicerReportFilter(workbook: WorkbookModel, reports: readonly SlicerConnectionReport[], items: string[] | PivotRecordSelection | null): boolean {
  const legacyItems = Array.isArray(items) ? items : undefined
  const selected = Array.isArray(items) ? undefined : normalizeSlicerRecordSelection(items)
  const changes = reports.filter(report => {
    if (legacyItems) return report.field.recordSelection !== undefined || report.field.groupedItems !== undefined
      || !sameSlicerSelection(report.field.selectedItems, legacyItems)
    const current = getPivotSourceRecordSelection(workbook, report.pivot, report.field.id)
    if (!sameSlicerRecordSelection(current, selected)) return true
    if (selected !== undefined) {
      const sourceId = pivotFilterSourceKey(report.pivot, report.field)
      const siblings = report.pivot.fields.filter(field => field.area !== 'value' && pivotFilterSourceKey(report.pivot, field) === sourceId)
      if (siblings.some(field => field.id !== report.field.id && getPivotFieldFilterCategories(field).has('member'))) return true
      if (report.pivot.options.allowMultipleFilters !== true && siblings.some(field =>
        [...getPivotFieldFilterCategories(field)].some(category => category !== 'member'))) return true
    }
    // Keep a grouped report's own display rules. Legacy strings must acquire raw identity.
    return selected !== undefined && report.field.recordSelection === undefined && report.field.groupedItems === undefined
  })
  if (!changes.length) return false
  const nextReports = new Map<PivotTableModel, PivotTableModel>()
  for (const report of changes) {
    const next = cloneModel(report.pivot)
    const field = next.fields.find(candidate => candidate.id === report.field.id && candidate.area !== 'value')
    if (!field) throw new Error('切片器报告字段已更改。')
    if (legacyItems) setPivotSourceMemberFilter(workbook, next, field.id, [...legacyItems], 'deferred')
    else setPivotRecordSelection(workbook, next, field.id, selected ?? null, 'deferred')
    nextReports.set(report.pivot, next)
  }
  const outputSheetIds = new Set(changes.filter(report => report.kind === 'pivotTable').map(report => report.subUnitId))
  const staged: WorkbookModel = { ...workbook,
    sheets: workbook.sheets.map(sheet => outputSheetIds.has(sheet.id) ? { ...sheet, cells: { ...sheet.cells } } : sheet),
    pivotTables: workbook.pivotTables?.map(pivot => nextReports.get(pivot) ?? pivot),
    charts: workbook.charts?.map(chart => chart.ownedPivot && nextReports.has(chart.ownedPivot)
      ? { ...chart, ownedPivot: nextReports.get(chart.ownedPivot) } : chart),
  }
  // Remove every previous output first so one report's old footprint cannot
  // incorrectly reject another report's new footprint. New overlaps still fail.
  for (const report of changes) if (report.kind === 'pivotTable') {
    const next = nextReports.get(report.pivot)!, sheet = staged.sheets.find(candidate => candidate.id === report.subUnitId)
    if (!sheet) throw new Error('切片器报告工作表不存在。')
    for (const key of next.generatedCellKeys ?? []) if (sheet.cells[key]?.metadata?.pivotTableId === next.id) delete sheet.cells[key]
    next.generatedCellKeys = []
    next.outputRange = undefined
  }
  try {
    for (const report of changes) {
      const next = nextReports.get(report.pivot)!
      if (report.kind === 'pivotTable') renderPivotTable(staged, next)
      else calculatePivotReport(staged, next)
    }
  } catch (error) {
    if (error instanceof PivotOutputConflictError) throw new SlicerConnectionError('invalid-target', error.message)
    throw error
  }
  for (const sheet of workbook.sheets) if (outputSheetIds.has(sheet.id)) sheet.cells = staged.sheets.find(candidate => candidate.id === sheet.id)!.cells
  for (const report of changes) {
    Object.assign(report.pivot, nextReports.get(report.pivot)!)
    if (report.chart) report.chart.revision += 1
  }
  return true
}
