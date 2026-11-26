import { FOffice, type FPivotTable } from '../../../api/facade.ts'
import type { OfficeStudio } from '../../../api/office-studio.ts'
import type { PivotTableModel, PivotTableOptions, WorkbookModel } from '../../../domain/model.ts'
import { cloneModel } from '../../../domain/serialization.ts'
import { parseSheetAddress } from '../../sheet-workbench/sheet-workbench.ts'
import { assertPivotReportEditPermission, assertPivotReportReadPermission } from './pivot-report-permissions.ts'
import { capturePivotChartPanelOwner, pivotChartPanelOwnerCurrent, type PivotChartPanelOwner } from './pivot-chart-panel-owner.ts'
import { getPivotFieldFilterCategories } from './pivot-filter-predicate.ts'
import type { PivotPanelDraft } from './pivot-panel-state.ts'
import type { PivotReportCommandContext } from './pivot-report-api-types.ts'
import { pivotChartFieldFilterEdit, type PivotChartFieldButton, type PivotChartFieldButtonAction,
  type PivotChartFieldFilterDraft } from './pivot-chart-field-button-state.ts'
import { pivotChartDateConditionFilter, pivotChartLabelConditionFilter, type PivotChartFieldConditionsDraft } from './pivot-chart-field-filter-condition.ts'

export interface PivotTableFieldFilterTarget { studio: OfficeStudio; unitId: string; sheetId: string; pivotId: string }
export interface PivotTableFieldFilterSnapshot {
  owner: PivotChartPanelOwner
  field: PivotChartFieldButton
  editable: boolean
  options: PivotTableOptions
}
export type PivotTableFieldFilterAction = Exclude<PivotChartFieldButtonAction, 'remove'>

export function pivotTableFieldFilterFacade(target: PivotTableFieldFilterTarget): FPivotTable {
  const pivot = FOffice.newAPI(target.studio).getWorkbook(target.unitId)?.getPivotTables()
    .find(candidate => candidate.getPivotTableId() === target.pivotId)
  if (!pivot) throw new Error('数据透视表已被删除或工作簿已关闭。')
  return pivot
}

/** Compare the authored source and local field identity; a pending destination/layout draft stays in the panel. */
export function pivotTableFieldFilterDraftCompatible(workbook: WorkbookModel, draft: PivotPanelDraft,
  pivotId: string, fieldId: string): boolean {
  const report = workbook.pivotTables?.find(report => report.id === pivotId), field = report?.fields.find(field => field.id === fieldId)
  const local = draft.fields.find(field => field.fieldId === fieldId)
  if (!report || !field || !local || draft.pivotId !== pivotId || local.area === 'unused' || local.area === 'value'
    || local.area !== field.area || local.sourceColumn !== field.sourceColumn || draft.sourceSheetId !== report.source.subUnitId) return false
  const parsed = parseSheetAddress(draft.sourceRange, { rowCount: 1_048_576, columnCount: 16_384 })?.range
  const range = report.source.range
  return Boolean(parsed && /^\$?[A-Z]+\$?[1-9]\d*(?::\$?[A-Z]+\$?[1-9]\d*)?$/i.test(draft.sourceRange.trim())
    && parsed.startRow === range.startRow && parsed.endRow === range.endRow
    && parsed.startColumn === range.startCol && parsed.endColumn === range.endCol)
}

export function pivotTableFieldFilterPanelSignature(draft: PivotPanelDraft): string {
  return JSON.stringify([draft.pivotId, draft.sourceSheetId, draft.sourceRange, draft.targetSheetId, draft.targetAddress, draft.fields, draft.options])
}

export function readPivotTableFieldFilterSnapshot(target: PivotTableFieldFilterTarget, fieldId: string): PivotTableFieldFilterSnapshot {
  const owner = capturePivotChartPanelOwner(target.studio, target.unitId, target.sheetId)
  const workbook = target.studio.getWorkbookModel(target.unitId), report = workbook?.pivotTables?.find(report => report.id === target.pivotId)
  if (!workbook || !report) throw new Error('数据透视表已被删除。')
  assertPivotReportReadPermission(workbook, report, owner.userId)
  const field = report.fields.find(field => field.id === fieldId)
  if (!field || field.area === 'value') throw new Error('当前字段不提供成员筛选。')
  let editable = true
  try { assertPivotReportEditPermission(workbook, report, owner.userId) } catch { editable = false }
  const result = { owner, editable, options: cloneModel(pivotTableFieldFilterFacade(target).getOptions()),
    field: { fieldId: field.id, label: field.displayName ?? field.name, area: field.area,
      filtered: getPivotFieldFilterCategories(field).size > 0, field: cloneModel(field) } }
  if (!pivotChartPanelOwnerCurrent(owner)) throw new Error('透视字段读取已过期。')
  return result
}

/** The same public typed setters reach the report transaction; UI never writes the report or output cells. */
export async function applyPivotTableFieldFilterAction(pivot: FPivotTable, fieldId: string,
  action: PivotTableFieldFilterAction, context: PivotReportCommandContext, draft?: PivotChartFieldFilterDraft,
  conditions?: PivotChartFieldConditionsDraft): Promise<void> {
  switch (action) {
    case 'applyFilter': {
      if (!draft) throw new Error('筛选成员尚未读取。')
      const edit = pivotChartFieldFilterEdit(fieldId, draft)
      await pivot.setLabelManualFilter(fieldId, edit.items, edit.isAll, context, edit.partialItems); return
    }
    case 'clearFilter': await pivot.clearFilters(fieldId, context); return
    case 'sortAscending': await pivot.setLabelSort(fieldId, { type: 'ascending' }, context); return
    case 'sortDescending': await pivot.setLabelSort(fieldId, { type: 'descending' }, context); return
    case 'clearSort': await pivot.clearLabelSort(fieldId, context); return
    case 'applyDateFilter':
      if (!conditions) throw new Error('日期筛选条件尚未读取。')
      await pivot.setDateFilter(fieldId, pivotChartDateConditionFilter(conditions), context); return
    case 'clearDateFilter': await pivot.setDateFilter(fieldId, null, context); return
    case 'applyLabelCondition':
      if (!conditions) throw new Error('标签筛选条件尚未读取。')
      await pivot.setLabelConditionFilter(fieldId, pivotChartLabelConditionFilter(conditions), context); return
    case 'clearLabelCondition': await pivot.setLabelConditionFilter(fieldId, null, context); return
  }
}

export function setPivotTablePanelMultipleFilters(draft: PivotPanelDraft, enabled: boolean): void {
  if (typeof enabled !== 'boolean') throw new TypeError('允许多个筛选的开关值无效。')
  draft.options = { ...draft.options, allowMultipleFilters: enabled }
}

/** Merge only the committed sort into the existing panel draft, keeping source, target, fields and options. */
export function synchronizePivotTableFieldFilterDraft(draft: PivotPanelDraft, report: PivotTableModel,
  fieldId: string, action: PivotTableFieldFilterAction): void {
  if (!['sortAscending', 'sortDescending', 'clearSort'].includes(action)) return
  const local = draft.fields.find(field => field.fieldId === fieldId), field = report.fields.find(field => field.id === fieldId)
  if (local && field) { if (field.sort) local.sort = field.sort; else delete local.sort }
}
