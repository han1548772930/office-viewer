import type { OfficeStudio } from './office-studio.ts'
import type { PivotChartHierarchyTarget } from '../features/pivot/services/pivot-hierarchy-api-types.ts'
import type { PivotAxisFieldState } from '../features/pivot/services/pivot-axis-state.ts'
import { cloneModel } from '../domain/serialization.ts'
import { getPivotChartReport } from '../features/pivot/services/pivot-chart-engine.ts'

export interface PivotChartAxisSelection extends PivotChartHierarchyTarget {
  subUnitId: string
  fieldId: string
  item: string
  collapsed: boolean
  fields: PivotAxisFieldState[]
  expectedRevision: number
  expectedHistoryGeneration: number
  expectedUserId: string
  host: 'workbench'
  chartSnapshot: string
  reportSnapshot: string
}

const selections = new WeakMap<OfficeStudio, { owner: object; selection: PivotChartAxisSelection }>()

/** Frame selection is ephemeral and cannot create a second content history. */
export function setPivotChartAxisSelection(studio: OfficeStudio, owner: object, selection: PivotChartAxisSelection): void {
  // Vue's view selection also carries a Studio owner and native frame target.
  // Copy only the control DTO; cloning those runtime handles is invalid.
  const value: PivotChartAxisSelection = { unitId: selection.unitId, subUnitId: selection.subUnitId, pivotChartId: selection.pivotChartId,
    fieldId: selection.fieldId, item: selection.item, collapsed: selection.collapsed, fields: cloneModel(selection.fields),
    expectedRevision: selection.expectedRevision, expectedHistoryGeneration: selection.expectedHistoryGeneration,
    expectedUserId: selection.expectedUserId, host: selection.host, chartSnapshot: selection.chartSnapshot, reportSnapshot: selection.reportSnapshot }
  selections.set(studio, { owner, selection: value })
}

export function clearPivotChartAxisSelection(studio: OfficeStudio, owner: object): void {
  if (selections.get(studio)?.owner === owner) selections.delete(studio)
}

export function getPivotChartAxisSelection(studio: OfficeStudio): PivotChartAxisSelection | undefined {
  const entry = selections.get(studio), input = entry?.selection
  if (!input) return undefined
  const workbook = studio.getWorkbookModel(input.unitId), chart = workbook?.charts?.find(chart => chart.id === input.pivotChartId && chart.sheetId === input.subUnitId)
  const current = !studio.isDisposed && !!workbook && !!chart && studio.activeWorkbookId === input.unitId
    && studio.editor.state.activeKind === 'workbook' && workbook.activeSheetId === input.subUnitId
    && studio.getWorkbookRevision(input.unitId) === input.expectedRevision && studio.currentUserId === input.expectedUserId
    && studio.editor.history.captureSavepoint(input.unitId).generation === input.expectedHistoryGeneration
    && JSON.stringify(chart) === input.chartSnapshot
  if (!current) { selections.delete(studio); return undefined }
  try {
    const report = getPivotChartReport(workbook!, chart!)
    if (JSON.stringify(report) !== input.reportSnapshot) { selections.delete(studio); return undefined }
    const fields = report.fields.filter(field => field.area === 'row').sort((a, b) => a.position - b.position)
    if (!fields.slice(0, -1).some(field => field.id === input.fieldId)) { selections.delete(studio); return undefined }
  } catch { selections.delete(studio); return undefined }
  return cloneModel(input)
}
