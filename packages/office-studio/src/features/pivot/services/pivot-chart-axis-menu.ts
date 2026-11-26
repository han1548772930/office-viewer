import type { PivotChartAxisSelection } from '../../../api/pivot-chart-axis-selection.ts'
import { PIVOT_HIERARCHY_COMMANDS, type PivotChartHierarchyTarget } from './pivot-hierarchy-api-types.ts'

export interface PivotChartAxisMenuEntry {
  key: string
  label: string
  commandId: string
  params?: PivotChartHierarchyTarget & { fieldId: string; collapse?: boolean }
  disabled: boolean
}

/** Fixed v1.6.0 or() selector: four static children, then one child per row field. */
export function pivotChartAxisMenu(selection: PivotChartAxisSelection, disabled = false): PivotChartAxisMenuEntry[] {
  const fields = selection.fields, selected = fields.findIndex(field => field.id === selection.fieldId), field = fields[selected]
  if (!field || selected < 0 || selected >= fields.length - 1) return []
  const target: PivotChartHierarchyTarget = { unitId: selection.unitId, subUnitId: selection.subUnitId,
    pivotChartId: selection.pivotChartId, host: selection.host, expectedRevision: selection.expectedRevision,
    expectedHistoryGeneration: selection.expectedHistoryGeneration, expectedUserId: selection.expectedUserId }
  return [
    { key: 'expand-item', label: '展开', commandId: PIVOT_HIERARCHY_COMMANDS.chartExpandItem, disabled: disabled || !selection.collapsed },
    { key: 'collapse-item', label: '折叠', commandId: PIVOT_HIERARCHY_COMMANDS.chartCollapseItem, disabled: disabled || selection.collapsed },
    { key: 'expand-field', label: '展开整个字段', commandId: PIVOT_HIERARCHY_COMMANDS.chartSetCollapse,
      params: { ...target, fieldId: field.id, collapse: false }, disabled: disabled || !field.canExpand },
    { key: 'collapse-field', label: '折叠整个字段', commandId: PIVOT_HIERARCHY_COMMANDS.chartSetCollapse,
      params: { ...target, fieldId: field.id, collapse: true }, disabled: disabled || !field.canCollapse },
    ...fields.map((field, index): PivotChartAxisMenuEntry => ({ key: `category-level:${field.id}`,
      label: `${index > selected ? '展开到' : '折叠到'}“${field.name}”`, commandId: PIVOT_HIERARCHY_COMMANDS.chartSetCategoryLevel,
      params: { ...target, fieldId: field.id },
      disabled: disabled || fields.slice(0, -1).every((field, level) => level < index ? !field.canExpand : !field.canCollapse) })),
  ]
}
