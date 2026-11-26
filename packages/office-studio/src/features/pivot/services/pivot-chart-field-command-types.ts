import type { PivotDateGroupType, PivotFieldArea, PivotFieldFilter, PivotSortDirection, PivotTableOptions } from '../../../domain/model.ts'
import type { PivotChartCommandContext } from './pivot-chart-api-types.ts'
import type { PivotChartFieldsConfig, PivotChartFieldSetting } from './pivot-chart-engine.ts'
import type { PivotDateFilterInput, PivotLabelConditionInput } from './pivot-filter-api.ts'

export interface PivotChartFieldCommandTarget extends PivotChartCommandContext {
  unitId: string
  pivotChartId: string
  subUnitId?: string
  sheetId?: string
}

export const PIVOT_CHART_FIELD_COMMANDS = {
  add: 'sheet.command.add-pivot-chart-field',
  move: 'sheet.command.move-pivot-chart-field',
  remove: 'sheet.command.remove-pivot-chart-field',
  setOptions: 'sheet.command.set-pivot-chart-options',
  setFieldsConfig: 'sheet.command.set-pivot-chart-fields-config',
  setFieldSetting: 'sheet.command.set-pivot-chart-field-setting',
  setFilter: 'sheet.command.set-pivot-chart-filter',
  setDateFilter: 'sheet.command.set-pivot-chart-date-filter',
  setLabelConditionFilter: 'sheet.command.set-pivot-chart-label-condition-filter',
  setSort: 'sheet.command.set-pivot-chart-sort',
  setValueFilter: 'sheet.command.set-pivot-chart-value-filter',
  setDateGroup: 'sheet.command.set-pivot-chart-date-group',
} as const

export interface PivotChartFieldCommandParams {
  add: { dataFieldId: string | number; fieldArea: PivotFieldArea; index?: number }
  move: { fieldId: string; fieldArea: PivotFieldArea; index: number }
  remove: { fieldIds: string[] }
  setOptions: { options: Partial<PivotTableOptions> }
  setFieldsConfig: { pivotTableConfig: PivotChartFieldsConfig }
  setFieldSetting: { fieldId: string; setting: PivotChartFieldSetting }
  setFilter: { fieldId: string; items: string[] | null; isAll?: boolean; partialItems?: string[]; clearAllFilters?: boolean }
  setDateFilter: { fieldId: string; filterInfo?: PivotDateFilterInput | null }
  setLabelConditionFilter: { fieldId: string; filterInfo?: PivotLabelConditionInput | null }
  setSort: { fieldId: string; info?: { type: PivotSortDirection } }
  setValueFilter: { fieldId: string; valueFilterInfo?: PivotFieldFilter | null }
  setDateGroup: { fieldId: string; dateType: PivotDateGroupType | null }
}

export type PivotChartFieldCommandInput<K extends keyof PivotChartFieldCommandParams> = PivotChartFieldCommandTarget & PivotChartFieldCommandParams[K]
