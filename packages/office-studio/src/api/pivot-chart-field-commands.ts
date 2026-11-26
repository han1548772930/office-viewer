import type { OfficeStudio } from './office-studio.ts'
import type { PivotChartReportEdit } from '../features/pivot/services/pivot-chart-engine.ts'
import { PIVOT_CHART_FIELD_COMMANDS, type PivotChartFieldCommandInput, type PivotChartFieldCommandParams } from '../features/pivot/services/pivot-chart-field-command-types.ts'
import { commitPivotChartReportEdit, resolvePivotChartTarget } from './pivot-chart-commands.ts'
import { normalizePivotDateFilterInput, normalizePivotLabelConditionInput } from '../features/pivot/services/pivot-filter-api.ts'

/** Fixed action IDs keep their Before/Changed identity and use the one report transaction. */
export function registerPivotChartFieldCommands(studio: OfficeStudio): void {
  function register<K extends keyof PivotChartFieldCommandParams>(key: K, edit: (input: PivotChartFieldCommandInput<K>) => PivotChartReportEdit): void {
    const commandId = PIVOT_CHART_FIELD_COMMANDS[key]
    studio.registerCommand<PivotChartFieldCommandInput<K>, Promise<boolean>>(commandId, raw => {
      const input = resolvePivotChartTarget(raw)
      return commitPivotChartReportEdit(studio, input, edit(input), commandId)
    }, { asynchronous: true })
  }
  register('add', input => ({ type: 'addField', dataFieldId: input.dataFieldId, fieldArea: input.fieldArea, index: input.index }))
  register('move', input => ({ type: 'moveField', fieldId: input.fieldId, fieldArea: input.fieldArea, index: input.index }))
  register('remove', input => ({ type: 'removeFields', fieldIds: input.fieldIds }))
  register('setOptions', input => ({ type: 'setOptions', options: input.options }))
  register('setFieldsConfig', input => ({ type: 'setFieldsConfig', config: input.pivotTableConfig }))
  register('setFieldSetting', input => ({ type: 'setFieldSetting', fieldId: input.fieldId, setting: input.setting }))
  register('setFilter', input => {
    return { type: 'setFilter', fieldId: input.fieldId, items: input.items, isAll: input.isAll,
      partialItems: input.partialItems, clearAllFilters: input.clearAllFilters }
  })
  register('setDateFilter', input => ({ type: 'setDateFilter', fieldId: input.fieldId, filter: normalizePivotDateFilterInput(input.filterInfo) }))
  register('setLabelConditionFilter', input => ({ type: 'setLabelConditionFilter', fieldId: input.fieldId, filter: normalizePivotLabelConditionInput(input.filterInfo) }))
  register('setSort', input => {
    if (input.info !== undefined && (!input.info || typeof input.info !== 'object' || Array.isArray(input.info)
      || Object.keys(input.info).some(key => key !== 'type'))) throw new TypeError('透视图标签排序信息无效。')
    return { type: 'setSort', fieldId: input.fieldId, direction: input.info?.type ?? null }
  })
  register('setValueFilter', input => ({ type: 'setValueFilter', fieldId: input.fieldId, filter: input.valueFilterInfo ?? null }))
  register('setDateGroup', input => ({ type: 'setDateGroup', fieldId: input.fieldId, dateType: input.dateType }))
}
