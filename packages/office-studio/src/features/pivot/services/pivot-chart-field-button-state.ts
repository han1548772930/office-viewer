import { FOffice, type FSheetPivotChart } from '../../../api/facade.ts'
import type { OfficeStudio } from '../../../api/office-studio.ts'
import type { PivotChartControls, PivotFieldArea, PivotFieldModel } from '../../../domain/model.ts'
import { cloneModel } from '../../../domain/serialization.ts'
import type { PivotChartCommandContext } from './pivot-chart-api-types.ts'
import type { PivotChartFieldsConfig, PivotChartReportEdit } from './pivot-chart-engine.ts'
import { canConfigurePivotChartPanel, canViewPivotChartPanelSource } from './pivot-chart-panel-state.ts'
import { capturePivotChartPanelOwner, pivotChartPanelOwnerCurrent, type PivotChartPanelOwner } from './pivot-chart-panel-owner.ts'
import { getPivotFieldFilterCategories } from './pivot-filter-predicate.ts'
import { pivotChartDateConditionFilter, pivotChartLabelConditionFilter, type PivotChartFieldConditionsDraft } from './pivot-chart-field-filter-condition.ts'
import { readPivotValuesUiState, type PivotValuesUiState } from './pivot-values-position-ui.ts'

export interface PivotChartFieldButtonsTarget { studio: OfficeStudio; unitId: string; sheetId: string; chartId: string }
export const PIVOT_CHART_BUTTON_AREAS: ReadonlyArray<{ area: PivotFieldArea; label: string }> = [
  { area: 'filter', label: '筛选' }, { area: 'column', label: '图例（系列）' },
  { area: 'row', label: '轴（类别）' }, { area: 'value', label: '值' },
]

export interface PivotChartFieldButton {
  fieldId: string
  label: string
  area: PivotFieldArea
  filtered: boolean
  field: PivotFieldModel
}
export interface PivotChartFieldButtonSection { area: PivotFieldArea; label: string; fields: PivotChartFieldButton[] }
export interface PivotChartFieldButtonsSnapshot {
  owner: PivotChartPanelOwner
  controls: PivotChartControls
  editable: boolean
  sections: PivotChartFieldButtonSection[]
  values: PivotValuesUiState
}
export interface PivotChartManualFilterItemState { checkedItems: string[]; partialItems: string[] }
export interface PivotChartFieldFilterDraft { items: string[]; selectedKeys: string[]; partialKeys: string[]; search: string }
export type PivotChartFieldButtonAction = 'applyFilter' | 'clearFilter' | 'sortAscending' | 'sortDescending' | 'clearSort' | 'remove'
  | 'applyDateFilter' | 'clearDateFilter' | 'applyLabelCondition' | 'clearLabelCondition'

/** Report-local identity and authored order survive duplicate source/value instances. */
export function pivotChartFieldButtonSections(config: PivotChartFieldsConfig): PivotChartFieldButtonSection[] {
  return PIVOT_CHART_BUTTON_AREAS.map(section => ({ ...section,
    fields: config.fields.filter(field => field.area === section.area).map((field, index) => ({ field, index }))
      .sort((left, right) => left.field.position - right.field.position || left.index - right.index)
      .map(({ field }) => ({ fieldId: field.id, label: field.displayName ?? field.name, area: field.area,
        filtered: getPivotFieldFilterCategories(field).size > 0, field: cloneModel(field) })),
  }))
}

export function pivotChartFieldButtonsFacade(target: PivotChartFieldButtonsTarget): FSheetPivotChart {
  const pivot = new FOffice(target.studio).getWorkbook(target.unitId)?.getSheetBySheetId(target.sheetId)?.getPivotChartByChartId(target.chartId)
  if (!pivot || pivot.getSubUnitId() !== target.sheetId) throw new Error('透视图已经不在原工作表中。')
  return pivot
}

/** Reads only; no report, cell, drawing or content history is registered. */
export function readPivotChartFieldButtonsSnapshot(target: PivotChartFieldButtonsTarget): PivotChartFieldButtonsSnapshot {
  const owner = capturePivotChartPanelOwner(target.studio, target.unitId, target.sheetId)
  const workbook = target.studio.getWorkbookModel(target.unitId), chart = workbook?.charts?.find(chart => chart.id === target.chartId)
  const report = chart?.ownedPivot ?? workbook?.pivotTables?.find(report => report.id === chart?.pivotTableId)
  if (!workbook || !chart || chart.sheetId !== target.sheetId || !report
    || !canViewPivotChartPanelSource(workbook, report, target.sheetId, chart, owner.userId, Boolean(chart.pivotTableId))) {
    throw new Error('当前用户无法查看透视图字段、来源或报告。')
  }
  const pivot = pivotChartFieldButtonsFacade(target), config = pivot.getFieldsConfig()
  // Field settings come from the same Facade so permission and local field lookup are shared.
  const fields = config.fields.map(field => {
    const setting = pivot.getFieldSetting(field.id)
    if (!setting) throw new Error('透视图字段已经被移除。')
    return setting
  })
  const values = pivot.getValuesPosition()
  const snapshot = { owner, controls: pivot.getControls(), sections: pivotChartFieldButtonSections({ ...config, fields }),
    values: readPivotValuesUiState({ valuePosition: values.position, valueIndex: values.index }, fields),
    editable: canConfigurePivotChartPanel(workbook, report, target.sheetId, chart, owner.userId, Boolean(chart.pivotTableId)) }
  if (!pivotChartPanelOwnerCurrent(owner)) throw new Error('透视图字段读取已过期。')
  return snapshot
}

export function createPivotChartFieldFilterDraft(items: string[], state: PivotChartManualFilterItemState): PivotChartFieldFilterDraft {
  if (!Array.isArray(items) || items.some(item => typeof item !== 'string') || new Set(items).size !== items.length) {
    throw new TypeError('透视图筛选成员必须是互不重复的字符串键。')
  }
  for (const keys of [state.checkedItems, state.partialItems]) {
    if (!Array.isArray(keys) || keys.some(key => typeof key !== 'string' || !items.includes(key)) || new Set(keys).size !== keys.length) {
      throw new TypeError('透视图筛选状态必须引用当前分组成员。')
    }
  }
  if (state.checkedItems.some(key => state.partialItems.includes(key))) throw new TypeError('筛选成员不能同时全选和部分选中。')
  const selected = new Set(state.checkedItems), partial = new Set(state.partialItems)
  return { items: [...items], selectedKeys: items.filter(item => selected.has(item)), partialKeys: items.filter(item => partial.has(item)), search: '' }
}

export function pivotChartFilterMemberLabel(key: string): string { return key === '' ? '（空白）' : key }
export function pivotChartFilterVisibleMembers(draft: PivotChartFieldFilterDraft): string[] {
  const search = draft.search.trim().toLocaleLowerCase()
  return draft.items.filter(item => !search || pivotChartFilterMemberLabel(item).toLocaleLowerCase().includes(search))
}
export function setPivotChartFilterSearch(draft: PivotChartFieldFilterDraft, search: string): PivotChartFieldFilterDraft {
  if (typeof search !== 'string') throw new TypeError('筛选搜索必须是文本。')
  return { ...draft, search }
}
export function setPivotChartFilterSelection(draft: PivotChartFieldFilterDraft, key: string, checked: boolean): PivotChartFieldFilterDraft {
  if (typeof checked !== 'boolean' || !draft.items.includes(key)) throw new TypeError('所选筛选成员已失效。')
  const selected = new Set(draft.selectedKeys)
  if (checked) selected.add(key)
  else selected.delete(key)
  return { ...draft, selectedKeys: draft.items.filter(item => selected.has(item)), partialKeys: draft.partialKeys.filter(item => item !== key) }
}
/** Select-all applies to the whole member catalog, including off-page/search-hidden keys. */
export function setPivotChartFilterAll(draft: PivotChartFieldFilterDraft, checked: boolean): PivotChartFieldFilterDraft {
  if (typeof checked !== 'boolean') throw new TypeError('全选状态无效。')
  return { ...draft, selectedKeys: checked ? [...draft.items] : [], partialKeys: [] }
}
export function pivotChartFieldFilterEdit(fieldId: string, draft: PivotChartFieldFilterDraft): Extract<PivotChartReportEdit, { type: 'setFilter' }> {
  if (!fieldId || !draft.selectedKeys.length && !draft.partialKeys.length
    || [draft.selectedKeys, draft.partialKeys].some(keys => keys.some(key => !draft.items.includes(key)) || new Set(keys).size !== keys.length)
    || draft.selectedKeys.some(key => draft.partialKeys.includes(key))) throw new Error('请至少选择一个仍然存在的筛选成员。')
  const isAll = !draft.partialKeys.length && draft.selectedKeys.length === draft.items.length
  return { type: 'setFilter', fieldId, items: isAll ? null : [...draft.selectedKeys], isAll, partialItems: [...draft.partialKeys] }
}

/** Every content action reaches the existing report edit command and its transaction/history. */
export function applyPivotChartFieldButtonAction(pivot: FSheetPivotChart, fieldId: string, action: PivotChartFieldButtonAction,
  context: PivotChartCommandContext, draft?: PivotChartFieldFilterDraft, conditions?: PivotChartFieldConditionsDraft): Promise<boolean> {
  switch (action) {
    case 'applyFilter': {
      if (!draft) throw new Error('筛选成员尚未读取。')
      const edit = pivotChartFieldFilterEdit(fieldId, draft)
      return pivot.applyReportEdit(edit, context)
    }
    case 'clearFilter': return pivot.applyReportEdit({ type: 'setFilter', fieldId, items: null, isAll: true, clearAllFilters: true }, context)
    case 'sortAscending': return pivot.setLabelSort(fieldId, { type: 'ascending' }, context)
    case 'sortDescending': return pivot.setLabelSort(fieldId, { type: 'descending' }, context)
    case 'clearSort': return pivot.clearLabelSort(fieldId, context)
    case 'remove': return pivot.removeField([fieldId], context)
    case 'applyDateFilter': {
      if (!conditions) throw new Error('日期筛选条件尚未读取。')
      return pivot.setDateFilter(fieldId, pivotChartDateConditionFilter(conditions), context)
    }
    case 'clearDateFilter': return pivot.setDateFilter(fieldId, null, context)
    case 'applyLabelCondition': {
      if (!conditions) throw new Error('标签筛选条件尚未读取。')
      return pivot.setLabelConditionFilter(fieldId, pivotChartLabelConditionFilter(conditions), context)
    }
    case 'clearLabelCondition': return pivot.setLabelConditionFilter(fieldId, null, context)
  }
}
