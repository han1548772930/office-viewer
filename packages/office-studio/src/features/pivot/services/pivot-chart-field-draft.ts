import type { PivotAggregation, PivotFieldArea, PivotFieldModel, PivotShowDataAsInfo, PivotTableModel } from '../../../domain/model.ts'
import { cloneModel } from '../../../domain/serialization.ts'
import type { PivotSourceFieldInfo } from './pivot-engine.ts'
import type { PivotChartFieldsConfig } from './pivot-chart-engine.ts'
import { getPivotShowDataAsRule } from './pivot-show-data-as.ts'
import { adjustPivotChartValuesUiFieldEdit } from './pivot-values-position-ui.ts'

export type PivotChartDraftFieldSetting = Pick<Partial<PivotFieldModel>, 'displayName' | 'format' | 'aggregation' | 'selectedItems' | 'sort'>
  & { showDataAs?: PivotShowDataAsInfo | null }

/** The editable report uses the real source and grouping, without registering or rendering output cells. */
export function pivotChartDraftShowDataAsReport(report: PivotTableModel, config: PivotChartFieldsConfig): PivotTableModel {
  const projected = { ...report, fields: cloneModel(config.fields), options: { ...report.options, ...config.options },
    dateGroupingMode: config.dateGroupingMode ?? report.dateGroupingMode }
  if (config.collapseInfo === null) delete projected.collapseInfo
  else if (config.collapseInfo !== undefined) projected.collapseInfo = cloneModel(config.collapseInfo)
  else if (report.collapseInfo) projected.collapseInfo = cloneModel(report.collapseInfo)
  return projected
}

export const PIVOT_CHART_FIELD_AREAS: ReadonlyArray<{ value: PivotFieldArea; label: string }> = [
  { value: 'filter', label: '筛选' }, { value: 'row', label: '轴（类别）' }, { value: 'column', label: '图例（系列）' }, { value: 'value', label: '值' },
]
export const PIVOT_CHART_AGGREGATION_OPTIONS: ReadonlyArray<{ value: PivotAggregation; label: string }> = [
  { value: 'sum', label: '求和' }, { value: 'count', label: '计数' }, { value: 'countNums', label: '数值计数' }, { value: 'average', label: '平均值' },
  { value: 'max', label: '最大值' }, { value: 'min', label: '最小值' }, { value: 'product', label: '乘积' },
  { value: 'stdDev', label: '样本标准差' }, { value: 'stdDevp', label: '总体标准差' }, { value: 'var', label: '样本方差' }, { value: 'varp', label: '总体方差' },
]

export function pivotChartFieldsInArea(config: PivotChartFieldsConfig, area: PivotFieldArea): PivotFieldModel[] {
  return config.fields.filter(field => field.area === area).sort((a, b) => a.position - b.position)
}

function renumber(config: PivotChartFieldsConfig): PivotChartFieldsConfig {
  for (const { value: area } of PIVOT_CHART_FIELD_AREAS) pivotChartFieldsInArea(config, area).forEach((field, index) => { field.position = index })
  return config
}

export function addPivotChartDraftField(config: PivotChartFieldsConfig, info: PivotSourceFieldInfo, area: PivotFieldArea): PivotChartFieldsConfig {
  if (!PIVOT_CHART_FIELD_AREAS.some(item => item.value === area) || !info.fieldId || !Number.isSafeInteger(info.sourceColumn)) throw new TypeError('请选择有效的来源字段及区域。')
  const existing = config.fields.find(field => (field.sourceFieldId ?? String(field.sourceColumn)) === info.fieldId || field.sourceColumn === info.sourceColumn)
  if (existing) return movePivotChartDraftField(config, existing.id, area, pivotChartFieldsInArea(config, area).filter(field => field.id !== existing.id).length)
  const next = cloneModel(config)
  let id = info.fieldId, suffix = 1
  while (next.fields.some(field => field.id === id)) id = `${info.fieldId}:${suffix++}`
  next.fields.push({ id, sourceFieldId: info.fieldId, sourceColumn: info.sourceColumn, name: info.fieldName, area,
    position: pivotChartFieldsInArea(next, area).length, ...(area === 'value' ? { aggregation: 'sum' } : {}) })
  return adjustPivotChartValuesUiFieldEdit(renumber(next), config.fields, id)
}

export function movePivotChartDraftField(config: PivotChartFieldsConfig, id: string, area: PivotFieldArea, index: number): PivotChartFieldsConfig {
  const next = cloneModel(config), field = next.fields.find(item => item.id === id)
  if (!field || !PIVOT_CHART_FIELD_AREAS.some(item => item.value === area) || !Number.isSafeInteger(index)) throw new TypeError('透视图字段或目标顺序无效。')
  const siblings = pivotChartFieldsInArea(next, area).filter(item => item !== field)
  if (index < 0 || index > siblings.length) throw new RangeError('字段移动位置超出了区域。')
  field.area = area
  if (area === 'value') {
    field.aggregation ??= 'sum'
    delete field.selectedItems; delete field.valueFilter; delete field.sort
    delete field.numberGroup; delete field.dateGroupType; delete field.manualGroups
  } else { delete field.aggregation; delete field.showDataAs }
  siblings.splice(index, 0, field)
  siblings.forEach((item, position) => { item.position = position })
  return adjustPivotChartValuesUiFieldEdit(renumber(next), config.fields, id)
}

export function removePivotChartDraftFields(config: PivotChartFieldsConfig, ids: readonly string[]): PivotChartFieldsConfig {
  if (ids.some(id => !config.fields.some(field => field.id === id))) throw new TypeError('要移除的透视图字段已不存在。')
  const next = cloneModel(config), removed = new Set(ids)
  next.fields = next.fields.filter(field => !removed.has(field.id))
  return adjustPivotChartValuesUiFieldEdit(renumber(next), config.fields)
}

export function setPivotChartDraftField(config: PivotChartFieldsConfig, id: string, patch: PivotChartDraftFieldSetting): PivotChartFieldsConfig {
  const next = cloneModel(config), field = next.fields.find(item => item.id === id)
  if (!field) throw new Error('要设置的透视图字段已不存在。')
  if (patch.aggregation !== undefined && (field.area !== 'value' || !PIVOT_CHART_AGGREGATION_OPTIONS.some(item => item.value === patch.aggregation))) throw new TypeError('该字段不支持所选汇总方式。')
  if (Object.hasOwn(patch, 'selectedItems') && field.area === 'value') throw new TypeError('请对类别、系列或筛选字段设置标签筛选。')
  if (Object.hasOwn(patch, 'showDataAs')) {
    if (field.area !== 'value') throw new TypeError('值显示方式只能用于值字段。')
    const value = patch.showDataAs
    if (value === undefined) throw new TypeError('请提供值显示方式或显式恢复无计算。')
    if (value === null || value.type === 'normal') delete field.showDataAs
    else {
      getPivotShowDataAsRule(value.type)
      if (typeof value.baseFieldId !== 'string' || typeof value.baseItem !== 'string'
        || value.baseItemType !== undefined && !['item', 'previous', 'next'].includes(value.baseItemType)) throw new TypeError('值显示方式草稿无效。')
      // Incomplete parameters are visible drafts. The common rule service validates before configure commits.
      field.showDataAs = { ...value }
    }
  }
  for (const key of ['displayName', 'format', 'aggregation', 'sort'] as const) if (Object.hasOwn(patch, key)) {
    if (patch[key] === undefined) delete field[key]
    else Object.assign(field, { [key]: patch[key] })
  }
  if (Object.hasOwn(patch, 'selectedItems')) {
    if (patch.selectedItems === undefined) delete field.selectedItems
    else field.selectedItems = [...new Set(patch.selectedItems)]
  }
  return next
}
