import type { PivotShowDataAsInfo, PivotShowDataAsType, PivotTableModel } from '../../../domain/model.ts'
import { getPivotShowDataAsBaseFieldIds, getPivotShowDataAsRule, normalizePivotShowDataAs, PIVOT_SHOW_DATA_AS_RULES } from './pivot-show-data-as.ts'

/** Fixed v1.6.0 picker entries. Imported historical rules remain readable without becoming new UI entries. */
export const PIVOT_SHOW_DATA_AS_OPTIONS = PIVOT_SHOW_DATA_AS_RULES.filter(rule => !rule.compatibilityOnly)
  .map(rule => ({ value: rule.type, label: rule.label }))

export type PivotShowDataAsBaseItemsLoader = (reportFieldId: string) => readonly string[] | Promise<readonly string[]>

export function createPivotShowDataAsDraft(value?: PivotShowDataAsInfo | null): PivotShowDataAsInfo {
  return value ? { ...value } : { type: 'normal', baseFieldId: '', baseItem: '' }
}

export function pivotShowDataAsBaseFieldOptions(report: PivotTableModel, type: PivotShowDataAsType): Array<{ value: string; label: string }> {
  const ids = getPivotShowDataAsBaseFieldIds(report, type)
  const fields = ids.map(id => report.fields.find(field => field.id === id)!)
  return fields.map(field => {
    const name = field.displayName ?? field.name
    const duplicate = fields.some(other => other.id !== field.id && (other.displayName ?? other.name) === name)
    return { value: field.id, label: duplicate ? `${name} · ${field.area === 'row' ? '行' : '列'} · ${field.id}` : name }
  })
}

/** Changing a mode keeps only an eligible report-local base field; old member/relative parameters are cleared. */
export function changePivotShowDataAsType(report: PivotTableModel, current: PivotShowDataAsInfo, type: PivotShowDataAsType): PivotShowDataAsInfo {
  if (!PIVOT_SHOW_DATA_AS_OPTIONS.some(option => option.value === type)) throw new TypeError('请选择固定版本支持的值显示方式。')
  const rule = getPivotShowDataAsRule(type)
  const ids = getPivotShowDataAsBaseFieldIds(report, type)
  return { type, baseFieldId: rule.requiresParameters && ids.includes(current.baseFieldId) ? current.baseFieldId : '', baseItem: '' }
}

export function changePivotShowDataAsBaseField(report: PivotTableModel, current: PivotShowDataAsInfo, baseFieldId: string): PivotShowDataAsInfo {
  if (!getPivotShowDataAsBaseFieldIds(report, current.type).includes(baseFieldId)) throw new TypeError('请选择已配置的行或列基准字段。')
  return { type: current.type, baseFieldId, baseItem: '' }
}

export function changePivotShowDataAsBaseItemType(current: PivotShowDataAsInfo, type: 'item' | 'previous' | 'next'): PivotShowDataAsInfo {
  const rule = getPivotShowDataAsRule(current.type)
  if (rule.baseItemMode !== 'item-or-relative') throw new TypeError('当前显示方式不支持相对基准项。')
  return { type: current.type, baseFieldId: current.baseFieldId, baseItem: '', ...(type === 'item' ? {} : { baseItemType: type }) }
}

/** UI validation delegates the DTO semantics to the one Pivot rule service. Actual commands also check scalar-key ambiguity. */
export function preparePivotShowDataAsDraft(report: PivotTableModel, value: PivotShowDataAsInfo | null,
  items?: readonly string[]): PivotShowDataAsInfo | null {
  return normalizePivotShowDataAs(report, value, items ? () => items : undefined) ?? null
}

export function pivotShowDataAsDraftError(report: PivotTableModel | undefined, value: PivotShowDataAsInfo | null,
  items?: readonly string[]): string {
  if (!report) return '请先读取仍可访问的报告字段。'
  try { preparePivotShowDataAsDraft(report, value, items); return '' }
  catch (reason) { return reason instanceof Error ? reason.message : String(reason) }
}

/** Only field structure/grouping affects base-member reads, not a value field's unsaved Show As choice. */
export function pivotShowDataAsMemberRevision(report?: PivotTableModel): string {
  return report ? JSON.stringify([report.id, report.source, report.dateGroupingMode, report.fields.map(field => ({
    id: field.id, sourceFieldId: field.sourceFieldId, sourceColumn: field.sourceColumn, area: field.area, position: field.position,
    numberGroup: field.numberGroup, dateGroupType: field.dateGroupType, manualGroups: field.manualGroups,
  }))]) : ''
}
