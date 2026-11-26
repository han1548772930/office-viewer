import type { PivotDateFilter, PivotDateFilterOperator, PivotLabelConditionFilter, PivotLabelConditionOperator } from '../../../domain/model.ts'
import { parseTimelineDate, timelineCalendarToSerial, timelineDateToInput } from '../../slicer/timeline-date.ts'
import { PIVOT_DATE_FILTER_OPERATORS, PIVOT_LABEL_CONDITION_OPERATORS, normalizePivotDateFilter,
  normalizePivotLabelConditionFilter } from './pivot-filter-predicate.ts'

export interface PivotChartFieldFilterCapabilities { supportsDateFilter: boolean; date1904: boolean }
export interface PivotChartDateConditionDraft { operator: PivotDateFilterOperator; first?: number; second?: number; useWholeDay: boolean }
export interface PivotChartLabelConditionDraft { operator: PivotLabelConditionOperator; first: string; second: string }
export interface PivotChartFieldConditionsDraft {
  capabilities: PivotChartFieldFilterCapabilities
  date: PivotChartDateConditionDraft
  label: PivotChartLabelConditionDraft
}

const dateLabels: Record<PivotDateFilterOperator, string> = { dateBetween: '介于', dateEqual: '等于', dateNewerThan: '晚于',
  dateNewerThanOrEqual: '晚于或等于', dateNotBetween: '不介于', dateNotEqual: '不等于', dateOlderThan: '早于', dateOlderThanOrEqual: '早于或等于' }
const labelLabels: Record<PivotLabelConditionOperator, string> = { captionBeginsWith: '开头是', captionBetween: '介于',
  captionContains: '包含', captionEndsWith: '结尾是', captionEqual: '等于', captionGreaterThan: '大于',
  captionGreaterThanOrEqual: '大于或等于', captionLessThan: '小于', captionLessThanOrEqual: '小于或等于',
  captionNotBeginsWith: '开头不是', captionNotBetween: '不介于', captionNotContains: '不包含', captionNotEndsWith: '结尾不是', captionNotEqual: '不等于' }
export const PIVOT_CHART_DATE_CONDITIONS = PIVOT_DATE_FILTER_OPERATORS.map(value => ({ value, label: dateLabels[value] }))
export const PIVOT_CHART_LABEL_CONDITIONS = PIVOT_LABEL_CONDITION_OPERATORS.map(value => ({ value, label: labelLabels[value] }))
export const pivotChartDateConditionBetween = (value: PivotDateFilterOperator): boolean => value === 'dateBetween' || value === 'dateNotBetween'
export const pivotChartLabelConditionBetween = (value: PivotLabelConditionOperator): boolean => value === 'captionBetween' || value === 'captionNotBetween'

/** Consume the Facade's source capability and typed conditions; do not infer date fields from names or member keys. */
export function createPivotChartFieldConditionsDraft(capabilities: PivotChartFieldFilterCapabilities,
  date?: PivotDateFilter, label?: PivotLabelConditionFilter): PivotChartFieldConditionsDraft {
  if (typeof capabilities.supportsDateFilter !== 'boolean' || typeof capabilities.date1904 !== 'boolean') throw new TypeError('透视图字段筛选能力无效。')
  const dateFilter = date ? normalizePivotDateFilter(date) : undefined, labelFilter = label ? normalizePivotLabelConditionFilter(label) : undefined
  const dates = Array.isArray(dateFilter?.expected) ? dateFilter.expected : [dateFilter?.expected, dateFilter?.expected]
  const labels = Array.isArray(labelFilter?.expected) ? labelFilter.expected : [labelFilter?.expected ?? '', labelFilter?.expected ?? '']
  return { capabilities: { ...capabilities },
    date: { operator: dateFilter?.operator ?? 'dateEqual', first: dates[0], second: dates[1], useWholeDay: dateFilter?.useWholeDay === true },
    label: { operator: labelFilter?.operator ?? 'captionEqual', first: labels[0]!, second: labels[1]! } }
}

export function setPivotChartDateConditionDraft(draft: PivotChartFieldConditionsDraft,
  patch: Partial<PivotChartDateConditionDraft>): PivotChartFieldConditionsDraft {
  if (Object.keys(patch).some(key => !['operator', 'first', 'second', 'useWholeDay'].includes(key))
    || patch.operator !== undefined && !PIVOT_DATE_FILTER_OPERATORS.includes(patch.operator)
    || patch.useWholeDay !== undefined && typeof patch.useWholeDay !== 'boolean'
    || [patch.first, patch.second].some(value => value !== undefined && (typeof value !== 'number' || !Number.isFinite(value)))) {
    throw new TypeError('日期筛选草稿无效。')
  }
  return { ...draft, date: { ...draft.date, ...patch } }
}
export function setPivotChartLabelConditionDraft(draft: PivotChartFieldConditionsDraft,
  patch: Partial<PivotChartLabelConditionDraft>): PivotChartFieldConditionsDraft {
  if (Object.keys(patch).some(key => !['operator', 'first', 'second'].includes(key))
    || patch.operator !== undefined && !PIVOT_LABEL_CONDITION_OPERATORS.includes(patch.operator)
    || [patch.first, patch.second].some(value => value !== undefined && typeof value !== 'string')) throw new TypeError('标签筛选草稿无效。')
  return { ...draft, label: { ...draft.label, ...patch } }
}
export function pivotChartDateConditionFilter(draft: PivotChartFieldConditionsDraft): PivotDateFilter {
  if (!draft.capabilities.supportsDateFilter) throw new Error('当前来源字段不支持日期筛选。')
  const { operator, first, second, useWholeDay } = draft.date
  if (first === undefined || pivotChartDateConditionBetween(operator) && second === undefined) throw new Error('请选择有效日期。')
  if (pivotChartDateConditionBetween(operator) && first > second!) throw new Error('结束日期必须晚于或等于开始日期。')
  return normalizePivotDateFilter({ operator, expected: pivotChartDateConditionBetween(operator) ? [first, second!] : first, useWholeDay })
}
export function pivotChartLabelConditionFilter(draft: PivotChartFieldConditionsDraft): PivotLabelConditionFilter {
  const { operator, first, second } = draft.label
  return normalizePivotLabelConditionFilter({ operator, expected: pivotChartLabelConditionBetween(operator) ? [first, second] : first })
}

export function pivotChartFilterDateCompatibilityDay(serial: number | undefined, date1904: boolean): boolean {
  return !date1904 && serial !== undefined && Math.floor(serial) === 60
}
/** Calendar text is an adapter only. Unchanged numeric operands keep all original fractional precision. */
export function pivotChartFilterDateCalendarInput(serial: number | undefined, date1904: boolean): string {
  if (serial === undefined || !parseTimelineDate(serial, date1904)) return ''
  const milliseconds = Math.min(86_399_999, Math.round((serial - Math.floor(serial)) * 86_400_000))
  const hour = Math.floor(milliseconds / 3_600_000), minute = Math.floor(milliseconds % 3_600_000 / 60_000)
  const second = Math.floor(milliseconds % 60_000 / 1000), fraction = milliseconds % 1000
  const pad = (value: number, length = 2): string => String(value).padStart(length, '0')
  return `${timelineDateToInput(serial, date1904)} ${pad(hour)}:${pad(minute)}:${pad(second)}.${pad(fraction, 3)}`
}
/** Reuse the workbook epoch/compatibility-day converter and interpret time as calendar time, independent of the host timezone. */
export function pivotChartFilterDateFromCalendarInput(value: string, date1904: boolean): number | undefined {
  const parts = /^(\d{4}-\d{2}-\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?)?$/.exec(value.trim())
  if (!parts) return undefined
  const day = timelineCalendarToSerial(parts[1]!, date1904), hour = Number(parts[2] ?? 0), minute = Number(parts[3] ?? 0), second = Number(parts[4] ?? 0)
  if (day === undefined || hour > 23 || minute > 59 || second > 59) return undefined
  const fraction = Number((parts[5] ?? '').padEnd(3, '0'))
  return day + (hour * 3_600_000 + minute * 60_000 + second * 1000 + fraction) / 86_400_000
}
