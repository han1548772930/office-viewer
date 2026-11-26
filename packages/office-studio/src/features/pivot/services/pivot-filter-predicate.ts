import type { CellScalar, PivotDateFilter, PivotDateFilterOperator, PivotFieldFilter, PivotFieldModel, PivotLabelConditionFilter,
  PivotLabelConditionOperator, PivotRecordSelection, PivotTableModel } from '../../../domain/model.ts'
import { getRecordValueKey, isRecordValue, normalizeRecordValue, pivotRecordScalar, pivotRecordValueKey, recordValueFromScalar,
  type PivotRecordValue, type RecordValue } from '../../../domain/record-value.ts'
import { compileTimelineSerialCondition, timelineDateConditionFromPivot } from '../../slicer/timeline-date-filter.ts'

export const PIVOT_DATE_FILTER_OPERATORS: readonly PivotDateFilterOperator[] = ['dateBetween', 'dateEqual', 'dateNewerThan',
  'dateNewerThanOrEqual', 'dateNotBetween', 'dateNotEqual', 'dateOlderThan', 'dateOlderThanOrEqual']
export const PIVOT_LABEL_CONDITION_OPERATORS: readonly PivotLabelConditionOperator[] = ['captionBeginsWith', 'captionBetween',
  'captionContains', 'captionEndsWith', 'captionEqual', 'captionGreaterThan', 'captionGreaterThanOrEqual', 'captionLessThan',
  'captionLessThanOrEqual', 'captionNotBeginsWith', 'captionNotBetween', 'captionNotContains', 'captionNotEndsWith', 'captionNotEqual']

const object = (value: unknown): value is Record<string, unknown> => !!value && typeof value === 'object' && !Array.isArray(value)

export function normalizePivotDateFilter(value: PivotDateFilter): PivotDateFilter {
  if (!object(value) || Object.keys(value).some(key => !['operator', 'expected', 'useWholeDay'].includes(key))
    || !PIVOT_DATE_FILTER_OPERATORS.includes(value.operator)
    || value.useWholeDay !== undefined && typeof value.useWholeDay !== 'boolean') throw new TypeError('Invalid pivot date filter')
  const between = value.operator === 'dateBetween' || value.operator === 'dateNotBetween'
  const operands = between ? value.expected : [value.expected]
  if (!Array.isArray(operands) || operands.length !== (between ? 2 : 1)
    || operands.some(operand => typeof operand !== 'number' || !Number.isFinite(operand))
    || between && Number(operands[0]) > Number(operands[1])) throw new RangeError('Pivot date operands must be finite and ordered')
  return { operator: value.operator, expected: between ? [Number(operands[0]), Number(operands[1])] : Number(operands[0]),
    ...(value.useWholeDay !== undefined ? { useWholeDay: value.useWholeDay } : {}) }
}

/** Excel serial arithmetic retains the workbook epoch and the distinct compatibility day 60. */
export function compilePivotDateFilter(value: PivotDateFilter): (serial: number) => boolean {
  return compileTimelineSerialCondition(timelineDateConditionFromPivot(normalizePivotDateFilter(value)))
}

export function normalizePivotLabelConditionFilter(value: PivotLabelConditionFilter): PivotLabelConditionFilter {
  if (!object(value) || Object.keys(value).some(key => !['operator', 'expected'].includes(key))
    || !PIVOT_LABEL_CONDITION_OPERATORS.includes(value.operator)) throw new TypeError('Invalid pivot label condition')
  const between = value.operator === 'captionBetween' || value.operator === 'captionNotBetween'
  const operands = between ? value.expected : [value.expected]
  if (!Array.isArray(operands) || operands.length !== (between ? 2 : 1) || operands.some(operand => typeof operand !== 'string')) {
    throw new TypeError('Pivot caption operands must be strings')
  }
  return { operator: value.operator, expected: between ? [String(operands[0]), String(operands[1])] : String(operands[0]) }
}

/** Fixed caption operators compare case-sensitive keys with JavaScript ordinal ordering. */
export function compilePivotLabelConditionFilter(value: PivotLabelConditionFilter): (label: string) => boolean {
  const filter = normalizePivotLabelConditionFilter(value), operands = Array.isArray(filter.expected) ? filter.expected : [filter.expected]
  const first = operands[0]!, second = operands[1] ?? first
  return label => {
    switch (filter.operator) {
      case 'captionBeginsWith': return label.startsWith(first)
      case 'captionBetween': return label >= first && label <= second
      case 'captionContains': return label.includes(first)
      case 'captionEndsWith': return label.endsWith(first)
      case 'captionEqual': return label === first
      case 'captionGreaterThan': return label > first
      case 'captionGreaterThanOrEqual': return label >= first
      case 'captionLessThan': return label < first
      case 'captionLessThanOrEqual': return label <= first
      case 'captionNotBeginsWith': return !label.startsWith(first)
      case 'captionNotBetween': return label < first || label > second
      case 'captionNotContains': return !label.includes(first)
      case 'captionNotEndsWith': return !label.endsWith(first)
      case 'captionNotEqual': return label !== first
    }
    return false
  }
}

export function normalizePivotValueFilter(value: PivotFieldFilter): PivotFieldFilter {
  if (!object(value) || Object.keys(value).some(key => !['operator', 'expected', 'expected2', 'valueFieldId'].includes(key))
    || !['valueGreaterThan', 'valueGreaterThanOrEqual', 'valueLessThan', 'valueLessThanOrEqual', 'valueEqual', 'valueNotEqual', 'valueBetween', 'valueNotBetween'].includes(value.operator)
    || typeof value.valueFieldId !== 'string' || !value.valueFieldId.trim()) throw new TypeError('Invalid pivot value filter')
  if (!Number.isFinite(value.expected) || value.expected2 !== undefined && !Number.isFinite(value.expected2)
    || ['valueBetween', 'valueNotBetween'].includes(value.operator) && value.expected2 === undefined) throw new TypeError('Pivot value filter requires finite operands')
  return { ...value }
}

const scalar = (value: unknown): value is CellScalar => value === null || typeof value === 'string' || typeof value === 'boolean'
  || typeof value === 'number' && Number.isFinite(value)

export function normalizePivotRecordSelection(value: PivotRecordSelection): PivotRecordSelection {
  if (!object(value) || !['all', 'include', 'exclude'].includes(value.mode)
    || Object.keys(value).some(key => !['mode', 'values'].includes(key))) throw new TypeError('Invalid pivot record selection')
  if (value.mode === 'all') {
    if (Object.hasOwn(value, 'values')) throw new TypeError('All record selection cannot contain values')
    return { mode: 'all' }
  }
  if (!Array.isArray(value.values)) throw new TypeError('Pivot record selection requires values')
  const seen = new Set<string>(), values: PivotRecordValue[] = []
  for (const member of value.values) {
    if (!isRecordValue(member) && !scalar(member)) throw new TypeError('Pivot record values must be finite scalars or typed records')
    const normalized = isRecordValue(member) ? normalizeRecordValue(member) : member, key = pivotRecordValueKey(normalized)
    if (!seen.has(key)) { seen.add(key); values.push(normalized) }
  }
  return { mode: value.mode, values }
}

export function compilePivotRecordSelection(value: PivotRecordSelection): (member: PivotRecordValue, typedValue?: RecordValue) => boolean {
  const selection = normalizePivotRecordSelection(value)
  if (selection.mode === 'all') return () => true
  const raw = new Set(selection.values.filter(value => !isRecordValue(value)).map(value => getRecordValueKey(recordValueFromScalar(pivotRecordScalar(value)))))
  const typed = new Set(selection.values.filter(isRecordValue).map(getRecordValueKey))
  return (member, typedValue) => (raw.has(getRecordValueKey(recordValueFromScalar(pivotRecordScalar(member))))
    || typed.has(getRecordValueKey(typedValue ?? (isRecordValue(member) ? member : recordValueFromScalar(member))))) === (selection.mode === 'include')
}

export type PivotFieldFilterCategory = 'member' | 'condition' | 'value'
export type PivotFieldFilterState = Pick<PivotFieldModel, 'selectedItems' | 'groupedItems' | 'partialItems' | 'recordSelection' | 'labelFilter' | 'dateFilter' | 'timelineDateRange' | 'timelineDateCondition' | 'valueFilter'>
const properties = {
  member: ['selectedItems', 'groupedItems', 'partialItems', 'recordSelection'], condition: ['labelFilter', 'dateFilter', 'timelineDateRange', 'timelineDateCondition'], value: ['valueFilter'],
} as const

export function pivotFilterSourceKey(pivot: PivotTableModel, field: PivotFieldModel): string {
  return field.sourceFieldId ?? `column:${field.sourceColumn - pivot.source.range.startCol}`
}

export function getPivotFieldFilterCategories(field: PivotFieldModel): Set<PivotFieldFilterCategory> {
  const categories = new Set<PivotFieldFilterCategory>()
  const member = field.groupedItems !== undefined || (field.recordSelection
    ? field.recordSelection.mode === 'include' || field.recordSelection.mode === 'exclude' && field.recordSelection.values.length > 0
    : field.selectedItems !== undefined)
  if (member) categories.add('member')
  if (field.labelFilter || field.dateFilter || field.timelineDateRange || field.timelineDateCondition) categories.add('condition')
  if (field.valueFilter) categories.add('value')
  return categories
}

/** Categories are counted per source identity, including repeated report-local dimensions. */
export function hasPivotMultipleFilters(pivot: PivotTableModel): boolean {
  const sources = new Map<string, Set<PivotFieldFilterCategory>>()
  for (const field of pivot.fields.filter(field => field.area !== 'value')) {
    const key = pivotFilterSourceKey(pivot, field), categories = sources.get(key) ?? new Set<PivotFieldFilterCategory>()
    for (const category of getPivotFieldFilterCategories(field)) categories.add(category)
    if (categories.size > 1) return true
    sources.set(key, categories)
  }
  return false
}

function labelField(pivot: PivotTableModel, fieldId: string): PivotFieldModel {
  const field = pivot.fields.find(field => field.id === fieldId)
  if (!field || field.area === 'value') throw new Error('Pivot filters require a configured label field')
  return field
}

function removeCategories(field: PivotFieldModel, categories: readonly PivotFieldFilterCategory[]): void {
  for (const category of categories) for (const key of properties[category]) delete field[key]
}

function publishFields(pivot: PivotTableModel, fields: PivotFieldModel[]): boolean {
  if (JSON.stringify(pivot.fields) === JSON.stringify(fields)) return false
  pivot.fields = fields
  return true
}

/** Latest active category replaces its source siblings; disabled multiple mode also retires other categories. */
export function replacePivotFilterCategory(pivot: PivotTableModel, fieldId: string, category: PivotFieldFilterCategory,
  state: PivotFieldFilterState, restrictive: boolean): boolean {
  const source = pivotFilterSourceKey(pivot, labelField(pivot, fieldId))
  const clear: PivotFieldFilterCategory[] = restrictive && pivot.options.allowMultipleFilters !== true ? ['member', 'condition', 'value'] : [category]
  const fields = pivot.fields.map(field => {
    if (field.area === 'value' || pivotFilterSourceKey(pivot, field) !== source) return field
    const next = { ...field }; removeCategories(next, clear)
    if (next.id === fieldId) Object.assign(next, state)
    return next
  })
  return publishFields(pivot, fields)
}

/** A date/label reset removes its own condition only; clear-all operates across every instance of this source. */
export function clearPivotFieldFilterState(pivot: PivotTableModel, fieldId: string, kind: 'date' | 'label' | 'value' | 'all'): boolean {
  const field = labelField(pivot, fieldId), source = pivotFilterSourceKey(pivot, field)
  const fields = pivot.fields.map(candidate => {
    if (candidate.area === 'value' || (kind === 'all' ? pivotFilterSourceKey(pivot, candidate) !== source : candidate.id !== fieldId)) return candidate
    const next = { ...candidate }
    if (kind === 'all') removeCategories(next, ['member', 'condition', 'value'])
    else if (kind === 'date') { delete next.dateFilter; delete next.timelineDateRange; delete next.timelineDateCondition }
    else if (kind === 'label') delete next.labelFilter
    else delete next.valueFilter
    return next
  })
  return publishFields(pivot, fields)
}

export function clearPivotReportFilterState(pivot: PivotTableModel): boolean {
  return publishFields(pivot, pivot.fields.map(field => {
    if (field.area === 'value') return field
    const next = { ...field }; removeCategories(next, ['member', 'condition', 'value']); return next
  }))
}
