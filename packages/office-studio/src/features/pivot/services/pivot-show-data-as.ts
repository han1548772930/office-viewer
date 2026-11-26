import type { CellScalar, PivotFieldModel, PivotShowDataAsInfo, PivotShowDataAsType, PivotTableModel } from '../../../domain/model.ts'

export interface PivotShowDataAsRule {
  type: PivotShowDataAsType
  label: string
  requiresParameters: boolean
  baseItemMode: 'none' | 'item' | 'item-or-relative'
  parentFieldOnly?: boolean
  percentage?: boolean
  compatibilityOnly?: boolean
}

/** First fifteen rules are the fixed 1.6.0 UI contract; the last has only a historical 1.5.2 enum source. */
export const PIVOT_SHOW_DATA_AS_RULES: readonly PivotShowDataAsRule[] = [
  { type: 'normal', label: '无计算', requiresParameters: false, baseItemMode: 'none' },
  { type: 'percentOfGrandTotal', label: '总计的百分比', requiresParameters: false, baseItemMode: 'none', percentage: true },
  { type: 'percentOfRowTotal', label: '行汇总的百分比', requiresParameters: false, baseItemMode: 'none', percentage: true },
  { type: 'percentOfColumnTotal', label: '列汇总的百分比', requiresParameters: false, baseItemMode: 'none', percentage: true },
  { type: 'percentOfParentRowTotal', label: '父行汇总的百分比', requiresParameters: false, baseItemMode: 'none', percentage: true },
  { type: 'percentOfParentColumnTotal', label: '父列汇总的百分比', requiresParameters: false, baseItemMode: 'none', percentage: true },
  { type: 'index', label: '指数', requiresParameters: false, baseItemMode: 'none' },
  { type: 'percentOfParent', label: '指定父项汇总的百分比', requiresParameters: true, baseItemMode: 'none', parentFieldOnly: true, percentage: true },
  { type: 'percentOf', label: '指定项的百分比', requiresParameters: true, baseItemMode: 'item', percentage: true },
  { type: 'differenceFrom', label: '与指定项的差值', requiresParameters: true, baseItemMode: 'item-or-relative' },
  { type: 'percentDifferenceFrom', label: '与指定项差值的百分比', requiresParameters: true, baseItemMode: 'item-or-relative', percentage: true },
  { type: 'runningTotal', label: '按字段累计', requiresParameters: true, baseItemMode: 'none' },
  { type: 'percentRunningTotal', label: '按字段累计百分比', requiresParameters: true, baseItemMode: 'none', percentage: true },
  { type: 'rankAscending', label: '升序排名', requiresParameters: true, baseItemMode: 'none' },
  { type: 'rankDescending', label: '降序排名', requiresParameters: true, baseItemMode: 'none' },
  { type: 'percentOfParentGrandTotal', label: '父行列汇总的百分比', requiresParameters: false, baseItemMode: 'none', percentage: true, compatibilityOnly: true },
]

export type PivotShowDataAsScope = Record<string, CellScalar>

/** The aggregation callback always belongs to the existing Pivot engine; this service only transforms its results. */
export interface PivotShowDataAsCellContext {
  value: CellScalar
  rowFields: readonly PivotFieldModel[]
  columnFields: readonly PivotFieldModel[]
  rowFilters: PivotShowDataAsScope
  columnFilters: PivotShowDataAsScope
  aggregate: (filters: PivotShowDataAsScope) => CellScalar
  baseItems: (fieldId: string, filters: PivotShowDataAsScope) => readonly CellScalar[]
}

const record = (value: unknown): value is Record<string, unknown> => !!value && typeof value === 'object' && !Array.isArray(value)
const own = (value: object, key: PropertyKey): boolean => Object.hasOwn(value, key)
const key = (value: CellScalar): string => String(value ?? '')

export function getPivotShowDataAsRule(type: PivotShowDataAsType): PivotShowDataAsRule {
  const rule = PIVOT_SHOW_DATA_AS_RULES.find(rule => rule.type === type)
  if (!rule) throw new TypeError('透视值显示方式无效。')
  return rule
}

export function getPivotShowDataAsBaseFieldIds(pivot: PivotTableModel, type: PivotShowDataAsType): string[] {
  const rule = getPivotShowDataAsRule(type)
  if (!rule.requiresParameters) return []
  const axis = (area: 'row' | 'column') => pivot.fields.filter(field => field.area === area).sort((a, b) => a.position - b.position)
  const rows = axis('row'), columns = axis('column')
  return (rule.parentFieldOnly ? [...rows.slice(0, -1), ...columns.slice(0, -1)] : [...rows, ...columns]).map(field => field.id)
}

/** baseFieldId is report-local in the public DTO; a unique source ID remains a fallback among axis labels. */
export function resolvePivotShowDataAsBaseField(pivot: PivotTableModel, id: string): PivotFieldModel {
  if (typeof id !== 'string' || !id.trim()) throw new TypeError('值显示方式需要有效的基准字段。')
  const fields = pivot.fields.filter(field => field.area === 'row' || field.area === 'column')
  const exact = fields.find(field => field.id === id)
  if (exact) return exact
  const source = fields.filter(field => (field.sourceFieldId ?? String(field.sourceColumn - pivot.source.range.startCol)) === id)
  if (source.length === 1) return source[0]!
  if (source.length > 1) throw new Error('基准来源存在多个标签实例，请指定报告字段 ID。')
  throw new Error('值显示方式的基准字段必须是已配置的行或列字段。')
}

export function normalizePivotShowDataAs(pivot: PivotTableModel, input: PivotShowDataAsInfo | null,
  baseItems?: (field: PivotFieldModel) => readonly CellScalar[]): PivotShowDataAsInfo | undefined {
  if (input === null) return undefined
  if (!record(input) || Object.keys(input).some(property => !['type', 'baseFieldId', 'baseItem', 'baseItemType'].includes(property))
    || typeof input.baseFieldId !== 'string' || typeof input.baseItem !== 'string') throw new TypeError('透视值显示配置无效。')
  const rule = getPivotShowDataAsRule(input.type)
  if (input.baseItemType !== undefined && !['item', 'previous', 'next'].includes(input.baseItemType)) throw new TypeError('基准项类型无效。')
  if (!rule.requiresParameters) {
    if (input.baseFieldId !== '' || input.baseItem !== '' || input.baseItemType !== undefined) throw new TypeError('当前值显示方式不接受基准字段或基准项。')
    return rule.type === 'normal' ? undefined : { type: rule.type, baseFieldId: '', baseItem: '' }
  }
  const field = resolvePivotShowDataAsBaseField(pivot, input.baseFieldId)
  if (!getPivotShowDataAsBaseFieldIds(pivot, rule.type).includes(field.id)) throw new Error('该基准字段没有可用的子层级。')
  if (rule.baseItemMode === 'none') {
    if (input.baseItem !== '' || input.baseItemType !== undefined) throw new TypeError('当前值显示方式不接受基准项。')
    return { type: rule.type, baseFieldId: field.id, baseItem: '' }
  }
  const relative = input.baseItemType === 'previous' || input.baseItemType === 'next'
  if (relative) {
    if (rule.baseItemMode !== 'item-or-relative' || input.baseItem !== '') throw new TypeError('相对基准只支持差值，并且不能同时指定基准项。')
    return { type: rule.type, baseFieldId: field.id, baseItem: '', baseItemType: input.baseItemType }
  }
  if (input.baseItem === '') throw new TypeError('当前值显示方式需要一个基准项。')
  if (baseItems) {
    const matching = baseItems(field).filter(item => key(item) === input.baseItem)
    if (!matching.length) throw new Error('值显示方式的基准项不在真实来源成员中。')
    if (matching.length > 1) throw new Error('基准项的显示键不唯一，无法明确选择成员。')
  }
  return { type: rule.type, baseFieldId: field.id, baseItem: input.baseItem }
}

export function pivotShowDataAsNumberFormat(value: PivotShowDataAsInfo | undefined): string | undefined {
  if (!value) return undefined
  const rule = getPivotShowDataAsRule(value.type)
  return rule.percentage ? '0.00%' : value.type === 'rankAscending' || value.type === 'rankDescending' ? '0' : undefined
}

function divide(value: number, denominator: CellScalar): CellScalar {
  if (typeof denominator !== 'number') return denominator === null ? '#N/A' : denominator
  if (!Number.isFinite(value) || !Number.isFinite(denominator)) return '#NUM!'
  if (denominator === 0) return '#DIV/0!'
  const result = value / denominator
  return Number.isFinite(result) ? result : '#NUM!'
}

function parentScope(filters: PivotShowDataAsScope, fields: readonly PivotFieldModel[]): PivotShowDataAsScope {
  const result = { ...filters }, leaf = [...fields].reverse().find(field => own(result, field.id))
  if (leaf) delete result[leaf.id]
  return result
}

function scopeThrough(filters: PivotShowDataAsScope, fields: readonly PivotFieldModel[], baseFieldId: string): PivotShowDataAsScope {
  const result = { ...filters }, index = fields.findIndex(field => field.id === baseFieldId)
  for (const field of fields.slice(index + 1)) delete result[field.id]
  return result
}

/** Derive one display value from original aggregates; totals never sum already transformed cells. */
export function applyPivotShowDataAs(value: PivotShowDataAsInfo | undefined, context: PivotShowDataAsCellContext): CellScalar {
  if (!value || value.type === 'normal') return context.value
  getPivotShowDataAsRule(value.type)
  if (context.value === null) return null
  if (typeof context.value !== 'number') return context.value
  if (!Number.isFinite(context.value)) return '#NUM!'
  const current = context.value, row = context.rowFilters, column = context.columnFilters, scope = { ...row, ...column }
  switch (value.type) {
    case 'percentOfGrandTotal': return divide(current, context.aggregate({}))
    case 'percentOfRowTotal': return divide(current, context.aggregate(row))
    case 'percentOfColumnTotal': return divide(current, context.aggregate(column))
    case 'percentOfParentRowTotal': return divide(current, context.aggregate({ ...parentScope(row, context.rowFields), ...column }))
    case 'percentOfParentColumnTotal': return divide(current, context.aggregate({ ...row, ...parentScope(column, context.columnFields) }))
    case 'percentOfParentGrandTotal': return divide(current, context.aggregate({ ...parentScope(row, context.rowFields), ...parentScope(column, context.columnFields) }))
    case 'index': {
      const grand = context.aggregate({}), rowTotal = context.aggregate(row), columnTotal = context.aggregate(column)
      if (typeof grand !== 'number' || typeof rowTotal !== 'number' || typeof columnTotal !== 'number') return '#N/A'
      return divide(current * grand, rowTotal * columnTotal)
    }
    case 'percentOfParent': {
      const fields = context.rowFields.some(field => field.id === value.baseFieldId) ? context.rowFields : context.columnFields
      return divide(current, context.aggregate({ ...scopeThrough(scope, fields, value.baseFieldId) }))
    }
    case 'percentOf':
    case 'differenceFrom':
    case 'percentDifferenceFrom': {
      if (!own(scope, value.baseFieldId)) return null
      const items = context.baseItems(value.baseFieldId, scope), item = scope[value.baseFieldId]!
      let reference: CellScalar | undefined
      if (value.baseItemType === 'previous' || value.baseItemType === 'next') {
        const index = items.findIndex(candidate => Object.is(candidate, item))
        if (index < 0) return '#REF!'
        const offset = value.baseItemType === 'previous' ? -1 : 1
        if (index + offset < 0 || index + offset >= items.length) return '#N/A'
        reference = items[index + offset]
      } else {
        const matching = items.filter(candidate => key(candidate) === value.baseItem)
        if (matching.length !== 1) return '#N/A'
        reference = matching[0]
      }
      if (reference === undefined) return '#N/A'
      if (value.type !== 'percentOf' && Object.is(reference, item)) return null
      const base = context.aggregate({ ...scope, [value.baseFieldId]: reference })
      if (value.type === 'percentOf') return divide(current, base)
      if (typeof base !== 'number') return base === null ? '#N/A' : base
      if (!Number.isFinite(base)) return '#NUM!'
      const difference = current - base
      if (!Number.isFinite(difference)) return '#NUM!'
      return value.type === 'differenceFrom' ? difference : divide(difference, base)
    }
    case 'runningTotal':
    case 'percentRunningTotal':
    case 'rankAscending':
    case 'rankDescending': {
      const items = context.baseItems(value.baseFieldId, scope)
      const peers = items.map(item => ({ item, value: context.aggregate({ ...scope, [value.baseFieldId]: item }) }))
      if (peers.some(peer => typeof peer.value === 'number' && !Number.isFinite(peer.value))) return '#NUM!'
      const numeric = peers.filter((peer): peer is { item: CellScalar; value: number } => typeof peer.value === 'number')
      if (!numeric.length) return null
      if (value.type === 'rankAscending' || value.type === 'rankDescending') {
        if (!own(scope, value.baseFieldId)) return null
        return 1 + numeric.filter(peer => value.type === 'rankAscending' ? peer.value < current : peer.value > current).length
      }
      const index = own(scope, value.baseFieldId) ? items.findIndex(item => Object.is(item, scope[value.baseFieldId])) : items.length - 1
      if (index < 0) return '#REF!'
      const cumulative = peers.slice(0, index + 1).reduce((total, peer) => total + (typeof peer.value === 'number' ? peer.value : 0), 0)
      if (value.type === 'runningTotal') return Number.isFinite(cumulative) ? cumulative : '#NUM!'
      return divide(cumulative, numeric.reduce((total, peer) => total + peer.value, 0))
    }
  }
}
