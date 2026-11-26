import type { PivotDateFilter, PivotDateFilterOperator, PivotLabelConditionFilter, PivotLabelConditionOperator } from '../../../domain/model.ts'
import { normalizePivotDateFilter, normalizePivotLabelConditionFilter } from './pivot-filter-predicate.ts'

/** Fixed v1.6.0 ST_PivotFilterOperatorEnum and CustomFilter=1 at the public boundary. */
export type PivotDateFilterCode = 16 | 17 | 18 | 19 | 20 | 21 | 22 | 23
export type PivotLabelConditionCode = 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 | 10 | 11 | 12 | 13 | 14
export interface PivotDateFilterApiInfo {
  type?: 1
  operator: PivotDateFilterCode
  expected: number | [number, number]
  useWholeDay?: boolean
}
export interface PivotLabelConditionApiInfo {
  type?: 1
  operator: PivotLabelConditionCode
  expected: string | [string, string]
}
export type PivotDateFilterInput = PivotDateFilter | PivotDateFilterApiInfo
export type PivotLabelConditionInput = PivotLabelConditionFilter | PivotLabelConditionApiInfo

const dateOperators: Readonly<Partial<Record<number, PivotDateFilterOperator>>> = {
  16: 'dateBetween', 17: 'dateEqual', 18: 'dateNewerThan', 19: 'dateNewerThanOrEqual',
  20: 'dateNotBetween', 21: 'dateNotEqual', 22: 'dateOlderThan', 23: 'dateOlderThanOrEqual',
}
const labelOperators: Readonly<Partial<Record<number, PivotLabelConditionOperator>>> = {
  1: 'captionBeginsWith', 2: 'captionBetween', 3: 'captionContains', 4: 'captionEndsWith', 5: 'captionEqual',
  6: 'captionGreaterThan', 7: 'captionGreaterThanOrEqual', 8: 'captionLessThan', 9: 'captionLessThanOrEqual',
  10: 'captionNotBeginsWith', 11: 'captionNotBetween', 12: 'captionNotContains', 13: 'captionNotEndsWith', 14: 'captionNotEqual',
}

function validateFilterInfo(value: object, keys: readonly string[]): void {
  if (Array.isArray(value) || Object.keys(value).some(key => !keys.includes(key))
    || 'type' in value && value.type !== 1) throw new TypeError('透视条件筛选信息无效。')
}

export function normalizePivotDateFilterInput(value: PivotDateFilterInput | null | undefined): PivotDateFilter | null {
  if (value == null) return null
  if (typeof value !== 'object') throw new TypeError('透视日期筛选必须为条件对象。')
  validateFilterInfo(value, ['type', 'operator', 'expected', 'useWholeDay'])
  const operator = typeof value.operator === 'number' ? dateOperators[value.operator] : value.operator
  if (!operator) throw new TypeError('透视日期筛选操作符无效。')
  return normalizePivotDateFilter({ operator, expected: value.expected, useWholeDay: value.useWholeDay })
}

export function normalizePivotLabelConditionInput(value: PivotLabelConditionInput | null | undefined): PivotLabelConditionFilter | null {
  if (value == null) return null
  if (typeof value !== 'object') throw new TypeError('透视标签筛选必须为条件对象。')
  validateFilterInfo(value, ['type', 'operator', 'expected'])
  const operator = typeof value.operator === 'number' ? labelOperators[value.operator] : value.operator
  if (!operator) throw new TypeError('透视标签筛选操作符无效。')
  return normalizePivotLabelConditionFilter({ operator, expected: value.expected })
}
