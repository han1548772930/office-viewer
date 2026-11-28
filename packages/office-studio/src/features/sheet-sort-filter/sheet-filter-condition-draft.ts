import type { SheetConditionFilter, SheetCompoundFilter, SheetFilterConditionOperator } from '../../domain/model.ts'
import { parseDataPanelScalar } from './sheet-data-panel-state.ts'

export interface SheetFilterConditionDraft {
  operator: SheetFilterConditionOperator | 'none'
  valueType: 'text' | 'number' | 'boolean'
  value: string
  secondValue: string
  caseSensitive: boolean
}
export const SHEET_FILTER_CONDITION_OPTIONS: Array<{ label: string; value: SheetFilterConditionOperator }> = [
  { label: '等于', value: 'equals' }, { label: '不等于', value: 'doesNotEqual' },
  { label: '大于', value: 'greaterThan' }, { label: '大于或等于', value: 'greaterThanOrEqual' },
  { label: '小于', value: 'lessThan' }, { label: '小于或等于', value: 'lessThanOrEqual' },
  { label: '介于', value: 'between' }, { label: '不介于', value: 'notBetween' },
  { label: '包含', value: 'contains' }, { label: '不包含', value: 'doesNotContain' },
  { label: '开头是', value: 'beginsWith' }, { label: '结尾是', value: 'endsWith' },
  { label: '为空', value: 'isBlank' }, { label: '非空', value: 'isNotBlank' }, { label: '为错误值', value: 'isError' },
]
export function sheetFilterConditionDraft(condition?: SheetConditionFilter): SheetFilterConditionDraft {
  return { operator: condition?.operator ?? 'none', valueType: typeof condition?.value === 'number' ? 'number'
    : typeof condition?.value === 'boolean' ? 'boolean' : 'text', value: condition?.value == null ? '' : String(condition.value),
    secondValue: condition?.secondValue == null ? '' : String(condition.secondValue), caseSensitive: condition?.caseSensitive === true }
}
export function sheetFilterConditionFromDraft(draft: SheetFilterConditionDraft): SheetConditionFilter | undefined {
  if (draft.operator === 'none') return undefined
  const needsValue = !['isBlank', 'isNotBlank', 'isError'].includes(draft.operator)
  const needsSecond = draft.operator === 'between' || draft.operator === 'notBetween'
  return { kind: 'condition', operator: draft.operator, ...(needsValue ? { value: parseDataPanelScalar(draft.value, draft.valueType) } : {}),
    ...(needsSecond ? { secondValue: parseDataPanelScalar(draft.secondValue, draft.valueType) } : {}), caseSensitive: draft.caseSensitive }
}
/** The fixed custom editor clears two None rows, and emits a single condition when only one row is configured. */
export function sheetCompoundFilterFromDrafts(join: 'and' | 'or', drafts: readonly SheetFilterConditionDraft[]): SheetConditionFilter | SheetCompoundFilter | undefined {
  if (join !== 'and' && join !== 'or' || drafts.length !== 2) throw new TypeError('自定义筛选需要两行条件和有效的连接方式。')
  const first = sheetFilterConditionFromDraft(drafts[0]!), second = sheetFilterConditionFromDraft(drafts[1]!)
  return first && second ? { kind: 'compound', join, conditions: [first, second] } : first ?? second
}
