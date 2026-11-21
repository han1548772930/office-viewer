import type { CellScalar, PivotDateGroupType, PivotFieldFilter, PivotFieldModel, PivotFilterOperator, PivotManualGroup, PivotNumberGroup, PivotTableModel, WorkbookModel } from '../../../domain/model'

export type PivotFieldToolChange =
  | { kind: 'valueFilter'; filter: PivotFieldFilter | null }
  | { kind: 'numberGroup'; settings: PivotNumberGroup | null }
  | { kind: 'dateGroup'; type: PivotDateGroupType | null }
  | { kind: 'manualGroup'; groups: PivotManualGroup[] }

export const PIVOT_VALUE_FILTER_OPTIONS: Array<{ label: string; value: PivotFilterOperator }> = [
  { label: '等于', value: 'valueEqual' }, { label: '不等于', value: 'valueNotEqual' },
  { label: '大于', value: 'valueGreaterThan' }, { label: '大于或等于', value: 'valueGreaterThanOrEqual' },
  { label: '小于', value: 'valueLessThan' }, { label: '小于或等于', value: 'valueLessThanOrEqual' },
  { label: '介于', value: 'valueBetween' }, { label: '不介于', value: 'valueNotBetween' },
]

export interface PivotMemberOption { key: string; label: string; value: CellScalar }

/** Keys retain scalar types so a text digit and a numeric digit remain distinct. */
export function readPivotFieldMembers(workbook: WorkbookModel, pivot: PivotTableModel, field: PivotFieldModel): PivotMemberOption[] {
  const sheet = workbook.sheets.find(candidate => candidate.id === pivot.source.subUnitId)
  if (!sheet) throw new Error('透视表数据源工作表已不存在。')
  const members = new Map<string, PivotMemberOption>()
  for (let row = pivot.source.range.startRow + 1; row <= pivot.source.range.endRow; row++) {
    const value = sheet.cells[`${row}:${field.sourceColumn}`]?.value ?? null
    const key = JSON.stringify([typeof value, value])
    if (!members.has(key)) members.set(key, { key, label: value == null || value === '' ? '（空白）' : String(value), value })
  }
  return [...members.values()]
}

export function defaultPivotNumberGroup(members: PivotMemberOption[]): PivotNumberGroup | undefined {
  const values = members.map(member => member.value).filter(value => value !== null && value !== '')
  if (!values.length || values.some(value => typeof value !== 'number' || !Number.isFinite(value))) return undefined
  const numbers = values as number[]
  const start = numbers.reduce((minimum, value) => Math.min(minimum, value), Infinity)
  const maximum = numbers.reduce((max, value) => Math.max(max, value), -Infinity)
  const end = maximum === start ? start + 1 : maximum
  return { start, end, interval: Math.max(1, Math.ceil((end - start) / 10)) }
}

export function pivotValueFilterInput(operator: PivotFilterOperator, valueFieldId: string, expected: number | null, expected2: number | null): PivotFieldFilter {
  if (!PIVOT_VALUE_FILTER_OPTIONS.some(option => option.value === operator)) throw new Error('请选择有效的筛选条件。')
  if (!valueFieldId) throw new Error('请选择用于筛选的值字段。')
  if (expected === null || !Number.isFinite(expected)) throw new Error('请输入有效的筛选数值。')
  const between = operator === 'valueBetween' || operator === 'valueNotBetween'
  if (between && (expected2 === null || !Number.isFinite(expected2))) throw new Error('请输入有效的第二个筛选数值。')
  return { operator, valueFieldId, expected, ...(between && expected2 !== null ? { expected2 } : {}) }
}
