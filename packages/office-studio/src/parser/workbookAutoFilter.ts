import type { CellRange, CellScalar, SheetAutoFilter, SheetFilterCriteria, SheetFilterConditionOperator } from '../domain/model.ts'
import { normalizePivotRecordSelection } from '../features/pivot/services/pivot-filter-predicate.ts'
import type { PivotRecordSelection } from '../domain/model.ts'

const operators = new Set<SheetFilterConditionOperator>([
  'equals', 'doesNotEqual', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual',
  'between', 'notBetween', 'contains', 'doesNotContain', 'beginsWith', 'endsWith', 'isBlank', 'isNotBlank', 'isError',
])

function object(value: unknown, label: string): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError(`${label}必须是对象。`)
  return value as Record<string, unknown>
}

function index(value: unknown, maximum: number, label: string): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 0 || value > maximum) throw new RangeError(`${label}超出工作表范围。`)
  return value
}

function scalar(value: unknown): CellScalar {
  if (value === null || typeof value === 'string' || typeof value === 'boolean' || (typeof value === 'number' && Number.isFinite(value))) return value
  throw new TypeError('筛选条件值必须是有限数字、文本、布尔值或空值。')
}

function criterion(value: unknown): SheetFilterCriteria {
  const raw = object(value, '筛选条件')
  if (raw.kind === 'values') {
    if (!Array.isArray(raw.selectedItems)) throw new TypeError('值筛选缺少成员数组。')
    if (raw.selectionMode !== undefined && raw.selectionMode !== 'include' && raw.selectionMode !== 'exclude') throw new TypeError('值筛选选择模式无效。')
    if (raw.recordDate1904 !== undefined && typeof raw.recordDate1904 !== 'boolean') throw new TypeError('原始成员日期系统无效。')
    return { kind: 'values', selectedItems: raw.selectedItems.map(scalar), ...(raw.selectionMode === 'exclude' ? { selectionMode: 'exclude' as const } : {}),
      ...(raw.recordSelection !== undefined ? { recordSelection: normalizePivotRecordSelection(raw.recordSelection as PivotRecordSelection) } : {}),
      ...(raw.recordDate1904 !== undefined ? { recordDate1904: raw.recordDate1904 } : {}) }
  }
  if (raw.kind !== 'condition' || typeof raw.operator !== 'string' || !operators.has(raw.operator as SheetFilterConditionOperator)) throw new TypeError('筛选条件运算符无效。')
  const operator = raw.operator as SheetFilterConditionOperator
  const noOperand = operator === 'isBlank' || operator === 'isNotBlank' || operator === 'isError'
  if (!noOperand && !Object.hasOwn(raw, 'value')) throw new TypeError('筛选条件缺少比较值。')
  if ((operator === 'between' || operator === 'notBetween') && !Object.hasOwn(raw, 'secondValue')) throw new TypeError('区间筛选缺少第二个比较值。')
  if (raw.caseSensitive !== undefined && typeof raw.caseSensitive !== 'boolean') throw new TypeError('筛选大小写标志无效。')
  return {
    kind: 'condition', operator,
    ...(Object.hasOwn(raw, 'value') ? { value: scalar(raw.value) } : {}),
    ...(Object.hasOwn(raw, 'secondValue') ? { secondValue: scalar(raw.secondValue) } : {}),
    ...(typeof raw.caseSensitive === 'boolean' ? { caseSensitive: raw.caseSensitive } : {}),
  }
}

/** Decode the parser boundary without losing scalar types or explicit null. */
export function normalizeWorkbookAutoFilter(value: unknown): SheetAutoFilter | undefined {
  if (value === undefined || value === null) return undefined
  const raw = object(value, '工作表筛选'), sourceRange = object(raw.range, '筛选范围')
  const range: CellRange = {
    startRow: index(sourceRange.startRow, 1_048_575, '筛选起始行'), endRow: index(sourceRange.endRow, 1_048_575, '筛选结束行'),
    startCol: index(sourceRange.startCol, 16_383, '筛选起始列'), endCol: index(sourceRange.endCol, 16_383, '筛选结束列'),
  }
  if (range.startRow >= range.endRow || range.startCol > range.endCol) throw new RangeError('筛选范围必须包含标题行和数据行。')
  const sourceColumns = object(raw.columns, '筛选列'), columns: Record<number, SheetFilterCriteria> = {}
  for (const [key, value] of Object.entries(sourceColumns)) {
    if (!/^(?:0|[1-9]\d*)$/.test(key)) throw new TypeError('筛选列索引无效。')
    const column = index(Number(key), 16_383, '筛选列')
    if (column < range.startCol || column > range.endCol) throw new RangeError('筛选列不在筛选范围内。')
    columns[column] = criterion(value)
  }
  if (!Array.isArray(raw.filteredRows)) throw new TypeError('筛选隐藏行必须是数组。')
  const rows = new Set<number>()
  for (const value of raw.filteredRows) {
    const row = index(value, 1_048_575, '筛选隐藏行')
    if (row <= range.startRow || row > range.endRow || rows.has(row)) throw new RangeError('筛选隐藏行重复或不在数据范围内。')
    rows.add(row)
  }
  return { range, columns, filteredRows: [...rows].sort((left, right) => left - right) }
}
