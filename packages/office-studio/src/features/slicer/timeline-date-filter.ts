import type { PivotDateFilter, PivotDateFilterOperator, TimelineLevel } from '../../domain/model.ts'
import { normalizeRecordValue, type RecordValue } from '../../domain/record-value.ts'
import { TimelineDateFilterOperator, type ITimelineDateCondition, type ITimelineRange, type TimelinePublicLevel } from '../../domain/timeline-filters.ts'

const record = (value: unknown): value is Record<string, unknown> => value !== null && typeof value === 'object' && !Array.isArray(value)
const finite = (value: unknown): value is number => typeof value === 'number' && Number.isFinite(value)
const levels: readonly TimelineLevel[] = ['year', 'quarter', 'month', 'day']
const operators = Object.values(TimelineDateFilterOperator)
export function normalizeTimelineLevel(value: unknown): TimelineLevel {
  if (typeof value === 'number' && Number.isInteger(value) && value >= 0 && value < levels.length) return levels[value]!
  if (typeof value === 'string' && (levels as readonly string[]).includes(value)) return value as TimelineLevel
  throw new TypeError('Invalid Timeline display or selection level')
}
export function toTimelinePublicLevel(value: TimelineLevel | TimelinePublicLevel): TimelinePublicLevel {
  return levels.indexOf(normalizeTimelineLevel(value)) as TimelinePublicLevel
}
function range(value: unknown, allowEmpty: boolean): ITimelineRange {
  if (!record(value) || Object.keys(value).some(key => !['startDate', 'endDate', 'useWholeDay'].includes(key))
    || !finite(value.startDate) || !finite(value.endDate)) throw new TypeError('Timeline range requires finite Excel serial values')
  if (value.useWholeDay !== undefined && typeof value.useWholeDay !== 'boolean') throw new TypeError('Invalid Timeline whole-day mode')
  if (allowEmpty ? value.startDate > value.endDate : value.startDate >= value.endDate) throw new RangeError('Timeline range startDate must precede endDate')
  return { startDate: value.startDate, endDate: value.endDate, ...(value.useWholeDay !== undefined ? { useWholeDay: value.useWholeDay } : {}) }
}
/** Public setters and persisted selection require a nonempty half-open range. Author fractional bounds remain retained. */
export function normalizeTimelineDateRange(value: unknown): ITimelineRange { return range(value, false) }
export function effectiveTimelineDateRange(value: ITimelineRange): { startDate: number; endDate: number } {
  const normalized = normalizeTimelineDateRange(value)
  return { startDate: normalized.useWholeDay === true ? Math.floor(normalized.startDate) : normalized.startDate,
    endDate: normalized.useWholeDay === true ? Math.ceil(normalized.endDate) : normalized.endDate }
}
/** The standalone published compiler also accepts equal bounds as an empty predicate; setters still reject them. */
export function compileTimelineSerialRange(value: ITimelineRange): (serial: number) => boolean {
  const normalized = range(value, true), empty = normalized.startDate === normalized.endDate
  const start = normalized.useWholeDay === true ? Math.floor(normalized.startDate) : normalized.startDate
  const end = normalized.useWholeDay === true ? Math.ceil(normalized.endDate) : normalized.endDate
  return serial => !empty && Number.isFinite(serial) && serial >= start && serial < end
}
export function compileTimelineRange(value: ITimelineRange): (member: RecordValue) => boolean {
  const matches = compileTimelineSerialRange(value)
  return input => { const member = normalizeRecordValue(input); return member.type === 'date' && matches(member.value) }
}
export function normalizeTimelineDateCondition(value: unknown): ITimelineDateCondition {
  if (!record(value) || Object.keys(value).some(key => !['operator', 'expected', 'useWholeDay'].includes(key))
    || typeof value.operator !== 'string' || !(operators as readonly string[]).includes(value.operator)
    || value.useWholeDay !== undefined && typeof value.useWholeDay !== 'boolean') throw new TypeError('Invalid Timeline date condition')
  const between = value.operator === 'between' || value.operator === 'notBetween'
  let expected: number | [number, number]
  if (between) {
    if (!Array.isArray(value.expected) || value.expected.length !== 2 || !value.expected.every(finite)
      || value.expected[0] > value.expected[1]) throw new RangeError('Timeline date operands must be finite and ordered')
    expected = [value.expected[0], value.expected[1]]
  } else {
    if (!finite(value.expected)) throw new RangeError('Timeline single-date condition requires a finite operand')
    expected = value.expected
  }
  return { operator: value.operator as TimelineDateFilterOperator, expected,
    ...(value.useWholeDay !== undefined ? { useWholeDay: value.useWholeDay } : {}) }
}
/** Shared numeric predicate for Pivot original date fields and the public typed Timeline compiler. */
export function compileTimelineSerialCondition(value: ITimelineDateCondition): (serial: number) => boolean {
  const filter = normalizeTimelineDateCondition(value), wholeDay = filter.useWholeDay === true
  const operands = Array.isArray(filter.expected) ? filter.expected : [filter.expected, filter.expected]
  const lower = wholeDay ? Math.floor(operands[0]!) : operands[0]!, upper = wholeDay ? Math.floor(operands[1]!) : operands[1]!
  return serial => {
    if (!Number.isFinite(serial)) return false
    const actual = wholeDay ? Math.floor(serial) : serial
    switch (filter.operator) {
      case 'between': return actual >= lower && actual <= upper
      case 'notBetween': return actual < lower || actual > upper
      case 'equal': return actual === lower
      case 'notEqual': return actual !== lower
      case 'newerThan': return actual > lower
      case 'newerThanOrEqual': return actual >= lower
      case 'olderThan': return actual < lower
      case 'olderThanOrEqual': return actual <= lower
    }
    return false
  }
}
export function compileTimelineCondition(value: ITimelineDateCondition): (member: RecordValue) => boolean {
  const matches = compileTimelineSerialCondition(value)
  return input => { const member = normalizeRecordValue(input); return member.type === 'date' && matches(member.value) }
}

/** Calendar block painting uses interval intersection, never an end-minus-epsilon sample. */
export function timelineConditionIntersectsPeriod(value: ITimelineDateCondition, period: ITimelineRange): boolean {
  const filter = normalizeTimelineDateCondition(value), block = normalizeTimelineDateRange(period), whole = filter.useWholeDay === true
  const operands = Array.isArray(filter.expected) ? filter.expected : [filter.expected, filter.expected]
  const lower = whole ? Math.floor(operands[0]!) : operands[0]!, upper = whole ? Math.floor(operands[1]!) : operands[1]!
  switch (filter.operator) {
    case 'between': return block.endDate > lower && (whole ? block.startDate < upper + 1 : block.startDate <= upper)
    case 'notBetween': return block.startDate < lower || block.endDate > (whole ? upper + 1 : upper)
    case 'equal': return whole ? block.endDate > lower && block.startDate < lower + 1 : block.startDate <= lower && block.endDate > lower
    case 'notEqual': return !whole || block.startDate < lower || block.endDate > lower + 1
    case 'newerThan': return block.endDate > lower + (whole ? 1 : 0)
    case 'newerThanOrEqual': return block.endDate > lower
    case 'olderThan': return block.startDate < lower
    case 'olderThanOrEqual': return whole ? block.startDate < lower + 1 : block.startDate <= lower
  }
}
const pivotOperators: Record<TimelineDateFilterOperator, PivotDateFilterOperator> = {
  between: 'dateBetween', notBetween: 'dateNotBetween', equal: 'dateEqual', notEqual: 'dateNotEqual',
  newerThan: 'dateNewerThan', newerThanOrEqual: 'dateNewerThanOrEqual', olderThan: 'dateOlderThan', olderThanOrEqual: 'dateOlderThanOrEqual',
}
export function timelineDateConditionToPivot(value: ITimelineDateCondition): PivotDateFilter {
  const normalized = normalizeTimelineDateCondition(value)
  return { ...normalized, operator: pivotOperators[normalized.operator] }
}
export function timelineDateConditionFromPivot(value: PivotDateFilter): ITimelineDateCondition {
  const operator = operators.find(candidate => pivotOperators[candidate] === value.operator)
  if (operator === undefined) throw new TypeError('Unsupported Timeline date operator')
  return normalizeTimelineDateCondition({ ...value, operator })
}
