import assert from 'node:assert/strict'
import { test } from 'vitest'
import { TimelineDateFilterOperator, TimelinePublicLevel } from '../src/domain/timeline-filters.ts'
import { compileTimelineCondition, compileTimelineRange, compileTimelineSerialCondition, compileTimelineSerialRange,
  effectiveTimelineDateRange, normalizeTimelineDateCondition, normalizeTimelineDateRange, normalizeTimelineLevel,
  timelineDateConditionFromPivot, timelineDateConditionToPivot, toTimelinePublicLevel } from '../src/features/slicer/timeline-date-filter.ts'

// Pure public DTO/predicate drafts. No test, SDK or browser execution has taken place.
const modes = [undefined, false, true]
const withMode = (value, useWholeDay) => ({ ...value, ...(useWholeDay === undefined ? {} : { useWholeDay }) })
const operators = [
  { operator: 'between', expected: [60.25, 60.75], exact: [false, false, true, true, false, false], days: [false, true, true, true, false, false] },
  { operator: 'notBetween', expected: [60.25, 60.75], exact: [true, true, false, false, true, true], days: [true, false, false, false, true, true] },
  { operator: 'equal', expected: 60.25, exact: [false, false, true, false, false, false], days: [false, true, true, true, false, false] },
  { operator: 'notEqual', expected: 60.25, exact: [true, true, false, true, true, true], days: [true, false, false, false, true, true] },
  { operator: 'newerThan', expected: 60.25, exact: [false, false, false, true, true, true], days: [false, false, false, false, true, true] },
  { operator: 'newerThanOrEqual', expected: 60.25, exact: [false, false, true, true, true, true], days: [false, true, true, true, true, true] },
  { operator: 'olderThan', expected: 60.25, exact: [true, true, false, false, false, false], days: [true, false, false, false, false, false] },
  { operator: 'olderThanOrEqual', expected: 60.25, exact: [true, true, true, false, false, false], days: [true, true, true, true, false, false] },
]

test('published Timeline levels and operator values map to retained grains without reversing Day/Year', () => {
  assert.deepEqual(TimelinePublicLevel, { Year: 0, Quarter: 1, Month: 2, Day: 3 })
  for (const [level, grain] of ['year', 'quarter', 'month', 'day'].entries()) {
    assert.equal(normalizeTimelineLevel(level), grain); assert.equal(normalizeTimelineLevel(grain), grain)
    assert.equal(toTimelinePublicLevel(level), level); assert.equal(toTimelinePublicLevel(grain), level)
  }
  assert.deepEqual(Object.values(TimelineDateFilterOperator).sort(), operators.map(item => item.operator).sort())
})

for (const invalid of [-1, 4, .5, NaN, Infinity, 'Month', '', null, undefined, {}]) test(`invalid level ${String(invalid)} cannot acquire a model grain`, () => {
  assert.throws(() => normalizeTimelineLevel(invalid), TypeError)
})

for (const useWholeDay of modes) test(`range mode ${String(useWholeDay)} preserves author fractional bounds and optional presence`, () => {
  const authored = withMode({ startDate: 59.25, endDate: 60.75 }, useWholeDay), normalized = normalizeTimelineDateRange(authored)
  assert.deepEqual(normalized, authored); assert.equal(Object.hasOwn(normalized, 'useWholeDay'), useWholeDay !== undefined)
  assert.deepEqual(effectiveTimelineDateRange(authored), useWholeDay === true ? { startDate: 59, endDate: 61 } : { startDate: 59.25, endDate: 60.75 })
  normalized.startDate = 0; assert.equal(authored.startDate, 59.25)
})

test('precise range includes its start and excludes its end and the representable value just above it', () => {
  const predicate = compileTimelineSerialRange({ startDate: 59.25, endDate: 59.5 }), step = 2 ** -30
  assert.deepEqual([59.25 - step, 59.25, 59.4, 59.5 - step, 59.5, 59.5 + step].map(predicate), [false, true, true, true, false, false])
  assert.equal(predicate(NaN), false); assert.equal(predicate(Infinity), false)
})

test('whole-day range expands negative fractions with floor/ceil while the upper day remains excluded', () => {
  const predicate = compileTimelineSerialRange({ startDate: -.25, endDate: .25, useWholeDay: true })
  assert.deepEqual([-1.25, -1, -.75, -.25, 0, .75, 1].map(predicate), [false, true, true, true, true, true, false])
})

test('typed Timeline range never treats a colliding raw number, calendar string, boolean, blank or error as Date', () => {
  const predicate = compileTimelineRange({ startDate: 60, endDate: 61 })
  assert.equal(predicate({ type: 'date', value: 60.5 }), true)
  for (const member of [{ type: 'number', value: 60.5 }, { type: 'string', value: '1900-02-29' }, { type: 'boolean', value: true },
    { type: 'blank' }, { type: 'error', value: '#N/A' }]) assert.equal(predicate(member), false)
  assert.equal(predicate({ type: 'date', value: 61 }), false)
})

for (const useWholeDay of modes) test(`standalone mode ${String(useWholeDay)} accepts empty bounds but persisted setters reject them`, () => {
  const empty = withMode({ startDate: 60.25, endDate: 60.25 }, useWholeDay)
  assert.throws(() => normalizeTimelineDateRange(empty), RangeError)
  assert.ok([60, 60.25, 60.5, 61].every(value => !compileTimelineSerialRange(empty)(value)))
})

for (const invalid of [null, [], {}, { startDate: 1 }, { startDate: '1', endDate: 2 }, { startDate: 1, endDate: '2' },
  { startDate: 3, endDate: 2 }, { startDate: NaN, endDate: 2 }, { startDate: 1, endDate: Infinity },
  { startDate: 1, endDate: 2, useWholeDay: 1 }, { startDate: 1, endDate: 2, level: 'day' }]) {
  test(`invalid range ${JSON.stringify(invalid)} is refused before a predicate can be used`, () => {
    assert.throws(() => normalizeTimelineDateRange(invalid)); assert.throws(() => compileTimelineSerialRange(invalid))
  })
}

for (const item of operators) for (const useWholeDay of modes) test(`${item.operator}/${String(useWholeDay)} has fixed inclusive condition semantics`, () => {
  const condition = withMode({ operator: item.operator, expected: item.expected }, useWholeDay)
  assert.deepEqual([59.75, 60, 60.25, 60.75, 61, 61.25].map(compileTimelineSerialCondition(condition)), useWholeDay === true ? item.days : item.exact)
  assert.equal(compileTimelineSerialCondition(condition)(NaN), false)
})

for (const item of operators) test(`${item.operator} typed condition rejects every non-Date identity even when its scalar would match`, () => {
  const predicate = compileTimelineCondition({ operator: item.operator, expected: item.expected })
  assert.equal(predicate({ type: 'date', value: 60.25 }), item.exact[2])
  for (const member of [{ type: 'number', value: 60.25 }, { type: 'string', value: '60.25' }, { type: 'boolean', value: true },
    { type: 'blank' }, { type: 'error', value: '#N/A' }]) assert.equal(predicate(member), false)
})

for (const item of operators) for (const useWholeDay of modes) test(`${item.operator}/${String(useWholeDay)} converts to and from Pivot without losing operand or mode presence`, () => {
  const condition = withMode({ operator: item.operator, expected: item.expected }, useWholeDay)
  const pivot = timelineDateConditionToPivot(condition), restored = timelineDateConditionFromPivot(pivot)
  assert.deepEqual(restored, condition); assert.equal(Object.hasOwn(restored, 'useWholeDay'), useWholeDay !== undefined)
  if (Array.isArray(pivot.expected)) { pivot.expected[0] = 0; assert.equal(condition.expected[0], 60.25) }
})

for (const invalid of [null, [], {}, { operator: 'after', expected: 1 }, { operator: 'equal', expected: [1] },
  { operator: 'between', expected: 1 }, { operator: 'between', expected: [2, 1] }, { operator: 'notBetween', expected: [1] },
  { operator: 'between', expected: [1, 2, 3] }, { operator: 'between', expected: [1, NaN] },
  { operator: 'equal', expected: Infinity }, { operator: 'equal', expected: '1' },
  { operator: 'equal', expected: 1, useWholeDay: 'yes' }, { operator: 'equal', expected: 1, startDate: 0 }]) {
  test(`invalid condition ${JSON.stringify(invalid)} cannot compile or become a cached fixed predicate`, () => {
    assert.throws(() => normalizeTimelineDateCondition(invalid)); assert.throws(() => compileTimelineCondition(invalid))
  })
}

test('compiled predicates retain their captured operands and preserve distinct Excel serial compatibility day 60', () => {
  const range = { startDate: 60, endDate: 61 }, between = { operator: 'between', expected: [60, 60.75] }
  const inRange = compileTimelineSerialRange(range), condition = compileTimelineSerialCondition(between)
  range.startDate = 0; range.endDate = 1; between.expected[0] = 0; between.expected[1] = 1
  assert.deepEqual([59, 60, 60.5, 61].map(inRange), [false, true, true, false])
  assert.deepEqual([59, 60, 60.75, 61].map(condition), [false, true, true, false])
  assert.deepEqual(normalizeTimelineDateCondition({ operator: 'between', expected: [60, 60] }), { operator: 'between', expected: [60, 60] })
})
