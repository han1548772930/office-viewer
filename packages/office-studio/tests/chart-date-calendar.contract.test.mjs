import assert from 'node:assert/strict'
import { test } from 'vitest'
import { chartDateAxis, chartDateAxisLabel } from '../src/features/charts/chart-date-axis.ts'
import { chartDateAdvance, chartDateCoordinate, chartDateCoordinateSerial, chartDateTicks, validateChartDateCalendarOptions, CHART_DATE_AXIS_MAX_TICKS } from '../src/features/charts/chart-date-calendar.ts'

// P22 implementation drafts only. No runner, build, browser, benchmark or Office sample has run.
const JAN_2024 = 45292, JAN_31_2024 = 45322, FEB_2024 = 45323, FEB_29_2024 = 45351,
  MAR_2024 = 45352, MAR_31_2024 = 45382, APR_2024 = 45383, APR_30_2024 = 45412
const month = (year, number) => year * 12 + number - 1
const extent = (sourceMin, sourceMax, date1904 = false) => ({ min: sourceMin - 0.5, max: sourceMax + 0.5, sourceMin, sourceMax, date1904 })
function emptyChart(xAxis) {
  return { type: 'line', xAxis: { scaleType: 'linear', ...xAxis }, options: { mapping: { useDateAxis: true } } }
}

test('monthly ticks preserve the original month-end anchor and fractional day across leap February', () => {
  const anchor = JAN_31_2024 + 0.375
  assert.deepEqual([0, 1, 2, 3].map(amount => chartDateAdvance(anchor, amount, 'months')),
    [JAN_31_2024, FEB_29_2024, MAR_31_2024, APR_30_2024].map(serial => serial + 0.375))
  const axis = { majorTimeUnit: 'months', majorUnit: 1 }, before = structuredClone(axis)
  assert.deepEqual(chartDateTicks(extent(JAN_31_2024, APR_30_2024), axis).map(tick => tick.serial),
    [JAN_31_2024, FEB_29_2024, MAR_31_2024, APR_30_2024])
  assert.deepEqual(axis, before)
})

test('annual leap-day stepping always returns to the original day in a later leap year', () => {
  assert.deepEqual([0, 1, 4].map(amount => chartDateAdvance(FEB_29_2024, amount, 'years')), [45351, 45716, 46812])
  assert.deepEqual([0, 1, 2].map(amount => chartDateAdvance(45688, amount, 'months')), [45688, 45716, 45747])
})

test('1900 fictional February, serial zero and the 1904 epoch remain separate calendar identities', () => {
  assert.deepEqual([0, 1, 2].map(amount => chartDateAdvance(31, amount, 'months')), [31, 60, 91])
  assert.deepEqual([0, 1, 2].map(amount => chartDateAdvance(30, amount, 'months', true)), [30, 59, 90])
  assert.equal(chartDateAdvance(0, 0, 'months'), 0)
  assert.equal(chartDateAdvance(0, 0, 'years', true), 0)
  assert.equal(chartDateCoordinate(0, 'months'), month(1900, 1))
  assert.equal(chartDateCoordinateSerial(month(1900, 1), 'months'), 1)
  assert.equal(chartDateCoordinateSerial(month(1904, 1), 'months', true), 0)
  assert.equal(chartDateAxisLabel(60, {}), '1900-02-29')
  assert.equal(chartDateAxisLabel(59, {}, true), '1904-02-29')
})

test('month and year coordinates use calendar buckets while elapsed days and labels retain original serials', () => {
  assert.deepEqual([45291, JAN_31_2024, FEB_29_2024, MAR_31_2024].map(serial => chartDateCoordinate(serial, 'months')),
    [month(2023, 12), month(2024, 1), month(2024, 2), month(2024, 3)])
  assert.deepEqual([45291, JAN_31_2024, 45658].map(serial => chartDateCoordinate(serial, 'years')), [2023, 2024, 2025])
  assert.equal(chartDateCoordinate(FEB_29_2024 + 0.5, 'days'), FEB_29_2024 + 0.5)
  assert.deepEqual([month(2024, 1), month(2024, 2), month(2024, 3)].map(value => chartDateCoordinateSerial(value, 'months')),
    [JAN_2024, FEB_2024, MAR_2024])
  assert.equal(chartDateCoordinateSerial(2024, 'years'), JAN_2024)
})

test('basic monthly units change point coordinates and label bucket first days independently of major units', () => {
  const source = { categories: ['January', 'February', 'March', 'April'], categoryValues: [JAN_31_2024, FEB_29_2024, MAR_31_2024, APR_30_2024],
    series: [{ field: 1, name: 'Units', values: [1, 2, 3, 4] }] }
  const axis = chartDateAxis(source, emptyChart({ baseTimeUnit: 'months', majorTimeUnit: 'months', majorUnit: 2 }))
  assert.deepEqual(axis.values, source.categoryValues)
  assert.deepEqual(axis.coordinates, [0, 1, 2, 3].map(offset => month(2024, 1) + offset))
  assert.deepEqual([axis.coordinateMin, axis.coordinateMax], [month(2024, 1) - 0.5, month(2024, 4) + 0.5])
  assert.deepEqual(axis.ticks, [{ value: month(2024, 1), serial: JAN_2024 }, { value: month(2024, 3), serial: MAR_2024 }])
  const annual = chartDateAxis(source, emptyChart({ baseTimeUnit: 'years', majorUnit: 1 }))
  assert.deepEqual(annual.coordinates, [2024, 2024, 2024, 2024])
  assert.deepEqual(annual.ticks, [{ value: 2024, serial: JAN_2024 }])
})

test('authored elapsed-date boundaries anchor calendar advances and monthly boundaries preserve full buckets', () => {
  const source = { categories: ['Jan', 'Apr'], categoryValues: [JAN_31_2024, APR_30_2024], series: [{ field: 1, name: 'Units', values: [1, 2] }] }
  const model = emptyChart({ min: FEB_2024, max: APR_30_2024, majorTimeUnit: 'months', majorUnit: 1 }), before = structuredClone(model)
  const axis = chartDateAxis(source, model)
  assert.deepEqual(axis.ticks.map(tick => tick.serial), [FEB_2024, MAR_2024, APR_2024])
  assert.deepEqual([axis.sourceMin, axis.sourceMax], [FEB_2024, APR_30_2024])
  assert.deepEqual([axis.min, axis.max], [FEB_2024, APR_30_2024])
  const buckets = chartDateAxis(source, emptyChart({ min: JAN_31_2024, max: MAR_31_2024, baseTimeUnit: 'months', majorUnit: 1 }))
  assert.deepEqual([buckets.coordinateMin, buckets.coordinateMax], [month(2024, 1) - 0.5, month(2024, 3) + 0.5])
  assert.deepEqual(buckets.ticks.map(tick => tick.serial), [JAN_2024, FEB_2024, MAR_2024])
  assert.deepEqual(model, before)
})

test('fractional-day author intervals retain exact quarter-day ticks through the fictional leap day', () => {
  const ticks = chartDateTicks({ min: 59, max: 60, sourceMin: 59, sourceMax: 60, date1904: false }, { majorTimeUnit: 'days', majorUnit: 0.25 })
  assert.deepEqual(ticks, [59, 59.25, 59.5, 59.75, 60].map(serial => ({ value: serial, serial })))
  assert.doesNotThrow(() => validateChartDateCalendarOptions({ baseTimeUnit: 'months', majorTimeUnit: 'days', majorUnit: 0.25 }))
})

test('automatic monthly alignment crosses year boundaries using one continuous month ordinal', () => {
  const ticks = chartDateTicks({ min: month(2023, 12) - 0.5, max: month(2024, 4) + 0.5, sourceMin: 45261, sourceMax: APR_2024, date1904: false },
    { baseTimeUnit: 'months' }, 3)
  assert.deepEqual(ticks, [{ value: month(2024, 1), serial: JAN_2024 }, { value: month(2024, 3), serial: MAR_2024 }])
})

test('automatic subsecond and full supported-calendar spans produce finite nonempty ticks without rewriting dates', () => {
  const small = { min: JAN_2024 + 0.000001, max: JAN_2024 + 0.000002, sourceMin: JAN_2024 + 0.000001, sourceMax: JAN_2024 + 0.000002, date1904: false }
  const narrow = chartDateTicks(small, {}, NaN)
  assert.equal(narrow.length, 1)
  assert.equal(narrow[0].serial, small.sourceMin)
  assert.ok(narrow[0].value >= small.min && narrow[0].value <= small.max)
  const wide = chartDateTicks(extent(1, 2958465), {}, 5)
  assert.ok(wide.length > 0 && wide.length <= 32)
  assert.equal(wide.every(tick => Number.isFinite(tick.value) && chartDateAxisLabel(tick.serial, {}).length > 0), true)
  assert.equal(wide.every((tick, index) => !index || tick.value > wide[index - 1].value), true)
})

test('active date options reject invalid units or fractional month/year intervals even with an empty source', () => {
  for (const xAxis of [{ baseTimeUnit: 'weeks' }, { majorTimeUnit: 'quarters' }, { baseTimeUnit: 'months', majorUnit: 0.5 },
    { majorTimeUnit: 'years', majorUnit: Number.MAX_SAFE_INTEGER + 1 }, { majorUnit: 0 }, { majorUnit: -1 }, { majorUnit: NaN }, { majorUnit: Infinity }]) {
    const model = emptyChart(xAxis), before = structuredClone(model)
    assert.throws(() => chartDateAxis({ categories: [], series: [] }, model), /时间单位|主单位/)
    assert.deepEqual(model, before)
  }
  assert.throws(() => chartDateTicks({ min: 1, max: 1, sourceMin: 1, sourceMax: 1, date1904: false }, {}), /坐标范围/)
  assert.throws(() => chartDateTicks({ min: 1, max: 2, sourceMin: 2, sourceMax: 1, date1904: false }, {}), /源日期范围/)
})

test('author intervals that exceed the step budget or cannot advance fail explicitly and remain authored', () => {
  const axis = { majorTimeUnit: 'days', majorUnit: 0.001 }, before = structuredClone(axis)
  assert.equal(CHART_DATE_AXIS_MAX_TICKS, 10000)
  assert.throws(() => chartDateTicks(extent(JAN_2024, JAN_2024 + 11), axis), /超过10000/)
  assert.deepEqual(axis, before)
  assert.throws(() => chartDateTicks(extent(JAN_2024, JAN_2024 + 1), { majorUnit: Number.MIN_VALUE }), /无法.*推进/)
})

test('a single authored boundary that clips all observed dates still produces a valid visible date anchor', () => {
  const source = { categories: ['Observed'], categoryValues: [60], series: [{ field: 1, name: 'Units', values: [1] }] }
  const maximum = chartDateAxis(source, emptyChart({ max: 59.75 }))
  assert.deepEqual([maximum.sourceMin, maximum.sourceMax], [59.75, 59.75])
  assert.ok(maximum.ticks.length > 0)
  assert.equal(maximum.ticks.every(tick => tick.value >= maximum.coordinateMin && tick.value <= maximum.coordinateMax), true)
  const minimum = chartDateAxis(source, emptyChart({ min: 60.25 }))
  assert.deepEqual([minimum.sourceMin, minimum.sourceMax], [60.25, 60.25])
  assert.ok(minimum.ticks.length > 0)
})
