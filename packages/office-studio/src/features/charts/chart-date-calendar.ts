import type { ChartAxisOptions, ChartAxisTimeUnit } from '../../domain/model'
import { DateSystem } from '../sheet-workbench/number-format-engine/date-system'
import { excelDateTimePartsToSerial, excelSerialToDateTimeParts, type IExcelDateTimeParts } from '../sheet-workbench/number-format-engine/serial-date'

/** Tick coordinates and their calendar labels are separate from source rows. */
export interface ChartDateAxisTick { value: number; serial: number }
export interface ChartDateTickExtent {
  min: number
  max: number
  sourceMin: number
  sourceMax: number
  date1904: boolean
}

export const CHART_DATE_AXIS_MAX_TICKS = 10_000
type CalendarOptions = Pick<ChartAxisOptions, 'baseTimeUnit' | 'majorTimeUnit' | 'majorUnit'>
const dateSystem = (date1904: boolean) => ({ dateSystem: date1904 ? DateSystem.Date1904 : DateSystem.Date1900 })
const validTimeUnit = (unit: ChartAxisTimeUnit | undefined): boolean => unit === undefined || unit === 'days' || unit === 'months' || unit === 'years'

/** Activity is decided by the caller. Numeric axes keep inactive date metadata
 * without treating their numeric interval as a month/year count. */
export function validateChartDateCalendarOptions(axis: CalendarOptions): void {
  if (!validTimeUnit(axis.baseTimeUnit) || !validTimeUnit(axis.majorTimeUnit)) throw new TypeError('日期轴时间单位只接受日、月或年。')
  const unit = axis.majorTimeUnit ?? axis.baseTimeUnit ?? 'days'
  if (axis.majorUnit !== undefined && (!Number.isFinite(axis.majorUnit) || axis.majorUnit <= 0 || unit !== 'days' && !Number.isSafeInteger(axis.majorUnit))) {
    throw new RangeError('日期轴主单位须为正数，月或年步进须为安全整数。')
  }
}

/** Calendar buckets use actual year/month fields, never a fixed 30/365-day divisor. */
export function chartDateCoordinate(serial: number, unit?: ChartAxisTimeUnit, date1904 = false): number | undefined {
  if (!validTimeUnit(unit)) return undefined
  const parts = excelSerialToDateTimeParts(serial, dateSystem(date1904))
  if (!parts) return undefined
  return unit === 'months' ? parts.year * 12 + parts.month - 1 : unit === 'years' ? parts.year : serial
}

/** The center of a calendar bucket labels its first day. Source dates stay separate. */
export function chartDateCoordinateSerial(value: number, unit?: ChartAxisTimeUnit, date1904 = false): number | undefined {
  if (!validTimeUnit(unit) || !Number.isFinite(value)) return undefined
  if (unit !== 'months' && unit !== 'years') return excelSerialToDateTimeParts(value, dateSystem(date1904)) ? value : undefined
  const position = Math.round(value), year = unit === 'months' ? Math.floor(position / 12) : position,
    month = unit === 'months' ? position - year * 12 + 1 : 1
  return excelDateTimePartsToSerial({ year, month, day: 1, hours: 0, minutes: 0, seconds: 0, fractionalSecond: 0 }, dateSystem(date1904)) ?? undefined
}

function monthLength(year: number, month: number, date1904: boolean): number {
  const leap = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0) || !date1904 && year === 1900
  return [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 1]!
}

/** Each tick advances from the original anchor, so Jan 31 → Feb 29 → Mar 31. */
export function chartDateAdvance(serial: number, amount: number, unit: ChartAxisTimeUnit, date1904 = false): number | undefined {
  const options = dateSystem(date1904), parts = excelSerialToDateTimeParts(serial, options)
  if (!parts || !validTimeUnit(unit) || unit === undefined || !Number.isFinite(amount) || amount < 0) return undefined
  if (amount === 0) return serial
  if (unit === 'days') {
    const next = serial + amount
    return excelSerialToDateTimeParts(next, options) ? next : undefined
  }
  if (!Number.isSafeInteger(amount)) return undefined
  const position = parts.year * 12 + parts.month - 1 + amount * (unit === 'years' ? 12 : 1)
  if (!Number.isSafeInteger(position)) return undefined
  const year = Math.floor(position / 12), month = position - year * 12 + 1
  return excelDateTimePartsToSerial({ ...parts, year, month,
    day: Math.min(Math.max(1, parts.day), monthLength(year, month, date1904)) }, options) ?? undefined
}

function roundedInterval(value: number, choices: readonly number[]): number {
  if (value <= 1) return 1
  const exact = choices.find(choice => choice >= value)
  if (exact !== undefined) return exact
  const magnitude = 10 ** Math.floor(Math.log10(value))
  return (choices.find(choice => choice >= value / magnitude) ?? choices.at(-1)!) * magnitude
}

interface TickPlan { unit: ChartAxisTimeUnit; step: number; align: boolean }
function automaticPlan(extent: ChartDateTickExtent, axis: CalendarOptions, count: number): TickPlan {
  const locked = axis.majorTimeUnit ?? axis.baseTimeUnit
  const span = Math.max(extent.sourceMax - extent.sourceMin, locked ? 0 : extent.max - extent.min), target = span / Math.max(1, count - 1)
  if (locked === 'months' || locked === 'years') {
    const min = chartDateCoordinate(extent.sourceMin, locked, extent.date1904)!, max = chartDateCoordinate(extent.sourceMax, locked, extent.date1904)!
    return { unit: locked, step: roundedInterval((max - min) / Math.max(1, count - 1), locked === 'months' ? [1, 2, 3, 6, 12] : [1, 2, 5, 10]), align: true }
  }
  if (locked === 'days') return { unit: 'days', step: target < 1 ? dayIntervals.find(value => value >= target) ?? 1 : roundedInterval(target, [1, 2, 5, 10]), align: true }
  const days = dayIntervals.find(value => value >= target)
  if (days !== undefined) return { unit: 'days', step: days, align: true }
  const month = [1, 2, 3, 6].find(value => value * 30.4375 >= target)
  if (month !== undefined) return { unit: 'months', step: month, align: true }
  return { unit: 'years', step: roundedInterval(target / 365.2425, [1, 2, 5, 10]), align: true }
}
const dayIntervals = [1, 5, 10, 15, 30].map(seconds => seconds / 86_400)
  .concat([1, 5, 15, 30].map(minutes => minutes / 1440), [1, 3, 6, 12].map(hours => hours / 24), [1, 2, 7, 14])

function calendarAnchor(parts: IExcelDateTimeParts, unit: ChartAxisTimeUnit, step: number, date1904: boolean): number | undefined {
  const position = Math.floor((parts.year * 12 + parts.month - 1) / step) * step,
    year = unit === 'years' ? Math.floor(parts.year / step) * step : Math.floor(position / 12),
    month = unit === 'years' ? 1 : position - year * 12 + 1
  return excelDateTimePartsToSerial({ year, month, day: 1, hours: 0, minutes: 0, seconds: 0, fractionalSecond: 0 }, dateSystem(date1904)) ?? undefined
}

/** Bounded calendar tick generation. An authored interval that cannot advance or
 * exceeds the budget fails before publication; it is never silently replaced. */
export function chartDateTicks(extent: ChartDateTickExtent, axis: CalendarOptions, preferredCount = 5): ChartDateAxisTick[] {
  if (!Number.isFinite(extent.min) || !Number.isFinite(extent.max) || extent.min >= extent.max) throw new RangeError('日期轴需要有限且递增的坐标范围。')
  validateChartDateCalendarOptions(axis)
  if (!Number.isFinite(extent.sourceMin) || !Number.isFinite(extent.sourceMax) || extent.sourceMin > extent.sourceMax
    || !excelSerialToDateTimeParts(extent.sourceMin, dateSystem(extent.date1904)) || !excelSerialToDateTimeParts(extent.sourceMax, dateSystem(extent.date1904))) {
    throw new RangeError('日期轴刻度需要当前日期系统中有效且递增的源日期范围。')
  }
  const calendar = axis.baseTimeUnit === 'months' || axis.baseTimeUnit === 'years'
  const firstCalendar = calendar ? chartDateCoordinateSerial(Math.ceil(extent.min), axis.baseTimeUnit, extent.date1904) : undefined
  const lastCalendar = calendar ? chartDateCoordinateSerial(Math.floor(extent.max), axis.baseTimeUnit, extent.date1904) : undefined
  const first = firstCalendar ?? extent.sourceMin, last = lastCalendar ?? extent.sourceMax
  const unit = axis.majorTimeUnit ?? axis.baseTimeUnit ?? 'days', authored = axis.majorUnit !== undefined
  const plan: TickPlan = authored ? { unit, step: axis.majorUnit!, align: false }
    : automaticPlan({ ...extent, sourceMin: first, sourceMax: last }, axis, Number.isFinite(preferredCount) ? Math.max(2, Math.min(32, Math.floor(preferredCount))) : 5)
  const parts = excelSerialToDateTimeParts(first, dateSystem(extent.date1904))
  if (!parts) return []
  const automaticStart = plan.unit === 'days' ? Math.floor((calendar ? first : extent.min) / plan.step) * plan.step
    : calendarAnchor(parts, plan.unit, plan.step, extent.date1904)
  const anchor = plan.align ? Math.max(0, automaticStart ?? first) : first,
    upper = authored || calendar ? last : extent.max
  const ticks: ChartDateAxisTick[] = []
  const finish = (): ChartDateAxisTick[] => {
    // Subsecond ranges can fall between automatic whole-second ticks. Keep an
    // observed anchor instead of manufacturing an empty axis or rounding data.
    if (!authored && !ticks.length) {
      const value = chartDateCoordinate(first, axis.baseTimeUnit, extent.date1904)
      if (value !== undefined && value >= extent.min && value <= extent.max) ticks.push({ value,
        serial: calendar ? chartDateCoordinateSerial(value, axis.baseTimeUnit, extent.date1904)! : first })
    }
    return ticks
  }
  let previous = -Infinity
  for (let index = 0; index <= CHART_DATE_AXIS_MAX_TICKS; index++) {
    const serial = chartDateAdvance(anchor, index * plan.step, plan.unit, extent.date1904)
    if (serial === undefined || serial > upper) return finish()
    if (serial <= previous) throw new RangeError('日期轴主单位无法在当前数值精度内推进。')
    if (index === CHART_DATE_AXIS_MAX_TICKS) throw new RangeError(`日期轴主刻度超过${CHART_DATE_AXIS_MAX_TICKS}个，请增大主单位。`)
    previous = serial
    const value = chartDateCoordinate(serial, axis.baseTimeUnit, extent.date1904)
    if (value !== undefined && value >= extent.min && value <= extent.max && ticks.at(-1)?.value !== value) {
      ticks.push({ value, serial: calendar ? chartDateCoordinateSerial(value, axis.baseTimeUnit, extent.date1904)! : serial })
    }
  }
  return finish()
}
