import type { CellScalar, TimelineLevel } from '../../domain/model'

const DAY_MS = 86_400_000
const EPOCH_1900 = Date.UTC(1899, 11, 31)
const EPOCH_1904 = Date.UTC(1904, 0, 1)

/** A workbook calendar day, including Excel's distinct fictitious serial day 60. */
export interface TimelineDate {
  year: number
  month: number
  day: number
  serial: number
}

function calendarTimestamp(year: number, month: number, day: number): number {
  // Date.UTC treats years 0–99 as 1900–1999; setUTCFullYear does not.
  const date = new Date(0)
  date.setUTCFullYear(year, month - 1, day)
  date.setUTCHours(0, 0, 0, 0)
  return date.getTime()
}

function calendarToSerial(year: number, month: number, day: number, date1904: boolean): number | undefined {
  if (year < 1 || year > 9999 || month < 1 || month > 12 || day < 1 || day > 31) return undefined
  // Preserve Excel's compatibility day instead of silently folding it into February 28.
  if (!date1904 && year === 1900 && month === 2 && day === 29) return 60
  const timestamp = calendarTimestamp(year, month, day)
  const date = new Date(timestamp)
  if (date.getUTCFullYear() !== year || date.getUTCMonth() + 1 !== month || date.getUTCDate() !== day) return undefined
  const days = (timestamp - (date1904 ? EPOCH_1904 : EPOCH_1900)) / DAY_MS
  return !date1904 && timestamp >= Date.UTC(1900, 2, 1) ? days + 1 : days
}

/** Strict calendar input shared by the panel and timeline command adapter. */
export function timelineCalendarToSerial(value: string, date1904: boolean): number | undefined {
  const parts = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value)
  return parts ? calendarToSerial(Number(parts[1]), Number(parts[2]), Number(parts[3]), date1904) : undefined
}

export function parseTimelineDate(value: CellScalar | undefined, date1904: boolean): TimelineDate | undefined {
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) return undefined
    const dayNumber = Math.floor(value)
    if (!date1904 && dayNumber === 60) return { year: 1900, month: 2, day: 29, serial: value }
    const timestamp = (date1904 ? EPOCH_1904 : EPOCH_1900) + (value - (!date1904 && value >= 61 ? 1 : 0)) * DAY_MS
    const date = new Date(timestamp)
    const year = date.getUTCFullYear()
    if (!Number.isFinite(date.getTime()) || year < 1 || year > 9999) return undefined
    return { year, month: date.getUTCMonth() + 1, day: date.getUTCDate(), serial: value }
  }
  if (typeof value !== 'string' || !value.trim()) return undefined
  const text = value.trim()
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    const serial = timelineCalendarToSerial(text, date1904)
    return serial === undefined ? undefined : parseTimelineDate(serial, date1904)
  }
  // Date.parse normalizes invalid ISO days (e.g. February 30) instead of rejecting them.
  const prefix = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[Tt ]|$)/.exec(text)
  if (prefix && calendarToSerial(Number(prefix[1]), Number(prefix[2]), Number(prefix[3]), true) === undefined) return undefined
  // An Office date-time without an offset is a calendar value, not host-local time.
  const localIso = /^\d{4}-\d{2}-\d{2}[Tt ]\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?$/.test(text)
  const timestamp = Date.parse(localIso ? `${text.replace(' ', 'T')}Z` : text)
  if (!Number.isFinite(timestamp)) return undefined
  const date = new Date(timestamp)
  const year = date.getUTCFullYear()
  const month = date.getUTCMonth() + 1
  const day = date.getUTCDate()
  const serial = calendarToSerial(year, month, day, date1904)
  if (serial === undefined) return undefined
  const timeOfDay = (timestamp - calendarTimestamp(year, month, day)) / DAY_MS
  return { year, month, day, serial: serial + timeOfDay }
}

export function timelineDateToInput(value: CellScalar | undefined, date1904: boolean): string {
  const date = parseTimelineDate(value, date1904)
  return date ? `${String(date.year).padStart(4, '0')}-${String(date.month).padStart(2, '0')}-${String(date.day).padStart(2, '0')}` : ''
}

/** Period keys include whole days and do not depend on the host timezone. */
export function timelinePeriodKey(date: TimelineDate, level: TimelineLevel): number {
  if (level === 'year') return date.year
  if (level === 'quarter') return date.year * 4 + Math.floor((date.month - 1) / 3)
  if (level === 'month') return date.year * 12 + date.month - 1
  return (date.year * 12 + date.month - 1) * 31 + date.day
}
