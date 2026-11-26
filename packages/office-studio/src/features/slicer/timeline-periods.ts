import type { CellScalar, SlicerTimelineRange, TimelineLevel } from '../../domain/model.ts'
import type { ITimelineRange, TimelineDateFilterState } from '../../domain/timeline-filters.ts'
import { normalizeTimelineDateRange, normalizeTimelineLevel } from './timeline-date-filter.ts'
import { parseTimelineDate, timelineCalendarToSerial } from './timeline-date.ts'

/** Calendar periods are real half-open Excel serial intervals, including compatibility day 60. */
export function timelinePeriodRange(value: CellScalar, inputLevel: TimelineLevel, date1904: boolean): ITimelineRange {
  const date = parseTimelineDate(value, date1904), level = normalizeTimelineLevel(inputLevel)
  if (!date) throw new TypeError('时间线周期需要有效工作簿日期。')
  if (level === 'day') return { startDate: Math.floor(date.serial), endDate: Math.floor(date.serial) + 1 }
  const month = level === 'year' ? 1 : level === 'quarter' ? Math.floor((date.month - 1) / 3) * 3 + 1 : date.month
  const width = level === 'year' ? 12 : level === 'quarter' ? 3 : 1
  const nextMonth = (month - 1 + width) % 12 + 1, nextYear = date.year + Math.floor((month - 1 + width) / 12)
  const calendar = (year: number, month: number) => `${String(year).padStart(4, '0')}-${String(month).padStart(2, '0')}-01`
  const startDate = timelineCalendarToSerial(calendar(date.year, month), date1904)
  let endDate = timelineCalendarToSerial(calendar(nextYear, nextMonth), date1904)
  if (startDate === undefined) throw new TypeError('时间线周期超出工作簿日期范围。')
  // The exclusive end of December/year 9999 is a finite boundary, although it is not a selectable calendar day.
  if (endDate === undefined && nextYear === 10000) {
    const start = new Date(0), end = new Date(0)
    start.setUTCFullYear(date.year, month - 1, 1); start.setUTCHours(0, 0, 0, 0)
    end.setUTCFullYear(nextYear, nextMonth - 1, 1); end.setUTCHours(0, 0, 0, 0)
    endDate = startDate + (end.getTime() - start.getTime()) / 86_400_000
  }
  if (endDate === undefined) throw new TypeError('时间线周期结束日期无效。')
  return normalizeTimelineDateRange({ startDate, endDate })
}

/** Legacy inclusive period authoring migrates without subtracting an epsilon or inferring from labels. */
export function timelineFilterStateFromLegacy(input: SlicerTimelineRange, date1904: boolean): TimelineDateFilterState {
  if (!input || typeof input !== 'object' || Array.isArray(input)
    || Object.keys(input).some(key => !['level', 'start', 'end'].includes(key))) throw new TypeError('旧时间线范围无效。')
  const level = normalizeTimelineLevel(input.level)
  const startDate = input.start == null ? undefined : parseTimelineDate(input.start, date1904)
  const endDate = input.end == null ? undefined : parseTimelineDate(input.end, date1904)
  if (input.start != null && !startDate || input.end != null && !endDate) throw new TypeError('旧时间线范围需要有效工作簿日期。')
  if (startDate && endDate && startDate.serial > endDate.serial) throw new RangeError('旧时间线开始日期不能晚于结束日期。')
  const first = input.start == null ? undefined : timelinePeriodRange(input.start, level, date1904)
  const last = input.end == null ? undefined : timelinePeriodRange(input.end, level, date1904)
  if (first && last) return { range: normalizeTimelineDateRange({ startDate: first.startDate, endDate: last.endDate }) }
  if (first) return { condition: { operator: 'newerThanOrEqual', expected: first.startDate } }
  if (last) return { condition: { operator: 'olderThan', expected: last.endDate } }
  return {}
}
