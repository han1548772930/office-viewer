import type { RecordValue } from '../../domain/record-value.ts'

function calendarOrder(locale: string): ReadonlyMap<string, number> {
  const order = new Map<string, number>()
  for (const format of ['long', 'short'] as const) {
    const month = new Intl.DateTimeFormat(locale, { month: format, timeZone: 'UTC' })
    const weekday = new Intl.DateTimeFormat(locale, { weekday: format, timeZone: 'UTC' })
    for (let index = 0; index < 12; index++) order.set(month.format(new Date(Date.UTC(2020, index, 1))), index)
    for (let index = 0; index < 7; index++) order.set(weekday.format(new Date(Date.UTC(2020, 10, 1 + index))), 12 + index)
  }
  return order
}

/** Locale adapter shared by synchronous display and Worker rank preparation.
 * Calendar aliases and collator-equivalent text have equal comparison ranks,
 * but exact raw identities still remain distinct in the Rust member index.
 */
export function createSlicerTextComparator(customListSort: boolean, locale = 'zh-CN'): (left: string, right: string) => number {
  const collator = new Intl.Collator(locale, { numeric: true, sensitivity: 'base' })
  const calendar = customListSort ? calendarOrder(locale) : undefined
  return (left, right) => {
    const firstOrder = calendar?.get(left), secondOrder = calendar?.get(right)
    return firstOrder !== undefined || secondOrder !== undefined
      ? (firstOrder ?? Number.MAX_SAFE_INTEGER) - (secondOrder ?? Number.MAX_SAFE_INTEGER)
      : collator.compare(left, right)
  }
}

export function createSlicerMemberValueComparator(sortOrder: 'ascending' | 'descending', customListSort: boolean, locale = 'zh-CN'):
  (first: { value: RecordValue }, second: { value: RecordValue }) => number {
  const compareText = createSlicerTextComparator(customListSort, locale), direction = sortOrder === 'descending' ? -1 : 1
  const rank: Record<RecordValue['type'], number> = { blank: 0, number: 1, date: 2, string: 3, boolean: 4, error: 5 }
  return (first, second) => {
    const left = first.value, right = second.value
    let order = rank[left.type] - rank[right.type]
    if (left.type === right.type && left.type !== 'blank' && right.type !== 'blank') {
      order = typeof left.value === 'number' && typeof right.value === 'number'
        ? left.value - right.value : compareText(String(left.value), String(right.value))
    }
    return direction * order
  }
}
