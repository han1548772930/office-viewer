import type { DocumentContentControl, DocumentContentControlListItem } from '../../../domain/model'

type CheckboxState = NonNullable<NonNullable<DocumentContentControl['checkbox']>['checkedState']>

export const CONTENT_CONTROL_CHECKBOX_STYLES = {
  Cross: { checkedState: { value: '2612', font: 'Segoe UI Symbol' }, uncheckedState: { value: '2610', font: 'Segoe UI Symbol' } },
  Check: { checkedState: { value: '2611', font: 'Segoe UI Symbol' }, uncheckedState: { value: '2610', font: 'Segoe UI Symbol' } },
  Minimal: { checkedState: { value: '2713', font: 'Segoe UI Symbol' }, uncheckedState: { value: '25A1', font: 'Segoe UI Symbol' } },
  Round: { checkedState: { value: '25CF', font: 'Segoe UI Symbol' }, uncheckedState: { value: '25CB', font: 'Segoe UI Symbol' } },
} as const satisfies Record<string, { checkedState: CheckboxState; uncheckedState: CheckboxState }>

export type ContentControlCheckboxStyle = keyof typeof CONTENT_CONTROL_CHECKBOX_STYLES
export const CONTENT_CONTROL_DATE_FORMATS = ['yyyy-MM-dd', 'yyyy/MM/dd', 'M/d/yyyy', 'dd MMM yyyy', 'MMMM d, yyyy'] as const

export function contentControlCheckboxStyle(control: DocumentContentControl): ContentControlCheckboxStyle | undefined {
  const raw = control.metadata?.checkboxStyle
  const defaults = CONTENT_CONTROL_CHECKBOX_STYLES.Cross
  const sameState = (actual: CheckboxState | undefined, fallback: CheckboxState, expected: CheckboxState) => {
    const value = actual?.value ?? fallback.value!
    return /^(?:0x)?[\da-f]{1,6}$/i.test(value) && Number.parseInt(value.replace(/^0x/i, ''), 16) === Number.parseInt(expected.value!, 16)
      && (actual?.font ?? fallback.font) === expected.font
  }
  const matches = (name: ContentControlCheckboxStyle) => {
    const style = CONTENT_CONTROL_CHECKBOX_STYLES[name]
    return sameState(control.checkbox?.checkedState, defaults.checkedState, style.checkedState)
      && sameState(control.checkbox?.uncheckedState, defaults.uncheckedState, style.uncheckedState)
  }
  if (typeof raw === 'string' && isContentControlCheckboxStyle(raw) && matches(raw)) return raw
  return (Object.keys(CONTENT_CONTROL_CHECKBOX_STYLES) as ContentControlCheckboxStyle[]).find(matches)
}

export function isContentControlCheckboxStyle(value: string): value is ContentControlCheckboxStyle {
  return Object.hasOwn(CONTENT_CONTROL_CHECKBOX_STYLES, value)
}

export function contentControlCheckboxGlyph(control: DocumentContentControl, checked = control.checked ?? control.checkbox?.checked ?? false): { text: string; fontFamily: string } {
  const defaults = CONTENT_CONTROL_CHECKBOX_STYLES.Cross
  const state = checked ? control.checkbox?.checkedState : control.checkbox?.uncheckedState
  const fallback = checked ? defaults.checkedState : defaults.uncheckedState
  const value = state?.value ?? fallback.value
  validateContentControlCheckboxState({ ...state, value })
  return { text: String.fromCodePoint(Number.parseInt(value.replace(/^0x/i, ''), 16)), fontFamily: state?.font ?? fallback.font }
}

export function validateContentControlCheckboxState(state: CheckboxState): void {
  if (state.value !== undefined) {
    if (typeof state.value !== 'string' || !/^(?:0x)?[\da-f]{1,6}$/i.test(state.value)) throw new RangeError('复选框符号必须是 Unicode 十六进制码点。')
    const code = Number.parseInt(state.value.replace(/^0x/i, ''), 16)
    if (code > 0x10ffff || code >= 0xd800 && code <= 0xdfff || code < 0x20 || code >= 0x7f && code <= 0x9f) throw new RangeError('复选框符号必须是可显示的 Unicode 字符。')
  }
  if (state.font !== undefined && (typeof state.font !== 'string' || !state.font.trim() || /[\r\n\u0000]/.test(state.font))) throw new RangeError('复选框字体名称无效。')
}

export function validateContentControlListItems(items: readonly DocumentContentControlListItem[]): void {
  if (!Array.isArray(items) || !items.length) throw new RangeError('列表控件至少需要一个选项。')
  const values = new Set<string>()
  for (const item of items) {
    if (!item || typeof item !== 'object' || typeof item.value !== 'string' || !item.value.trim()) throw new RangeError('列表选项的值不能为空。')
    if (item.displayText !== undefined && typeof item.displayText !== 'string') throw new TypeError('列表显示文字必须是文本。')
    if (values.has(item.value)) throw new RangeError(`列表选项的值不能重复：${item.value}`)
    values.add(item.value)
  }
}

export function validContentControlIsoDate(value: string): boolean {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value) || value.startsWith('0000-')) return false
  const date = new Date(`${value}T00:00:00.000Z`)
  return Number.isFinite(date.valueOf()) && date.toISOString().slice(0, 10) === value
}

type DatePart = { token: string } | { literal: string }

function dateFormatParts(format: string): DatePart[] {
  if (typeof format !== 'string' || format.length > 512) throw new RangeError('日期格式过长或无效。')
  const parts: DatePart[] = []
  for (let index = 0; index < format.length;) {
    const character = format[index]!
    if (character === '\\') {
      if (++index === format.length) throw new RangeError('日期格式的转义字符后需要文字。')
      parts.push({ literal: format[index++]! })
    } else if (character === "'" || character === '"') {
      const quote = character
      let literal = '', closed = false
      index++
      while (index < format.length) {
        if (format[index] !== quote) { literal += format[index++]!; continue }
        if (format[index + 1] === quote) { literal += quote; index += 2; continue }
        index++; closed = true; break
      }
      if (!closed) throw new RangeError('日期格式的引号未闭合。')
      parts.push({ literal })
    } else {
      const token = /^(yyyy|yy|MMMM|MMM|MM|M|dddd|ddd|dd|d|gg)/.exec(format.slice(index))?.[0]
      if (token) { parts.push({ token }); index += token.length }
      else { parts.push({ literal: character }); index++ }
    }
  }
  return parts
}

const officeCalendars: Record<string, string> = {
  gregorian: 'gregory', gregory: 'gregory', hijri: 'islamic', islamic: 'islamic',
  hebrew: 'hebrew', japan: 'japanese', japanese: 'japanese', taiwan: 'roc', roc: 'roc',
  korean: 'dangi', dangi: 'dangi', thai: 'buddhist', buddhist: 'buddhist', saka: 'indian', indian: 'indian',
}

function calendarName(value?: string): string { return value ? officeCalendars[value] ?? value : 'gregory' }

export function validateContentControlDateProperties(format?: string, locale?: string, calendar?: string): void {
  if (format !== undefined) dateFormatParts(format)
  if (locale !== undefined && typeof locale !== 'string') throw new TypeError('日期区域设置必须是文本。')
  if (calendar !== undefined && typeof calendar !== 'string') throw new TypeError('日历名称必须是文本。')
  try {
    const requested = calendarName(calendar)
    const resolved = new Intl.DateTimeFormat(locale || 'en-US', { calendar: requested }).resolvedOptions()
    if (resolved.calendar !== requested) throw new RangeError('日期日历不可用。')
  } catch { throw new RangeError('日期区域设置或日历无效。') }
}

/** ISO storage stays Gregorian; the display calendar, names and literals belong to presentation. */
export function formatContentControlDate(value: string, format = 'yyyy-MM-dd', locale = 'en-US', calendar?: string): string {
  if (!value) return ''
  if (!validContentControlIsoDate(value)) throw new RangeError('Date content control value must be YYYY-MM-DD')
  validateContentControlDateProperties(format, locale, calendar)
  const date = new Date(`${value}T00:00:00.000Z`)
  const options = { calendar: calendarName(calendar), timeZone: 'UTC', numberingSystem: 'latn' }
  const numeric = new Intl.DateTimeFormat(locale || 'en-US', { ...options, year: 'numeric', month: 'numeric', day: 'numeric', era: 'short' }).formatToParts(date)
  const number = (type: Intl.DateTimeFormatPartTypes, fallback: number): string => numeric.find(part => part.type === type)?.value ?? String(fallback)
  const year = number('year', date.getUTCFullYear()), month = number('month', date.getUTCMonth() + 1), day = number('day', date.getUTCDate())
  const name = (type: 'month' | 'weekday', style: 'long' | 'short') => new Intl.DateTimeFormat(locale || 'en-US', { ...options, [type]: style }).format(date)
  const tokens: Record<string, () => string> = {
    yyyy: () => year.padStart(4, '0'), yy: () => year.padStart(2, '0').slice(-2),
    MMMM: () => name('month', 'long'), MMM: () => name('month', 'short'), MM: () => month.padStart(2, '0'), M: () => month,
    dddd: () => name('weekday', 'long'), ddd: () => name('weekday', 'short'), dd: () => day.padStart(2, '0'), d: () => day,
    gg: () => numeric.find(part => part.type === 'era')?.value ?? '',
  }
  return dateFormatParts(format || 'yyyy-MM-dd').map(part => 'literal' in part ? part.literal : tokens[part.token]!()).join('')
}

/** Parse only an explicit ISO date or an unambiguous Gregorian display mask. Never use Date.parse. */
export function parseContentControlDateText(value: string, format = 'yyyy-MM-dd', locale = 'en-US', calendar?: string): string | undefined {
  if (typeof value !== 'string') throw new TypeError('日期内容必须是文本。')
  if (value === '') return ''
  if (validContentControlIsoDate(value)) return value
  try { validateContentControlDateProperties(format, locale, calendar) } catch { return undefined }
  if (calendarName(calendar) !== 'gregory') return undefined
  const parts = dateFormatParts(format || 'yyyy-MM-dd')
  if (parts.some(part => 'token' in part && (part.token === 'yy' || part.token === 'gg'))) return undefined
  const escape = (text: string): string => text.replace(/[.*+?^{}$()|[\]\\]/g, '\\$&')
  const fold = (text: string): string => text.toLocaleLowerCase(locale || 'en-US')
  const captures: Array<{ token: string; index: number; names?: Map<string, number> }> = []
  let pattern = '^', group = 0
  for (const part of parts) {
    if ('literal' in part) { pattern += escape(fold(part.literal)); continue }
    const token = part.token
    let names: Map<string, number> | undefined
    let expression: string
    if (token === 'MMM' || token === 'MMMM') {
      names = new Map(Array.from({ length: 12 }, (_entry, month) => [
        fold(new Intl.DateTimeFormat(locale || 'en-US', { calendar: 'gregory', numberingSystem: 'latn', month: token === 'MMM' ? 'short' : 'long', timeZone: 'UTC' }).format(new Date(Date.UTC(2000, month, 1)))), month + 1,
      ] as const))
      if (names.size !== 12) return undefined
      expression = [...names.keys()].sort((left, right) => right.length - left.length).map(escape).join('|')
    } else if (token === 'ddd' || token === 'dddd') {
      const weekdays = Array.from({ length: 7 }, (_entry, day) => fold(new Intl.DateTimeFormat(locale || 'en-US', {
        calendar: 'gregory', numberingSystem: 'latn', weekday: token === 'ddd' ? 'short' : 'long', timeZone: 'UTC',
      }).format(new Date(Date.UTC(2020, 10, 1 + day)))))
      expression = weekdays.sort((left, right) => right.length - left.length).map(escape).join('|')
    } else expression = token === 'yyyy' ? '[0-9]{4}' : token === 'MM' || token === 'dd' ? '[0-9]{2}' : '[0-9]{1,2}'
    captures.push({ token, index: ++group, ...(names ? { names } : {}) })
    pattern += '(' + expression + ')'
  }
  const match = new RegExp(pattern + '$', 'u').exec(fold(value))
  if (!match) return undefined
  const values = { year: [] as number[], month: [] as number[], day: [] as number[] }
  for (const capture of captures) {
    const text = match[capture.index]!
    if (capture.token === 'yyyy') values.year.push(Number(text))
    else if (['M', 'MM', 'MMM', 'MMMM'].includes(capture.token)) values.month.push(capture.names ? capture.names.get(fold(text))! : Number(text))
    else if (capture.token === 'd' || capture.token === 'dd') values.day.push(Number(text))
  }
  if (Object.values(values).some(numbers => !numbers.length || numbers.some(number => number !== numbers[0]))) return undefined
  const iso = String(values.year[0]).padStart(4, '0') + '-' + String(values.month[0]).padStart(2, '0') + '-' + String(values.day[0]).padStart(2, '0')
  if (!validContentControlIsoDate(iso)) return undefined
  // Reformatting also verifies a supplied weekday, repeated tokens, spacing and literals.
  return fold(formatContentControlDate(iso, format, locale, calendar)) === fold(value) ? iso : undefined
}

export function contentControlDisplayColor(control: DocumentContentControl, fallback: string): string {
  return control.color && /^#?[\da-f]{6}$/i.test(control.color) ? `#${control.color.replace(/^#/, '')}` : fallback
}
