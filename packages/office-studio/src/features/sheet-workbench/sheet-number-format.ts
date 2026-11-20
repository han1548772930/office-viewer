import type { CellScalar } from '../../domain/model.ts'
import { parseTimelineDate } from '../slicer/timeline-date.ts'

export type SheetNumberFormat = 'general' | 'number' | 'currency' | 'accounting' | 'percent' | 'date' | 'time' | 'datetime' | 'text'
export const SHEET_NUMBER_FORMATS: Readonly<Record<SheetNumberFormat, string>> = Object.freeze({
  general: 'General', number: '#,##0.00', currency: '¥#,##0.00;[Red]-¥#,##0.00',
  accounting: '¥* #,##0.00;¥* (#,##0.00);¥* "-"', percent: '0%',
  date: 'yyyy-mm-dd', time: 'hh:mm:ss', datetime: 'yyyy-mm-dd hh:mm:ss', text: '@',
})

export function classifySheetNumberFormat(pattern: string): SheetNumberFormat | 'custom' {
  if (!pattern || /^general$/i.test(pattern)) return 'general'
  if (pattern === '@') return 'text'
  const unquoted = pattern.replace(/"[^"]*"|\\.|\[[^\]]*\]/g, '')
  if (/[yd]/i.test(unquoted)) return /[hs]/i.test(unquoted) ? 'datetime' : 'date'
  if (/[hs]/i.test(unquoted)) return 'time'
  if (unquoted.includes('%')) return 'percent'
  if (/[$¥￥€£]/.test(unquoted)) return unquoted.includes('*') ? 'accounting' : 'currency'
  if (/^[#,0.]+(?:;[#,0.()\-]+)*$/.test(pattern)) return 'number'
  return 'custom'
}

/** Change only decimal placeholders; date/text/custom literals stay intact. */
export function changeSheetDecimalPlaces(pattern: string, delta: -1 | 1): string {
  const kind = classifySheetNumberFormat(pattern)
  if (kind === 'date' || kind === 'time' || kind === 'datetime' || kind === 'text') return pattern
  const source = kind === 'general' ? '0' : pattern
  return source.split(';').map(section => {
    let quoted = false, escaped = false, match: { start: number; end: number; count: number } | undefined
    for (let i = 0; i < section.length; i += 1) {
      const char = section[i]!
      if (escaped) { escaped = false; continue }
      if (char === '\\') { escaped = true; continue }
      if (char === '"') { quoted = !quoted; continue }
      if (quoted || !/[0#]/.test(char)) continue
      const token = /^[#,0]+(?:\.([0#]+))?/.exec(section.slice(i))
      if (token) { match = { start: i, end: i + token[0].length, count: token[1]?.length ?? 0 }; break }
    }
    if (!match) return section
    const count = Math.max(0, Math.min(15, match.count + delta))
    const integer = section.slice(match.start, match.end).split('.')[0]!
    return section.slice(0, match.start) + integer + (count ? `.${'0'.repeat(count)}` : '') + section.slice(match.end)
  }).join(';')
}

/** Returns undefined for an unsupported imported pattern so its preserved display can be used. */
export function formatSheetValue(value: CellScalar, pattern: string, date1904 = false, locale = 'zh-CN'): string | undefined {
  if (value === null) return ''
  if (typeof value !== 'number' || !Number.isFinite(value)) return typeof value === 'boolean' ? String(value).toUpperCase() : String(value)
  const kind = classifySheetNumberFormat(pattern)
  if (kind === 'general' || kind === 'text') return String(value)
  if (kind === 'custom') return undefined
  if (kind === 'date' || kind === 'time' || kind === 'datetime') {
    if (![SHEET_NUMBER_FORMATS.date, SHEET_NUMBER_FORMATS.time, SHEET_NUMBER_FORMATS.datetime].includes(pattern)) return undefined
    const day = parseTimelineDate(value, date1904)
    if (!day) return '########'
    const date = `${String(day.year).padStart(4, '0')}-${String(day.month).padStart(2, '0')}-${String(day.day).padStart(2, '0')}`
    const seconds = Math.floor((((value % 1) + 1) % 1) * 86_400 + 1e-7)
    const time = `${String(Math.floor(seconds / 3600)).padStart(2, '0')}:${String(Math.floor(seconds % 3600 / 60)).padStart(2, '0')}:${String(seconds % 60).padStart(2, '0')}`
    return kind === 'date' ? date : kind === 'time' ? time : `${date} ${time}`
  }
  const sections = pattern.split(';')
  const section = value < 0 ? sections[1] ?? sections[0]! : value === 0 ? sections[2] ?? sections[0]! : sections[0]!
  const decimals = Math.min(15, /[0#]\.([0#]+)/.exec(section)?.[1]?.length ?? 0)
  const numeric = new Intl.NumberFormat(locale, { useGrouping: section.includes(','), minimumFractionDigits: decimals, maximumFractionDigits: decimals })
    .format(Math.abs(value) * (kind === 'percent' ? 100 : 1))
  const negative = value < 0
  if (kind === 'currency' || kind === 'accounting') {
    const symbol = /[$¥￥€£]/.exec(section)?.[0] ?? '¥'
    const amount = kind === 'accounting' && value === 0 ? '—' : numeric
    return kind === 'accounting' ? `${symbol} ${negative ? `(${amount})` : amount}` : `${negative ? '-' : ''}${symbol}${amount}`
  }
  return `${negative ? '-' : ''}${numeric}${kind === 'percent' ? '%' : ''}`
}
