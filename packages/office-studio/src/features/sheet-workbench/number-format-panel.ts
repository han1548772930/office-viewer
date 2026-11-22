import type { OfficeLocale } from '../../i18n/regional-settings.ts'
import { currencySymbols } from './number-format-engine/constants.ts'
import { describeSheetNumberFormat } from './number-format-render.ts'
import { numberFormatDecimalPlaces, numberFormatWithoutDecimals, setNumberFormatDecimalPlaces } from './number-format-pattern.ts'
import { numberFormatCurrency, sheetCurrencySymbol } from './sheet-currency.ts'

/** SheetNumfmtPanel TYPE_OPTIONS, in fixed v1.6.0 order. */
export const SHEET_NUMBER_FORMAT_CATEGORIES = ['general', 'accounting', 'currency', 'date', 'grouped', 'custom'] as const
export type SheetNumberFormatCategory = typeof SHEET_NUMBER_FORMAT_CATEGORIES[number]
export function isSheetNumberFormatCategory(value: unknown): value is SheetNumberFormatCategory {
  return SHEET_NUMBER_FORMAT_CATEGORIES.some(category => category === value)
}
export interface SheetNumberFormatDraft {
  category: SheetNumberFormatCategory
  decimals: number
  currency: string
  negative: number
  pattern: string
}
export const SHEET_NUMBER_FORMAT_PANEL_CURRENCIES: readonly string[] = currencySymbols
export const SHEET_GROUPED_NUMBER_PATTERNS: readonly string[] = [
  '#,##0_);(#,##0)', '#,##0_);[Red](#,##0)', '#,##0.00_);#,##0.00',
  '#,##0.00_);[Red]#,##0.00', '#,##0.00_);-#,##0.00', '#,##0.00_);[Red]-#,##0.00',
]

export function sheetCurrencyNumberPatterns(currency: string): string[] {
  const symbol = `"${currency.replaceAll('"', '""')}"`
  return [
    `${symbol}#,##0.00_);${symbol}#,##0.00`, `${symbol}#,##0.00_);[Red]${symbol}#,##0.00`,
    `${symbol}#,##0.00_);(${symbol}#,##0.00)`, `${symbol}#,##0.00_);[Red](${symbol}#,##0.00)`,
    `${symbol}#,##0.00_);-${symbol}#,##0.00`, `${symbol}#,##0.00_);[Red]-${symbol}#,##0.00`,
  ]
}

interface DateFormatProfile { dates: readonly string[]; shortDate: string; periodBefore?: boolean }
const iso: DateFormatProfile = { dates: ['yyyy/MM/dd', 'MM-dd'], shortDate: 'MM-dd' }
const mdy: DateFormatProfile = { dates: ['m/d/yyyy', 'mm/dd/yyyy', 'mmm d, yyyy', 'm/d', 'mmmm d'], shortDate: 'm/d' }
const dmy: DateFormatProfile = { dates: ['dd/mm/yyyy', 'd/m/yyyy', 'd mmm yyyy', 'dd/mm', 'd mmmm'], shortDate: 'dd/mm' }
const dot: DateFormatProfile = { dates: ['dd.mm.yyyy', 'd.m.yyyy', 'd mmm yyyy', 'dd.mm', 'd mmmm'], shortDate: 'dd.mm' }
const cjk: DateFormatProfile = { dates: ['yyyy/MM/dd', 'yyyy"年"MM"月"dd"日"', 'MM-dd', 'M"月"d"日"'], shortDate: 'MM-dd', periodBefore: true }
const dateProfiles: Readonly<Record<OfficeLocale, DateFormatProfile>> = {
  'en-US': mdy, 'fr-FR': dmy, 'zh-CN': cjk, 'ru-RU': dot, 'zh-TW': cjk,
  'zh-HK': { ...cjk, dates: ['dd/MM/yyyy', ...cjk.dates], shortDate: 'dd/MM' },
  'vi-VN': dmy, 'fa-IR': iso, 'ca-ES': { dates: ['dd/mm/yyyy', 'd/m/yyyy', 'dd/mm'], shortDate: 'dd/mm' },
  'ja-JP': cjk, 'ko-KR': { dates: ['yyyy. mm. dd.', 'yyyy"년" m"월" d"일"', 'mm-dd', 'm"월" d"일"'], shortDate: 'mm-dd', periodBefore: true },
  'es-ES': dmy, 'sk-SK': dot, 'pt-BR': dmy, 'de-DE': dot, 'it-IT': dmy, 'id-ID': dmy, 'pl-PL': dot,
  'ar-SA': { dates: ['dd/mm/yyyy', 'd/m/yyyy', 'dd/mm'], shortDate: 'dd/mm' },
}

/** Date options use UI language profiles; number presentation independently uses the region. */
export function sheetDateNumberPatterns(locale: string, currentPattern?: string): string[] {
  const profile = Object.entries(dateProfiles).find(([key]) => key.replace('-', '').toLowerCase() === locale.replace('-', '').toLowerCase())?.[1] ?? iso
  const times = ['hh:mm', 'h:mm', 'h:mm:ss'].map(pattern => profile.periodBefore ? `AM/PM ${pattern}` : `${pattern} AM/PM`)
  const result = [...new Set(['yyyy-MM-dd', ...profile.dates, 'h:mm:ss', 'h:mm', ...times, `${profile.shortDate} ${times[0]}`])]
  if (currentPattern && !result.includes(currentPattern)) result.unshift(currentPattern)
  return result
}

export const SHEET_DATE_FORMAT_PREVIEW_VALUE = (Date.UTC(1930, 7, 5, 13, 30, 30) - Date.UTC(1899, 11, 30)) / 86_400_000

export function sheetNumberFormatCategory(pattern: string): SheetNumberFormatCategory {
  if (!pattern || /^general$/i.test(pattern)) return 'general'
  if (numberFormatCurrency(pattern)) return pattern.startsWith('_(') ? 'accounting' : 'currency'
  const type = describeSheetNumberFormat(pattern)?.type
  if (type === 'date' || type === 'datetime' || type === 'time') return 'date'
  return type === 'grouped' ? 'grouped' : 'custom'
}

export function createSheetNumberFormatDraft(pattern: string, locale: string, currencies: readonly string[], category = sheetNumberFormatCategory(pattern)): SheetNumberFormatDraft {
  const source = !pattern || /^general$/i.test(pattern) ? '' : pattern
  const currency = numberFormatCurrency(source) ?? currencies[0] ?? sheetCurrencySymbol(locale)
  const choices = category === 'currency' ? sheetCurrencyNumberPatterns(currency) : SHEET_GROUPED_NUMBER_PATTERNS
  const negative = Math.max(0, choices.findIndex(candidate => numberFormatWithoutDecimals(candidate) === numberFormatWithoutDecimals(source)))
  const draft: SheetNumberFormatDraft = { category, currency, negative, decimals: Math.min(30, source ? numberFormatDecimalPlaces(source) : 2),
    pattern: category === 'date' ? sheetNumberFormatCategory(source) === 'date' ? source : sheetDateNumberPatterns(locale)[0]! : source }
  return draft
}

export function sheetNumberFormatDraftPattern(draft: SheetNumberFormatDraft): string {
  const decimals = Math.max(0, Math.min(30, Math.floor(draft.decimals)))
  if (draft.category === 'general') return 'General'
  if (draft.category === 'custom' || draft.category === 'date') return draft.pattern
  if (draft.category === 'accounting') return setNumberFormatDecimalPlaces(`_("${draft.currency}"* #,##0_)`, decimals)
  const options = draft.category === 'currency' ? sheetCurrencyNumberPatterns(draft.currency) : SHEET_GROUPED_NUMBER_PATTERNS
  return setNumberFormatDecimalPlaces(options[draft.negative] ?? options[0]!, decimals)
}

export function orderNumberFormatPreferences(candidates: readonly string[], recent: readonly string[]): string[] {
  const unique = [...new Set(candidates)]
  return unique.sort((a, b) => {
    const aIndex = recent.indexOf(a), bIndex = recent.indexOf(b)
    return (aIndex < 0 ? Infinity : aIndex) - (bIndex < 0 ? Infinity : bIndex) || 0
  })
}

export function sheetCustomNumberPatterns(locale: string, history: readonly string[]): string[] {
  return orderNumberFormatPreferences([...sheetCurrencyNumberPatterns('$'), ...sheetDateNumberPatterns(locale), ...SHEET_GROUPED_NUMBER_PATTERNS, ...history], history)
}

export interface SheetNumberFormatPreferences { history: string[]; currencies: string[] }
export const SHEET_NUMBER_FORMAT_PREFERENCE_KEY = 'office-studio.number-format-preferences'
type NumberFormatPreferenceReader = { getItem(key: string): string | null }
type NumberFormatPreferenceWriter = { setItem(key: string, value: string): void }

function strings(value: unknown): string[] {
  const list: unknown[] = Array.isArray(value) ? value : []
  return [...new Set(list.filter((item): item is string => typeof item === 'string' && item.length > 0 && item.length <= 255))]
}

export function readSheetNumberFormatPreferences(storage?: NumberFormatPreferenceReader): SheetNumberFormatPreferences {
  const serialized = storage?.getItem(SHEET_NUMBER_FORMAT_PREFERENCE_KEY)
  if (!serialized) return { history: [], currencies: [] }
  const value: unknown = JSON.parse(serialized)
  if (!value || typeof value !== 'object') return { history: [], currencies: [] }
  return { history: strings('history' in value ? value.history : undefined).slice(0, 10),
    currencies: strings('currencies' in value ? value.currencies : undefined).filter(symbol => currencySymbols.includes(symbol)) }
}

/** Preference history changes on confirmation; previews and cancellations never record a use. */
export function rememberSheetNumberFormat(preferences: SheetNumberFormatPreferences, pattern: string, custom: boolean): SheetNumberFormatPreferences {
  const currency = numberFormatCurrency(pattern)
  return { history: custom && pattern ? [pattern, ...preferences.history.filter(value => value !== pattern)].slice(0, 10) : [...preferences.history],
    currencies: currency && currencySymbols.includes(currency) ? [currency, ...preferences.currencies.filter(value => value !== currency)] : [...preferences.currencies] }
}

export function writeSheetNumberFormatPreferences(storage: NumberFormatPreferenceWriter, preferences: SheetNumberFormatPreferences): void {
  storage.setItem(SHEET_NUMBER_FORMAT_PREFERENCE_KEY, JSON.stringify(preferences))
}
