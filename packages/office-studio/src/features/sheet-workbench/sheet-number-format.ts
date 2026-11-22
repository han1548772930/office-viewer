import type { CellScalar } from '../../domain/model.ts'
import { isDateNumberFormatSection, numberFormatDecimalPlaces, numberFormatWithoutDecimals, readNumberFormatSections, setNumberFormatDecimalPlaces } from './number-format-pattern.ts'
import { numberFormatCurrency, sheetCurrencyPattern, sheetCurrencySymbol } from './sheet-currency.ts'
import { renderSheetNumberFormat } from './number-format-render.ts'

export type SheetNumberFormat = 'general' | 'number' | 'currency' | 'accounting' | 'percent' | 'scientific' | 'financial' | 'roundingCurrency' | 'date' | 'time' | 'datetime' | 'duration' | 'text'
/** Exact fixed-version menu defaults; percent's direct button intentionally uses zero decimals. */
export const SHEET_NUMBER_FORMATS: Readonly<Record<SheetNumberFormat, string>> = Object.freeze({
  general: 'General', number: '0', currency: sheetCurrencyPattern(),
  accounting: '"¥" #,##0.00_);[Red]("¥"#,##0.00)', percent: '0%', scientific: '0.00E+00',
  financial: '#,##0.00;[Red]#,##0.00', roundingCurrency: '"¥"#,##0;[Red]"¥"#,##0',
  date: 'yyyy-mm-dd;@', time: 'am/pm h":"mm":"ss', datetime: 'yyyy-m-d am/pm h:mm', duration: '[h]:mm:ss', text: '@',
})

export function sheetNumberFormatPattern(kind: SheetNumberFormat, region = 'zh-CN'): string {
  const symbol = sheetCurrencySymbol(region)
  return kind === 'currency' ? sheetCurrencyPattern(region)
    : kind === 'accounting' ? `"${symbol}" #,##0.00_);[Red]("${symbol}"#,##0.00)`
    : kind === 'roundingCurrency' ? `"${symbol}"#,##0;[Red]"${symbol}"#,##0`
    : SHEET_NUMBER_FORMATS[kind]
}

export function classifySheetNumberFormat(pattern: string): SheetNumberFormat | 'custom' {
  if (!pattern || /^general$/i.test(pattern)) return 'general'
  if (pattern === '@') return 'text'
  const first = readNumberFormatSections(pattern)?.[0]
  if (!first) return 'custom'
  if (isDateNumberFormatSection(first)) {
    if (first.tokens.some(token => token.kind === 'bracket' && /^(?:h+|m+|s+)$/i.test(token.value))) return 'duration'
    return /[yd]/i.test(first.mask) ? /[hs]/i.test(first.mask) ? 'datetime' : 'date'
      : /[hs]/i.test(first.mask) ? 'time' : 'date'
  }
  if (first.mask.includes('%')) return 'percent'
  if (/[eE][+-]?[0#?]+/.test(first.mask)) return 'scientific'
  if (numberFormatCurrency(pattern)) {
    const firstPlaceholder = /[0#?]/.exec(first.mask)?.index ?? -1
    return first.tokens.some(token => token.kind === 'fill') || firstPlaceholder > 0 && first.mask[firstPlaceholder - 1] === ' ' ? 'accounting' : 'currency'
  }
  if (numberFormatWithoutDecimals(pattern) === numberFormatWithoutDecimals(SHEET_NUMBER_FORMATS.financial)) return 'financial'
  const mask = first.mask.replaceAll('\u0000', '').trim()
  return /^[#,0?.()\-+ ]+$/.test(mask) && /[0#?]/.test(mask) ? 'number' : 'custom'
}

export const SHEET_NUMBER_FORMAT_LABELS: Readonly<Record<SheetNumberFormat | 'custom', string>> = Object.freeze({
  general: '常规', text: '文本', number: '数字', percent: '百分比', scientific: '科学计数',
  accounting: '会计专用', financial: '财务数值', currency: '货币', roundingCurrency: '整数货币',
  date: '日期', time: '时间', datetime: '日期和时间', duration: '持续时间', custom: '自定义格式',
})

export type SheetNumberFormatMenuEntry = { type: 'format'; id: SheetNumberFormat; label: string; pattern: string | null }
  | { type: 'divider'; id: string }
  | { type: 'more'; id: 'more'; label: string }

/** sheets-numfmt-ui/menu/number-format-options.ts, including separators and the distinct More action. */
export function sheetNumberFormatMenu(region = 'zh-CN'): SheetNumberFormatMenuEntry[] {
  const entries: Array<SheetNumberFormat | '|'> = ['general', 'text', '|', 'number', 'percent', 'scientific', '|', 'accounting', 'financial', 'currency', 'roundingCurrency', '|', 'date', 'time', 'datetime', 'duration', '|']
  return [...entries.map((id, index): SheetNumberFormatMenuEntry => id === '|'
    ? { type: 'divider', id: `divider-${index}` }
    : { type: 'format', id, label: SHEET_NUMBER_FORMAT_LABELS[id], pattern: id === 'general' ? null : id === 'percent' ? '0.00%' : sheetNumberFormatPattern(id, region) }),
  { type: 'more', id: 'more', label: '更多格式' }]
}

export function sheetNumberFormatKey(pattern: string | null, region = 'zh-CN'): SheetNumberFormat | 'mixed' | 'custom' {
  if (pattern === null) return 'mixed'
  if (!pattern || /^general$/i.test(pattern)) return 'general'
  const preset = sheetNumberFormatMenu(region).find(entry => entry.type === 'format' && entry.pattern && numberFormatWithoutDecimals(entry.pattern) === numberFormatWithoutDecimals(pattern))
  return preset?.type === 'format' ? preset.id : classifySheetNumberFormat(pattern)
}

export function sheetNumberFormatLabel(pattern: string | null, region = 'zh-CN'): string {
  const key = sheetNumberFormatKey(pattern, region)
  return key === 'mixed' ? '混合' : SHEET_NUMBER_FORMAT_LABELS[key]
}

export function changeSheetDecimalPlaces(pattern: string, delta: -1 | 1): string {
  const source = !pattern || /^general$/i.test(pattern) ? '0' : pattern
  return setNumberFormatDecimalPlaces(source, numberFormatDecimalPlaces(source) + delta)
}

/** Unsupported imported formats retain their source display; supported formats are always recalculated. */
export function formatSheetValue(value: CellScalar, pattern: string, date1904 = false, locale = 'zh-CN'): string | undefined {
  return renderSheetNumberFormat(value, pattern, date1904, locale)?.text
}

export { renderSheetNumberFormat }
