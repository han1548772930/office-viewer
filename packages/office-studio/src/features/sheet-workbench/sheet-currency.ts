import { currencySymbols } from './number-format-engine/constants.ts'
import { readNumberFormatSections, type NumberFormatSection } from './number-format-pattern.ts'

/** The currency region is independent of the UI language. */
export const SHEET_CURRENCY_SYMBOLS: Readonly<Record<string, string>> = Object.freeze({
  'en-US': '$', 'ru-RU': '₽', 'vi-VN': '₫', 'zh-CN': '¥', 'zh-TW': 'NT$', 'zh-HK': 'HK$',
  'fr-FR': '€', 'fa-IR': '﷼', 'ko-KR': '₩', 'es-ES': '€', 'ca-ES': '€', 'sk-SK': '€',
  'ja-JP': '¥', 'pt-BR': 'R$', 'de-DE': '€', 'it-IT': '€', 'id-ID': 'Rp', 'pl-PL': 'zł', 'ar-SA': '﷼',
})

export const SHEET_CURRENCY_CHOICES = Object.freeze([...new Set([...currencySymbols, ...Object.values(SHEET_CURRENCY_SYMBOLS)])])
const recognizedCurrencies = new Set([...SHEET_CURRENCY_CHOICES, '￥', 'US$', 'CA$', 'A$', 'NZ$'])
const multiCharacterCurrencies = [...recognizedCurrencies].filter(symbol => symbol.length > 1)

export function sheetCurrencySymbol(region = 'zh-CN'): string {
  const key = Object.keys(SHEET_CURRENCY_SYMBOLS).find(tag => tag.replace('-', '').toLowerCase() === region.replace('-', '').toLowerCase())
  return key ? SHEET_CURRENCY_SYMBOLS[key]! : '$'
}

export function sheetCurrencyPattern(region = 'zh-CN', decimals = 2): string {
  const count = Math.max(0, Math.min(30, Math.floor(decimals)))
  const fraction = count ? `.${'0'.repeat(count)}` : ''
  const symbol = sheetCurrencySymbol(region)
  return `"${symbol}"#,##0${fraction}_);[Red]("${symbol}"#,##0${fraction})`
}

export function numberFormatCurrency(source: string | NumberFormatSection): string | undefined {
  const section = typeof source === 'string' ? readNumberFormatSections(source)?.[0] : source
  if (!section) return undefined
  for (const token of section.tokens) {
    if (token.kind === 'bracket' && token.value.startsWith('$')) {
      const currency = token.value.slice(1).split('-', 1)[0]
      if (currency) return currency
    }
    if (token.kind !== 'literal' && token.kind !== 'code') continue
    const currency = token.kind === 'literal' ? token.value.trim()
      : multiCharacterCurrencies.find(symbol => section.mask.startsWith(symbol, token.start)) ?? token.value
    if (recognizedCurrencies.has(currency)) return currency
  }
  return undefined
}
