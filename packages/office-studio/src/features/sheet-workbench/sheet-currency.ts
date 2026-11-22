/** The currency region is independent of the UI language. */
export const SHEET_CURRENCY_SYMBOLS: Readonly<Record<string, string>> = Object.freeze({
  'en-US': '$', 'ru-RU': '₽', 'vi-VN': '₫', 'zh-CN': '¥', 'zh-TW': 'NT$', 'zh-HK': 'HK$',
  'fr-FR': '€', 'fa-IR': '﷼', 'ko-KR': '₩', 'es-ES': '€', 'ca-ES': '€', 'sk-SK': '€',
  'ja-JP': '¥', 'pt-BR': 'R$', 'de-DE': '€', 'it-IT': '€', 'id-ID': 'Rp', 'pl-PL': 'zł', 'ar-SA': '﷼',
})

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

export function numberFormatCurrency(pattern: string): string | undefined {
  const bracket = /\[\$([^\]-]+)(?:-[\da-f]+)?\]/i.exec(pattern)?.[1]
  if (bracket) return bracket
  // Multi-character symbols must win over the dollar sign they contain.
  return /NT\$|HK\$|R\$|Rp|zł|[$¥￥€£₽₫﷼₩]/.exec(pattern)?.[0]
}
