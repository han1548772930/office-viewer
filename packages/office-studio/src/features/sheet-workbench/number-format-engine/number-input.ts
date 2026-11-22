import type { ParseData } from './types'
import type { ParseOptions } from './parse-value'
import { currencySymbols } from './constants'
import { defaultLocale, getLocale } from './locale'

interface DecimalParts { negative: boolean; digits: string; exponent: number }

function decimalParts(source: string, shift = 0): DecimalParts | null {
  const match = /^([+-])?(\d*)(?:\.(\d*))?(?:e([+-]?\d+))?$/i.exec(source)
  if (!match || !match[2] && !match[3]) return null
  let digits = (match[2]! + (match[3] ?? '')).replace(/^0+/, '')
  let exponent = Number(match[4] ?? 0) - (match[3]?.length ?? 0) + shift
  if (!Number.isSafeInteger(exponent)) return null
  if (!digits) return { negative: false, digits: '0', exponent: 0 }
  const end = digits.replace(/0+$/, '').length
  exponent += digits.length - end
  digits = digits.slice(0, end)
  return { negative: match[1] === '-', digits, exponent }
}

/** Compare decimal intent with the shortest stored number without counting punctuation as precision. */
export function decimalInputKey(source: string, shift = 0): string | null {
  const parts = decimalParts(source, shift)
  return parts ? `${parts.negative ? '-' : ''}${parts.digits}e${parts.exponent}` : null
}

/** Decimal shifts use text arithmetic so percent editing never gains binary multiplication noise. */
export function shiftDecimalInput(source: string, shift: number): string | null {
  const parts = decimalParts(source, shift)
  if (!parts) return null
  const { digits, exponent, negative } = parts
  const sign = negative ? '-' : ''
  if (digits === '0') return '0'
  const point = digits.length + exponent
  if (point > 256 || point < -256) return `${sign}${digits}e${exponent}`
  if (point <= 0) return `${sign}0.${'0'.repeat(-point)}${digits}`
  if (point >= digits.length) return sign + digits + '0'.repeat(point - digits.length)
  return `${sign}${digits.slice(0, point)}.${digits.slice(point)}`
}

export function normalizeInputDigits(source: string): string {
  return source.replace(/[\u0660-\u0669\u06f0-\u06f9\uff10-\uff19]/g, digit => {
    const code = digit.charCodeAt(0)
    return String(code >= 0xff10 ? code - 0xff10 : code >= 0x06f0 ? code - 0x06f0 : code - 0x0660)
  })
}

const inputCurrencies = [...new Set([...currencySymbols, 'NT$', 'HK$', 'US$', 'CA$', 'A$', 'NZ$', 'R$', 'Rp', 'zł', '﷼'])]
  .sort((left, right) => right.length - left.length)
const isSpace = (value: string) => /^[\s\u00a0\u202f]$/.test(value)

export interface ParsedNumberInput extends ParseData<number> {
  kind: 'number' | 'grouped' | 'currency' | 'percent' | 'scientific'
}

/** Strict localized input; a rejected conversion keeps the complete original text in the editor model. */
export function parseNumber(value: string, options: ParseOptions = {}): ParsedNumberInput | null {
  const locale = getLocale(options.locale || '') ?? defaultLocale
  const source = normalizeInputDigits(value).trim()
  let first = 0, last = source.length, negative = false, explicitSign = false
  let parentheses = false, percent = false, currency = '', currencyTrailing = false
  const percentSigns = new Set(['%', '％', locale.percent])
  const positives = new Set(['+', locale.positive])
  const negatives = new Set(['-', '−', locale.negative])
  while (first < last) {
    const char = source[first]!
    const symbol = inputCurrencies.find(item => source.startsWith(item, first))
    if (isSpace(char)) first += 1
    else if (positives.has(char) || negatives.has(char)) {
      if (explicitSign || parentheses) return null
      explicitSign = true; negative = negatives.has(char); first += 1
    } else if (char === '(') {
      if (parentheses || explicitSign) return null
      parentheses = true; negative = true; first += 1
    } else if (percentSigns.has(char)) {
      if (percent) return null
      percent = true; first += 1
    } else if (symbol) {
      if (currency) return null
      currency = symbol; first += symbol.length
    } else break
  }
  let closed = false
  while (first < last) {
    const char = source[last - 1]!
    const symbol = inputCurrencies.find(item => source.slice(first, last).endsWith(item))
    if (isSpace(char)) last -= 1
    else if (char === ')') {
      if (!parentheses || closed) return null
      closed = true; last -= 1
    } else if (percentSigns.has(char)) {
      if (percent) return null
      percent = true; last -= 1
    } else if (symbol) {
      if (currency) return null
      currency = symbol; currencyTrailing = true; last -= symbol.length
    } else break
  }
  if (parentheses !== closed || currency && percent) return null
  const body = source.slice(first, last)
  const scientific = /^(.*?)[eE]([+-]?\d+)$/.exec(body)
  if (scientific && (currency || percent)) return null
  const mantissa = scientific ? scientific[1]! : body
  const decimal = locale.decimal
  const decimalParts = mantissa.split(decimal)
  if (decimalParts.length > 2) return null
  const integer = decimalParts[0]!, fraction = decimalParts[1]
  if (fraction !== undefined && !/^\d*$/.test(fraction)) return null
  const groupingChars = new Set([locale.group, ',', '.', '\u00a0', '\u202f', "'", '’', '٬'])
  groupingChars.delete(decimal)
  if (/^[\s\u00a0\u202f]$/.test(locale.group)) groupingChars.add(' ')
  else groupingChars.delete(' ')
  const usedGroups = [...new Set([...integer].filter(char => groupingChars.has(char)))]
  if (usedGroups.length > 1) return null
  const group = usedGroups[0]
  let digits = integer
  if (group) {
    const groups = integer.split(group)
    if (!/^\d{1,3}$/.test(groups[0]!) || groups.slice(1).some(item => !/^\d{3}$/.test(item)) || fraction === '') return null
    digits = groups.join('')
  }
  if (!/^\d*$/.test(digits) || !digits && !fraction) return null
  const normalized = `${negative ? '-' : ''}${digits || '0'}${fraction === undefined ? '' : '.' + fraction}${scientific ? 'e' + scientific[2] : ''}`
  const shifted = shiftDecimalInput(normalized, percent ? -2 : 0)
  if (shifted === null) return null
  const number = Number(shifted)
  if (!Number.isFinite(number) || decimalInputKey(shifted) !== decimalInputKey(String(number))) return null
  const baseFormat = fraction ? '#,##0.00' : '#,##0'
  // Keep multi-character currency names literal while single symbols use the normal format token.
  const symbolFormat = currency.length > 1 ? `"${currency}"` : currency
  const z = scientific ? '0.00E+00' : percent ? fraction ? '0.00%' : '0%'
    : currency ? currencyTrailing ? baseFormat + symbolFormat : symbolFormat + baseFormat
    : group ? baseFormat : undefined
  const kind = scientific ? 'scientific' : percent ? 'percent' : currency ? 'currency' : group ? 'grouped' : 'number'
  return { v: number, kind, ...(z ? { z } : {}) }
}
