interface DecimalParts {
  negative: boolean
  digits: string
  exponent: number
}

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

/** Compare decimal intent without losing leading/trailing zeros or scientific scale to a Number conversion. */
export function decimalInputKey(source: string, shift = 0): string | null {
  const parts = decimalParts(source, shift)
  return parts ? `${parts.negative ? '-' : ''}${parts.digits}e${parts.exponent}` : null
}

/** Shift decimal text directly so percent edit values do not gain binary multiplication noise. */
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

/** Return a finite scalar only when its shortest decimal text preserves the input's complete decimal intent. */
export function parseExactDecimalInput(source: string, shift = 0): number | null {
  const shifted = shiftDecimalInput(source, shift)
  if (shifted === null) return null
  const value = Number(shifted)
  return Number.isFinite(value) && decimalInputKey(shifted) === decimalInputKey(String(value)) ? value : null
}
