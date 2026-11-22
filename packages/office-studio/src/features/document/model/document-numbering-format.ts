export interface NumberingOrdinal {
  text: string
  recognized: boolean
}

const decimalZeroCodePoints = [
  0x30, 0x660, 0x6f0, 0x7c0, 0x966, 0x9e6, 0xa66, 0xae6, 0xb66, 0xbe6,
  0xc66, 0xce6, 0xd66, 0xde6, 0xe50, 0xed0, 0xf20, 0x1040, 0x1090, 0x17e0,
  0x1810, 0x1946, 0x19d0, 0x1a80, 0x1a90, 0x1b50, 0x1bb0, 0x1c40, 0x1c50,
  0xa620, 0xa8d0, 0xa900, 0xa9d0, 0xa9f0, 0xaa50, 0xabf0, 0xff10, 0x104a0,
  0x10d30, 0x11066, 0x110f0, 0x11136, 0x111d0, 0x112f0, 0x11450, 0x114d0,
  0x11650, 0x116c0, 0x11730, 0x118e0, 0x11950, 0x11c50, 0x11d50, 0x11da0,
  0x16a60, 0x16ac0, 0x16b50, 0x1d7ce, 0x1d7d8, 0x1d7e2, 0x1d7ec, 0x1d7f6,
  0x1e140, 0x1e2f0, 0x1e950, 0x1fbf0,
]

const customAlphabets: Record<string, string> = {
  A: 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', a: 'abcdefghijklmnopqrstuvwxyz',
  Α: 'ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ', α: 'αβγδεζηθικλμνξοπρστυφχψω',
  А: 'АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ', а: 'абвгдеёжзийклмнопрстуфхцчшщъыьэюя',
  ア: 'アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン',
  あ: 'あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわをん',
}

function roman(value: number, uppercase: boolean): string | undefined {
  if (value < 1 || value > 3_999_999) return
  let result = '', remainder = value
  for (const [amount, glyph] of [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']] as const) {
    while (remainder >= amount) { result += glyph; remainder -= amount }
  }
  return uppercase ? result.toUpperCase() : result
}

function alphabetic(value: number, alphabet: string): string {
  const glyphs = [...alphabet]
  let result = '', remainder = value
  while (remainder > 0) {
    remainder--
    result = glyphs[remainder % glyphs.length] + result
    remainder = Math.floor(remainder / glyphs.length)
  }
  return result
}

function customOrdinal(value: number, pattern: string): NumberingOrdinal {
  if (pattern.length > 4096) return { text: String(value), recognized: false }
  const tokens = pattern.match(/[\p{L}\p{N}]+|[^\p{L}\p{N}]+/gu) ?? []
  const index = tokens.findIndex(token => /^[\p{L}\p{N}]/u.test(token))
  if (index < 0) return { text: String(value), recognized: false }
  const token = tokens[index]!, characters = [...token]
  const prefix = index ? tokens[0]! : ''
  const last = tokens.at(-1) ?? ''
  const suffix = /^[^\p{L}\p{N}]/u.test(last) ? last : ''
  let formatted: string | undefined
  const final = characters.at(-1)?.codePointAt(0)
  const zero = decimalZeroCodePoints.find(code => final === code + 1 && characters.slice(0, -1).every(glyph => glyph.codePointAt(0) === code))
  if (zero !== undefined && characters.length <= 255) {
    formatted = String(value).padStart(characters.length, '0').replace(/[0-9]/g, digit => String.fromCodePoint(zero + Number(digit)))
  } else if (token === 'I' || token === 'i') formatted = value === 0 ? '0' : roman(value, token === 'I')
  else if (customAlphabets[token]) formatted = alphabetic(value, customAlphabets[token]!)
  if (formatted === undefined) return { text: String(value), recognized: false }
  // XSLT numbering converts zero to a string instead of applying a positive ordinal token.
  return { text: value === 0 ? '0' : prefix + formatted + suffix, recognized: true }
}

/** A custom token uses XSLT numbering syntax; the OOXML format remains its fallback. */
export function numberingOrdinal(value: number, format: string, minimumDigits = 0, customFormat?: string): NumberingOrdinal {
  if (!Number.isSafeInteger(value) || value < 0) return { text: String(value), recognized: false }
  if (customFormat !== undefined) {
    const custom = customOrdinal(value, customFormat)
    if (custom.recognized) return custom
  }
  let result: string | undefined
  if (format === 'lowerLetter' || format === 'upperLetter') {
    if (value > 0 && Math.ceil(value / 26) <= 4096) {
      const letter = String.fromCharCode((format === 'lowerLetter' ? 97 : 65) + (value - 1) % 26)
      result = letter.repeat(Math.ceil(value / 26))
    }
  } else if (format === 'lowerRoman' || format === 'upperRoman') result = roman(value, format === 'upperRoman')
  else if (format === 'ordinal') result = `${value}${value % 100 >= 11 && value % 100 <= 13 ? 'th' : value % 10 === 1 ? 'st' : value % 10 === 2 ? 'nd' : value % 10 === 3 ? 'rd' : 'th'}`
  else if (format === 'hex') result = value.toString(16).toUpperCase()
  else if (format === 'decimal' || format === 'decimalZero') {
    const width = Math.max(format === 'decimalZero' ? 2 : 0, Math.max(0, Math.min(255, Math.trunc(minimumDigits) || 0)))
    result = value === 0 ? '0' : String(value).padStart(width, '0')
  }
  if (value === 0 && ['lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman'].includes(format)) result = '0'
  return { text: result ?? String(value), recognized: result !== undefined && customFormat === undefined }
}
