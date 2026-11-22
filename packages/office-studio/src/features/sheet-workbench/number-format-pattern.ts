export interface NumberFormatToken {
  kind: 'code' | 'literal' | 'bracket' | 'space' | 'fill'
  value: string
  start: number
  end: number
}

export interface NumberFormatSection {
  text: string
  /** Literals retain their positions but cannot be mistaken for format placeholders. */
  mask: string
  tokens: NumberFormatToken[]
}

export const MAX_SHEET_DECIMALS = 30

/** OOXML sections split only on unescaped semicolons outside literals and brackets. */
export function readNumberFormatSections(pattern: string): NumberFormatSection[] | undefined {
  const sections: NumberFormatSection[] = []
  let start = 0
  let tokens: NumberFormatToken[] = []
  let mask = ''
  const masked = (length: number) => '\u0000'.repeat(length)
  for (let index = 0; index < pattern.length;) {
    const character = pattern[index]!
    if (character === ';') {
      sections.push({ text: pattern.slice(start, index), mask, tokens })
      index += 1; start = index; tokens = []; mask = ''
      continue
    }
    const tokenStart = index
    let token: NumberFormatToken
    if (character === '"') {
      index += 1
      let value = ''
      while (index < pattern.length && pattern[index] !== '"') {
        if (pattern[index] === '\\' && index + 1 < pattern.length) index += 1
        value += pattern[index]!
        index += 1
      }
      if (index === pattern.length) return undefined
      index += 1
      token = { kind: 'literal', value, start: tokenStart - start, end: index - start }
    } else if (character === '\\' || character === '_' || character === '*') {
      if (index + 1 >= pattern.length) return undefined
      index += 2
      token = { kind: character === '\\' ? 'literal' : character === '_' ? 'space' : 'fill', value: pattern[tokenStart + 1]!, start: tokenStart - start, end: index - start }
    } else if (character === '[') {
      const end = pattern.indexOf(']', index + 1)
      if (end < 0 || pattern.slice(index + 1, end).includes('[')) return undefined
      index = end + 1
      token = { kind: 'bracket', value: pattern.slice(tokenStart + 1, end), start: tokenStart - start, end: index - start }
    } else {
      if (character === ']' || character.charCodeAt(0) < 32) return undefined
      index += 1
      token = { kind: 'code', value: character, start: tokenStart - start, end: index - start }
    }
    tokens.push(token)
    mask += token.kind === 'code' ? token.value : masked(token.end - token.start)
  }
  sections.push({ text: pattern.slice(start), mask, tokens })
  return sections.length <= 4 ? sections : undefined
}

export function isDateNumberFormatSection(section: NumberFormatSection): boolean {
  return /[ymdhs]/i.test(section.mask) || section.tokens.some(token => token.kind === 'bracket' && /^(?:h+|m+|s+)$/i.test(token.value))
}

export function numberFormatDecimalPlaces(pattern: string): number {
  const section = readNumberFormatSections(pattern)?.[0]
  if (!section || isDateNumberFormatSection(section)) return 0
  return /\.([0#?]+)/.exec(section.mask)?.[1]?.length ?? 0
}

/** Fixed v1.6.0 decimal rules: copy optional placeholders and leave date/time sections intact. */
export function setNumberFormatDecimalPlaces(pattern: string, decimalPlaces: number): string {
  const sections = readNumberFormatSections(pattern)
  if (!sections) return pattern
  const count = Math.min(MAX_SHEET_DECIMALS, Math.max(0, Math.floor(decimalPlaces)))
  return sections.map(section => {
    if (isDateNumberFormatSection(section)) return section.text
    const dot = section.mask.indexOf('.')
    if (dot < 0) {
      if (!count) return section.text
      // A fractional format or an exponent must not receive a decimal in its denominator/exponent.
      const integerMask = section.mask.split(/[eE/]/)[0] ?? ''
      const zero = /0(?!0)/.exec(integerMask)
      if (!zero) return section.text
      const insertion = zero.index + 1
      return `${section.text.slice(0, insertion)}.${'0'.repeat(count)}${section.text.slice(insertion)}`
    }
    let end = dot + 1
    while (/[0#?]/.test(section.mask[end] ?? '\u0000')) end += 1
    const fraction = section.text.slice(dot + 1, end)
    const next = count <= fraction.length ? fraction.slice(0, count)
      : fraction + (fraction.at(-1) ?? '0').repeat(count - fraction.length)
    return section.text.slice(0, dot) + (next ? `.${next}` : '') + section.text.slice(end)
  }).join(';')
}

/** Preset labels ignore fractional precision, but preserve currencies, literals and section order. */
export function numberFormatWithoutDecimals(pattern: string): string {
  return setNumberFormatDecimalPlaces(pattern, 0).toLowerCase()
}

export function scientificPatternForValue(value: number): string | undefined {
  const text = String(value)
  if (!Number.isFinite(value) || !/e/i.test(text)) return undefined
  const mantissa = text.split(/e/i)[0] ?? ''
  const decimalPlaces = Math.max(2, mantissa.split('.')[1]?.length ?? 0)
  return `0.${'0'.repeat(decimalPlaces)}E+00`
}

export function rawNumberDecimalPlaces(value: number): number {
  const scientific = scientificPatternForValue(value)
  return scientific ? numberFormatDecimalPlaces(scientific) : /\.(\d*)$/.exec(String(value))?.[1]?.length ?? 0
}
