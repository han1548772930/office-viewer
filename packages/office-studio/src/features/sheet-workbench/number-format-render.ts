import type { CellScalar } from '../../domain/model.ts'
import { formatColor, formatValue, getValueFormatSection } from './number-format-engine/format-number.ts'
import { info } from './number-format-engine/format-info.ts'
import { parsePattern } from './number-format-engine/parse-pattern.ts'
import { defaultOptions } from './number-format-engine/options.ts'
import { DateSystem } from './number-format-engine/date-system.ts'
import type { FormatInfo, ParsedPattern, ResolvedFormatOptions } from './number-format-engine/types.ts'

export type SheetNumberFormatRun = { kind: 'text'; text: string } | { kind: 'space' | 'fill'; character: string }
export interface SheetNumberPresentation { text: string; color?: string; runs?: SheetNumberFormatRun[] }
export interface PositionedNumberFormatRun { text: string; offset: number }

const MAX_CACHED_PATTERNS = 256
const patterns = new Map<string, ParsedPattern>()
const FILL_MARKER = '\u0001', SPACE_MARKER = '\u0002'

/** Fixed-version numeric locales, including its documented fa-IR and ca-ES fallbacks. */
export function sheetNumberFormatLocale(locale = 'zh-CN'): string {
  const tag = locale.replace(/[-_]/g, '').toLowerCase()
  const locales: Readonly<Record<string, string>> = {
    zhcn: 'zh-CN', zhtw: 'zh-TW', zhhk: 'zh-HK', enus: 'en', frfr: 'fr', ruru: 'ru',
    vivn: 'vi', kokr: 'ko', eses: 'es', caes: 'es', sksk: 'sk', jajp: 'ja', ptbr: 'pt',
    dede: 'de', itit: 'it', idid: 'id', plpl: 'pl', arsa: 'ar', fair: 'en',
  }
  return locales[tag] ?? locale
}

function parsedFormat(pattern: string): ParsedPattern {
  const source = pattern || 'General'
  const cached = patterns.get(source)
  if (cached) {
    patterns.delete(source)
    patterns.set(source, cached)
    return cached
  }
  const parsed = parsePattern(source)
  patterns.set(source, parsed)
  if (patterns.size > MAX_CACHED_PATTERNS) {
    const oldest = patterns.keys().next().value
    if (oldest !== undefined) patterns.delete(oldest)
  }
  return parsed
}

export function isValidSheetNumberFormat(pattern: string): boolean {
  try { parsedFormat(pattern); return true } catch { return false }
}

export function describeSheetNumberFormat(pattern: string): FormatInfo | undefined {
  try { return info(parsedFormat(pattern).partitions) } catch { return undefined }
}

function formatRuns(value: number, parsed: ParsedPattern, options: ResolvedFormatOptions): SheetNumberFormatRun[] | undefined {
  const part = getValueFormatSection(value, parsed.partitions)
  if (!part || !('scale' in part)) return undefined
  const widthTokens = part.tokens.filter(token => token.type === 'fill' || token.type === 'skip')
  if (!widthTokens.length || widthTokens.some(token => typeof token.value !== 'string')) return undefined
  const encoded = formatValue(value, parsed, { ...options, fillChar: FILL_MARKER, skipChar: SPACE_MARKER })
  const runs: SheetNumberFormatRun[] = []
  let text = '', tokenIndex = 0
  for (let index = 0; index < encoded.length;) {
    const marker = encoded[index]!
    if (marker !== FILL_MARKER && marker !== SPACE_MARKER) { text += marker; index += 1; continue }
    const token = widthTokens[tokenIndex++]
    if (!token || typeof token.value !== 'string') return undefined
    if (text) { runs.push({ kind: 'text', text }); text = '' }
    runs.push({ kind: marker === FILL_MARKER ? 'fill' : 'space', character: token.value.replace(/^\\/, '') })
    index += token.value.length + 1
  }
  if (text) runs.push({ kind: 'text', text })
  return runs
}

/** Raw values and formula caches stay in the model; presentation is derived for the current region. */
export function renderSheetNumberFormat(value: CellScalar, pattern: string, date1904 = false, locale = 'zh-CN'): SheetNumberPresentation | undefined {
  try {
    const parsed = parsedFormat(pattern === '@@@' ? '@' : pattern)
    const options: ResolvedFormatOptions = { ...defaultOptions, locale: sheetNumberFormatLocale(locale),
      dateSystem: date1904 ? DateSystem.Date1904 : DateSystem.Date1900, throws: true }
    const text = formatValue(value, parsed, options)
    const color = formatColor(value, parsed, options)
    const runs = typeof value === 'number' && Number.isFinite(value) ? formatRuns(value, parsed, options) : undefined
    return { text, ...(typeof color === 'string' ? { color } : {}), ...(runs ? { runs } : {}) }
  } catch {
    // An unsupported imported pattern retains the parser's source display and original code.
    return undefined
  }
}

/** Canvas accounts for _width reservations and *fill without storing padding in a cell's value. */
export function positionSheetNumberFormatRuns(runs: readonly SheetNumberFormatRun[], availableWidth: number,
  measure: (text: string) => number): { runs: PositionedNumberFormatRun[]; width: number } {
  const widthOf = (text: string) => { const width = measure(text); return Number.isFinite(width) ? Math.max(0, width) : 0 }
  const fixedWidth = runs.reduce((width, run) => width + (run.kind === 'fill' ? 0 : widthOf(run.kind === 'text' ? run.text : run.character)), 0)
  const fills = runs.filter(run => run.kind === 'fill').length
  const fillWidth = fills ? Math.max(0, availableWidth - fixedWidth) / fills : 0
  const result: PositionedNumberFormatRun[] = []
  let offset = 0
  for (const run of runs) {
    if (run.kind === 'text') { result.push({ text: run.text, offset }); offset += widthOf(run.text) }
    else if (run.kind === 'space') offset += widthOf(run.character)
    else {
      const characterWidth = widthOf(run.character)
      const count = characterWidth ? Math.min(2048, Math.floor(fillWidth / characterWidth)) : 0
      if (run.character.trim() && count) result.push({ text: run.character.repeat(count), offset })
      offset += fillWidth
    }
  }
  return { runs: result, width: offset }
}
