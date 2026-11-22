import type { CellScalar } from '../../domain/model.ts'
import { sheetNumberFormatLocale, describeSheetNumberFormat } from './number-format-render.ts'
import { DateSystem } from './number-format-engine/date-system.ts'
import { createDateTimeEditPattern } from './number-format-engine/date-time-edit-pattern.ts'
import { dateInfo, info } from './number-format-engine/format-info.ts'
import { getValueFormatSection } from './number-format-engine/format-number.ts'
import { defaultLocale, getLocale } from './number-format-engine/locale.ts'
import { shiftDecimalInput } from './number-format-engine/number-input.ts'
import { parseBool, parseDate, parseDuration, parseNumber, parseTime, type ParseOptions } from './number-format-engine/parse-value.ts'
import { parsePattern } from './number-format-engine/parse-pattern.ts'
import { excelSerialToDateTimeParts, getDateSystemMaxSerial } from './number-format-engine/serial-date.ts'
import type { FormatDateInfo, FormatInfoType, ParsedPattern, ParseData } from './number-format-engine/types.ts'
import { formatSheetValue } from './sheet-number-format.ts'

export interface SheetNumberEditOptions {
  date1904?: boolean
  region?: string
  /** Supplying the year makes partial-date input deterministic for callers and contract fixtures. */
  currentYear?: number
}

export interface SheetNumberEditResult {
  value: number | boolean
  numberFormat?: string
}

interface ActiveNumberFormat {
  parsed: ParsedPattern
  type: FormatInfoType
  date: FormatDateInfo
}

function parseOptions(options: SheetNumberEditOptions): ParseOptions & { dateSystem: DateSystem } {
  return { locale: sheetNumberFormatLocale(options.region), dateSystem: options.date1904 ? DateSystem.Date1904 : DateSystem.Date1900,
    ...(options.currentYear === undefined ? {} : { currentYear: options.currentYear }) }
}

function activeFormat(pattern: string, value: number): ActiveNumberFormat | undefined {
  try {
    const parsed = parsePattern(pattern || 'General')
    const section = getValueFormatSection(value, parsed.partitions)
    if (!section || !('scale' in section)) return undefined
    return { parsed, type: info([section]).type, date: dateInfo([section]) }
  } catch {
    // Imported, unsupported format codes remain in the model and use the raw value for editing.
    return undefined
  }
}

function decimalSeparator(options: SheetNumberEditOptions): string {
  return (getLocale(sheetNumberFormatLocale(options.region)) ?? defaultLocale).decimal
}

function localizeDecimal(value: string, options: SheetNumberEditOptions): string {
  return value.replace('.', decimalSeparator(options))
}

function trimSecondsFraction(value: string, options: SheetNumberEditOptions): string {
  const decimal = decimalSeparator(options)
  const secondsStart = value.lastIndexOf(':')
  if (secondsStart < 0) return value
  const decimalIndex = value.indexOf(decimal, secondsStart)
  if (decimalIndex < 0) return value
  let end = decimalIndex + decimal.length
  while (end < value.length && /\d/.test(value[end]!)) end += 1
  if (end === decimalIndex + decimal.length) return value
  let fractionEnd = end
  while (fractionEnd > decimalIndex + decimal.length && value[fractionEnd - 1] === '0') fractionEnd -= 1
  if (fractionEnd === decimalIndex + decimal.length) fractionEnd = decimalIndex
  return value.slice(0, fractionEnd) + value.slice(end)
}

function serializeDuration(value: number, unit: FormatDateInfo['elapsedUnit'], options: SheetNumberEditOptions): string | undefined {
  const absoluteSeconds = Math.abs(value) * 86400
  if (!Number.isFinite(absoluteSeconds) || absoluteSeconds > Number.MAX_SAFE_INTEGER) return undefined
  // Round the final subsecond component once, allowing a carry into seconds, minutes and hours.
  let wholeSeconds = Math.floor(absoluteSeconds)
  let fraction = Math.round((absoluteSeconds - wholeSeconds) * 1e10)
  if (fraction === 1e10) { wholeSeconds += 1; fraction = 0 }
  const suffix = fraction ? decimalSeparator(options) + String(fraction).padStart(10, '0').replace(/0+$/, '') : ''
  const sign = value < 0 ? '-' : ''
  if (unit === 'seconds') return sign + wholeSeconds + suffix
  const seconds = String(wholeSeconds % 60).padStart(2, '0') + suffix
  if (unit === 'minutes') return `${sign}${Math.floor(wholeSeconds / 60)}:${seconds}`
  return `${sign}${Math.floor(wholeSeconds / 3600)}:${String(Math.floor(wholeSeconds / 60) % 60).padStart(2, '0')}:${seconds}`
}

/** Editable text exposes precision and hidden calendar fields without changing a cell's persisted format. */
export function sheetNumberEditText(value: number, pattern: string, options: SheetNumberEditOptions = {}): string | undefined {
  if (!Number.isFinite(value)) return undefined
  const format = activeFormat(pattern, value)
  if (!format) return localizeDecimal(String(value), options)
  if (format.date.isDuration) return serializeDuration(value, format.date.elapsedUnit, options) ?? localizeDecimal(String(value), options)
  if (format.type === 'date' || format.type === 'time' || format.type === 'datetime') {
    const parts = excelSerialToDateTimeParts(value, { dateSystem: parseOptions(options).dateSystem })
    if (!parts) return localizeDecimal(String(value), options)
    const editPattern = createDateTimeEditPattern(format.parsed.partitions, value, 10)
    const text = editPattern ? formatSheetValue(value, editPattern, options.date1904, options.region) : undefined
    return text === undefined ? localizeDecimal(String(value), options) : trimSecondsFraction(text, options)
  }
  if (format.type === 'percent') {
    const percent = shiftDecimalInput(String(value), 2)
    return percent === null ? undefined : localizeDecimal(percent, options) + '%'
  }
  return localizeDecimal(String(value), options)
}

function temporalEdit(content: string, originalValue: number, pattern: string, options: SheetNumberEditOptions): number | undefined {
  const format = activeFormat(pattern, originalValue)
  if (!format) return undefined
  const settings = parseOptions(options)
  if (format.date.isDuration) return parseDuration(content, { ...settings, unit: format.date.elapsedUnit })?.v ?? undefined
  if (originalValue < 0) return undefined
  let value: number | undefined
  if (format.type === 'time') {
    const parsed = parseTime(content, settings)
    if (parsed) value = Math.floor(originalValue) + parsed.v
  } else if (format.type === 'date' || format.type === 'datetime') {
    const parsed = parseDate(content, settings)
    // A date-only format retains the hidden time even when the input includes a clock.
    if (parsed) value = format.type === 'date' ? Math.floor(parsed.v) + (originalValue - Math.floor(originalValue)) : parsed.v
  }
  if (value === undefined || value < 0 || value >= getDateSystemMaxSerial(settings.dateSystem) + 1) return undefined
  return value
}

function detectedResult(parsed: ParseData<number | boolean>, originalValue: CellScalar | undefined, pattern: string): SheetNumberEditResult {
  const originalFormat = typeof originalValue === 'number' && Number.isFinite(originalValue) ? activeFormat(pattern, originalValue) : undefined
  const detectedFormat = typeof parsed.v === 'number' && parsed.z ? activeFormat(parsed.z, parsed.v) : undefined
  const originalType = originalFormat?.date.isDuration ? 'duration' : originalFormat?.type ?? describeSheetNumberFormat(pattern)?.type
  const detectedType = detectedFormat?.date.isDuration ? 'duration' : detectedFormat?.type ?? (parsed.z ? describeSheetNumberFormat(parsed.z)?.type : undefined)
  return { value: parsed.v, ...(parsed.z && originalType !== detectedType ? { numberFormat: parsed.z } : {}) }
}

/** Calendar edits retain the hidden half of a numeric value; newly detected values and formats form one edit. */
export function parseSheetNumberEdit(content: string, originalValue: CellScalar | undefined, pattern: string,
  options: SheetNumberEditOptions = {}): SheetNumberEditResult | undefined {
  if (typeof originalValue === 'number' && Number.isFinite(originalValue)) {
    if (content === sheetNumberEditText(originalValue, pattern, options)) return { value: originalValue }
    const temporal = temporalEdit(content, originalValue, pattern, options)
    if (temporal !== undefined) return { value: temporal }
  }
  const settings = parseOptions(options)
  // A strict date parser precedes numeric grouping so a complete dotted date retains its calendar meaning.
  const parsed = parseDate(content, settings) ?? parseTime(content, settings) ?? parseDuration(content, settings)
    ?? parseNumber(content, settings) ?? parseBool(content, settings)
  return parsed ? detectedResult(parsed, originalValue, pattern) : undefined
}
