import type { DocumentPageField } from '../../../domain/model'

export type DocumentDateTimeFieldFormat = 'short' | 'long' | 'time'
export interface DocumentDateTimeFieldData {
  field: 'DATE' | 'TIME'
  instruction: string
  cachedResult: string
}

type DateTimeField = Pick<DocumentPageField, 'field' | 'instruction' | 'cachedResult'>
type PicturePart = { literal: string } | { pattern: string }
type ResultTransform = 'upper' | 'lower'

function dateLocale(locale: string, date: Date): string | undefined {
  if (!(date instanceof Date) || !Number.isFinite(date.getTime()) || date.getFullYear() < 1) return undefined
  try {
    const canonical = Intl.getCanonicalLocales(locale)[0]
    return canonical && Intl.DateTimeFormat.supportedLocalesOf([canonical]).length ? canonical : undefined
  } catch { return undefined }
}

/** Quoted instruction arguments retain picture escapes for the picture parser. */
function readArgument(instruction: string, start: number): { value: string; end: number } | undefined {
  while (/\s/.test(instruction[start] ?? '') && start < instruction.length) start++
  if (start === instruction.length) return undefined
  if (instruction[start] !== '"') {
    let end = start
    while (end < instruction.length && !/[\s\\]/.test(instruction[end]!)) end++
    return end > start ? { value: instruction.slice(start, end), end } : undefined
  }
  let cursor = start + 1
  while (cursor < instruction.length) {
    if (instruction[cursor] === '"') return { value: instruction.slice(start + 1, cursor), end: cursor + 1 }
    if (instruction[cursor] === '\\') {
      if (cursor + 1 === instruction.length) return undefined
      cursor += 2
    } else cursor++
  }
  return undefined
}

function parseInstruction(field: DateTimeField): { picture: string; transforms: ResultTransform[] } | undefined {
  if (field.field !== 'DATE' && field.field !== 'TIME' || typeof field.instruction !== 'string') return undefined
  const name = /^\s*(DATE|TIME)(?=\s|$)/i.exec(field.instruction)
  if (!name || name[1]!.toUpperCase() !== field.field) return undefined
  let cursor = name[0].length, picture: string | undefined
  const transforms: ResultTransform[] = []
  while (cursor < field.instruction.length) {
    while (cursor < field.instruction.length && /\s/.test(field.instruction[cursor]!)) cursor++
    if (cursor === field.instruction.length) break
    const kind = field.instruction.slice(cursor, cursor + 2)
    if (kind !== '\\@' && kind !== '\\*') return undefined
    const argument = readArgument(field.instruction, cursor + 2)
    if (!argument) return undefined
    cursor = argument.end
    if (kind === '\\@') {
      if (picture !== undefined) return undefined
      picture = argument.value
    } else {
      const format = argument.value.toUpperCase()
      if (format === 'UPPER') transforms.push('upper')
      else if (format === 'LOWER') transforms.push('lower')
      else if (format !== 'MERGEFORMAT') return undefined
    }
  }
  // Native regional defaults and unsupported switches keep their source cache.
  return picture === undefined || picture.length === 0 ? undefined : { picture, transforms }
}

function pictureParts(picture: string): PicturePart[] | undefined {
  const parts: PicturePart[] = []
  const literal = (value: string) => {
    const last = parts.at(-1)
    if (last && 'literal' in last) last.literal += value
    else parts.push({ literal: value })
  }
  let cursor = 0
  const escaped = () => {
    const point = picture.codePointAt(cursor + 1)
    if (point === undefined) return false
    const value = String.fromCodePoint(point)
    literal(value); cursor += 1 + value.length
    return true
  }
  while (cursor < picture.length) {
    const char = picture[cursor]!
    if (char === '\\') { if (!escaped()) return undefined; continue }
    if (char === "'") {
      cursor++
      let closed = false
      while (cursor < picture.length) {
        if (picture[cursor] === "'") {
          if (picture[cursor + 1] === "'") { literal("'"); cursor += 2 }
          else { cursor++; closed = true; break }
        } else if (picture[cursor] === '\\') { if (!escaped()) return undefined }
        else {
          const value = String.fromCodePoint(picture.codePointAt(cursor)!)
          literal(value); cursor += value.length
        }
      }
      if (!closed) return undefined
      continue
    }
    const period = /^am\/pm/i.exec(picture.slice(cursor))
    if (period) { parts.push({ pattern: period[0] }); cursor += period[0].length; continue }
    if (/[dMyhHms]/.test(char)) {
      let end = cursor + 1
      while (picture[end] === char) end++
      const size = end - cursor
      if (char === 'y' ? size !== 2 && size !== 4 : char === 'd' || char === 'M' ? size > 4 : size > 2) return undefined
      parts.push({ pattern: picture.slice(cursor, end) }); cursor = end
      continue
    }
    // Unquoted ASCII words are not silently treated as custom date tokens.
    if (/[A-Za-z]/.test(char)) return undefined
    const value = String.fromCodePoint(picture.codePointAt(cursor)!)
    literal(value); cursor += value.length
  }
  return parts
}

function formatPicture(parts: PicturePart[], locale: string, date: Date): string {
  const dayNumber = parts.some(part => 'pattern' in part && /^d{1,2}$/.test(part.pattern))
  const named = new Map<string, string>()
  const partValue = (type: 'month' | 'weekday' | 'dayPeriod', width: 'short' | 'long') => {
    const key = type + ':' + width
    const previous = named.get(key)
    if (previous !== undefined) return previous
    const options: Intl.DateTimeFormatOptions = { calendar: 'gregory', numberingSystem: 'latn' }
    if (type === 'month') { options.month = width; if (dayNumber) options.day = 'numeric' }
    else if (type === 'weekday') options.weekday = width
    else { options.hour = 'numeric'; options.hourCycle = 'h12' }
    const value = new Intl.DateTimeFormat(locale, options).formatToParts(date).find(part => part.type === type)?.value
    if (value === undefined) throw new RangeError('日期时间字段的语言缺少所需格式。')
    named.set(key, value)
    return value
  }
  const numeric = (value: number, size: number) => String(value).padStart(size, '0')
  return parts.map(part => {
    if ('literal' in part) return part.literal
    const mask = part.pattern, size = mask.length
    if (/^am\/pm$/i.test(mask)) {
      const value = partValue('dayPeriod', 'short')
      return mask === mask.toLowerCase() ? value.toLocaleLowerCase(locale) : value.toLocaleUpperCase(locale)
    }
    switch (mask[0]) {
      case 'd': return size <= 2 ? numeric(date.getDate(), size) : partValue('weekday', size === 3 ? 'short' : 'long')
      case 'M': return size <= 2 ? numeric(date.getMonth() + 1, size) : partValue('month', size === 3 ? 'short' : 'long')
      case 'y': return numeric(size === 2 ? date.getFullYear() % 100 : date.getFullYear(), size)
      case 'h': return numeric(date.getHours() % 12 || 12, size)
      case 'H': return numeric(date.getHours(), size)
      case 'm': return numeric(date.getMinutes(), size)
      case 's': return numeric(date.getSeconds(), size)
      default: throw new RangeError('日期时间字段含有未知格式。')
    }
  }).join('')
}

/** Local time and the supplied language drive explicit Word date-time pictures. */
export function formatDocumentDateTimePicture(picture: string, locale: string, date: Date): string | undefined {
  const canonical = dateLocale(locale, date), parts = pictureParts(picture)
  if (!canonical || !parts?.length) return undefined
  try { return formatPicture(parts, canonical, date) }
  catch { return undefined }
}

export function resolveDocumentDateTimeFieldResult(field: DateTimeField, locale: string, date: Date): string | undefined {
  const instruction = parseInstruction(field), canonical = dateLocale(locale, date)
  if (!instruction || !canonical) return undefined
  let result = formatDocumentDateTimePicture(instruction.picture, canonical, date)
  if (result === undefined) return undefined
  for (const transform of instruction.transforms) result = transform === 'upper' ? result.toLocaleUpperCase(canonical) : result.toLocaleLowerCase(canonical)
  return result
}

export function createDocumentDateTimeField(format: DocumentDateTimeFieldFormat, locale: string, date: Date): DocumentDateTimeFieldData {
  const canonical = dateLocale(locale, date)
  if (!canonical) throw new RangeError('日期时间字段的日期或语言无效。')
  if (format !== 'short' && format !== 'long' && format !== 'time') throw new RangeError('日期时间字段的格式无效。')
  const options: Intl.DateTimeFormatOptions = format === 'time'
    ? { hourCycle: 'h23', hour: '2-digit', minute: '2-digit', second: '2-digit' }
    : { year: 'numeric', month: format === 'long' ? 'long' : '2-digit', day: '2-digit' }
  const formatter = new Intl.DateTimeFormat(canonical, { ...options, calendar: 'gregory', numberingSystem: 'latn' })
  const tokens: Record<string, string> = { year: 'yyyy', month: format === 'long' ? 'MMMM' : 'MM', day: 'dd', hour: 'HH', minute: 'mm', second: 'ss' }
  const mask = formatter.formatToParts(date).map(part => tokens[part.type] ??
    "'" + part.value.replaceAll('\\', '\\\\').replaceAll("'", "''").replaceAll('"', '\\"') + "'").join('')
  const field = format === 'time' ? 'TIME' : 'DATE'
  return { field, instruction: field + ' \\@ "' + mask + '"', cachedResult: formatter.format(date) }
}
