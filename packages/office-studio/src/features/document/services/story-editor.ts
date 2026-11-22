import type { DocumentPageField, DocumentSectionHeaderFooter } from '../../../domain/model'

export type StoryKey = DocumentPageField['story']
export interface StoryLocation { page: number; area: 'header' | 'footer' }

/** A field is a single selectable object in the input bridge, independent of its displayed width. */
export function storyBuffer(section: DocumentSectionHeaderFooter, story: StoryKey) {
  const literal = section[story]
  let value = '', pos = 0
  const objects = new Map<number, DocumentPageField>()
  const fields = (section.pageFields ?? []).filter(f => f.story === story).sort((a, b) => (a.offset ?? Infinity) - (b.offset ?? Infinity))
  for (const field of fields) {
    const offset = Math.min(field.offset ?? literal.length, literal.length)
    value += literal.slice(pos, offset) + (field.offset === undefined ? ' ' : '')
    objects.set(value.length, field)
    value += '\uFFFC'
    pos = offset
  }
  return { value: value + literal.slice(pos), objects }
}

export function replaceStoryRange(section: DocumentSectionHeaderFooter, story: StoryKey, start: number, end: number, insert: string | Omit<DocumentPageField, 'offset' | 'story'>) {
  const { value, objects } = storyBuffer(section, story)
  const a = Math.max(0, Math.min(start, value.length)), b = Math.max(a, Math.min(end, value.length))
  const char = typeof insert === 'string' ? insert : '\uFFFC'
  const updated = value.slice(0, a) + char + value.slice(b)
  const fields = (section.pageFields ?? []).filter(f => f.story !== story)
  let text = ''
  for (let i = 0; i < updated.length; i++) {
    const field = i < a ? objects.get(i) : i >= a + char.length ? objects.get(i - char.length + b - a) : typeof insert === 'string' ? undefined : insert
    if (field) fields.push({ ...field, story, offset: text.length })
    else text += updated[i]
  }
  return { [story]: text, pageFields: fields }
}

export function editStoryBuffer(section: DocumentSectionHeaderFooter, story: StoryKey, updated: string, caret?: number) {
  const old = storyBuffer(section, story).value
  let a = 0, b = old.length, c = updated.length
  // Preserve the actual edit point when adjacent fields or repeated characters are identical.
  while (b > 0 && c > (caret ?? updated.length) && old[b - 1] === updated[c - 1]) { b--; c-- }
  while (a < b && a < c && old[a] === updated[a]) a++
  while (b > a && c > a && old[b - 1] === updated[c - 1]) { b--; c-- }
  return replaceStoryRange(section, story, a, b, updated.slice(a, c))
}

export function storyDateField(format: 'short' | 'long' | 'time', locale: string, date: Date) {
  const options: Intl.DateTimeFormatOptions = format === 'time' ? { hourCycle: 'h23', hour: '2-digit', minute: '2-digit', second: '2-digit' } : { year: 'numeric', month: format === 'long' ? 'long' : '2-digit', day: '2-digit' }
  const fmt = new Intl.DateTimeFormat(locale, { ...options, calendar: 'gregory', numberingSystem: 'latn' })
  const tokens: Record<string, string> = { year: 'yyyy', month: format === 'long' ? 'MMMM' : 'MM', day: 'dd', hour: 'HH', minute: 'mm', second: 'ss' }
  const mask = fmt.formatToParts(date).map(p => tokens[p.type] ?? `'${p.value.replaceAll("'", "''")}'`).join('')
  const field = format === 'time' ? 'TIME' as const : 'DATE' as const
  return { field, instruction: `${field} \\@ "${mask}"`, cachedResult: fmt.format(date) }
}
