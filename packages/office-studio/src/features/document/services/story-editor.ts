import type { DocumentPageField, DocumentSectionHeaderFooter } from '../../../domain/model'
import { documentStoryContent, documentStoryDefaultStyle, documentStoryTextStyle, editDocumentStoryLocks, editDocumentStoryParagraphs, isDocumentTextBoundary, validateDocumentStoryLocks, type StoryDefaultsSource } from '../model/document-story'
import { documentRunStyleAt, replaceDocumentRunRange } from '../model/document-run-range'

export type StoryKey = DocumentPageField['story']
export interface StoryLocation { page: number; area: 'header' | 'footer' }

/** A field is a single selectable object in the input bridge, independent of its displayed width. */
export function storyBuffer(section: DocumentSectionHeaderFooter, story: StoryKey) {
  const literal = section[story]
  let value = '', pos = 0
  const objects = new Map<number, DocumentPageField>()
  const literalOffsets = [0]
  const append = (end: number) => {
    value += literal.slice(pos, end)
    for (let index = pos + 1; index <= end; index++) literalOffsets.push(index)
    pos = end
  }
  const fields = (section.pageFields ?? []).filter(f => f.story === story).sort((a, b) => (a.offset ?? Infinity) - (b.offset ?? Infinity))
  for (const field of fields) {
    const offset = field.offset ?? literal.length
    if (!isDocumentTextBoundary(literal, offset)) throw new RangeError('页面字段的 UTF-16 位置无效。')
    append(offset)
    if (field.offset === undefined) { value += ' '; literalOffsets.push(offset) }
    objects.set(value.length, field)
    value += '\uFFFC'
    literalOffsets.push(offset)
    pos = offset
  }
  append(literal.length)
  return { value, objects, literalOffsets }
}

export function replaceStoryRange(section: DocumentSectionHeaderFooter, story: StoryKey, start: number, end: number, insert: string | Omit<DocumentPageField, 'offset' | 'story'>, model?: StoryDefaultsSource) {
  const { value, objects, literalOffsets } = storyBuffer(section, story)
  if (!isDocumentTextBoundary(value, start) || !isDocumentTextBoundary(value, end)) throw new RangeError('页眉页脚选区的 UTF-16 位置无效。')
  const a = Math.min(start, end), b = Math.max(start, end)
  const defaults = documentStoryDefaultStyle(model)
  const content = documentStoryContent(section, story, model)
  validateDocumentStoryLocks(content.locks, value)
  const char = typeof insert === 'string' ? insert : '\uFFFC'
  const updated = value.slice(0, a) + char + value.slice(b)
  const fields = (section.pageFields ?? []).filter(f => f.story !== story)
  let text = ''
  for (let i = 0; i < updated.length; i++) {
    const field = i < a ? objects.get(i) : i >= a + char.length ? objects.get(i - char.length + b - a) : typeof insert === 'string' ? undefined : insert
    if (field) fields.push({ ...field, story, offset: text.length,
      ...(typeof insert !== 'string' && i >= a && i < a + char.length ? { textStyle: insert.textStyle ?? documentStoryTextStyle(documentRunStyleAt(content.runs, literalOffsets[a]!, literalOffsets[b]!) ?? {}, defaults.text) } : {}) })
    else text += updated[i]
  }
  let from = literalOffsets[a]!, to = literalOffsets[b]!, insertion = typeof insert === 'string' ? insert : ''
  // Editing an old append-only field materializes its compatibility spacer.
  // Prefer exact selection offsets; diff only this legacy representation change.
  if (section[story].slice(0, from) + insertion + section[story].slice(to) !== text) {
    from = 0; to = section[story].length
    let last = text.length
    while (from < to && from < last && section[story][from] === text[from]) from++
    while (to > from && last > from && section[story][to - 1] === text[last - 1]) { to--; last-- }
    while (!isDocumentTextBoundary(section[story], from)) from--
    while (!isDocumentTextBoundary(section[story], to)) { to++; last++ }
    insertion = text.slice(from, last)
  }
  const rich = {
    ...content,
    runs: text === section[story] ? content.runs : replaceDocumentRunRange(content.runs, from, to, insertion, defaults.text),
    paragraphs: text === section[story] ? content.paragraphs : editDocumentStoryParagraphs(content.paragraphs, section[story], from, to, insertion),
    ...(content.locks ? { locks: editDocumentStoryLocks(content.locks, value, a, b, char.length) } : {}),
  }
  return { [story]: text, pageFields: fields, storyContent: { ...section.storyContent, [story]: rich } }
}

export function editStoryBuffer(section: DocumentSectionHeaderFooter, story: StoryKey, updated: string, caret?: number, model?: StoryDefaultsSource) {
  const old = storyBuffer(section, story).value
  let a = 0, b = old.length, c = updated.length
  // Preserve the actual edit point when adjacent fields or repeated characters are identical.
  while (b > 0 && c > (caret ?? updated.length) && old[b - 1] === updated[c - 1]) { b--; c-- }
  while (a < b && a < c && old[a] === updated[a]) a++
  while (b > a && c > a && old[b - 1] === updated[c - 1]) { b--; c-- }
  while (!isDocumentTextBoundary(old, a)) a--
  if ((!isDocumentTextBoundary(old, b) || !isDocumentTextBoundary(updated, c)) && b < old.length && c < updated.length) { b++; c++ }
  if (!isDocumentTextBoundary(old, b) || !isDocumentTextBoundary(updated, c)) throw new RangeError('页眉页脚输入包含无效的 UTF-16 替换范围。')
  return replaceStoryRange(section, story, a, b, updated.slice(a, c), model)
}

export function storyDateField(format: 'short' | 'long' | 'time', locale: string, date: Date) {
  const options: Intl.DateTimeFormatOptions = format === 'time' ? { hourCycle: 'h23', hour: '2-digit', minute: '2-digit', second: '2-digit' } : { year: 'numeric', month: format === 'long' ? 'long' : '2-digit', day: '2-digit' }
  const fmt = new Intl.DateTimeFormat(locale, { ...options, calendar: 'gregory', numberingSystem: 'latn' })
  const tokens: Record<string, string> = { year: 'yyyy', month: format === 'long' ? 'MMMM' : 'MM', day: 'dd', hour: 'HH', minute: 'mm', second: 'ss' }
  const mask = fmt.formatToParts(date).map(p => tokens[p.type] ?? `'${p.value.replaceAll("'", "''")}'`).join('')
  const field = format === 'time' ? 'TIME' as const : 'DATE' as const
  return { field, instruction: `${field} \\@ "${mask}"`, cachedResult: fmt.format(date) }
}
