import type { DocumentPageField, DocumentSectionHeaderFooter } from '../../../domain/model'
import { documentStoryContent, documentStoryDefaultStyle, documentStoryInputEdit, documentStoryTextStyle, editDocumentStoryLocks, editDocumentStoryParagraphs, isDocumentTextBoundary, validateDocumentStoryLocks, type DocumentStoryRange, type StoryDefaultsSource } from '../model/document-story'
import { documentStoryFieldDisplayRuns } from '../model/document-story'
import { applyRunPatch, createRun, type RunPatch } from '../model/document-model'
import { documentRunStyleAt, replaceDocumentRunRange, validateDocumentRunPatch } from '../model/document-run-range'
import { projectDocumentSectionStories } from '../model/document-story-blocks'

export type StoryKey = DocumentPageField['story']
export interface StoryLocation { page: number; area: 'header' | 'footer' }

/** A field is a single selectable object in the input bridge, independent of its displayed width. */
export function storyBuffer(section: DocumentSectionHeaderFooter, story: StoryKey) {
  section = projectDocumentSectionStories(section)
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

/** Fields are atomic input objects; replacement inherits the first selected object or character. */
export function storyRunStyleAt(section: DocumentSectionHeaderFooter, story: StoryKey, start: number, end = start, model?: StoryDefaultsSource): RunPatch {
  const buffer = storyBuffer(section, story), defaults = documentStoryDefaultStyle(model)
  if (!isDocumentTextBoundary(buffer.value, start) || !isDocumentTextBoundary(buffer.value, end) || end < start) throw new RangeError('页眉页脚选区的 UTF-16 位置无效。')
  const position = end > start ? start : Math.max(0, start - 1)
  const field = buffer.objects.get(position)
  if (field) {
    const displayed = documentStoryFieldDisplayRuns(field, field.cachedResult ?? '', defaults.text)
    const characters = displayed.filter(run => run.text.length > 0)
    const style = (end > start || start === 0 ? characters[0] ?? displayed[0] : characters.at(-1) ?? displayed.at(-1)) ?? field.textStyle ?? defaults.text
    return { ...documentStoryTextStyle(style, { ...defaults.text, eastAsiaFontFamily: undefined, highlight: undefined, verticalAlign: undefined }),
      eastAsiaFontFamily: style.eastAsiaFontFamily, highlight: style.highlight, verticalAlign: style.verticalAlign, href: undefined, metadata: undefined }
  }
  const literal = buffer.literalOffsets[position] ?? 0
  const content = documentStoryContent(section, story, model)
  return documentRunStyleAt(content.runs, literal, literal + (buffer.value.length ? 1 : 0)) ?? defaults.text
}

export function replaceStoryRange(section: DocumentSectionHeaderFooter, story: StoryKey, start: number, end: number, insert: string | Omit<DocumentPageField, 'offset' | 'story'>, model?: StoryDefaultsSource, marks: RunPatch = {}) {
  // Validate supplied input overrides rather than rejecting a valid imported font size.
  validateDocumentRunPatch(marks)
  const { value, objects, literalOffsets } = storyBuffer(section, story)
  if (!isDocumentTextBoundary(value, start) || !isDocumentTextBoundary(value, end)) throw new RangeError('页眉页脚选区的 UTF-16 位置无效。')
  const a = Math.min(start, end), b = Math.max(start, end)
  const defaults = documentStoryDefaultStyle(model)
  const content = documentStoryContent(section, story, model)
  const { id: _id, text: _text, ...inherited } = applyRunPatch(createRun('', storyRunStyleAt(section, story, a, b, model)), marks)
  const inputStyle = documentStoryTextStyle(inherited, { ...defaults.text, eastAsiaFontFamily: undefined, highlight: undefined, verticalAlign: undefined })
  validateDocumentStoryLocks(content.locks, value)
  const char = typeof insert === 'string' ? insert : '\uFFFC'
  const updated = value.slice(0, a) + char + value.slice(b)
  const fields = (section.pageFields ?? []).filter(f => f.story !== story)
  let text = ''
  for (let i = 0; i < updated.length; i++) {
    const field = i < a ? objects.get(i) : i >= a + char.length ? objects.get(i - char.length + b - a) : typeof insert === 'string' ? undefined : insert
    if (field) fields.push({ ...field, story, offset: text.length,
      ...(typeof insert !== 'string' && i >= a && i < a + char.length ? { textStyle: insert.textStyle ?? inputStyle } : {}) })
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
    runs: from === to && !insertion.length ? content.runs : replaceDocumentRunRange(content.runs, from, to, insertion, defaults.text, {}, inherited),
    paragraphs: text === section[story] ? content.paragraphs : editDocumentStoryParagraphs(content.paragraphs, section[story], from, to, insertion),
    ...(content.locks ? { locks: editDocumentStoryLocks(content.locks, value, a, b, char.length) } : {}),
  }
  return { [story]: text, pageFields: fields, storyContent: { ...section.storyContent, [story]: rich } }
}

export function editStoryBuffer(section: DocumentSectionHeaderFooter, story: StoryKey, updated: string, caret?: number, model?: StoryDefaultsSource, marks: RunPatch = {}, sourceRange?: Pick<DocumentStoryRange, 'start' | 'end'>) {
  const old = storyBuffer(section, story).value
  if (sourceRange) {
    if (caret === undefined) throw new RangeError('页眉页脚输入缺少提交后的光标位置。')
    const edit = documentStoryInputEdit(old, updated, sourceRange, caret)
    return replaceStoryRange(section, story, edit.start, edit.end, edit.text, model, marks)
  }
  let a = 0, b = old.length, c = updated.length
  // Preserve the actual edit point when adjacent fields or repeated characters are identical.
  while (b > 0 && c > (caret ?? updated.length) && old[b - 1] === updated[c - 1]) { b--; c-- }
  while (a < b && a < c && old[a] === updated[a]) a++
  while (b > a && c > a && old[b - 1] === updated[c - 1]) { b--; c-- }
  while (!isDocumentTextBoundary(old, a)) a--
  if ((!isDocumentTextBoundary(old, b) || !isDocumentTextBoundary(updated, c)) && b < old.length && c < updated.length) { b++; c++ }
  if (!isDocumentTextBoundary(old, b) || !isDocumentTextBoundary(updated, c)) throw new RangeError('页眉页脚输入包含无效的 UTF-16 替换范围。')
  return replaceStoryRange(section, story, a, b, updated.slice(a, c), model, marks)
}

export { createDocumentDateTimeField as storyDateField } from './document-date-time-field'
