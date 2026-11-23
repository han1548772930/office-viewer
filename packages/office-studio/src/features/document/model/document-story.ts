import type { DocumentDefaultStyle, DocumentModel, DocumentPageField, DocumentSectionHeaderFooter, DocumentStoryContent, DocumentStoryLock, DocumentStoryParagraph } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { documentDefaultStyle } from './document-default-style'
import { createRun, runText } from './document-model'
import { documentRunStyleAt, splitDocumentRunRange } from './document-run-range'
import type { DocumentRun, RunPatch } from './document-model'
import { projectDocumentStoryBlocks } from './document-story-blocks'

export const DOCUMENT_STORY_KEYS = ['header', 'footer', 'firstPageHeader', 'firstPageFooter', 'evenPageHeader', 'evenPageFooter'] as const
export type DocumentStoryKey = DocumentPageField['story']
export interface DocumentStoryRange { sectionId: string; story: DocumentStoryKey; start: number; end: number }
export type StoryDefaultsSource = Pick<DocumentModel, 'defaultStyle'>
export interface DocumentStoryFieldCache { version: 1; runs: DocumentRun[] }

/** A field remains one input object while its cached result keeps many runs. */
export function documentStoryFieldCache(field: DocumentPageField): DocumentStoryFieldCache | undefined {
  const raw = field.metadata?.officeStoryFieldCache
  if (raw === undefined) return undefined
  if (!raw || typeof raw !== 'object' || Array.isArray(raw) || Reflect.get(raw, 'version') !== 1) throw new TypeError('页眉页脚字段缓存版本无效。')
  const runs: unknown = Reflect.get(raw, 'runs')
  if (!Array.isArray(runs)) throw new TypeError('页眉页脚字段缓存范围无效。')
  for (const run of runs) {
    if (!run || typeof run !== 'object' || Array.isArray(run) || typeof Reflect.get(run, 'id') !== 'string' || typeof Reflect.get(run, 'text') !== 'string') throw new TypeError('页眉页脚字段缓存字符格式无效。')
  }
  // The structural narrowing above and full story validation below form the
  // boundary contract; metadata never bypasses the style/range validator.
  const cache = raw as DocumentStoryFieldCache
  const text = runText(cache.runs)
  if (text !== (field.cachedResult ?? '')) throw new RangeError('页眉页脚字段缓存文本与格式范围不一致。')
  validateDocumentStoryContent({ runs: cache.runs, paragraphs: documentStoryParagraphs(text, documentStoryDefaultStyle().paragraph) }, text)
  return cache
}

/** Changing a field's modeled style overlays only the properties that changed. */
export function documentStoryFieldDisplayRuns(field: DocumentPageField, text: string, defaults: DocumentDefaultStyle['text']): DocumentRun[] {
  const cache = documentStoryFieldCache(field)
  if (!cache?.runs.length) return [createRun(text, field.textStyle ?? defaults)]
  const cleared = field.metadata?.officeStoryTextFormatCleared === true
  const style = field.textStyle ?? (cleared ? defaults : undefined)
  const baseline = field.metadata?.officeStoryStyleBaseline
  const displayed = !style ? cache.runs : cache.runs.map(run => {
    const next = { ...run, metadata: { ...run.metadata } }
    for (const key of ['bold', 'italic', 'underline', 'strike', 'color', 'fontFamily', 'eastAsiaFontFamily', 'fontSize', 'highlight', 'verticalAlign'] as const) {
      const current = style[key]
      // A materialized cache already contains a clear. Later one-property edits
      // retain its other ranges while the source-replay clear marker stays set.
      if (baseline && typeof baseline === 'object' && !Array.isArray(baseline) && current === Reflect.get(baseline, key)) continue
      if (current === undefined) {
        if (key === 'eastAsiaFontFamily' || key === 'highlight' || key === 'verticalAlign') delete next[key]
      }
      else Object.assign(next, { [key]: current })
      if (key === 'highlight') delete next.metadata.highlight
    }
    return next
  })
  return remapDocumentStoryFieldRuns(displayed, text, defaults)
}

/** Preserve cached format spans while a live or refreshed field result changes. */
export function remapDocumentStoryFieldRuns(runs: readonly DocumentRun[], text: string, defaults: DocumentDefaultStyle['text']): DocumentRun[] {
  const previous = runText(runs)
  if (text === previous) return cloneModel([...runs])
  const edit = documentStoryTextEdit(previous, text)
  const { before, after } = splitDocumentRunRange(runs, edit.start, edit.end)
  const spans: Array<{ run: DocumentRun; length: number }> = []
  let offset = 0
  for (const run of runs) {
    const from = offset, to = from + run.text.length
    const length = Math.max(0, Math.min(to, edit.end) - Math.max(from, edit.start))
    if (length || !run.text.length && from > edit.start && from < edit.end) spans.push({ run, length })
    offset = to
  }
  if (!spans.some(span => span.length > 0)) {
    const source = createRun('', documentRunStyleAt(runs, edit.start, edit.end) ?? defaults)
    spans.push({ run: source, length: 0 })
  }
  let finalSpan = spans.length - 1
  for (let index = spans.length - 1; index >= 0; index--) if (spans[index]!.length > 0) { finalSpan = index; break }
  const changed: DocumentRun[] = []
  let cursor = 0
  for (const [index, span] of spans.entries()) {
    let end = index === finalSpan ? edit.text.length : Math.min(edit.text.length, cursor + span.length)
    while (end > cursor && !isDocumentTextBoundary(edit.text, end)) end--
    const { id: _id, text: _text, ...style } = span.run
    changed.push(createRun(edit.text.slice(cursor, end), style))
    cursor = end
  }
  const result = [...before, ...changed, ...after]
  if (runText(result) !== text) throw new RangeError('页眉页脚字段结果与格式范围不一致。')
  return cloneModel(result)
}

/** Fixed Univer Docs uses 9 pt for header/footer defaults while keeping the document's family. */
export function documentStoryDefaultStyle(model?: StoryDefaultsSource): DocumentDefaultStyle {
  const defaults = documentDefaultStyle(model)
  return { version: 1, text: { ...defaults.text, fontSize: 9 }, paragraph: { ...defaults.paragraph } }
}

export function documentStoryTextStyle(style: RunPatch, defaults: DocumentDefaultStyle['text']): DocumentDefaultStyle['text'] {
  const source = { ...defaults, ...style }
  const { bold, italic, underline, strike, color, fontFamily, eastAsiaFontFamily, fontSize, verticalAlign, highlight } = source
  return { bold, italic, underline, strike, color, fontFamily, fontSize,
    ...(eastAsiaFontFamily !== undefined ? { eastAsiaFontFamily } : {}),
    ...(verticalAlign !== undefined ? { verticalAlign } : {}), ...(highlight !== undefined ? { highlight } : {}) }
}

export function isDocumentTextBoundary(text: string, offset: number): boolean {
  if (!Number.isSafeInteger(offset) || offset < 0 || offset > text.length) return false
  const previous = text.charCodeAt(offset - 1), next = text.charCodeAt(offset)
  return !(previous >= 0xD800 && previous <= 0xDBFF && next >= 0xDC00 && next <= 0xDFFF)
}

/** UI restoration may follow an older snapshot; API and command ranges remain strict. */
export function clampDocumentTextRange(text: string, range: { start: number; end: number }): { start: number; end: number } {
  const clamp = (offset: number) => Math.max(0, Math.min(text.length, Number.isSafeInteger(offset) ? offset : 0))
  let start = clamp(Math.min(range.start, range.end)), end = clamp(Math.max(range.start, range.end))
  if (start === end) {
    if (!isDocumentTextBoundary(text, start)) start--
    return { start, end: start }
  }
  if (!isDocumentTextBoundary(text, start)) start--
  if (!isDocumentTextBoundary(text, end)) end++
  return { start, end }
}

export function validateDocumentStoryContent(content: DocumentStoryContent, text: string): void {
  if (content?.blocks !== undefined) content = projectDocumentStoryBlocks(content).content
  if (!content || !Array.isArray(content.runs) || !Array.isArray(content.paragraphs) || !content.paragraphs.length) throw new TypeError('页眉页脚富文本数据无效。')
  let runOffset = 0
  for (const run of content.runs) {
    if (!run || typeof run.id !== 'string' || typeof run.text !== 'string') throw new TypeError('页眉页脚字符范围无效。')
    runOffset += run.text.length
    if (!isDocumentTextBoundary(text, runOffset) || run.metadata !== undefined && (!run.metadata || typeof run.metadata !== 'object' || Array.isArray(run.metadata))) throw new TypeError('页眉页脚字符格式边界无效。')
    documentDefaultStyle({ defaultStyle: { version: 1, text: run, paragraph: { align: 'left', indent: 0, lineHeight: 1, spaceBefore: 0, spaceAfter: 0 } } })
  }
  if (runText(content.runs) !== text) throw new RangeError('页眉页脚文本与格式范围不一致。')
  let next = 0
  for (const [index, paragraph] of content.paragraphs.entries()) {
    if (!paragraph || paragraph.startOffset !== next || !isDocumentTextBoundary(text, paragraph.startOffset)
      || !isDocumentTextBoundary(text, paragraph.endOffset) || paragraph.endOffset < paragraph.startOffset
      || index < content.paragraphs.length - 1 && text[paragraph.endOffset] !== '\n'
      || index === content.paragraphs.length - 1 && paragraph.endOffset !== text.length
      || text.slice(paragraph.startOffset, paragraph.endOffset).includes('\n')
      || paragraph.metadata !== undefined && (!paragraph.metadata || typeof paragraph.metadata !== 'object' || Array.isArray(paragraph.metadata))) throw new RangeError('页眉页脚段落范围无效。')
    documentDefaultStyle({ defaultStyle: { version: 1, text: documentStoryDefaultStyle().text, paragraph: paragraph.style } })
    next = paragraph.endOffset + 1
  }
  if (content.metadata !== undefined && (!content.metadata || typeof content.metadata !== 'object' || Array.isArray(content.metadata))) throw new TypeError('页眉页脚来源数据无效。')
}

export function validateDocumentStoryLocks(locks: readonly DocumentStoryLock[] | undefined, buffer: string): void {
  if (locks === undefined) return
  if (!Array.isArray(locks)) throw new TypeError('页眉页脚内容锁无效。')
  const ids = new Set<string>()
  for (const lock of locks) {
    if (!lock || typeof lock.id !== 'string' || !lock.id || ids.has(lock.id)
      || !['contentLocked', 'sdtLocked', 'sdtContentLocked'].includes(lock.lock)
      || !isDocumentTextBoundary(buffer, lock.startOffset) || !isDocumentTextBoundary(buffer, lock.endOffset)
      || lock.endOffset < lock.startOffset) throw new RangeError('页眉页脚内容锁范围无效。')
    ids.add(lock.id)
  }
  for (let a = 0; a < locks.length; a++) for (let b = a + 1; b < locks.length; b++) {
    const first = locks[a]!, second = locks[b]!
    if (first.startOffset < second.startOffset && second.startOffset < first.endOffset && first.endOffset < second.endOffset
      || second.startOffset < first.startOffset && first.startOffset < second.endOffset && second.endOffset < first.endOffset) throw new RangeError('页眉页脚内容锁不能交叉重叠。')
  }
}

export function documentStoryTextEdit(before: string, after: string): { start: number; end: number; text: string } {
  let start = 0, end = before.length, last = after.length
  while (start < end && start < last && before[start] === after[start]) start++
  while (end > start && last > start && before[end - 1] === after[last - 1]) { end--; last-- }
  while (start > 0 && (!isDocumentTextBoundary(before, start) || !isDocumentTextBoundary(after, start))) start--
  while (end < before.length && last < after.length && (!isDocumentTextBoundary(before, end) || !isDocumentTextBoundary(after, last))) { end++; last++ }
  return { start, end, text: after.slice(start, last) }
}

/** Resolve a native input against its captured selection rather than a repeated-text diff. */
export function documentStoryInputEdit(before: string, after: string, range: Pick<DocumentStoryRange, 'start' | 'end'>, caret: number): { start: number; end: number; text: string } {
  if (typeof before !== 'string' || typeof after !== 'string' || !range || !isDocumentTextBoundary(before, range.start)
    || !isDocumentTextBoundary(before, range.end) || range.end < range.start || !isDocumentTextBoundary(after, caret)) {
    throw new RangeError('页眉页脚输入的选区或光标位置无效。')
  }
  const delta = after.length - before.length
  const splice = (start: number, end: number, length: number) => length >= 0 && caret === start + length
    && after.slice(0, start) === before.slice(0, start) && after.slice(caret) === before.slice(end)
    ? { start, end, text: after.slice(start, caret) } : undefined
  const replacement = splice(range.start, range.end, delta + range.end - range.start)
  if (replacement) return replacement
  if (range.start === range.end && delta < 0) {
    const backward = isDocumentTextBoundary(before, range.start + delta) ? splice(range.start + delta, range.start, 0) : undefined
    if (backward) return backward
    const forward = isDocumentTextBoundary(before, range.start - delta) ? splice(range.start, range.start - delta, 0) : undefined
    if (forward) return forward
  }
  throw new RangeError('页眉页脚输入与提交前的选区不一致，请重新输入。')
}

/** Plan every lock before mutating text or anchors. Wrapper-only locks retain an empty wrapper. */
export function editDocumentStoryLocks(locks: readonly DocumentStoryLock[] | undefined, buffer: string, start: number, end: number, replacementLength: number): DocumentStoryLock[] | undefined {
  validateDocumentStoryLocks(locks, buffer)
  if (!locks) return undefined
  const delta = replacementLength - (end - start)
  return locks.map(lock => {
    const insertion = start === end
    const intersects = insertion ? start > lock.startOffset && start < lock.endOffset
      || lock.startOffset === lock.endOffset && start === lock.startOffset && replacementLength > 0
      : start < lock.endOffset && lock.startOffset < end
    if (intersects && lock.lock !== 'sdtLocked') throw new Error(`页眉页脚内容已锁定：${lock.id}`)
    if (end <= lock.startOffset) return { ...lock, startOffset: lock.startOffset + delta, endOffset: lock.endOffset + delta }
    if (start >= lock.endOffset) return { ...lock }
    return { ...lock, startOffset: lock.startOffset < start ? lock.startOffset : start,
      endOffset: lock.endOffset > end ? lock.endOffset + delta : start + replacementLength }
  })
}

export function documentStoryParagraphs(text: string, style: DocumentDefaultStyle['paragraph']): DocumentStoryParagraph[] {
  const result: DocumentStoryParagraph[] = []
  let startOffset = 0
  for (const part of text.split('\n')) {
    result.push({ startOffset, endOffset: startOffset + part.length, style: { ...style } })
    startOffset += part.length + 1
  }
  return result
}

export function documentStoryContent(section: DocumentSectionHeaderFooter, story: DocumentStoryKey, model?: StoryDefaultsSource): DocumentStoryContent {
  const text = section[story], content = section.storyContent?.[story]
  if (content !== undefined) {
    const prepared = content.blocks === undefined ? content : projectDocumentStoryBlocks(content).content
    validateDocumentStoryContent(prepared, content.blocks === undefined ? text : runText(prepared.runs))
    return prepared
  }
  const defaults = documentStoryDefaultStyle(model)
  // Keep legacy stories centered until the user applies paragraph formatting.
  return { runs: [createRun(text, defaults.text)], paragraphs: documentStoryParagraphs(text, { ...defaults.paragraph, align: 'center', lineHeight: 1.25, spaceBefore: 0, spaceAfter: 0 }) }
}

/** Surviving paragraph separators keep their style; inserted paragraphs inherit the edit point. */
export function editDocumentStoryParagraphs(paragraphs: readonly DocumentStoryParagraph[], before: string, start: number, end: number, insert: string): DocumentStoryParagraph[] {
  const after = before.slice(0, start) + insert + before.slice(end)
  const delta = insert.length - (end - start)
  const separators = paragraphs.slice(0, -1).map(paragraph => paragraph.endOffset)
    .filter(offset => offset < start || offset >= end).map(offset => offset < start ? offset : offset + delta)
  for (let index = 0; index < insert.length; index++) if (insert[index] === '\n') separators.push(start + index)
  separators.sort((a, b) => a - b)
  const result: DocumentStoryParagraph[] = []
  let next = 0
  for (const to of [...new Set(separators), after.length]) {
    const sourceOffset = next < start ? next : next >= start + insert.length ? next - delta : start
    const source = paragraphs.find(paragraph => sourceOffset >= paragraph.startOffset && sourceOffset <= paragraph.endOffset) ?? paragraphs.at(-1)!
    result.push({ ...source, startOffset: next, endOffset: to, style: { ...source.style }, ...(source.metadata ? { metadata: { ...source.metadata } } : {}) })
    next = to + 1
  }
  return result
}
