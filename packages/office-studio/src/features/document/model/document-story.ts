import type { DocumentDefaultStyle, DocumentModel, DocumentPageField, DocumentSectionHeaderFooter, DocumentStoryContent, DocumentStoryLock, DocumentStoryParagraph } from '../../../domain/model'
import { documentDefaultStyle } from './document-default-style'
import { createRun, runText } from './document-model'
import type { DocumentRun, RunPatch } from './document-model'

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
  if (!cache?.runs.length || text !== field.cachedResult) return [createRun(text, field.textStyle ?? defaults)]
  const cleared = field.metadata?.officeStoryTextFormatCleared === true
  const style = field.textStyle ?? (cleared ? defaults : undefined)
  if (!style) return cache.runs.map(run => ({ ...run }))
  const baseline = field.metadata?.officeStoryStyleBaseline
  return cache.runs.map(run => {
    const next = { ...run, metadata: { ...run.metadata } }
    for (const key of ['bold', 'italic', 'underline', 'strike', 'color', 'fontFamily', 'eastAsiaFontFamily', 'fontSize', 'highlight', 'verticalAlign'] as const) {
      const current = style[key]
      if (!cleared && baseline && typeof baseline === 'object' && !Array.isArray(baseline) && current === Reflect.get(baseline, key)) continue
      if (current === undefined) {
        if (key === 'eastAsiaFontFamily' || key === 'highlight' || key === 'verticalAlign') delete next[key]
      }
      else Object.assign(next, { [key]: current })
      if (key === 'highlight') delete next.metadata.highlight
    }
    return next
  })
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

export function validateDocumentStoryContent(content: DocumentStoryContent, text: string): void {
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
  if (content !== undefined) { validateDocumentStoryContent(content, text); return content }
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
