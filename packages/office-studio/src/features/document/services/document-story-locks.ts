import type { DocumentSectionHeaderFooter, DocumentStoryContent } from '../../../domain/model'
import { documentStoryContent, documentStoryTextEdit, editDocumentStoryLocks, validateDocumentStoryLocks, type DocumentStoryKey, type StoryDefaultsSource } from '../model/document-story'
import { storyBuffer } from './story-editor'

function stableAppearance(value: unknown): string {
  return JSON.stringify(value, (_key, entry: unknown) => {
    if (!entry || typeof entry !== 'object' || Array.isArray(entry)) return entry
    return Object.fromEntries(Object.entries(entry).sort(([first], [second]) => first < second ? -1 : first > second ? 1 : 0))
  })
}

function sourceFormat(metadata: Record<string, unknown> | undefined): unknown {
  const source = metadata?.source
  return source && typeof source === 'object' && !Array.isArray(source) ? Reflect.get(source, 'sourceFormat') : undefined
}

function lockedAppearance(section: DocumentSectionHeaderFooter, story: DocumentStoryKey, content: DocumentStoryContent, start: number, end: number): string {
  const buffer = storyBuffer(section, story), from = buffer.literalOffsets[start]!, to = buffer.literalOffsets[end]!
  const runs: Array<{ text: string; style: string }> = []
  let offset = 0
  for (const run of content.runs) {
    const a = Math.max(from, offset), b = Math.min(to, offset + run.text.length)
    if (a < b) {
      const { id: _id, text: _text, metadata: rawMetadata, ...properties } = run
      const metadata = { ...rawMetadata }
      delete metadata.source
      const style = stableAppearance({ ...properties, sourceFormat: sourceFormat(rawMetadata), ...(Object.keys(metadata).length ? { metadata } : {}) })
      const text = run.text.slice(a - offset, b - offset), previous = runs.at(-1)
      if (previous?.style === style) previous.text += text
      else runs.push({ text, style })
    }
    offset += run.text.length
  }
  const fields = [...buffer.objects].filter(([position]) => position >= start && position < end).map(([, field]) => {
    const { story: _story, offset: _offset, ...value } = field
    return value
  })
  const paragraphs = content.paragraphs.filter(paragraph => (to > from ? paragraph.startOffset < to : paragraph.startOffset <= to) && paragraph.endOffset >= from)
    .map(paragraph => ({ style: paragraph.style, sourceFormat: sourceFormat(paragraph.metadata),
      cleared: paragraph.metadata?.officeStoryParagraphFormatCleared, baseline: paragraph.metadata?.officeStoryParagraphStyleBaseline }))
  return stableAppearance({ text: buffer.value.slice(start, end), runs, fields, paragraphs })
}

/** Apply the same content-lock rules to direct section API patches and editor commands. */
export function assertDocumentStoryLocksPreserved(current: DocumentSectionHeaderFooter, next: DocumentSectionHeaderFooter, story: DocumentStoryKey, model?: StoryDefaultsSource): void {
  const content = current.storyContent?.[story]
  if (!content?.locks?.length) return
  const before = storyBuffer(current, story), after = storyBuffer(next, story)
  const edit = documentStoryTextEdit(before.value, after.value)
  const expected = editDocumentStoryLocks(content.locks, before.value, edit.start, edit.end, edit.text.length)!
  const prepared = documentStoryContent(next, story, model)
  validateDocumentStoryLocks(prepared.locks, after.value)
  for (const [index, lock] of content.locks.entries()) {
    const planned = expected[index]!, actual = prepared.locks?.find(candidate => candidate.id === lock.id)
    if (!actual || actual.lock !== lock.lock || actual.startOffset !== planned.startOffset || actual.endOffset !== planned.endOffset) throw new Error(`页眉页脚补丁不能移除或绕过内容锁：${lock.id}`)
    if (lock.lock !== 'sdtLocked' && lockedAppearance(current, story, content, lock.startOffset, lock.endOffset)
      !== lockedAppearance(next, story, prepared, actual.startOffset, actual.endOffset)) throw new Error(`页眉页脚内容已锁定：${lock.id}`)
  }
}

export function assertDocumentStoryFormattingEditable(content: DocumentStoryContent, buffer: string, start: number, end: number): void {
  validateDocumentStoryLocks(content.locks, buffer)
  // The next paragraph's empty wrapper belongs to that paragraph. The final
  // story endpoint remains part of its paragraph's protected formatting mark.
  const locked = content.locks?.find(lock => lock.lock !== 'sdtLocked'
    && (start < lock.endOffset && lock.startOffset < end || lock.startOffset === lock.endOffset && lock.startOffset >= start
      && (lock.startOffset < end || end === buffer.length && lock.startOffset === end)))
  if (locked) throw new Error(`页眉页脚内容已锁定：${locked.id}`)
}
