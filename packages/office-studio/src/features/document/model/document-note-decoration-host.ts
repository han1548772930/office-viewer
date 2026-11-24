import type { DocumentModel, DocumentStoryContent } from '../../../domain/model'
import type { DocumentNoteDecorationHost } from './document-block-host'

export const DOCUMENT_NOTE_DECORATION_HOST_KEYS = ['separator', 'continuationSeparator', 'continuationNotice'] as const

/** Dormant special stories retain their native identity and are never ordinary notes. */
export function documentDormantNoteDecorations(model: DocumentModel, kind: 'footnote' | 'endnote'): DocumentStoryContent[] {
  const value = model.noteSettings?.[kind]?.metadata?.officeDormantNoteDecorations
  if (value === undefined) return []
  if (!value || typeof value !== 'object' || Array.isArray(value) || Reflect.get(value, 'version') !== 1 || !Array.isArray(Reflect.get(value, 'contents'))) {
    throw new TypeError('休眠注释分隔内容的版本或结构无效。')
  }
  return Reflect.get(value, 'contents') as DocumentStoryContent[]
}

export function documentDormantNoteDecorationHost(model: DocumentModel, kind: 'footnote' | 'endnote', index: number): DocumentNoteDecorationHost {
  if (!Number.isSafeInteger(index) || index < 0) throw new RangeError('休眠注释分隔内容的索引无效。')
  const content = documentDormantNoteDecorations(model, kind)[index], source = content?.metadata?.decorationSource
  if (!source || typeof source !== 'object' || Array.isArray(source) || Reflect.get(source, 'version') !== 1
    || Reflect.get(source, 'kind') !== kind || !DOCUMENT_NOTE_DECORATION_HOST_KEYS.includes(Reflect.get(source, 'type'))) {
    throw new TypeError('休眠注释分隔内容缺少匹配的原生来源身份。')
  }
  return { kind: 'noteDecoration', noteKind: kind, decoration: Reflect.get(source, 'type'), dormantIndex: index }
}

/** Borrow one stored story. Default display separators are created by the layout/draft adapter. */
export function documentNoteDecorationContent(model: DocumentModel, host: DocumentNoteDecorationHost): DocumentStoryContent | undefined {
  if (host.dormantIndex === undefined) return model.noteSettings?.[host.noteKind]?.[host.decoration]
  const actual = documentDormantNoteDecorationHost(model, host.noteKind, host.dormantIndex)
  if (actual.decoration !== host.decoration) throw new RangeError('休眠注释分隔内容的类型与地址不匹配。')
  return documentDormantNoteDecorations(model, host.noteKind)[host.dormantIndex]
}

/** Publish only one story into a caller-owned candidate, preserving all source metadata. */
export function setDocumentNoteDecorationContent(model: DocumentModel, host: DocumentNoteDecorationHost, content: DocumentStoryContent): void {
  const settings = { ...model.noteSettings?.[host.noteKind] }
  if (host.dormantIndex === undefined) settings[host.decoration] = content
  else {
    documentNoteDecorationContent(model, host)
    const contents = [...documentDormantNoteDecorations(model, host.noteKind)]
    contents[host.dormantIndex] = content
    settings.metadata = { ...settings.metadata, officeDormantNoteDecorations: { version: 1, contents } }
  }
  model.noteSettings = { ...model.noteSettings, [host.noteKind]: settings }
}
