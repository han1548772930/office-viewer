import type { DocumentModel, DocumentStoryContent } from '../../../domain/model'
import type { DocumentNoteDecorationHost } from './document-block-host'

export const DOCUMENT_NOTE_DECORATION_HOST_KEYS = ['separator', 'continuationSeparator', 'continuationNotice'] as const

export interface DocumentNoteDecorationSourceIdentity {
  format: 'docx' | 'univer'
  kind: 'footnote' | 'endnote'
  type: typeof DOCUMENT_NOTE_DECORATION_HOST_KEYS[number]
  noteId: string | number
}

/** Public special notes have their own source identity; they never impersonate native DOCX XML. */
export function documentNoteDecorationSourceIdentity(content: DocumentStoryContent, kind: 'footnote' | 'endnote',
  type?: typeof DOCUMENT_NOTE_DECORATION_HOST_KEYS[number]): DocumentNoteDecorationSourceIdentity | undefined {
  const read = (value: unknown, format: 'docx' | 'univer'): DocumentNoteDecorationSourceIdentity | undefined => {
    if (value === undefined) return undefined
    if (!value || typeof value !== 'object' || Array.isArray(value) || Reflect.get(value, 'version') !== 1
      || Reflect.get(value, 'kind') !== kind || !DOCUMENT_NOTE_DECORATION_HOST_KEYS.includes(Reflect.get(value, 'type'))
      || type !== undefined && Reflect.get(value, 'type') !== type) throw new TypeError('注释分隔内容的来源类型或版本无效。')
    const noteId = Reflect.get(value, 'noteId')
    if (format === 'docx') {
      if (typeof noteId !== 'number' || !Number.isSafeInteger(noteId) || typeof Reflect.get(value, 'partName') !== 'string'
        || !Reflect.get(value, 'partName') || typeof Reflect.get(value, 'sourceXml') !== 'string') {
        throw new TypeError('原生注释分隔内容需要真实的数字 ID、部件名与来源 XML。')
      }
    } else if (typeof noteId !== 'string' || !noteId.trim() || noteId.length > 4096 || /[\r\n\u0000]/.test(noteId)) {
      throw new TypeError('公开注释分隔内容需要非空的字符串来源 ID。')
    }
    return { format, kind, type: Reflect.get(value, 'type'), noteId }
  }
  // Validate both when a converted file retains both provenance records; an invalid native record cannot fall back to a public one.
  const native = read(content.metadata?.decorationSource, 'docx'), publicSource = read(content.metadata?.univerNoteDecorationSource, 'univer')
  if (native && publicSource && native.type !== publicSource.type) throw new TypeError('公开与原生注释分隔来源的类型不一致。')
  return publicSource ?? native
}

/** Dormant special stories retain their source identity and are never ordinary notes. */
export function documentDormantNoteDecorations(model: Pick<DocumentModel, 'noteSettings'>, kind: 'footnote' | 'endnote'): DocumentStoryContent[] {
  const value = model.noteSettings?.[kind]?.metadata?.officeDormantNoteDecorations
  if (value === undefined) return []
  if (!value || typeof value !== 'object' || Array.isArray(value) || Reflect.get(value, 'version') !== 1 || !Array.isArray(Reflect.get(value, 'contents'))) {
    throw new TypeError('休眠注释分隔内容的版本或结构无效。')
  }
  const contents = Reflect.get(value, 'contents') as DocumentStoryContent[]
  if (contents.length > 100_000) throw new RangeError('休眠注释分隔内容超过单次故事预算。')
  return contents
}

export function documentDormantNoteDecorationHost(model: Pick<DocumentModel, 'noteSettings'>, kind: 'footnote' | 'endnote', index: number): DocumentNoteDecorationHost {
  if (!Number.isSafeInteger(index) || index < 0) throw new RangeError('休眠注释分隔内容的索引无效。')
  const content = documentDormantNoteDecorations(model, kind)[index], source = content && documentNoteDecorationSourceIdentity(content, kind)
  if (!source) throw new TypeError('休眠注释分隔内容缺少匹配的公开或原生来源身份。')
  return { kind: 'noteDecoration', noteKind: kind, decoration: source.type, dormantIndex: index }
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
