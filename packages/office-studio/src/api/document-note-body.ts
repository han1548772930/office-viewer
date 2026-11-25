import type { DocumentDefaultStyle, DocumentEndnoteSettings, DocumentFootnoteSettings,
  DocumentModel, DocumentNote, DocumentStoryContent } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { documentDefaultStyle } from '../features/document/model/document-default-style'
import { createParagraph, createRun } from '../features/document/model/document-model'
import { assertDocumentNoteSerializable } from '../features/document/notes/document-note-content'
import { DOCUMENT_NOTE_DECORATION_KEYS, type DocumentNoteSettingsValue } from '../features/document/notes/document-note-settings'
import type { DocumentNoteDecorationControls, DocumentNoteInsertOptions } from '../features/document/notes/document-note-types'
export type { DocumentNoteDecorationControls } from '../features/document/notes/document-note-types'
import { decodeDocumentBody, documentBodyContentControls, encodeDocumentBody } from './document-body-codec'
import { bodyRecord, documentRunStyleFromBody, documentRunStyleToBody, type UniverBodyRecord } from './document-body-style'
import { documentBodyOfficeSource, documentBodySourceMetadata } from './document-body-source'
import type { DocumentBodyDecodeOptions, UniverDocumentBody, UniverDocumentBodyData, UniverDocumentBodyResources } from './document-body-types'

/** Fixed core IDocumentNote input; domain-shaped input remains available for local clients. */
export interface PublicDocumentNoteInput extends UniverDocumentBodyResources {
  body?: UniverDocumentBody
  customMark?: string
  referenceTextStyle?: UniverBodyRecord
  showBodyReference?: boolean
  officeStudio?: UniverBodyRecord
}
export type DocumentNoteContentInput = PublicDocumentNoteInput | DocumentNoteInsertOptions
export interface PublicDocumentNoteDecoration extends UniverDocumentBody, UniverDocumentBodyResources {}
export type PublicDocumentFootnoteSettings = Omit<DocumentFootnoteSettings, typeof DOCUMENT_NOTE_DECORATION_KEYS[number]>
  & Partial<Record<typeof DOCUMENT_NOTE_DECORATION_KEYS[number], PublicDocumentNoteDecoration>>
export type PublicDocumentEndnoteSettings = Omit<DocumentEndnoteSettings, typeof DOCUMENT_NOTE_DECORATION_KEYS[number]>
  & Partial<Record<typeof DOCUMENT_NOTE_DECORATION_KEYS[number], PublicDocumentNoteDecoration>>
export type DocumentNoteSettingsInput = DocumentNoteSettingsValue | PublicDocumentFootnoteSettings | PublicDocumentEndnoteSettings
export interface DocumentNoteInfo extends Omit<DocumentNote, 'referenceTextStyle'>, UniverDocumentBodyData {
  noteId: string
  type: 'footnote' | 'endnote'
  referenceTextStyle?: DocumentDefaultStyle['text'] & UniverBodyRecord
}

function resourcesFrom(value: UniverBodyRecord): UniverDocumentBodyResources {
  const resources: UniverDocumentBodyResources = {}
  for (const key of ['tableSource', 'drawings', 'drawingsOrder', 'lists'] as const) {
    if (value[key] !== undefined) Object.assign(resources, { [key]: value[key] })
  }
  return resources
}

/** Reference formatting is independent of note-body formatting and retains unknown public keys. */
function referenceStyleFrom(value: unknown, defaults: DocumentDefaultStyle): DocumentDefaultStyle['text'] {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError('注释引用样式必须为对象。')
  const raw = bodyRecord(value), result = { ...defaults.text }
  for (const key of ['fontFamily', 'eastAsiaFontFamily', 'fontSize', 'bold', 'italic', 'underline', 'strike', 'color', 'highlight', 'verticalAlign'] as const) {
    if (raw[key] !== undefined) Object.assign(result, { [key]: raw[key] })
  }
  const patch = documentRunStyleFromBody(raw)
  for (const key of ['fontFamily', 'eastAsiaFontFamily', 'fontSize', 'bold', 'italic', 'underline', 'strike', 'color', 'highlight', 'verticalAlign'] as const) {
    if (patch[key] !== undefined) Object.assign(result, { [key]: patch[key] })
  }
  return result
}

export function documentNoteInputFromPublic(model: DocumentModel, input?: DocumentNoteContentInput, note?: DocumentBodyDecodeOptions['note'], bodyOptions: Partial<DocumentBodyDecodeOptions> = {}): DocumentNoteInsertOptions | undefined {
  if (input === undefined) return undefined
  assertDocumentNoteSerializable(input)
  if (!input || typeof input !== 'object' || Array.isArray(input)) throw new TypeError('注释输入必须为对象。')
  const raw = bodyRecord(input), defaults = documentDefaultStyle(model)
  const next = cloneModel(raw)
  const provenance = documentBodyOfficeSource(raw, '注释')
  if (Object.keys(provenance).length) next.metadata = { ...documentBodySourceMetadata(provenance.metadata, '注释', ['univerDocumentSource', 'univerBodySource']), ...bodyRecord(raw.metadata) }
  delete next.officeStudio
  if (raw.body !== undefined) {
    if (raw.content !== undefined || raw.text !== undefined) throw new TypeError('公开注释 body 与本地文字/块输入不能同时提供。')
    const decoded = decodeDocumentBody(raw.body, resourcesFrom(raw), { ...bodyOptions, defaults, note })
    next.content = decoded.content
    if (decoded.contentControls.length) next.contentControls = decoded.contentControls
    delete next.body
  } else if (['tableSource', 'drawings', 'drawingsOrder', 'lists'].some(key => raw[key] !== undefined)) {
    throw new TypeError('注释资源必须随对应 body 一起提供。')
  }
  for (const key of ['tableSource', 'drawings', 'drawingsOrder', 'lists']) delete next[key]
  if (raw.referenceTextStyle !== undefined) {
    next.referenceTextStyle = referenceStyleFrom(raw.referenceTextStyle, defaults)
    next.metadata = { ...bodyRecord(next.metadata), univerReferenceTextStyle: cloneModel(raw.referenceTextStyle) }
  }
  if (raw.settings !== undefined) {
    const decoded = documentNoteSettingsFromPublic(model, raw.settings as DocumentNoteSettingsInput | null)
    next.settings = decoded.settings
    if (decoded.decorationContentControls) next.decorationContentControls = decoded.decorationContentControls
  }
  // The domain command performs the full property, content, mode and permission validation.
  return next as DocumentNoteInsertOptions
}

export function documentNoteToPublic(model: DocumentModel, note: DocumentNote, type: 'footnote' | 'endnote'): DocumentNoteInfo {
  const defaults = documentDefaultStyle(model), detached = cloneModel(note)
  const content = detached.content ?? { blocks: [createParagraph(note.text, defaults.text)], runs: [], paragraphs: [] }
  const publicData = encodeDocumentBody(content, defaults, documentBodyContentControls(content, model.contentControls), { note: { kind: type, id: note.id } })
  const style = detached.referenceTextStyle === undefined ? undefined : {
    ...cloneModel(bodyRecord(detached.metadata?.univerReferenceTextStyle)), ...detached.referenceTextStyle,
    ...documentRunStyleToBody(createRun('', detached.referenceTextStyle)),
  }
  return { ...detached, ...publicData, noteId: note.id, type, ...(style === undefined ? {} : { referenceTextStyle: style }) }
}

/** Decorations use a direct IDocumentBody with tableSource, rather than a nested body wrapper. */
export function documentNoteSettingsFromPublic(model: DocumentModel, input: DocumentNoteSettingsInput | null, bodyOptions: Partial<DocumentBodyDecodeOptions> = {}): {
  settings: DocumentNoteSettingsValue | null
  decorationContentControls?: DocumentNoteDecorationControls
} {
  if (input === null) return { settings: null }
  assertDocumentNoteSerializable(input)
  if (!input || typeof input !== 'object' || Array.isArray(input)) throw new TypeError('注释设置必须为对象。')
  const settings = cloneModel(input) as DocumentNoteSettingsValue, controls: DocumentNoteDecorationControls = {}
  for (const key of DOCUMENT_NOTE_DECORATION_KEYS) {
    const value = input[key]
    if (value === undefined) continue
    if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError('注释分隔内容必须为对象。')
    const raw = bodyRecord(value)
    if (Object.hasOwn(raw, 'dataStream')) {
      const decoded = decodeDocumentBody(raw, resourcesFrom(raw), { ...bodyOptions, defaults: documentDefaultStyle(model) })
      settings[key] = decoded.content
      if (decoded.contentControls.length) controls[key] = decoded.contentControls
    }
  }
  return { settings, ...(Object.keys(controls).length ? { decorationContentControls: controls } : {}) }
}

export function documentNoteSettingsToPublic(model: DocumentModel, settings: DocumentNoteSettingsValue): PublicDocumentFootnoteSettings | PublicDocumentEndnoteSettings {
  const result = cloneModel(settings) as PublicDocumentFootnoteSettings | PublicDocumentEndnoteSettings
  for (const key of DOCUMENT_NOTE_DECORATION_KEYS) {
    const content: DocumentStoryContent | undefined = settings[key]
    if (content === undefined) continue
    const { body, ...resources } = encodeDocumentBody(content, documentDefaultStyle(model), documentBodyContentControls(content, model.contentControls))
    result[key] = { ...body, ...resources }
  }
  return result
}
