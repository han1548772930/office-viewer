import type { DocumentModel } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { documentDefaultStyle } from '../features/document/model/document-default-style'
import { documentDormantNoteDecorations, documentNoteDecorationSourceIdentity } from '../features/document/model/document-note-decoration-host'
import { DOCUMENT_NOTE_DECORATION_KEYS } from '../features/document/notes/document-note-settings'
import { documentBodyContentControls, encodeDocumentBody } from './document-body-codec'
import { bodyRecord, type UniverBodyRecord } from './document-body-style'
import { documentBodySourceMetadata } from './document-body-source'
import { documentNoteSettingsToPublic, documentNoteToPublic, type PublicDocumentEndnoteSettings, type PublicDocumentFootnoteSettings } from './document-note-body'
import type { UniverDocumentNoteData, UniverDocumentNoteSource } from './document-snapshot-types'

export function writePublicDocumentNotes(model: DocumentModel, source: UniverBodyRecord): {
  notes: Record<string, UniverDocumentNoteData>
  footnotes: Record<string, UniverDocumentNoteSource>
  endnotes: Record<string, UniverDocumentNoteSource>
  noteSettings: { footnote?: PublicDocumentFootnoteSettings; endnote?: PublicDocumentEndnoteSettings }
} {
  const notes: Record<string, UniverDocumentNoteData> = {}, footnotes: Record<string, UniverDocumentNoteSource> = {}, endnotes: Record<string, UniverDocumentNoteSource> = {}
  const noteSettings: { footnote?: PublicDocumentFootnoteSettings; endnote?: PublicDocumentEndnoteSettings } = {}
  for (const kind of ['footnote', 'endnote'] as const) {
    const property = kind === 'footnote' ? 'footnotes' : 'endnotes', target = kind === 'footnote' ? footnotes : endnotes
    const ordinary = kind === 'footnote' ? model.footnotes : model.endnotes
    for (const [id, note] of Object.entries(ordinary ?? {})) {
      if (notes[id]) throw new RangeError('公开 notes 不能包含同名脚注与尾注：' + id)
      const data = documentNoteToPublic(model, note, kind), original = bodyRecord(bodyRecord(source.notes)[id])
      const value: UniverDocumentNoteData = { ...cloneModel(original), noteId: id, type: kind, body: data.body,
        tableSource: data.tableSource, drawings: data.drawings, drawingsOrder: data.drawingsOrder, lists: data.lists, callouts: data.callouts, codes: data.codes,
        officeStudio: { version: 1, metadata: documentBodySourceMetadata(note.metadata, '注释', ['univerBodySource', 'univerDocumentSource']) },
        ...(note.customMark === undefined ? {} : { customMark: note.customMark }),
        ...(data.referenceTextStyle === undefined ? {} : { referenceTextStyle: data.referenceTextStyle }),
        ...(note.showBodyReference === undefined ? {} : { showBodyReference: note.showBodyReference }) }
      for (const key of ['customMark', 'referenceTextStyle', 'showBodyReference'] as const) if (note[key] === undefined) delete value[key]
      for (const key of ['text', 'content', 'referenceBlockId', 'referenceRunId']) Reflect.deleteProperty(value, key)
      notes[id] = value
      target[id] = { ...cloneModel(bodyRecord(bodyRecord(source[property])[id])), ...value, type: 'normal' }
    }
    const settings = model.noteSettings?.[kind]
    if (!settings) continue
    const publicSettings = documentNoteSettingsToPublic(model, settings)
    if (publicSettings.metadata?.officeDormantNoteDecorations !== undefined) delete publicSettings.metadata.officeDormantNoteDecorations
    // Kind was validated when entering the domain; the output keeps its
    // distinct public position/restart type and independently editable bodies.
    if (kind === 'footnote') noteSettings.footnote = publicSettings as PublicDocumentFootnoteSettings
    else noteSettings.endnote = publicSettings as PublicDocumentEndnoteSettings
    const contents = [...DOCUMENT_NOTE_DECORATION_KEYS.flatMap(key => settings[key] ? [settings[key]!] : []), ...documentDormantNoteDecorations(model, kind)]
    for (const content of contents) {
      const identity = documentNoteDecorationSourceIdentity(content, kind)
      if (!identity) continue
      const id = String(identity.noteId)
      if (target[id]) throw new RangeError('普通注释与保留的分隔资源身份冲突：' + id)
      const data = encodeDocumentBody(content, documentDefaultStyle(model), documentBodyContentControls(content, model.contentControls))
      target[id] = { ...cloneModel(bodyRecord(bodyRecord(source[property])[id])), noteId: id, type: identity.type, ...data,
        tableSource: data.tableSource, drawings: data.drawings, drawingsOrder: data.drawingsOrder, lists: data.lists, callouts: data.callouts, codes: data.codes }
    }
  }
  return { notes, footnotes, endnotes, noteSettings }
}
