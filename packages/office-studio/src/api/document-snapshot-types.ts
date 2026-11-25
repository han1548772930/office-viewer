import type { DocumentCustomXmlPart, DocumentPermissionRule, DocumentThreadComment } from '../domain/model'
import type { UniverBodyRecord } from './document-body-style'
import type { UniverDocumentBodyData } from './document-body-types'
import type { PublicDocumentEndnoteSettings, PublicDocumentFootnoteSettings, PublicDocumentNoteInput } from './document-note-body'

export interface UniverDocumentNoteData extends PublicDocumentNoteInput {
  noteId: string
  type: 'footnote' | 'endnote'
  body: UniverDocumentBodyData['body']
}
export interface UniverDocumentNoteSource extends UniverDocumentBodyData, UniverBodyRecord {
  noteId: string
  type: 'normal' | 'separator' | 'continuationSeparator' | 'continuationNotice'
}
/** Project provenance and services survive a public snapshot without becoming
 * a second text model. Body/section/story data remain the editable authority. */
export interface OfficeDocumentSnapshotSource {
  version: 1
  metadata?: UniverBodyRecord
  permissionRules?: DocumentPermissionRule[]
  comments?: Record<string, DocumentThreadComment>
}
export interface UniverDocumentData extends UniverDocumentBodyData, UniverBodyRecord {
  id: string
  rev?: number
  title?: string
  documentStyle: UniverBodyRecord
  styles?: Record<string, UniverBodyRecord>
  headers?: Record<string, UniverDocumentBodyData & UniverBodyRecord>
  footers?: Record<string, UniverDocumentBodyData & UniverBodyRecord>
  notes?: Record<string, UniverDocumentNoteData>
  footnotes?: Record<string, UniverDocumentNoteSource>
  endnotes?: Record<string, UniverDocumentNoteSource>
  noteSettings?: { footnote?: PublicDocumentFootnoteSettings; endnote?: PublicDocumentEndnoteSettings }
  customXmlParts?: DocumentCustomXmlPart[]
  officeStudio?: OfficeDocumentSnapshotSource
}
