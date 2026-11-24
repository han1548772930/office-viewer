import type { DocumentContentControl, DocumentDefaultStyle, DocumentStoryContent } from '../../../domain/model'
import type { DocumentSelection } from '../model/document-model'
import type { DocumentNoteSettingsValue } from './document-note-settings'

export type DocumentNoteKind = 'footnote' | 'endnote'
export type DocumentNoteDecorationControls = Partial<Record<'separator' | 'continuationSeparator' | 'continuationNotice', DocumentContentControl[]>>

export interface DocumentNoteInsertOptions {
  noteId?: string
  text?: string
  content?: DocumentStoryContent
  contentControls?: DocumentContentControl[]
  customMark?: string
  referenceTextStyle?: DocumentDefaultStyle['text']
  showBodyReference?: boolean
  metadata?: Record<string, unknown>
  settings?: DocumentNoteSettingsValue | null
  decorationContentControls?: DocumentNoteDecorationControls
  sectionId?: string
  suppressEndnotes?: boolean
}

/** null clears an optional explicit value; content/text changes share the ordinary story edit rules. */
export interface DocumentNotePatch {
  text?: string
  content?: DocumentStoryContent
  contentControls?: DocumentContentControl[]
  customMark?: string | null
  referenceTextStyle?: DocumentDefaultStyle['text'] | null
  showBodyReference?: boolean
  metadata?: Record<string, unknown> | null
}

export type DocumentNoteMutationRequest =
  | { type: 'insert'; kind: DocumentNoteKind; selection: DocumentSelection; options?: DocumentNoteInsertOptions }
  | { type: 'delete'; kind: DocumentNoteKind; noteId: string }
  | { type: 'update'; kind: DocumentNoteKind; noteId: string; patch: DocumentNotePatch }
  | { type: 'update-text'; kind: DocumentNoteKind; values: Record<string, string> }
  | { type: 'settings'; kind: DocumentNoteKind; settings: DocumentNoteSettingsValue | null; sectionId?: string; suppressEndnotes?: boolean;
      decorationContentControls?: DocumentNoteDecorationControls }
