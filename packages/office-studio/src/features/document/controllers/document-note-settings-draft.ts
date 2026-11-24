import { cloneModel } from '../../../domain/serialization'
import { DOCUMENT_NOTE_DECORATION_KEYS, validateDocumentNoteSettings, type DocumentNoteSettingsValue } from '../notes/document-note-settings'
import type { DocumentNoteKind } from '../notes/document-note-types'

export type DocumentNoteDecorationKey = typeof DOCUMENT_NOTE_DECORATION_KEYS[number]
export interface DocumentNoteSettingsForm {
  position?: string; restart?: string; numberFormat?: string; startNumber?: number | null; columnCount?: number | null
}

/** Empty fields remove explicit overrides; document-level decorations and source metadata are retained. */
export function documentNoteSettingsFromForm(base: DocumentNoteSettingsValue | undefined, form: DocumentNoteSettingsForm,
  kind: DocumentNoteKind, sectionId?: string): DocumentNoteSettingsValue {
  const settings: DocumentNoteSettingsValue = cloneModel(base ?? {})
  for (const key of ['position', 'restart', 'numberFormat', 'startNumber'] as const) {
    const value = form[key]
    if (value === undefined || value === null || value === '') Reflect.deleteProperty(settings, key)
    else Reflect.set(settings, key, value)
  }
  if (kind === 'footnote') {
    if (form.columnCount == null) Reflect.deleteProperty(settings, 'columnCount')
    else Reflect.set(settings, 'columnCount', form.columnCount)
  }
  validateDocumentNoteSettings(settings, kind, sectionId)
  return settings
}

