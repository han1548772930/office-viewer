import type { DocumentModel, DocumentStoryContent } from '../../../domain/model'
import { documentDefaultStyle } from '../model/document-default-style'
import { projectDocumentStoryBlocks } from '../model/document-story-blocks'
import type { DocumentNoteKind } from './document-note-types'
import type { DOCUMENT_NOTE_DECORATION_KEYS } from './document-note-settings'

export type DocumentNoteDecorationKind = typeof DOCUMENT_NOTE_DECORATION_KEYS[number]

/** The default line belongs to a real editable paragraph; its label/geometry is not persisted body text. */
export function defaultDocumentNoteDecoration(model: DocumentModel, kind: DocumentNoteKind, decoration: DocumentNoteDecorationKind,
  forEditing = false): DocumentStoryContent | undefined {
  if (decoration === 'continuationNotice' && !forEditing) return undefined
  const defaults = documentDefaultStyle(model), id = `default-note-decoration:${kind}:${decoration}:paragraph`
  const paragraph: DocumentModel['blocks'][number] = { id, type: 'paragraph', level: 0,
    runs: [{ id: `${id}:run`, text: '', ...defaults.text, fontSize: 10 }],
    style: { ...defaults.paragraph, spaceBefore: 0, lineHeight: 1, spacingRule: 'auto', spaceAfter: 0 },
    metadata: decoration === 'continuationNotice' ? {}
      : { noteSeparator: { version: 1, fieldType: decoration === 'separator' ? 'FOOTNOTE_SEPARATOR' : 'FOOTNOTE_CONTINUATION_SEPARATOR',
        widthFraction: decoration === 'separator' ? 1 / 3 : 1, color: '#1f2937' } } }
  return projectDocumentStoryBlocks({ blocks: [paragraph], runs: [], paragraphs: [] }, 'display').content
}
