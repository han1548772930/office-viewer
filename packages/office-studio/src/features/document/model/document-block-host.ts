import type { DocumentBlock, DocumentModel } from '../../../domain/model'
import { documentBlockAncestry } from './document-model'
import type { DocumentStoryKey } from './document-story'

/** A transient editing address, never an alternate persisted document body. */
export type DocumentAuxiliaryBlockHost =
  | { kind: 'story'; sectionId: string; story: DocumentStoryKey }
  | { kind: 'note'; noteKind: 'footnote' | 'endnote'; noteId: string }

export function documentBlockHostId(host: DocumentAuxiliaryBlockHost): string {
  return host.kind === 'story' ? `${host.sectionId}:${host.story}` : `${host.noteKind}:${host.noteId}`
}

export function sameDocumentBlockHost(left?: DocumentAuxiliaryBlockHost, right?: DocumentAuxiliaryBlockHost): boolean {
  if (!left || !right) return left === right
  return left.kind === 'story' ? right.kind === 'story' && left.sectionId === right.sectionId && left.story === right.story
    : right.kind === 'note' && left.noteKind === right.noteKind && left.noteId === right.noteId
}

/** Stored hosts only. Resolved linked stories use their source IDs and never duplicate block ownership. */
export function documentStoredBlockHosts(model: DocumentModel): Array<{ host: DocumentAuxiliaryBlockHost; blocks: DocumentBlock[] }> {
  const result: Array<{ host: DocumentAuxiliaryBlockHost; blocks: DocumentBlock[] }> = []
  for (const [sectionId, section] of Object.entries(model.sectionHeaderFooters ?? {})) {
    for (const [story, content] of Object.entries(section.storyContent ?? {})) {
      if (content?.blocks) result.push({ host: { kind: 'story', sectionId, story: story as DocumentStoryKey }, blocks: content.blocks })
    }
  }
  for (const noteKind of ['footnote', 'endnote'] as const) {
    const notes = noteKind === 'footnote' ? model.footnotes : model.endnotes
    for (const [noteId, note] of Object.entries(notes ?? {})) {
      if (note.content?.blocks) result.push({ host: { kind: 'note', noteKind, noteId }, blocks: note.content.blocks })
    }
  }
  return result
}

export function documentBlockHostForId(model: DocumentModel, blockId: string): DocumentAuxiliaryBlockHost | undefined {
  return documentStoredBlockHosts(model).find(entry => documentBlockAncestry({ blocks: entry.blocks }, blockId))?.host
}

/** Only scope adapters add this marker; committing a scope must not copy it into document metadata. */
export function documentEditingBlockHost(model: DocumentModel): DocumentAuxiliaryBlockHost | undefined {
  const value = model.metadata?.documentEditingBlockHost
  if (!value || typeof value !== 'object' || Array.isArray(value)) return undefined
  if (Reflect.get(value, 'kind') === 'story' && typeof Reflect.get(value, 'sectionId') === 'string' && typeof Reflect.get(value, 'story') === 'string') return value as DocumentAuxiliaryBlockHost
  if (Reflect.get(value, 'kind') === 'note' && ['footnote', 'endnote'].includes(Reflect.get(value, 'noteKind')) && typeof Reflect.get(value, 'noteId') === 'string') return value as DocumentAuxiliaryBlockHost
  return undefined
}
