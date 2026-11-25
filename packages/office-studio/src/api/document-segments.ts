import type { DocumentModel } from '../domain/model'
import { documentBlockHostId, documentStoredBlockHosts, type DocumentAuxiliaryBlockHost } from '../features/document/model/document-block-host'
import { documentNoteDecorationSourceIdentity } from '../features/document/model/document-note-decoration-host'
import { resolveDocumentBlockHost } from '../features/document/services/document-block-host'
import { documentSectionIds } from '../features/document/services/section-header-footer'
import { DOCUMENT_STORY_KEYS } from '../features/document/model/document-story'
import { bodyRecord } from './document-body-style'
import { writePublicDocumentStories } from './document-output-stories'

const STORY_REFERENCES = {
  header: 'defaultHeaderId', footer: 'defaultFooterId', firstPageHeader: 'firstPageHeaderId', firstPageFooter: 'firstPageFooterId',
  evenPageHeader: 'evenPageHeaderId', evenPageFooter: 'evenPageFooterId',
} as const

/** Use the outgoing snapshot's ID assignment, including independently copied source parts. */
function publicStorySegments(model: DocumentModel): Map<string, string> {
  const ids = documentSectionIds(model), sections = ids.map(() => ({} as Record<string, unknown>))
  writePublicDocumentStories(model, bodyRecord(bodyRecord(model.metadata?.univerDocumentSource).snapshot), sections)
  const result = new Map<string, string>()
  for (const [index, sectionId] of ids.entries()) for (const story of DOCUMENT_STORY_KEYS) {
    const id = sections[index]![STORY_REFERENCES[story]]
    if (typeof id === 'string') result.set(documentBlockHostId({ kind: 'story', sectionId, story }), id)
  }
  return result
}

function localStorySegmentId(host: Extract<DocumentAuxiliaryBlockHost, { kind: 'story' }>): string {
  const identity = host.story.toLowerCase().includes('header') ? 'headerId' : 'footerId'
  return `office:${identity}:${host.sectionId}:${host.story}`
}

/** Public header/footer IDs and local note IDs address the same stored hosts. */
export function documentSegmentIdForHost(model: DocumentModel, host: DocumentAuxiliaryBlockHost): string {
  if (host.kind === 'note') return host.noteId
  const view = resolveDocumentBlockHost(model, host)
  if (host.kind === 'noteDecoration') {
    const identity = documentNoteDecorationSourceIdentity(view.content, host.noteKind)
    return identity ? String(identity.noteId) : documentBlockHostId(view.sourceHost)
  }
  return publicStorySegments(model).get(documentBlockHostId(host))
    ?? localStorySegmentId(view.sourceHost.kind === 'story' ? view.sourceHost : host)
}

export function resolveDocumentSegmentHost(model: DocumentModel, segmentId: string): DocumentAuxiliaryBlockHost | undefined {
  if (typeof segmentId !== 'string') throw new TypeError('文档内容区标识必须为字符串。')
  if (!segmentId) return undefined
  const matches = new Map<string, DocumentAuxiliaryBlockHost>(), aliases = new Map<string, DocumentAuxiliaryBlockHost>()
  const stories = publicStorySegments(model)
  for (const { host } of documentStoredBlockHosts(model)) {
    const source = resolveDocumentBlockHost(model, host).sourceHost
    const publicId = host.kind === 'story' ? stories.get(documentBlockHostId(host)) ?? localStorySegmentId(host) : documentSegmentIdForHost(model, host)
    if (publicId === segmentId) matches.set(documentBlockHostId(source), source)
    if (documentBlockHostId(host) === segmentId || host.kind === 'story' && localStorySegmentId(host) === segmentId) aliases.set(documentBlockHostId(source), source)
  }
  if (!matches.size) for (const [id, host] of aliases) matches.set(id, host)
  // Legacy note content can precede its first block migration.
  for (const noteKind of ['footnote', 'endnote'] as const) if ((noteKind === 'footnote' ? model.footnotes : model.endnotes)?.[segmentId]) {
    const host = { kind: 'note' as const, noteKind, noteId: segmentId }
    matches.set(documentBlockHostId(host), host)
  }
  if (matches.size !== 1) throw new RangeError('文档内容区不存在或标识不唯一：' + segmentId)
  return [...matches.values()][0]!
}
