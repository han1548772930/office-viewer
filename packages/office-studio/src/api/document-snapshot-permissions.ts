import type { DocumentModel } from '../domain/model'
import { documentBlockHostId, documentStoredBlockHosts } from '../features/document/model/document-block-host'
import { documentTableDescendantIds } from '../features/document/model/document-table-selection'
import { DOCUMENT_STORY_KEYS } from '../features/document/model/document-story'
import { documentModelForBlockHost } from '../features/document/services/document-block-host'
import { documentSectionIds, resolveSectionHeaderFooter } from '../features/document/services/section-header-footer'
import { assertDocumentPermission, permissionTargetsForBlock, type DocumentPermissionTarget } from '../features/document/permissions/document-permissions'
import { documentNotePermissionTargets } from '../features/document/permissions/document-note-permissions'

/** A public snapshot exposes every body and auxiliary segment, including hidden
 * first/even stories and dormant special notes. Read permission applies to the
 * same real block/entity/section targets used by editing commands. */
export function assertDocumentSnapshotViewPermission(model: DocumentModel, userId: string): void {
  const targets = new Map<string, DocumentPermissionTarget>()
  const add = (target: DocumentPermissionTarget) => { targets.set(`${target.scope}:${target.id}`, target) }
  for (const id of documentSectionIds(model)) add({ scope: 'section', id })
  for (const id of documentTableDescendantIds(model.blocks)) for (const target of permissionTargetsForBlock(model, id)) add(target)
  const hosts = documentStoredBlockHosts(model)
  // Resolved legacy and linked stories may not have their own stored block tree.
  for (const sectionId of documentSectionIds(model)) {
    const section = resolveSectionHeaderFooter(model, sectionId)
    for (const story of DOCUMENT_STORY_KEYS) {
      if (section.storyContent?.[story] === undefined && !section[story] && !section.pageFields?.some(field => field.story === story)) continue
      const host = { kind: 'story' as const, sectionId, story }, scoped = documentModelForBlockHost(model, host)
      add({ scope: 'entity', id: documentBlockHostId(host) })
      add({ scope: 'entity', id: `${sectionId}:${story.toLowerCase().includes('header') ? 'header' : 'footer'}` })
      for (const id of documentTableDescendantIds(scoped.blocks)) for (const target of permissionTargetsForBlock(scoped, id)) add(target)
    }
  }
  for (const kind of ['footnote', 'endnote'] as const) {
    for (const id of Object.keys((kind === 'footnote' ? model.footnotes : model.endnotes) ?? {})) {
      for (const target of documentNotePermissionTargets(model, kind, id)) add(target)
    }
  }
  for (const { host } of hosts) if (host.kind === 'noteDecoration') {
    const scoped = documentModelForBlockHost(model, host)
    add({ scope: 'entity', id: documentBlockHostId(host) })
    for (const id of documentTableDescendantIds(scoped.blocks)) for (const target of permissionTargetsForBlock(scoped, id)) add(target)
  }
  assertDocumentPermission(model.permissionRules ?? [], { documentId: model.id!, userId }, 'view', [...targets.values()])
}
