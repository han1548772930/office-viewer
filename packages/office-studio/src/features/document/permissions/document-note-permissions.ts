import type { DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { documentBlockHostId, type DocumentNoteDecorationHost } from '../model/document-block-host'
import { documentDefaultStyle } from '../model/document-default-style'
import { createRun } from '../model/document-model'
import { documentStoryParagraphs } from '../model/document-story'
import { migrateDocumentStoryBlocks } from '../model/document-story-blocks'
import { documentTableDescendantIds } from '../model/document-table-selection'
import { documentNoteReferenceRecords } from '../notes/document-note-references'
import type { DocumentNoteKind } from '../notes/document-note-types'
import { DEFAULT_DOCUMENT_SECTION_ID, documentSectionIds } from '../services/section-header-footer'
import { permissionTargetsForBlock, permissionTargetsForChangedBlocks, type DocumentPermissionTarget } from './document-permissions'
import { documentModelForBlockHost } from '../services/document-block-host'

/** Separators are document settings but retain paragraph, entity and nested-table permissions. */
export function documentNoteDecorationPermissionTargets(model: DocumentModel, host: DocumentNoteDecorationHost): DocumentPermissionTarget[] {
  const scoped = documentModelForBlockHost(model, host), result = new Map<string, DocumentPermissionTarget>()
  const add = (target: DocumentPermissionTarget) => { result.set(`${target.scope}:${target.id}`, target) }
  add({ scope: 'entity', id: documentBlockHostId(host) })
  for (const id of documentSectionIds(model)) add({ scope: 'section', id })
  for (const id of documentTableDescendantIds(scoped.blocks)) for (const target of permissionTargetsForBlock(scoped, id)) add(target)
  return [...result.values()]
}

/** A note addresses its actual reference, entity, full body and imported nested note resources. */
export function documentNotePermissionTargets(model: DocumentModel, kind: DocumentNoteKind, id: string): DocumentPermissionTarget[] {
  const references = documentNoteReferenceRecords(model), result = new Map<string, DocumentPermissionTarget>()
  const add = (target: DocumentPermissionTarget) => { result.set(`${target.scope}:${target.id}`, target) }
  const pending = [{ kind, id }], visited = new Set<string>(), defaults = documentDefaultStyle(model)
  while (pending.length) {
    const current = pending.pop()!, key = `${current.kind}:${current.id}`
    if (visited.has(key)) continue
    visited.add(key)
    const note = (current.kind === 'footnote' ? model.footnotes : model.endnotes)?.[current.id]
    if (!note) continue
    const host = { kind: 'note' as const, noteKind: current.kind, noteId: current.id }
    const reference = references.find(reference => reference.kind === current.kind && reference.note.id === current.id)
    const sectionId = reference?.sectionId ?? DEFAULT_DOCUMENT_SECTION_ID
    add({ scope: 'entity', id: documentBlockHostId(host) }); add({ scope: 'section', id: sectionId })
    if (reference) for (const target of permissionTargetsForBlock(model, reference.block.id)) add(target)
    const legacy = note.content ?? { runs: [createRun(note.text, defaults.text)], paragraphs: documentStoryParagraphs(note.text, defaults.paragraph) }
    const content = migrateDocumentStoryBlocks(cloneModel(legacy), [], host, defaults)
    const scoped = { ...model, blocks: content.blocks!, metadata: { ...model.metadata, documentEditingBlockHost: host, documentEditingSectionId: sectionId } }
    for (const blockId of documentTableDescendantIds(scoped.blocks)) for (const target of permissionTargetsForBlock(scoped, blockId)) add(target)
    for (const nested of references) if (nested.host?.kind === 'note' && nested.host.noteKind === current.kind && nested.host.noteId === current.id) pending.push({ kind: nested.kind, id: nested.note.id })
  }
  return [...result.values()]
}

/** Generated numbering and settings participate even when only an entity property changes. */
export function documentNoteMutationPermissionTargets(before: DocumentModel, after: DocumentModel): DocumentPermissionTarget[] {
  const result = new Map(permissionTargetsForChangedBlocks(before, after).map(target => [`${target.scope}:${target.id}`, target]))
  const add = (target: DocumentPermissionTarget) => { result.set(`${target.scope}:${target.id}`, target) }
  for (const kind of ['footnote', 'endnote'] as const) {
    const property = kind === 'footnote' ? 'footnotes' : 'endnotes'
    for (const id of new Set([...Object.keys(before[property] ?? {}), ...Object.keys(after[property] ?? {})])) {
      if (JSON.stringify(before[property]?.[id]) === JSON.stringify(after[property]?.[id])) continue
      for (const model of [before, after]) for (const target of documentNotePermissionTargets(model, kind, id)) add(target)
    }
  }
  if (JSON.stringify(before.noteSettings) !== JSON.stringify(after.noteSettings)) {
    for (const id of new Set([...documentSectionIds(before), ...documentSectionIds(after)])) add({ scope: 'section', id })
  }
  for (const id of new Set([...documentSectionIds(before), ...documentSectionIds(after)])) {
    const setup = (model: DocumentModel) => id === DEFAULT_DOCUMENT_SECTION_ID ? model.sectionSetup : model.blocks.find(block => block.id === id)?.sectionBreak?.setup
    const previous = setup(before), next = setup(after)
    if (JSON.stringify(previous?.noteProperties) !== JSON.stringify(next?.noteProperties) || previous?.suppressEndnotes !== next?.suppressEndnotes) add({ scope: 'section', id })
  }
  return [...result.values()]
}
