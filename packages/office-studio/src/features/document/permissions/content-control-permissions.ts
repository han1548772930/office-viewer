import type { DocumentContentControl, DocumentModel } from '../../../domain/model'
import { documentContentControlOwnedBlockIds, resolveDocumentContentControlBlocks } from '../model/content-control-block-range'
import { documentContentControlTableRowBlockIds, resolveDocumentContentControlTableRows } from '../model/content-control-table-rows'
import { documentTableDescendantIds } from '../model/document-table-selection'
import { documentDefaultStyle } from '../model/document-default-style'
import { documentNoteReferenceRunMatches } from '../model/document-note-reference'
import { createRun } from '../model/document-model'
import { migrateDocumentStoryBlocks } from '../model/document-story-blocks'
import { documentStoryParagraphs } from '../model/document-story'
import { documentTextFieldParagraphs } from '../model/document-text-fields'
import { permissionTargetsForBlock, type DocumentPermissionTarget } from './document-permissions'

/** Resolve one editing host before calling. A fallback paragraph never narrows an owned range. */
export function documentContentControlPermissionTargets(model: DocumentModel, control: DocumentContentControl,
  resources: DocumentModel = model): DocumentPermissionTarget[] {
  const owned = control.blockRange ? documentContentControlOwnedBlockIds(resolveDocumentContentControlBlocks(model, control.blockRange))
    : control.tableRowRange ? documentContentControlTableRowBlockIds(resolveDocumentContentControlTableRows(model, control.tableRowRange))
      : new Set([control.blockId])
  const blocks = new Set(owned)
  if (control.tableRowRange) blocks.add(control.tableRowRange.tableId)
  // Editing or copying a reference also addresses the actual note content. Nested references
  // are traversed once, so permissions cannot be bypassed through a second note story.
  const notes = (['footnotes', 'endnotes'] as const).flatMap(kind => Object.values(resources[kind] ?? {}).map(note => ({ kind, note })))
  const visited = new Set<string>()
  const result = new Map<string, DocumentPermissionTarget>()
  const add = (targets: DocumentPermissionTarget[]) => { for (const target of targets) result.set(`${target.scope}:${target.id}`, target) }
  for (const id of blocks) add(permissionTargetsForBlock(model, id))
  const hosts = new Map(documentTextFieldParagraphs(model).map(paragraph => [paragraph.block.id, { block: paragraph.block, model }]))
  const inline = !control.blockRange && !control.tableRowRange
  const defaults = documentDefaultStyle(resources)
  let added = true
  while (added) {
    added = false
    for (const { kind, note } of notes) {
      const key = `${kind}:${note.id}`
      if (visited.has(key)) continue
      const noteKind = kind === 'footnotes' ? 'footnote' as const : 'endnote' as const
      const owners = [...hosts.values()].filter(owner => owner.block.runs.some(run => documentNoteReferenceRunMatches(run, note, noteKind)))
      const addressed = owners.filter(owner => blocks.has(owner.block.id))
      if (!addressed.length) continue
      if (owners.length !== 1) throw new RangeError('注释引用没有唯一的真实内容宿主：' + note.id)
      const owner = addressed[0]!
      if (owner.block.runs.filter(run => documentNoteReferenceRunMatches(run, note, noteKind)).length !== 1) {
        throw new RangeError('注释引用没有唯一的真实字符身份：' + note.id)
      }
      if (inline && owner.block.id === control.blockId) {
        let offset = 0
        const included = owner.block.runs.some(run => {
          const start = offset, end = start + run.text.length
          offset = end
          return (start === end ? start > control.startOffset && end < control.endOffset : start >= control.startOffset && end <= control.endOffset)
            && documentNoteReferenceRunMatches(run, note, noteKind)
        })
        if (!included) continue
      }
      visited.add(key)
      const host = { kind: 'note' as const, noteKind, noteId: note.id }
      // A legacy note still has an independently protected entity before block migration.
      const entity = { scope: 'entity' as const, id: `${host.noteKind}:${note.id}` }
      result.set(`entity:${entity.id}`, entity)
      const legacy = note.content ?? { runs: [createRun(note.text, defaults.text)], paragraphs: documentStoryParagraphs(note.text, defaults.paragraph) }
      const content = migrateDocumentStoryBlocks(legacy, [], host, defaults)
      const noteModel = { ...resources, blocks: content.blocks!, metadata: { ...resources.metadata, documentEditingBlockHost: host,
        documentEditingSectionId: permissionTargetsForBlock(owner.model, owner.block.id).find(target => target.scope === 'section')?.id } }
      for (const paragraph of documentTextFieldParagraphs(noteModel)) {
        if (hosts.has(paragraph.block.id)) throw new RangeError('注释内容块被多个内容区共享：' + paragraph.block.id)
        hosts.set(paragraph.block.id, { block: paragraph.block, model: noteModel })
      }
      for (const id of documentTableDescendantIds(noteModel.blocks)) {
        if (!blocks.has(id)) { blocks.add(id); added = true }
        add(permissionTargetsForBlock(noteModel, id))
      }
    }
  }
  return [...result.values()]
}
