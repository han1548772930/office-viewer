import type { DocumentModel } from '../../../domain/model'
import { documentBlockAncestry, tableForBlock } from '../model/document-model'
import { documentNoteReferenceRunMatches } from '../model/document-note-reference'
import { transformDocumentTextFieldsForBlockEdit } from '../model/document-text-fields'
import { transformDocumentCommentAnchorsForTextEdit } from '../comments/thread-comments'
import { applyDocumentBlockHostModel, documentModelForBlockHost } from '../services/document-block-host'
import { documentTableCellContentContainers } from '../model/document-table-continuations'
import { documentNoteReferenceHosts } from './document-note-references'
import { resolveDocumentNoteLabels, type DocumentNoteLabel } from './document-note-numbering'

/** Generated labels resize enclosing wrappers without deleting their identities. */
function remapControlOffsets(model: DocumentModel, blockId: string, start: number, end: number, length: number): void {
  const delta = length - (end - start)
  for (const control of model.contentControls ?? []) {
    if (control.blockId !== blockId) continue
    if (end <= control.startOffset) { control.startOffset += delta; control.endOffset += delta }
    else if (start < control.endOffset && end > control.startOffset || start === end && start >= control.startOffset && start <= control.endOffset) {
      control.startOffset = Math.min(control.startOffset, start)
      control.endOffset = Math.max(control.startOffset, control.endOffset >= end ? control.endOffset + delta : start + length)
    }
  }
}

/** Source text is a compatibility cache. eachPage uses physical-page labels in the layout DTO. */
export function renumberDocumentNotes(model: DocumentModel): void {
  const labels = resolveDocumentNoteLabels(model)
  const byBlock = new Map<string, DocumentNoteLabel[]>()
  for (const label of labels.values()) {
    const list = byBlock.get(label.reference.block.id) ?? []
    list.push(label); byBlock.set(label.reference.block.id, list)
  }
  for (const entry of documentNoteReferenceHosts(model)) {
    if (entry.linked) continue
    const scoped = entry.host ? documentModelForBlockHost(model, entry.host, true) : model
    let changed = false
    const visit = (blocks: DocumentModel['blocks']) => {
      for (const block of blocks) {
        let offset = 0
        for (const run of block.runs) {
          const label = byBlock.get(block.id)?.find(label => documentNoteReferenceRunMatches(run, label.reference.note, label.reference.kind))
          if (label) {
            const note = (label.reference.kind === 'footnote' ? model.footnotes : model.endnotes)![label.reference.note.id]!
            note.referenceBlockId = block.id; note.referenceRunId = run.id
            if (label.label !== run.text) {
              const end = offset + run.text.length
              transformDocumentTextFieldsForBlockEdit(block, offset, end, label.label, undefined, scoped)
              remapControlOffsets(scoped, block.id, offset, end, label.label.length)
              run.text = label.label
              transformDocumentCommentAnchorsForTextEdit(scoped, block.id, offset, end, label.label)
              for (const ancestor of documentBlockAncestry(scoped, block.id) ?? []) delete ancestor.metadata.html
              changed = true
            }
          }
          offset += run.text.length
        }
        for (const row of tableForBlock(block)?.rows ?? []) for (const cell of row.cells) for (const container of documentTableCellContentContainers(cell)) visit(container)
      }
    }
    visit(scoped.blocks)
    if (changed && entry.host) applyDocumentBlockHostModel(model, entry.host, scoped)
  }
}
