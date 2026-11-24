import type { DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { documentTextFieldParagraphs, documentTextFieldRanges, documentTextFields, setDocumentTextFields, documentFieldLocked, type DocumentTextField, type DocumentTextFieldParagraph } from '../model/document-text-fields'
import { assertContentControlBlockRemovable, validateDocumentContentControlRanges } from './content-controls'
import { transformDocumentCommentAnchorsForParagraphEdits, type DocumentCommentParagraphEdit } from '../comments/document-comment-anchors'
import { blockAtPoint, documentBlockAncestry, runText, type DocumentSelection } from '../model/document-model'
import { reconcileDocumentTocFieldParents } from '../toc/document-toc-edit'
import { reconcileDocumentNotesAfterStructure } from '../notes/document-notes'
import { reconcileDocumentContentControlBlockEdits } from '../model/content-control-block-edits'

/**
 * Cell structure can move paragraphs without changing their identities. Keep
 * local anchors on those IDs; only deleted paragraphs collapse to a surviving
 * boundary. Invalid crossing fields or locked results reject the detached plan.
 */
export function planDocumentTableFieldAnchors(before: DocumentModel, destination: readonly DocumentTextFieldParagraph[], transformedFieldIds: ReadonlySet<string> = new Set()): Map<string, DocumentTextField[]> {
  const source = documentTextFieldParagraphs(before)
  const sourceById = new Map(source.map((paragraph, index) => [paragraph.block.id, { paragraph, index }]))
  const destinationById = new Map(destination.map((paragraph, index) => [paragraph.block.id, { paragraph, index }]))
  const fieldsByOwner = new Map<string, DocumentTextField[]>()
  const fields = documentTextFieldRanges(before)
  const oldIds = new Set(fields.map(range => range.field.id))
  const current = new Map(destination.flatMap(paragraph => documentTextFields(paragraph.block).map(field => [field.id, { field, paragraph }] as const)))
  const append = (owner: string, field: DocumentTextField) => {
    const values = fieldsByOwner.get(owner) ?? []
    values.push(field); fieldsByOwner.set(owner, values)
  }
  for (const [id, entry] of current) if (!oldIds.has(id)) append(entry.paragraph.block.id, cloneModel(entry.field))
  for (const range of fields) {
    const firstIndex = sourceById.get(range.startParagraph.block.id)!.index, lastIndex = sourceById.get(range.endParagraph.block.id)!.index
    const previousResult = source.slice(firstIndex, lastIndex + 1).map((paragraph, index, paragraphs) => paragraph.text.slice(
      index === 0 ? range.field.startOffset : 0, index === paragraphs.length - 1 ? range.field.endOffset : undefined)).join('\n')
    if (transformedFieldIds.has(range.field.id)) {
      const entry = current.get(range.field.id)
      if (!entry) {
        if (documentFieldLocked(range.field)) throw new RangeError(`表格编辑不能删除锁定字段：${range.field.id}`)
        continue
      }
      const from = destinationById.get(entry.paragraph.block.id)!, to = destinationById.get(entry.field.endBlockId ?? entry.paragraph.block.id)
      if (!to || from.index > to.index || from.index === to.index && entry.field.startOffset > entry.field.endOffset) throw new RangeError(`表格移动会反转跨段字段：${range.field.id}`)
      const nextResult = destination.slice(from.index, to.index + 1).map((paragraph, index, paragraphs) => paragraph.text.slice(
        index === 0 ? entry.field.startOffset : 0, index === paragraphs.length - 1 ? entry.field.endOffset : undefined)).join('\n')
      if (documentFieldLocked(range.field) && nextResult !== previousResult) throw new RangeError(`表格编辑不能改变锁定字段结果：${range.field.id}`)
      const field = cloneModel(entry.field)
      if (field.cachedResult !== undefined && nextResult !== previousResult) field.cachedResult = nextResult
      append(from.paragraph.block.id, field)
      continue
    }
    const surviving = source.slice(firstIndex, lastIndex + 1).flatMap(paragraph => {
      const entry = destinationById.get(paragraph.block.id)
      return entry ? [entry] : []
    })
    if (!surviving.length) {
      if (documentFieldLocked(range.field)) throw new RangeError(`表格编辑不能删除锁定字段：${range.field.id}`)
      continue
    }
    const from = destinationById.get(range.startParagraph.block.id) ?? surviving[0]!
    const to = destinationById.get(range.endParagraph.block.id) ?? surviving.at(-1)!
    const startOffset = from.paragraph.block.id === range.startParagraph.block.id ? range.field.startOffset : 0
    const endOffset = to.paragraph.block.id === range.endParagraph.block.id ? range.field.endOffset : to.paragraph.text.length
    if (from.index > to.index || from.index === to.index && startOffset > endOffset) {
      throw new RangeError(`表格移动会反转跨段字段：${range.field.id}`)
    }
    const nextResult = destination.slice(from.index, to.index + 1).map((paragraph, index, paragraphs) => paragraph.text.slice(
      index === 0 ? startOffset : 0, index === paragraphs.length - 1 ? endOffset : undefined)).join('\n')
    if (documentFieldLocked(range.field) && nextResult !== previousResult) throw new RangeError(`表格编辑不能改变锁定字段结果：${range.field.id}`)
    const definition = cloneModel(range.field)
    delete definition.endBlockId
    const field: DocumentTextField = { ...definition, startOffset, endOffset,
      ...(from.paragraph.block.id === to.paragraph.block.id ? {} : { endBlockId: to.paragraph.block.id }) }
    if (field.cachedResult !== undefined && nextResult !== previousResult) field.cachedResult = nextResult
    append(from.paragraph.block.id, field)
  }
  // Validate the projected field collection on paragraph shells, without
  // cloning or changing the document for a toolbar capability query.
  const projected = destination.map(paragraph => ({ ...paragraph.block, metadata: { ...paragraph.block.metadata } }))
  for (const paragraph of projected) setDocumentTextFields(paragraph, fieldsByOwner.get(paragraph.id) ?? [])
  documentTextFieldRanges({ blocks: projected })
  return fieldsByOwner
}

export function reconcileDocumentTableAnchors(before: DocumentModel, after: DocumentModel, selection: DocumentSelection, transformedFieldIds: ReadonlySet<string> = new Set()): void {
  reconcileDocumentContentControlBlockEdits(before, after, { absorbAdjacentInsertions: false })
  const source = documentTextFieldParagraphs(before), destination = documentTextFieldParagraphs(after)
  const sourceById = new Map(source.map(paragraph => [paragraph.block.id, { paragraph }]))
  const destinationById = new Map(destination.map(paragraph => [paragraph.block.id, { paragraph }]))
  const removed = new Set(source.filter(paragraph => !destinationById.has(paragraph.block.id)).map(paragraph => paragraph.block.id))
  for (const blockId of removed) assertContentControlBlockRemovable(after, blockId)
  const fieldsByOwner = planDocumentTableFieldAnchors(before, destination, transformedFieldIds)
  for (const paragraph of destination) {
    if (paragraph.block.metadata.documentTextFields !== undefined || fieldsByOwner.has(paragraph.block.id)) {
      setDocumentTextFields(paragraph.block, fieldsByOwner.get(paragraph.block.id) ?? [])
    }
  }
  reconcileDocumentTocFieldParents(before, after)
  documentTextFieldRanges(after)
  const preferred = blockAtPoint(after, selection.focus)
  const boundary = (preferred ? destinationById.get(preferred.id)?.paragraph : undefined) ?? destination[0]
  if (removed.size && !boundary) throw new RangeError('删除表格后必须保留可编辑段落')
  const edits = new Map<string, DocumentCommentParagraphEdit>()
  if (boundary) for (const id of removed) {
    edits.set(id, { blockId: boundary.block.id, startOffset: 0, endOffset: sourceById.get(id)!.paragraph.text.length,
      insertedText: '', targetOffset: Math.max(0, Math.min(selection.focus.offset, boundary.text.length)) })
  }
  transformDocumentCommentAnchorsForParagraphEdits(after, source, destination, edits)
  reconcileDocumentNotesAfterStructure(after, removed)
  validateDocumentContentControlRanges(after)
  documentTextFieldRanges(after)
  // Parent source HTML is no longer an authoritative representation after a
  // structural mutation. The model and retained provenance remain available.
  for (const paragraph of destination) if (sourceById.get(paragraph.block.id)?.paragraph.text !== runText(paragraph.block.runs)) {
    for (const ancestor of documentBlockAncestry(after, paragraph.block.id) ?? []) delete ancestor.metadata.html
  }
}
