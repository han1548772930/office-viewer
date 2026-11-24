import type { DocumentBlock, DocumentContentControl, DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { createParagraph, documentBlockAncestry, findDocumentBlock, type DocumentImage, type RunPatch } from '../model/document-model'
import { documentContentControlBlocksSelection, documentContentControlOwnedBlockIds, resolveDocumentContentControlBlocks } from '../model/content-control-block-range'
import { documentContentControlContains, documentContentControlTableRowBlockIds, reconcileDocumentContentControlTableRows, resolveDocumentContentControlTableRows } from '../model/content-control-table-rows'
import { reconcileDocumentContentControlBlockEdits } from '../model/content-control-block-edits'
import { documentTableDescendantIds } from '../model/document-table-selection'
import { documentFieldLocked, documentTextFieldParagraphs, documentTextFieldRanges, transformDocumentTextFieldsForBlockRangeEdit } from '../model/document-text-fields'
import { documentRunStyleAt } from '../model/document-run-range'
import { reconcileDocumentTableAnchors } from './document-table-anchors'
import { removeDeletedDocumentRepeatingNoteScopes } from './repeating-content-control-copy'

function hasAncestor(model: DocumentModel, control: DocumentContentControl, id: string): boolean {
  const seen = new Set([control.id])
  let parentId = control.parentId
  while (parentId) {
    if (seen.has(parentId)) throw new RangeError('完整块内容控件父级关系循环。')
    if (parentId === id) return true
    seen.add(parentId)
    parentId = model.contentControls?.find(control => control.id === parentId)?.parentId
  }
  return false
}

/** The caller owns a detached candidate; text and pictures use the same owned-tree splice. */
function replaceDocumentBlockContentControlContents(model: DocumentModel, control: DocumentContentControl, output: DocumentBlock[]): void {
  if (!control.blockRange || control.tableRowRange) throw new TypeError('完整块设值需要唯一的块范围。')
  if (['contentLocked', 'sdtContentLocked'].includes(control.lock ?? '')) throw new RangeError('完整块内容控件已锁定：' + control.id)
  const before = cloneModel(model), range = resolveDocumentContentControlBlocks(model, control.blockRange)
  const owned = documentContentControlOwnedBlockIds(range), removing = new Set<string>(), ancestors: DocumentContentControl[] = []
  for (const other of model.contentControls ?? []) {
    if (other.id === control.id || other.metadata?.officeStoryLockSource) continue
    const root = other.blockRange?.blockIds[0] ?? other.tableRowRange?.tableId ?? other.blockId
    if (!findDocumentBlock(model, root)) continue
    const enclosing = documentContentControlContains(model, other, control), inside = documentContentControlContains(model, control, other)
    if (enclosing && hasAncestor(model, control, other.id)) { ancestors.push(other); continue }
    const ids = other.blockRange ? documentContentControlOwnedBlockIds(resolveDocumentContentControlBlocks(model, other.blockRange))
      : other.tableRowRange ? documentContentControlTableRowBlockIds(resolveDocumentContentControlTableRows(model, other.tableRowRange)) : new Set([other.blockId])
    if (![...ids].some(id => owned.has(id))) continue
    if (!inside) throw new RangeError('完整块设值不能切开其他内容控件或绕过其父关系：' + other.id)
    if (['controlLocked', 'sdtLocked', 'contentLocked', 'sdtContentLocked'].includes(other.lock ?? '')) {
      throw new RangeError('完整块设值不能删除锁定的子控件：' + other.id)
    }
    removing.add(other.id)
  }
  const paragraphs = documentTextFieldParagraphs(model), fields = documentTextFieldRanges(model)
  for (const field of fields) {
    const first = paragraphs.findIndex(paragraph => paragraph.block.id === field.startParagraph.block.id)
    const last = paragraphs.findIndex(paragraph => paragraph.block.id === field.endParagraph.block.id)
    if (documentFieldLocked(field.field) && paragraphs.slice(first, last + 1).some(paragraph => owned.has(paragraph.block.id))) {
      throw new RangeError('完整块设值不能改变锁定字段：' + field.field.id)
    }
  }
  const inputs = documentTextFieldParagraphs({ blocks: range.blocks }), outputs = documentTextFieldParagraphs({ blocks: output })
  const last = inputs.at(-1), text = outputs.map(paragraph => paragraph.text).join('\n')
  if (!inputs.length || !last || !outputs.length) throw new RangeError('完整块设值缺少真实文字宿主。')
  if (text.length > 16 * 1024 * 1024 || output.length > 100_000
    || documentTableDescendantIds(model.blocks).size - owned.size + documentTableDescendantIds(output).size > 250_000) {
    throw new RangeError('完整块设值超过文档内容预算。')
  }
  const sameContainerAncestors = ancestors.filter(parent => parent.blockRange && resolveDocumentContentControlBlocks(model, parent.blockRange).container === range.container)
  // Non-text objects stay in the block tree; the field planner receives only real paragraph endpoints.
  transformDocumentTextFieldsForBlockRangeEdit(range.blocks, 0, last.text.length, text, outputs.map(paragraph => paragraph.block), undefined, model)
  range.container.splice(range.startIndex, range.blockIds.length, ...output)
  control.blockRange = { blockIds: output.map(block => block.id) }
  for (const parent of sameContainerAncestors) {
    const index = parent.blockRange!.blockIds.indexOf(range.blockIds[0]!)
    if (index < 0 || range.blockIds.some((id, offset) => parent.blockRange!.blockIds[index + offset] !== id)) {
      throw new RangeError('完整块设值的父控件范围已失效。')
    }
    parent.blockRange!.blockIds.splice(index, range.blockIds.length, ...control.blockRange.blockIds)
  }
  model.contentControls = model.contentControls?.filter(candidate => !removing.has(candidate.id))
  const sourceSelection = documentContentControlBlocksSelection(before, before.contentControls!.find(candidate => candidate.id === control.id)!.blockRange!)
  reconcileDocumentContentControlBlockEdits(before, model, { selection: sourceSelection, activeControlId: control.id, absorbAdjacentInsertions: false })
  reconcileDocumentContentControlTableRows(before, model)
  removeDeletedDocumentRepeatingNoteScopes(before, model)
  const live = model.contentControls?.find(candidate => candidate.id === control.id)
  if (!live?.blockRange) throw new RangeError('完整块设值移除了目标内容控件。')
  const selection = documentContentControlBlocksSelection(model, live.blockRange)
  reconcileDocumentTableAnchors(before, model, selection, new Set(fields.map(field => field.field.id)))
  const result = model.contentControls?.find(candidate => candidate.id === control.id)
  if (!result) throw new RangeError('完整块设值缺少最终内容控件。')
  // Existing setters retain this reference while setting their scalar/date/placeholder state.
  Object.assign(control, result)
  model.contentControls = model.contentControls!.map(candidate => candidate.id === control.id ? control : candidate)
  for (const block of output) for (const ancestor of documentBlockAncestry(model, block.id) ?? []) delete ancestor.metadata.html
}

function replacementSourceParagraph(model: DocumentModel, control: DocumentContentControl): DocumentBlock {
  const target = documentTextFieldParagraphs({ blocks: resolveDocumentContentControlBlocks(model, control.blockRange!).blocks })[0]
  if (!target) throw new RangeError('完整块设值缺少可继承格式的文字宿主。')
  return target.block
}

function replacementParagraph(original: DocumentBlock, text: string, marks: RunPatch, retainMetadata: boolean): DocumentBlock {
  const block = createParagraph(text, marks)
  block.type = original.type; block.level = original.level; block.style = cloneModel(original.style)
  if (original.checklist) block.checklist = cloneModel(original.checklist)
  if (retainMetadata) block.metadata = cloneModel(original.metadata)
  delete block.metadata.html; delete block.metadata.documentTextFields
  const code = original.metadata.codeBlock
  if (code && typeof code === 'object' && !Array.isArray(code)) block.metadata.codeBlock = { ...cloneModel(code), source: text }
  return block
}

/** Replace every owned block with real paragraphs, preserving the first input paragraph's format. */
export function replaceDocumentBlockContentControlText(model: DocumentModel, control: DocumentContentControl, value: string, marks: RunPatch = {}): void {
  if (value.length > 16 * 1024 * 1024) throw new RangeError('完整块设值超过文档内容预算。')
  const parts = value.replace(/\r\n?/g, '\n').split('\n')
  if (parts.length > 100_000) throw new RangeError('完整块设值超过文档内容预算。')
  const original = replacementSourceParagraph(model, control), inherited = documentRunStyleAt(original.runs, 0, original.runs.reduce((count, run) => count + run.text.length, 0)) ?? {}
  const output = parts.map((part, index) => replacementParagraph(original, part, { ...inherited, ...marks }, index === 0))
  replaceDocumentBlockContentControlContents(model, control, output)
}

/** The persisted picture ID must name an image physically owned by this complete block wrapper. */
export function replaceDocumentBlockContentControlImage(model: DocumentModel, control: DocumentContentControl, image: DocumentImage): string {
  if (!control.blockRange || control.tableRowRange || control.type !== 'picture') throw new TypeError('图片替换需要完整块图片内容控件。')
  if (['contentLocked', 'sdtContentLocked'].includes(control.lock ?? '')) throw new RangeError('完整块图片内容控件已锁定：' + control.id)
  const range = resolveDocumentContentControlBlocks(model, control.blockRange), owned = documentContentControlOwnedBlockIds(range)
  const pointer = control.metadata?.pictureBlockId
  if (pointer !== undefined && (typeof pointer !== 'string' || !pointer)) throw new TypeError('图片内容控件的对象身份无效：' + control.id)
  const pictureId = typeof pointer === 'string' ? pointer : undefined
  const existing = pictureId ? findDocumentBlock(model, pictureId) : undefined
  if (pictureId && (!existing || existing.type !== 'image')) throw new RangeError('图片内容控件的对象不是实际图片：' + control.id)
  if (existing && !owned.has(existing.id)) throw new RangeError('图片内容控件指向范围之外的对象：' + control.id)
  if (existing?.type === 'image') {
    for (const other of model.contentControls ?? []) {
      if (other.id === control.id || other.type !== 'picture' || other.metadata?.pictureBlockId !== existing.id) continue
      if (['contentLocked', 'sdtContentLocked'].includes(other.lock ?? '')) throw new RangeError('图片替换不能改变锁定子控件所引用的图片：' + other.id)
      throw new RangeError('图片对象不能被多个图片内容控件同时拥有：' + existing.id)
    }
    existing.metadata = { ...existing.metadata, image: cloneModel(image), pictureInputAnchor: true }
    control.metadata = { ...control.metadata, pictureImage: cloneModel(image) }
    for (const ancestor of documentBlockAncestry(model, existing.id) ?? []) delete ancestor.metadata.html
    return existing.id
  }
  const original = replacementSourceParagraph(model, control), inherited = documentRunStyleAt(original.runs, 0, original.runs.reduce((count, run) => count + run.text.length, 0)) ?? {}
  const anchor = replacementParagraph(original, '', inherited, true), picture = createParagraph()
  picture.type = 'image'; picture.runs = []
  picture.metadata = { image: cloneModel(image), contentControlId: control.id, pictureInputAnchor: true, anchorBlockId: anchor.id, anchorOffset: 0 }
  replaceDocumentBlockContentControlContents(model, control, [anchor, picture])
  const { pictureTailBlockId: _tail, ...metadata } = control.metadata ?? {}
  control.metadata = { ...metadata, pictureBlockId: picture.id, pictureAnchorBlockId: anchor.id, pictureImage: cloneModel(image) }
  return picture.id
}
