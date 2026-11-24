import type { DocumentContentControl, DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { createParagraph, documentId, findDocumentBlock, runText, runsForTarget, textTargets, type DocumentSelection } from '../model/document-model'
import { documentContentControlContains, reconcileDocumentContentControlTableRows } from '../model/content-control-table-rows'
import { documentContentControlOwnedBlockIds, resolveDocumentContentControlBlocks, type DocumentContentControlBlocks } from '../model/content-control-block-range'
import { reconcileDocumentContentControlBlockEdits } from '../model/content-control-block-edits'
import { documentContentControlBlocksMutationSelection } from '../model/content-control-block-mutations'
import { documentTableDescendantIds } from '../model/document-table-selection'
import { documentFieldLocked, documentTextFieldParagraphs, documentTextFieldRanges } from '../model/document-text-fields'
import { isDocumentTextBoundary } from '../model/document-story'
import { assertDocumentPermission, permissionTargetsForChangedBlocks,
  type DocumentPermissionContext, type DocumentPermissionTarget } from '../permissions/document-permissions'
import { documentContentControlPermissionTargets } from '../permissions/content-control-permissions'
import { documentContentControlGroupEditAllowed, documentContentControlGroupOwnsInput } from '../model/content-control-group-lock'
import { copyDocumentBlocksForPaste } from './document-block-paste'
import { reconcileDocumentTableAnchors } from './document-table-anchors'
import { assertDocumentRepeatingNoteContentRemovable, copyDocumentRepeatingContentControlComments, copyDocumentRepeatingContentControlNotes,
  createDocumentRepeatingContentControlCopyIds, indexDocumentRepeatingContentControlCopies, remapDocumentRepeatingContentControlMetadata,
  removeDeletedDocumentRepeatingNoteScopes, visitDocumentRepeatingContentControlBlocks, type DocumentRepeatingContentControlCopyIds } from './repeating-content-control-copy'

export type DocumentRepeatingBlocksAction = 'duplicate' | 'delete'
interface RepeatingBlocksContext { item: DocumentContentControl; range: DocumentContentControlBlocks; ancestors: DocumentContentControl[]; section: DocumentContentControl }
export interface DocumentRepeatingBlocksMutationPlan {
  model: DocumentModel
  selection: DocumentSelection
  controlId?: string
  sectionId: string
  identityMap: DocumentRepeatingContentControlCopyIds
}

const contentLocked = (control: DocumentContentControl) => ['contentLocked', 'sdtContentLocked'].includes(control.lock ?? '')
const wrapperLocked = (control: DocumentContentControl) => ['controlLocked', 'sdtLocked', 'sdtContentLocked'].includes(control.lock ?? '')

function context(model: DocumentModel, id: string): RepeatingBlocksContext {
  const item = model.contentControls?.find(control => control.id === id)
  if (item?.type !== 'repeatingSectionItem' || item.placement !== 'block' || !item.blockRange || item.tableRowRange) {
    throw new TypeError('重复块项目必须包含同一内容容器内连续的完整块范围。')
  }
  const range = resolveDocumentContentControlBlocks(model, item.blockRange), owned = documentContentControlOwnedBlockIds(range)
  const target = textTargets(model).find(target => (target.cellBlockId ?? target.blockId) === item.blockId)
  const text = target && runText(runsForTarget(model, target))
  if (!target || text === undefined || !owned.has(item.blockId) || item.startOffset > item.endOffset
    || !isDocumentTextBoundary(text, item.startOffset) || !isDocumentTextBoundary(text, item.endOffset)) {
    throw new RangeError('重复块项目的合法输入锚点已不存在。')
  }
  const ancestors: DocumentContentControl[] = [], seen = new Set([item.id])
  let parentId = item.parentId
  while (parentId) {
    if (seen.has(parentId)) throw new RangeError('重复块项目父级关系循环。')
    seen.add(parentId)
    const parent = model.contentControls?.find(control => control.id === parentId)
    if (!parent || !['group', 'repeatingSection', 'repeatingSectionItem'].includes(parent.type)
      || !documentContentControlContains(model, parent, item)) throw new RangeError('重复块项目不属于其父控件的真实范围。')
    ancestors.push(parent); parentId = parent.parentId
  }
  const section = ancestors.find(control => control.type === 'repeatingSection')
  if (!section?.blockRange || resolveDocumentContentControlBlocks(model, section.blockRange).container !== range.container) {
    throw new RangeError('重复块项目必须属于同一块容器内的重复节。')
  }
  if (section.repeatingSection?.doNotAllowInsertDelete) throw new RangeError('此重复节禁止增加或删除项目。')
  return { item, range, ancestors, section }
}

/** All rows, continuations and object blocks in the owned block tree participate in these checks. */
export function assertDocumentRepeatingContentControlBlocksEditable(model: DocumentModel, id: string, action: DocumentRepeatingBlocksAction): void {
  const source = context(model, id), owned = documentContentControlOwnedBlockIds(source.range)
  if (wrapperLocked(source.item) || contentLocked(source.item) || source.ancestors.some(control => contentLocked(control)
    && !documentContentControlGroupOwnsInput(model, control, id))) throw new RangeError('重复块项目或其父控件已锁定。')
  const ancestorIds = new Set(source.ancestors.map(control => control.id))
  for (const control of model.contentControls ?? []) {
    if (control.id === id || ancestorIds.has(control.id) || control.metadata?.officeStoryLockSource) continue
    if (!findDocumentBlock(model, control.blockRange?.blockIds[0] ?? control.tableRowRange?.tableId ?? control.blockId)) continue
    const inside = documentContentControlContains(model, source.item, control)
    const enclosing = documentContentControlContains(model, control, source.item)
    if (enclosing && contentLocked(control) && !documentContentControlGroupOwnsInput(model, control, id)) throw new RangeError('重复块项目处于锁定内容范围。')
    if (control.blockRange) {
      const blocks = documentContentControlOwnedBlockIds(resolveDocumentContentControlBlocks(model, control.blockRange))
      if ([...blocks].some(id => owned.has(id)) && !inside && !enclosing) throw new RangeError('重复项目编辑不能切开其他完整块内容控件。')
    }
    if (action === 'delete' && inside && (wrapperLocked(control) || contentLocked(control))) throw new RangeError('重复块项目包含锁定的子控件。')
  }
  const paragraphs = documentTextFieldParagraphs(model)
  for (const range of documentTextFieldRanges(model)) {
    const first = paragraphs.findIndex(paragraph => paragraph.block.id === range.startParagraph.block.id)
    const last = paragraphs.findIndex(paragraph => paragraph.block.id === range.endParagraph.block.id)
    const affected = paragraphs.slice(first, last + 1)
    if (!affected.some(paragraph => owned.has(paragraph.block.id))) continue
    if (action === 'delete' && documentFieldLocked(range.field)) throw new RangeError('删除重复块不能改变锁定字段。')
    if (action === 'duplicate' && affected.some(paragraph => !owned.has(paragraph.block.id))) throw new RangeError('复制重复块不能切开跨项目字段。')
  }
  if (action === 'delete') for (const kind of ['footnotes', 'endnotes'] as const) for (const note of Object.values(model[kind] ?? {})) {
    if (owned.has(note.referenceBlockId)) assertDocumentRepeatingNoteContentRemovable(model, note)
  }
  if (action === 'duplicate' && documentTableDescendantIds(model.blocks).size + owned.size > 250_000) throw new RangeError('重复块复制超过文档块预算。')
}

export function documentRepeatingContentControlBlocksPermissionTargets(model: DocumentModel, id: string): DocumentPermissionTarget[] {
  return documentContentControlPermissionTargets(model, context(model, id).item)
}

function duplicateBlocks(before: DocumentModel, model: DocumentModel, source: RepeatingBlocksContext, ids: DocumentRepeatingContentControlCopyIds): string {
  const range = resolveDocumentContentControlBlocks(model, source.item.blockRange!), copies = copyDocumentBlocksForPaste(range.blocks)
  indexDocumentRepeatingContentControlCopies(range.blocks, copies, ids)
  const bodyIds = new Set(ids.block.keys()), ancestorIds = new Set(source.ancestors.map(control => control.id))
  range.container.splice(range.endIndex + 1, 0, ...copies)
  copyDocumentRepeatingContentControlNotes(before, model, ids)
  const controls = (before.contentControls ?? []).filter(control => !control.metadata?.officeStoryLockSource && !ancestorIds.has(control.id)
    && (control.blockRange ? control.blockRange.blockIds.every(id => ids.block.has(id))
      : control.tableRowRange ? ids.block.has(control.tableRowRange.tableId) && control.tableRowRange.rowIds.every(id => ids.row.has(id)) : ids.block.has(control.blockId)))
  if (!controls.some(control => control.id === source.item.id)) throw new RangeError('复制范围不包含真实重复块项目。')
  for (const control of controls) ids.control.set(control.id, documentId('content-control'))
  const copyId = ids.control.get(source.item.id)!
  const added = controls.map(control => {
    const next = cloneModel(control)
    next.id = ids.control.get(control.id)!
    next.blockId = ids.block.get(control.blockId)!
    if (next.blockRange) next.blockRange = { blockIds: next.blockRange.blockIds.map(id => ids.block.get(id)!) }
    if (next.tableRowRange) next.tableRowRange = { tableId: ids.block.get(next.tableRowRange.tableId)!, rowIds: next.tableRowRange.rowIds.map(id => ids.row.get(id)!) }
    if (control.id === source.item.id) next.parentId = control.parentId
    else if (control.parentId && ids.control.has(control.parentId)) next.parentId = ids.control.get(control.parentId)!
    else if (bodyIds.has(control.blockId)) next.parentId = copyId
    next.metadata = remapDocumentRepeatingContentControlMetadata(next.metadata, ids)
    return next
  })
  model.contentControls = [...(model.contentControls ?? []), ...added]
  for (const ancestor of model.contentControls) {
    if (!ancestorIds.has(ancestor.id) || !ancestor.blockRange) continue
    const parent = resolveDocumentContentControlBlocks(before, ancestor.blockRange)
    if (parent.container !== source.range.container) continue
    const at = ancestor.blockRange.blockIds.indexOf(source.range.blockIds.at(-1)!) + 1
    if (!at) throw new RangeError('重复块复制的父范围已失效。')
    ancestor.blockRange.blockIds.splice(at, 0, ...copies.map(block => block.id))
  }
  const newIds = new Set(ids.block.values())
  visitDocumentRepeatingContentControlBlocks(model.blocks, block => {
    if (newIds.has(block.id)) block.metadata = remapDocumentRepeatingContentControlMetadata(block.metadata, ids)!
  })
  for (const kind of ['footnotes', 'endnotes'] as const) for (const note of Object.values(model[kind] ?? {})) {
    if (note.content?.blocks) visitDocumentRepeatingContentControlBlocks(note.content.blocks, block => {
      if (newIds.has(block.id)) block.metadata = remapDocumentRepeatingContentControlMetadata(block.metadata, ids)!
    })
  }
  copyDocumentRepeatingContentControlComments(before, model, ids)
  return copyId
}

function deleteBlocks(model: DocumentModel, source: RepeatingBlocksContext): void {
  const range = resolveDocumentContentControlBlocks(model, source.item.blockRange!), owned = documentContentControlOwnedBlockIds(range)
  range.container.splice(range.startIndex, range.blockIds.length)
  if (!textTargets({ ...model, blocks: range.container }).length) range.container.splice(Math.min(range.startIndex, range.container.length), 0, createParagraph())
  model.contentControls = model.contentControls?.filter(control => control.id !== source.item.id
    && (source.ancestors.some(parent => parent.id === control.id) || !owned.has(control.blockRange?.blockIds[0] ?? control.tableRowRange?.tableId ?? control.blockId)))
}

/** One detached candidate owns the real block splice, all resource remaps and the resulting caret. */
export function prepareDocumentRepeatingContentControlBlocksMutation(before: DocumentModel, id: string, action: DocumentRepeatingBlocksAction,
  permission?: DocumentPermissionContext): DocumentRepeatingBlocksMutationPlan {
  assertDocumentRepeatingContentControlBlocksEditable(before, id, action)
  const source = context(before, id), targets = documentRepeatingContentControlBlocksPermissionTargets(before, id)
  if (permission) {
    assertDocumentPermission(before.permissionRules ?? [], permission, 'edit', targets)
    if (action === 'duplicate') assertDocumentPermission(before.permissionRules ?? [], permission, 'copy', targets)
  }
  const model = cloneModel(before), identityMap = createDocumentRepeatingContentControlCopyIds()
  const controlId = action === 'duplicate' ? duplicateBlocks(before, model, source, identityMap) : undefined
  if (action === 'delete') deleteBlocks(model, source)
  reconcileDocumentContentControlBlockEdits(before, model, { absorbAdjacentInsertions: false })
  reconcileDocumentContentControlTableRows(before, model)
  removeDeletedDocumentRepeatingNoteScopes(before, model)
  const result = documentContentControlBlocksMutationSelection(before, model, id, action)
  if (!result || result.controlId !== controlId) throw new RangeError('重复块提交结果无法解析合法选区。')
  reconcileDocumentTableAnchors(before, model, result.selection)
  for (const group of before.contentControls ?? []) if (contentLocked(group) && group.type === 'group'
    && source.ancestors.some(parent => parent.id === group.id) && !documentContentControlGroupEditAllowed(before, model, group, { activeControlId: id })) {
    throw new RangeError('重复块项目编辑不能改变锁定分组的范围外内容。')
  }
  if (permission) assertDocumentPermission(before.permissionRules ?? [], permission, 'edit', permissionTargetsForChangedBlocks(before, model))
  model.wordCount = textTargets(model).reduce((count, target) => count + runText(runsForTarget(model, target)).replace(/\s/g, '').length, 0)
  return { model, selection: result.selection, sectionId: source.section.id, identityMap, ...(controlId ? { controlId } : {}) }
}
