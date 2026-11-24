import type { DocumentBlock, DocumentContentControl, DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { createParagraph, documentId, findDocumentBlock, findDocumentBlockContainer, runText, runsForTarget, tableForBlock, textTargets, type DocumentSelection, type DocumentTable } from '../model/document-model'
import { documentTableGrid } from '../model/document-table-grid'
import { deleteDocumentTableRows, insertCopiedDocumentTableRows } from '../model/document-table-row-structure'
import { documentContentControlContains, documentContentControlOwnedRowBlockIds, documentContentControlTableRowBlockIds,
  documentContentControlTableRowsMutationSelection, reconcileDocumentContentControlTableRows, resolveDocumentContentControlTableRows,
  type DocumentContentControlTableRows } from '../model/content-control-table-rows'
import { copyDocumentBlocksForPaste } from './document-block-paste'
import { reconcileDocumentTableAnchors } from './document-table-anchors'
import { isDocumentTextBoundary } from '../model/document-story'
import { documentFieldLocked, documentTextFieldParagraphs, documentTextFieldRanges } from '../model/document-text-fields'
import { reconcileDocumentContentControlBlockEdits } from '../model/content-control-block-edits'
import { assertDocumentRepeatingNoteContentRemovable as assertNoteContentRemovable, createDocumentRepeatingContentControlCopyIds,
  indexDocumentRepeatingContentControlCopies as indexCopies, copyDocumentRepeatingContentControlNotes as copyNotes,
  remapDocumentRepeatingContentControlMetadata as remapCopyMetadata, visitDocumentRepeatingContentControlBlocks as visitBlocks,
  copyDocumentRepeatingContentControlComments, removeDeletedDocumentRepeatingNoteScopes as removeDeletedNoteScopes,
  type DocumentRepeatingContentControlCopyIds } from './repeating-content-control-copy'
import { assertDocumentPermission, permissionTargetsForChangedBlocks,
  type DocumentPermissionContext, type DocumentPermissionTarget } from '../permissions/document-permissions'
import { documentContentControlPermissionTargets } from '../permissions/content-control-permissions'
import { documentContentControlGroupEditAllowed, documentContentControlGroupOwnsInput } from '../model/content-control-group-lock'

export type DocumentRepeatingRowsAction = 'duplicate' | 'delete'
interface RepeatingRowsContext { item: DocumentContentControl; rows: DocumentContentControlTableRows; ancestors: DocumentContentControl[] }
export interface DocumentRepeatingRowsMutationPlan { model: DocumentModel; selection: DocumentSelection; controlId?: string;
  sectionId: string; identityMap: DocumentRepeatingContentControlCopyIds }

const contentLocked = (control: DocumentContentControl) => ['contentLocked', 'sdtContentLocked'].includes(control.lock ?? '')
const wrapperLocked = (control: DocumentContentControl) => ['controlLocked', 'sdtLocked', 'sdtContentLocked'].includes(control.lock ?? '')

function context(model: DocumentModel, id: string): RepeatingRowsContext {
  const item = model.contentControls?.find(control => control.id === id)
  if (item?.type !== 'repeatingSectionItem' || item.placement !== 'row' || !item.tableRowRange) throw new TypeError('重复行项目必须包含完整的持久表格行范围。')
  const rows = resolveDocumentContentControlTableRows(model, item.tableRowRange), ancestors: DocumentContentControl[] = [], seen = new Set([item.id])
  const target = textTargets(model).find(target => (target.cellBlockId ?? target.blockId) === item.blockId)
  const text = target && runText(runsForTarget(model, target))
  if (!target || text === undefined || !documentContentControlTableRowBlockIds(rows).has(item.blockId)
    || item.startOffset > item.endOffset || !isDocumentTextBoundary(text, item.startOffset) || !isDocumentTextBoundary(text, item.endOffset)) {
    throw new RangeError('重复行项目的输入锚点不属于当前表格行或 UTF-16 范围无效。')
  }
  let parentId = item.parentId
  while (parentId) {
    if (seen.has(parentId)) throw new RangeError('重复行项目父级关系循环。')
    seen.add(parentId)
    const parent = model.contentControls?.find(control => control.id === parentId)
    if (!parent || !documentContentControlContains(model, parent, item)) throw new RangeError('重复行项目不在其父控件的持久范围内。')
    ancestors.push(parent); parentId = parent.parentId
  }
  const section = ancestors.find(control => control.type === 'repeatingSection')
  if (!section?.tableRowRange || section.tableRowRange.tableId !== item.tableRowRange.tableId) throw new RangeError('重复行项目必须属于同一表格中的重复节。')
  if (section.repeatingSection?.doNotAllowInsertDelete) throw new RangeError('此重复节禁止增加或删除项目。')
  return { item, rows, ancestors }
}

/** Capability and mutation use the same row scope, merge-boundary and wrapper/content-lock rules. */
export function assertDocumentRepeatingContentControlRowsEditable(model: DocumentModel, id: string, action: DocumentRepeatingRowsAction): void {
  const { item, rows, ancestors } = context(model, id), affected = documentContentControlTableRowBlockIds(rows)
  const owned = documentContentControlOwnedRowBlockIds(rows), paragraphs = documentTextFieldParagraphs(model)
  if (wrapperLocked(item) || contentLocked(item) || ancestors.some(control => contentLocked(control)
    && !documentContentControlGroupOwnsInput(model, control, id))) throw new RangeError('重复行项目或其父控件已锁定。')
  for (const control of model.contentControls ?? []) {
    if (control.id === item.id) continue
    const contains = control.tableRowRange && findDocumentBlock(model, control.tableRowRange.tableId) && documentContentControlContains(model, control, item)
    if (contains && contentLocked(control) && !documentContentControlGroupOwnsInput(model, control, id)) throw new RangeError('重复行项目处于锁定的内容范围。')
    if (action === 'delete' && (affected.has(control.blockId) || documentContentControlContains(model, item, control))
      && !ancestors.some(parent => parent.id === control.id) && (wrapperLocked(control) || contentLocked(control))) {
      throw new RangeError('重复行项目包含锁定的子控件。')
    }
  }
  for (const field of documentTextFieldRanges(model)) {
    const first = paragraphs.findIndex(paragraph => paragraph.block.id === field.startParagraph.block.id)
    const last = paragraphs.findIndex(paragraph => paragraph.block.id === field.endParagraph.block.id)
    const intersects = paragraphs.slice(first, last + 1).some(paragraph => owned.has(paragraph.block.id))
    if (action === 'delete' && intersects && documentFieldLocked(field.field)) throw new RangeError('删除重复行不能改变锁定字段。')
    if (action === 'duplicate' && intersects && paragraphs.slice(first, last + 1).some(paragraph => !owned.has(paragraph.block.id))) {
      throw new RangeError('复制重复行不能切开跨项目字段范围。')
    }
  }
  if (action === 'delete') for (const kind of ['footnotes', 'endnotes'] as const) for (const note of Object.values(model[kind] ?? {})) {
    if (!owned.has(note.referenceBlockId)) continue
    assertNoteContentRemovable(model, note)
  }
  if (action === 'duplicate') {
    if (rows.grid.anchors.some(owner => owner.row < rows.startRow && owner.row + owner.rowSpan > rows.startRow
      || owner.row <= rows.endRow && owner.row + owner.rowSpan > rows.endRow + 1)) throw new RangeError('复制重复行不能切开跨项目纵向合并单元格。')
    const amount = rows.endRow - rows.startRow + 1
    if (rows.grid.rowCount + amount > 10_000 || (rows.grid.rowCount + amount) * rows.grid.columnCount > 1_000_000) throw new RangeError('重复行复制超过表格预算。')
  }
}

export function documentRepeatingContentControlRowsPermissionTargets(model: DocumentModel, id: string): DocumentPermissionTarget[] {
  return documentContentControlPermissionTargets(model, context(model, id).item)
}

function duplicateRows(before: DocumentModel, model: DocumentModel, source: RepeatingRowsContext,
  ids: DocumentRepeatingContentControlCopyIds): string {
  const rows = resolveDocumentContentControlTableRows(model, source.item.tableRowRange!)
  const sourceTable: DocumentTable = { ...rows.table, rows: rows.table.rows.slice(rows.startRow, rows.endRow + 1),
    ...(rows.table.rowStyles ? { rowStyles: rows.table.rowStyles.slice(rows.startRow, rows.endRow + 1) } : {}),
    headerRowCount: Math.max(0, Math.min(rows.rowIds.length, (rows.table.headerRowCount ?? 0) - rows.startRow)),
    titleRow: rows.startRow === 0 && rows.table.titleRow === true }
  const shell: DocumentBlock = { ...rows.block, metadata: { ...rows.block.metadata, table: sourceTable } }
  const copies = copyDocumentBlocksForPaste([shell]), copied = tableForBlock(copies[0]!)!
  indexCopies([shell], copies, ids)
  // The temporary shell is never inserted; anchors to the real outer table keep its existing identity.
  ids.block.delete(rows.block.id)
  const bodyIds = new Set(ids.block.keys()), copyId = documentId('content-control')
  ids.control.set(source.item.id, copyId)
  const ancestorIds = new Set(source.ancestors.map(control => control.id))
  insertCopiedDocumentTableRows(rows.table, rows.grid, rows.endRow + 1, copied, rows.startRow)
  copyNotes(before, model, ids)
  const selected = (before.contentControls ?? []).filter(control => !control.metadata?.officeStoryLockSource && !ancestorIds.has(control.id)
    && (control.id === source.item.id || (control.blockRange ? control.blockRange.blockIds.every(id => ids.block.has(id))
      : control.tableRowRange ? documentContentControlContains(before, source.item, control)
        || ids.block.has(control.tableRowRange.tableId) && control.tableRowRange.rowIds.every(id => ids.row.has(id)) : ids.block.has(control.blockId))))
  for (const control of selected) if (!ids.control.has(control.id)) ids.control.set(control.id, documentId('content-control'))
  const added = selected.map(control => {
    const next = cloneModel(control)
    next.id = ids.control.get(control.id)!
    next.blockId = ids.block.get(control.blockId) ?? control.blockId
    if (next.tableRowRange) next.tableRowRange = { tableId: next.tableRowRange.tableId === rows.block.id ? rows.block.id
      : ids.block.get(next.tableRowRange.tableId) ?? next.tableRowRange.tableId,
      rowIds: next.tableRowRange.rowIds.map(id => ids.row.get(id) ?? id) }
    if (next.blockRange) next.blockRange = { blockIds: next.blockRange.blockIds.map(id => {
      const copiedId = ids.block.get(id)
      if (!copiedId) throw new RangeError('重复行复制不能切开子控件的完整块范围。')
      return copiedId
    }) }
    if (control.id === source.item.id) next.parentId = control.parentId
    else if (control.parentId && ids.control.has(control.parentId)) next.parentId = ids.control.get(control.parentId)!
    else if (bodyIds.has(control.blockId)) next.parentId = copyId
    next.metadata = remapCopyMetadata(next.metadata, ids)
    return next
  })
  model.contentControls = [...(model.contentControls ?? []), ...added]
  const newRows = copied.rows.map(row => row.id)
  for (const ancestor of model.contentControls) if (ancestorIds.has(ancestor.id) && ancestor.tableRowRange?.tableId === rows.block.id) {
    const at = ancestor.tableRowRange.rowIds.indexOf(source.rows.rowIds.at(-1)!) + 1
    if (!at) throw new RangeError('重复行复制的父范围已失效。')
    ancestor.tableRowRange.rowIds.splice(at, 0, ...newRows)
  }
  const newBlockIds = new Set(ids.block.values())
  visitBlocks(model.blocks, block => { if (newBlockIds.has(block.id)) block.metadata = remapCopyMetadata(block.metadata, ids)! })
  for (const kind of ['footnotes', 'endnotes'] as const) for (const note of Object.values(model[kind] ?? {})) {
    if (note.content?.blocks) visitBlocks(note.content.blocks, block => { if (newBlockIds.has(block.id)) block.metadata = remapCopyMetadata(block.metadata, ids)! })
  }
  copyDocumentRepeatingContentControlComments(before, model, ids)
  return copyId
}

function deleteRows(model: DocumentModel, item: DocumentContentControl): void {
  const rows = resolveDocumentContentControlTableRows(model, item.tableRowRange!)
  if (rows.startRow === 0 && rows.endRow === rows.grid.rowCount - 1) {
    const container = findDocumentBlockContainer(model, rows.block.id)
    if (!container) throw new RangeError('重复项目所属表格容器已不存在。')
    const index = container.indexOf(rows.block); container.splice(index, 1)
    if (!textTargets({ ...model, blocks: container }).length) container.splice(Math.min(index, container.length), 0, createParagraph())
  } else {
    deleteDocumentTableRows(rows.table, rows.grid, rows.startRow, rows.endRow)
    documentTableGrid(rows.table)
  }
  model.contentControls = model.contentControls?.filter(control => control.id !== item.id)
}

function reconnectSurvivingParents(before: DocumentModel, model: DocumentModel): void {
  const previous = new Map(before.contentControls?.map(control => [control.id, control])), live = new Map(model.contentControls?.map(control => [control.id, control]))
  for (const control of live.values()) {
    if (!findDocumentBlock(model, control.tableRowRange?.tableId ?? control.blockId)) continue
    const seen = new Set([control.id]); let parentId = control.parentId
    while (parentId && (!live.has(parentId) || !documentContentControlContains(model, live.get(parentId)!, control))) {
      if (seen.has(parentId)) throw new RangeError('删除重复行后的父关系循环。')
      seen.add(parentId); parentId = previous.get(parentId)?.parentId
    }
    if (parentId) control.parentId = parentId
    else delete control.parentId
  }
}

/** Build one detached candidate; fields, controls, comments, notes and permissions finish before publication. */
export function prepareDocumentRepeatingContentControlRowsMutation(before: DocumentModel, id: string, action: DocumentRepeatingRowsAction,
  permission?: DocumentPermissionContext): DocumentRepeatingRowsMutationPlan {
  assertDocumentRepeatingContentControlRowsEditable(before, id, action)
  const source = context(before, id), targets = documentRepeatingContentControlRowsPermissionTargets(before, id)
  if (permission) {
    assertDocumentPermission(before.permissionRules ?? [], permission, 'edit', targets)
    if (action === 'duplicate') assertDocumentPermission(before.permissionRules ?? [], permission, 'copy', targets)
  }
  const model = cloneModel(before), identityMap = createDocumentRepeatingContentControlCopyIds()
  const controlId = action === 'duplicate' ? duplicateRows(before, model, source, identityMap) : undefined
  if (action === 'delete') deleteRows(model, source.item)
  reconcileDocumentContentControlTableRows(before, model)
  reconcileDocumentContentControlBlockEdits(before, model, { absorbAdjacentInsertions: false })
  removeDeletedNoteScopes(before, model)
  reconnectSurvivingParents(before, model)
  const result = documentContentControlTableRowsMutationSelection(before, model, id, action)
  if (!result || controlId !== result.controlId) throw new RangeError('重复行提交结果无法解析合法选区。')
  reconcileDocumentTableAnchors(before, model, result.selection)
  reconcileDocumentContentControlTableRows(before, model)
  for (const group of before.contentControls ?? []) if (contentLocked(group) && group.type === 'group'
    && source.ancestors.some(parent => parent.id === group.id) && !documentContentControlGroupEditAllowed(before, model, group, { activeControlId: id })) {
    throw new RangeError('重复行项目编辑不能改变锁定分组的范围外内容。')
  }
  if (permission) assertDocumentPermission(before.permissionRules ?? [], permission, 'edit', permissionTargetsForChangedBlocks(before, model))
  model.wordCount = textTargets(model).reduce((count, target) => count + runText(runsForTarget(model, target)).replace(/\s/g, '').length, 0)
  return { model, selection: result.selection, sectionId: source.ancestors.find(control => control.type === 'repeatingSection')!.id,
    identityMap, ...(controlId ? { controlId } : {}) }
}
