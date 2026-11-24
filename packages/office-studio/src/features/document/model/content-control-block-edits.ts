import type { DocumentContentControl, DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { findDocumentBlock, findDocumentBlockContainer, type DocumentSelection } from './document-model'
import { documentContentControlContains, documentContentControlContainsSelection } from './content-control-table-rows'
import { documentContentControlBlockContentSignature, reanchorDocumentBlockContentControl, resolveDocumentContentControlBlocks } from './content-control-block-range'
import { documentTableDescendantIds } from './document-table-selection'
import { documentContentControlGroupEditAllowed } from './content-control-group-lock'

export interface DocumentContentControlBlockEditContext {
  /** The checked source selection. Blocks inserted while editing its contents belong to the wrapper. */
  selection?: DocumentSelection
  activeControlId?: string
  absorbAdjacentInsertions?: boolean
}

const contentLocked = (control: DocumentContentControl) => ['contentLocked', 'sdtContentLocked'].includes(control.lock ?? '')
const wrapperLocked = (control: DocumentContentControl) => ['controlLocked', 'sdtLocked', 'sdtContentLocked'].includes(control.lock ?? '')

/**
 * One editing host at a time. Existing block IDs keep their order; new blocks between
 * surviving endpoints, and a checked split at an endpoint, join the same wrapper.
 * A complete removal also removes its wrappers. Every plan is checked before publication.
 */
export function reconcileDocumentContentControlBlockEdits(before: DocumentModel, after: DocumentModel,
  context: DocumentContentControlBlockEditContext = {}): void {
  if (!(before.contentControls ?? []).some(control => control.blockRange)
    && !(after.contentControls ?? []).some(control => control.blockRange)) return
  const previous = new Map((before.contentControls ?? []).map(control => [control.id, control]))
  const planned = cloneModel(after.contentControls ?? []), current = new Map(planned.map(control => [control.id, control]))
  const model: DocumentModel = { ...after, contentControls: planned }, originalIds = documentTableDescendantIds(before.blocks)
  const removing = new Set<string>()
  for (const control of planned) {
    if (!control.blockRange) continue
    if (control.tableRowRange || control.placement !== 'block') throw new RangeError('完整块内容控件不能同时属于物理表格行。')
    const source = previous.get(control.id)
    const sourcePresent = source?.blockRange && !!findDocumentBlock(before, source.blockRange.blockIds[0]!)
    // The full document retains stored stories/notes; their own commands reconcile those block trees.
    if (!sourcePresent && !findDocumentBlock(model, control.blockRange.blockIds[0]!)) continue
    if (source?.blockRange && sourcePresent && JSON.stringify(source.blockRange) === JSON.stringify(control.blockRange)) {
      const oldRange = resolveDocumentContentControlBlocks(before, source.blockRange)
      const survivors = oldRange.blockIds.filter(id => findDocumentBlock(model, id))
      if (!survivors.length) { removing.add(control.id); continue }
      const container = findDocumentBlockContainer(model, survivors[0]!)!
      const positions = survivors.map(id => container.findIndex(block => block.id === id))
      if (positions.some((position, index) => position < 0 || index > 0 && position <= positions[index - 1]!)) {
        throw new RangeError('结构编辑不能反转完整块内容控件或把它拆到不同容器。')
      }
      let start = positions[0]!, end = positions.at(-1)!
      const sourceInside = context.selection && documentContentControlContainsSelection(before, source, context.selection)
      if (context.absorbAdjacentInsertions !== false && sourceInside) {
        while (start > 0 && !originalIds.has(container[start - 1]!.id)) start--
        while (end + 1 < container.length && !originalIds.has(container[end + 1]!.id)) end++
      }
      const blockIds = container.slice(start, end + 1).map(block => block.id)
      if (blockIds.some(id => originalIds.has(id) && !oldRange.blockIds.includes(id))) {
        throw new RangeError('结构编辑不能把其他已有块移入内容控件。')
      }
      control.blockRange = { blockIds }
    }
    resolveDocumentContentControlBlocks(model, control.blockRange)
    reanchorDocumentBlockContentControl(model, control)
    if (source?.blockRange && sourcePresent && contentLocked(source)
      && documentContentControlBlockContentSignature(before, source.blockRange) !== documentContentControlBlockContentSignature(model, control.blockRange)) {
      const insideChild = documentContentControlGroupEditAllowed(before, model, source, context)
      if (!insideChild) throw new RangeError('结构编辑不能改变锁定的完整块内容控件。')
    }
  }
  for (const source of previous.values()) {
    if (!source.blockRange || !findDocumentBlock(before, source.blockRange.blockIds[0]!)) continue
    if (!current.has(source.id)) removing.add(source.id)
  }
  let expanded = true
  while (expanded) {
    expanded = false
    for (const control of planned) if (control.parentId && removing.has(control.parentId) && !removing.has(control.id)) {
      removing.add(control.id); expanded = true
    }
  }
  for (const id of removing) {
    const source = previous.get(id)
    if (source && (wrapperLocked(source) || contentLocked(source))) throw new RangeError('结构编辑不能移除锁定的内容控件：' + id)
  }
  for (const control of planned) {
    if (removing.has(control.id) || !findDocumentBlock(model, control.blockRange?.blockIds[0] ?? control.tableRowRange?.tableId ?? control.blockId)) continue
    let parent = control.parentId
    const seen = new Set([control.id])
    while (parent && (removing.has(parent) || !current.has(parent))) {
      if (seen.has(parent)) throw new RangeError('结构编辑后的内容控件父级关系循环。')
      seen.add(parent); parent = previous.get(parent)?.parentId
    }
    if (parent && !documentContentControlContains(model, current.get(parent)!, control)) throw new RangeError('结构编辑把内容控件移到了其父级范围之外。')
    if (parent) control.parentId = parent
    else delete control.parentId
  }
  after.contentControls = planned.filter(control => !removing.has(control.id))
}
