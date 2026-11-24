import type { DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { documentBlockHostForId, type DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import { runText, runsForTarget, textTargets, type DocumentSelection } from '../model/document-model'
import { isDocumentTextBoundary } from '../model/document-story'
import { assertDocumentPermission, permissionTargetsForChangedBlocks, type DocumentPermissionContext } from '../permissions/document-permissions'
import { documentContentControlPermissionTargets } from '../permissions/content-control-permissions'
import { assertDocumentRepeatingSectionItemEditable, deleteRepeatingSectionItem, duplicateRepeatingSectionItem, validateDocumentContentControlRanges } from './content-controls'
import { applyDocumentBlockHostModel, documentModelForBlockHost } from './document-block-host'
import { assertDocumentStoryLocksPreserved } from './document-story-locks'
import { resolveSectionHeaderFooter } from './section-header-footer'
import { prepareDocumentRepeatingContentControlRowsMutation } from './repeating-content-control-rows'
import { prepareDocumentRepeatingContentControlBlocksMutation } from './repeating-content-control-blocks'
import { createDocumentRepeatingContentControlCopyIds, removeDeletedDocumentRepeatingNoteScopes, type DocumentRepeatingContentControlCopyIds } from './repeating-content-control-copy'
import { finalizeDocumentContentControlHostResources } from './content-control-host-resources'
import { documentContentControlTableRowsMutationSelection } from '../model/content-control-table-rows'
import { documentContentControlBlocksMutationSelection } from '../model/content-control-block-mutations'

export interface DocumentRepeatingContentControlMutationPlan {
  model: DocumentModel
  selection: DocumentSelection
  sectionId: string
  controlId?: string
  identityMap: DocumentRepeatingContentControlCopyIds
}

function sectionId(model: DocumentModel, id: string): string {
  const controls = new Map((model.contentControls ?? []).map(control => [control.id, control]))
  const seen = new Set([id])
  let parentId = controls.get(id)?.parentId
  while (parentId) {
    if (seen.has(parentId)) throw new RangeError('重复项目父级关系循环。')
    seen.add(parentId)
    const parent = controls.get(parentId)
    if (!parent) break
    if (parent.type === 'repeatingSection') return parent.id
    parentId = parent.parentId
  }
  throw new RangeError('重复项目所属重复节已不存在。')
}

/** Every entry uses the same real content planner, with its control identity map available to XML transactions. */
export function prepareDocumentRepeatingContentControlMutation(before: DocumentModel, id: string, action: 'duplicate' | 'delete',
  permission?: DocumentPermissionContext, requestedHost?: DocumentAuxiliaryBlockHost): DocumentRepeatingContentControlMutationPlan {
  const item = before.contentControls?.find(control => control.id === id)
  if (!item) throw new RangeError('重复项目已不存在。')
  const host = requestedHost ?? documentBlockHostForId(before, item.blockRange?.blockIds[0] ?? item.tableRowRange?.tableId ?? item.blockId)
  const source = host ? documentModelForBlockHost(before, host, true) : before
  if (permission) {
    const targets = documentContentControlPermissionTargets(source, item, before)
    assertDocumentPermission(before.permissionRules ?? [], permission, 'edit', targets)
    if (action === 'duplicate') assertDocumentPermission(before.permissionRules ?? [], permission, 'copy', targets)
  }
  let plan: DocumentRepeatingContentControlMutationPlan
  if (item.placement === 'row') plan = prepareDocumentRepeatingContentControlRowsMutation(source, id, action, permission)
  else if (item.blockRange) plan = prepareDocumentRepeatingContentControlBlocksMutation(source, id, action, permission)
  else {
    assertDocumentRepeatingSectionItemEditable(source, id, action)
    const model = cloneModel(source), identityMap = createDocumentRepeatingContentControlCopyIds()
    const targets = documentContentControlPermissionTargets(source, item)
    if (permission) {
      assertDocumentPermission(source.permissionRules ?? [], permission, 'edit', targets)
      if (action === 'duplicate') assertDocumentPermission(source.permissionRules ?? [], permission, 'copy', targets)
    }
    const copy = action === 'duplicate' ? duplicateRepeatingSectionItem(model, id, permission, identityMap) : undefined
    if (action === 'delete') deleteRepeatingSectionItem(model, id, permission)
    removeDeletedDocumentRepeatingNoteScopes(source, model)
    const target = textTargets(model).find(target => (target.cellBlockId ?? target.blockId) === (copy?.blockId ?? item.blockId))
    const start = copy?.startOffset ?? item.startOffset, end = copy?.endOffset ?? item.startOffset
    const text = target && runText(runsForTarget(model, target))
    if (!target || text === undefined || !isDocumentTextBoundary(text, start) || !isDocumentTextBoundary(text, end)) {
      throw new RangeError('重复项目提交结果没有合法的文字选区。')
    }
    validateDocumentContentControlRanges(model)
    plan = { model, sectionId: sectionId(source, id), identityMap,
      selection: { anchor: { ...target, offset: start }, focus: { ...target, offset: end } }, ...(copy ? { controlId: copy.id } : {}) }
  }
  if (host) {
    const candidate = cloneModel(before)
    applyDocumentBlockHostModel(candidate, host, plan.model)
    const inline = !item.blockRange && !item.tableRowRange ? { blockId: item.blockId, startOffset: item.startOffset, endOffset: item.endOffset, insertionOffset: item.endOffset } : undefined
    finalizeDocumentContentControlHostResources(before, candidate, action === 'duplicate' ? plan.identityMap : undefined, inline)
    validateDocumentContentControlRanges(candidate)
    const destination = documentModelForBlockHost(candidate, host)
    const structural = item.tableRowRange ? documentContentControlTableRowsMutationSelection(source, destination, id, action)
      : item.blockRange ? documentContentControlBlocksMutationSelection(source, destination, id, action) : undefined
    if (item.tableRowRange || item.blockRange) {
      if (!structural || structural.controlId !== plan.controlId) throw new RangeError('重复项目资源提交后的结构选区已失效。')
      plan.selection = structural.selection
    } else {
      const current = plan.controlId ? destination.contentControls?.find(control => control.id === plan.controlId) : undefined
      const target = textTargets(destination).find(target => (target.cellBlockId ?? target.blockId) === (current?.blockId ?? item.blockId))
      const start = current ? current.startOffset : item.startOffset, end = current ? current.endOffset : item.startOffset
      const text = target && runText(runsForTarget(destination, target))
      if (!target || text === undefined || !isDocumentTextBoundary(text, start) || !isDocumentTextBoundary(text, end)) throw new RangeError('重复项目资源提交后的文字选区已失效。')
      plan.selection = { anchor: { ...target, offset: start }, focus: { ...target, offset: end } }
    }
    if (host.kind === 'story') assertDocumentStoryLocksPreserved(resolveSectionHeaderFooter(before, host.sectionId),
      resolveSectionHeaderFooter(candidate, host.sectionId), host.story, before)
    plan = { ...plan, model: candidate }
  }
  if (permission) assertDocumentPermission(before.permissionRules ?? [], permission, 'edit', permissionTargetsForChangedBlocks(before, plan.model))
  plan.model.wordCount = textTargets(plan.model).reduce((count, target) => count + runText(runsForTarget(plan.model, target)).replace(/\s/g, '').length, 0)
  return plan
}
