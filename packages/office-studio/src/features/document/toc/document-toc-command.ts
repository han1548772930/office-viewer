import type { DocumentModel } from '../../../domain/model'
import { collapsedSelection, type DocumentSelection } from '../model/document-model'
import { assertDocumentPermission, permissionTargetsForChangedBlocks, permissionTargetsForSelection, type DocumentPermissionContext } from '../permissions/document-permissions'
import { documentTocActionUnavailableReason, documentTocIdAtBlock, documentTocIdForSelection, prepareDocumentTocDeletion, prepareDocumentTocFullUpdate,
  prepareDocumentTocPageNumberUpdate, supportsDocumentTableOfContents, type DocumentTocAction, type DocumentTocPageNumbers } from './document-toc'
import { documentTocMutationEntityIds, type DocumentTocMutationPlan } from './document-toc-edit'

export interface DocumentTocActionState { available: boolean; reason?: string }

/** Menu, toolbar and commands prepare the same candidate and check the same ACLs. */
export function prepareDocumentTocAction(
  model: DocumentModel,
  selection: DocumentSelection,
  tocId: string,
  action: DocumentTocAction,
  pageNumbers: DocumentTocPageNumbers,
  context: DocumentPermissionContext,
  timestamp = Date.now(),
): DocumentTocMutationPlan {
  if (!['pageNumbersOnly', 'all', 'delete'].includes(action)) throw new TypeError('目录操作无效。')
  if (!supportsDocumentTableOfContents(model)) throw new RangeError('当前文档模式尚未接入正文目录。')
  if (documentTocIdForSelection(model, selection) !== tocId) throw new RangeError('目录命令要求选区位于目标目录内。')
  const blocks = model.blocks.filter(block => documentTocIdAtBlock(block) === tocId)
  assertDocumentPermission(model.permissionRules ?? [], context, 'edit', blocks.flatMap(block => permissionTargetsForSelection(model,
    collapsedSelection({ blockId: block.id, offset: 0 }))))
  const plan = action === 'delete' ? prepareDocumentTocDeletion(model, tocId, selection, timestamp)
    : action === 'all' ? prepareDocumentTocFullUpdate(model, tocId, pageNumbers, selection, timestamp)
      : prepareDocumentTocPageNumberUpdate(model, tocId, pageNumbers, selection, timestamp)
  if (!plan) throw new RangeError(documentTocActionUnavailableReason(model, tocId, action, pageNumbers)
    ?? (action === 'pageNumbersOnly' ? '目录的布局页码尚未就绪，或原生页码范围尚未映射。' : '当前目录不能安全执行结构变更。'))
  if (action === 'pageNumbersOnly' && !plan.changed && plan.updatablePageFields === 0) throw new RangeError('目录中没有可更新的页码结果。')
  if (plan.changed) assertDocumentPermission(model.permissionRules ?? [], context, 'edit', [
    ...permissionTargetsForChangedBlocks(model, { ...model, blocks: plan.blocks }),
    ...documentTocMutationEntityIds(model, plan).map(id => ({ scope: 'entity' as const, id })),
  ])
  return plan
}

export function documentTocActionState(
  model: DocumentModel, selection: DocumentSelection, tocId: string, action: DocumentTocAction,
  pageNumbers: DocumentTocPageNumbers, context: DocumentPermissionContext,
): DocumentTocActionState {
  try {
    prepareDocumentTocAction(model, selection, tocId, action, pageNumbers, context)
    return { available: true }
  } catch (error) {
    return { available: false, reason: error instanceof Error ? error.message : String(error) }
  }
}
