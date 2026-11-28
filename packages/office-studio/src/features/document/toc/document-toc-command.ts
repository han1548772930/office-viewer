import type { DocumentBlock, DocumentModel } from '../../../domain/model'
import { blockAtPoint, collapsedSelection, selectionCollapsed, type DocumentSelection } from '../model/document-model'
import { assertDocumentPermission, permissionTargetsForChangedBlocks, permissionTargetsForSelection, type DocumentPermissionContext } from '../permissions/document-permissions'
import { assertDocumentTextSelectionEditable } from '../services/document-text-formatting'
import { documentFieldLocked, documentTextFieldParagraphs, documentTextFieldRangeIntersects, documentTextFieldRanges } from '../model/document-text-fields'
import { buildDocumentTocBlocks, collectDocumentTocHeadings, DEFAULT_DOCUMENT_TOC_OPTIONS, documentTocActionUnavailableReason, documentTocIdAtBlock,
  documentTocIdForSelection, prepareDocumentTocDeletion, prepareDocumentTocFullUpdate, prepareDocumentTocPageNumberUpdate, supportsDocumentTableOfContents,
  type DocumentTocAction, type DocumentTocOptions, type DocumentTocPageNumbers } from './document-toc'
import { documentTocMutationEntityIds, type DocumentTocMutationPlan } from './document-toc-edit'

export interface DocumentTocActionState { available: boolean; reason?: string }

export interface DocumentTocInsertionRequest extends Partial<DocumentTocOptions> {
  /** Automatic UI insertion supplies a localized title; custom/headless insertion can omit it. */
  title?: string
  /** Derived logical labels from the matching layout, never guessed physical indices. */
  pageNumbers?: DocumentTocPageNumbers
}

export function assertDocumentTocInsertionSelection(model: DocumentModel, selection: DocumentSelection, context: DocumentPermissionContext): void {
  if (!supportsDocumentTableOfContents(model)) throw new RangeError('当前文档模式尚未接入正文目录。')
  if (selection.tableSelection || selection.anchor.row != null || selection.focus.row != null || !selectionCollapsed(selection)) {
    throw new RangeError('插入目录需要正文中的折叠文字选区。')
  }
  assertDocumentTextSelectionEditable(model, selection)
  const block = blockAtPoint(model, selection.focus)
  if (!block || !model.blocks.includes(block) || documentTocIdAtBlock(block)) throw new RangeError('目录不能插入现有目录或嵌套内容区。')
  const paragraph = documentTextFieldParagraphs(model).find(candidate => candidate.block.id === block.id)
  const offset = paragraph && paragraph.absoluteStart + selection.focus.offset
  if (offset !== undefined && documentTextFieldRanges(model).some(range => documentFieldLocked(range.field)
    && documentTextFieldRangeIntersects(range, offset, offset))) throw new RangeError('目录不能插入锁定字段的结果内。')
  assertDocumentPermission(model.permissionRules ?? [], context, 'edit', permissionTargetsForSelection(model, selection))
}

export function documentTocInsertionState(model: DocumentModel, selection: DocumentSelection, context: DocumentPermissionContext): DocumentTocActionState {
  try {
    assertDocumentTocInsertionSelection(model, selection, context)
    return { available: true }
  } catch (error) { return { available: false, reason: error instanceof Error ? error.message : String(error) } }
}

/** UI and public commands use the same typed options, source permissions and builder. */
export function prepareDocumentTocInsertion(model: DocumentModel, selection: DocumentSelection, request: DocumentTocInsertionRequest,
  context: DocumentPermissionContext): DocumentBlock[] {
  assertDocumentTocInsertionSelection(model, selection, context)
  if (request.levels !== undefined && (typeof request.levels !== 'number' || !Number.isFinite(request.levels))
    || request.minimumLevel !== undefined && (typeof request.minimumLevel !== 'number' || !Number.isFinite(request.minimumLevel))
    || request.showPageNumbers !== undefined && typeof request.showPageNumbers !== 'boolean'
    || request.rightAlignPageNumbers !== undefined && typeof request.rightAlignPageNumbers !== 'boolean'
    || request.tabLeader !== undefined && !['none', 'dots', 'dashes', 'underline'].includes(request.tabLeader)
    || request.format !== undefined && !['fromTemplate', 'classic', 'modern', 'simple'].includes(request.format)
    || request.title !== undefined && typeof request.title !== 'string') throw new TypeError('目录插入选项无效。')
  const options: DocumentTocOptions = { ...DEFAULT_DOCUMENT_TOC_OPTIONS,
    ...(request.levels === undefined ? {} : { levels: Math.min(9, Math.max(1, Math.trunc(request.levels))) }),
    ...(request.minimumLevel === undefined ? {} : { minimumLevel: Math.min(9, Math.max(1, Math.trunc(request.minimumLevel))) }),
    ...(request.showPageNumbers === undefined ? {} : { showPageNumbers: request.showPageNumbers }),
    ...(request.rightAlignPageNumbers === undefined ? {} : { rightAlignPageNumbers: request.rightAlignPageNumbers }),
    ...(request.tabLeader === undefined ? {} : { tabLeader: request.tabLeader }),
    ...(request.format === undefined ? {} : { format: request.format }) }
  options.levels = Math.max(options.minimumLevel ?? 1, options.levels)
  const labels = request.pageNumbers === undefined ? {} : request.pageNumbers
  if (!labels || typeof labels !== 'object' || Array.isArray(labels)
    || Object.values(labels).some(value => value !== undefined && (!Number.isSafeInteger(value) || value < 1))) throw new TypeError('目录布局页码无效。')
  const headings = collectDocumentTocHeadings(model, options.levels, options.minimumLevel)
  if (!headings.length) throw new RangeError('文档中没有可用于目录的标题。')
  assertDocumentPermission(model.permissionRules ?? [], context, 'view', headings.flatMap(heading => permissionTargetsForSelection(model,
    collapsedSelection({ blockId: heading.id, offset: 0 }))))
  const pageNumbers = Object.fromEntries(headings.map(heading => [heading.id, Object.hasOwn(labels, heading.id) ? labels[heading.id] : undefined]))
  const title = request.title?.trim() ?? ''
  return buildDocumentTocBlocks(model, options, pageNumbers, undefined, [], title, title.length > 0)
}

/** Menu, toolbar and commands prepare the same candidate and check the same ACLs. */
export function prepareDocumentTocAction(
  model: DocumentModel,
  selection: DocumentSelection,
  tocId: string,
  action: DocumentTocAction,
  pageNumbers: DocumentTocPageNumbers,
  context: DocumentPermissionContext,
  timestamp = Date.now(),
  preserveInputSelection = false,
): DocumentTocMutationPlan {
  if (!['pageNumbersOnly', 'all', 'delete'].includes(action)) throw new TypeError('目录操作无效。')
  if (!supportsDocumentTableOfContents(model)) throw new RangeError('当前文档模式尚未接入正文目录。')
  if (!preserveInputSelection && documentTocIdForSelection(model, selection) !== tocId) throw new RangeError('目录命令要求选区位于目标目录内。')
  const blocks = model.blocks.filter(block => documentTocIdAtBlock(block) === tocId)
  assertDocumentPermission(model.permissionRules ?? [], context, 'edit', blocks.flatMap(block => permissionTargetsForSelection(model,
    collapsedSelection({ blockId: block.id, offset: 0 }))))
  if (action === 'all') {
    const options = blocks.map(block => block.metadata.documentToc as { options?: DocumentTocOptions } | undefined).find(metadata => metadata?.options)?.options
    if (options) assertDocumentPermission(model.permissionRules ?? [], context, 'view',
      collectDocumentTocHeadings(model, options.levels, options.minimumLevel).flatMap(heading => permissionTargetsForSelection(model,
        collapsedSelection({ blockId: heading.id, offset: 0 }))))
  }
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
