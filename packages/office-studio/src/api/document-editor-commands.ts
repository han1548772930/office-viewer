import type { DocumentModel } from '../domain/model.ts'
import type { DocumentFieldCacheState } from '../features/document/services/document-field-open-refresh.ts'
import { editCodeBlockText } from '../features/document/code-blocks/code-blocks.ts'
import { executeDocumentCommand, type DocumentCommand } from '../features/document/commands/document-commands.ts'
import { clampPoint, runText, runsForTarget, sameTarget, textTargets, type DocumentPoint, type DocumentSelection } from '../features/document/model/document-model.ts'
import { assertDocumentPermission, permissionTargetsForSelection, type DocumentPermissionContext } from '../features/document/permissions/document-permissions.ts'
import { assertDocumentTableViewSelectionShape, cloneDocumentSelection, resolveDocumentTableViewSelection } from '../features/document/model/document-table-selection.ts'
import type { DocumentAuxiliaryBlockHost } from '../features/document/model/document-block-host.ts'
import { applyDocumentBlockHostModel, documentModelForBlockHost, validateDocumentBlockHost } from '../features/document/services/document-block-host.ts'
import { cloneModel } from '../domain/serialization.ts'
import { permissionTargetsForChangedBlocks } from '../features/document/permissions/document-permissions.ts'

export interface DocumentEditorCommandRequest {
  unitId: string
  command: DocumentCommand
  selection: DocumentSelection
  /** Optional facade guard for content changes made by BeforeCommandExecute listeners. */
  expectedRevision?: number
  /** Reject a same-ID reopen or history reset even when the model revision matches. */
  expectedHistoryGeneration?: number
  inputGeneration?: number
}
export interface DocumentEditorCommandResult {
  changed: boolean
  selection: DocumentSelection
  inputChanged?: boolean
  /** False when navigation or another edit has superseded the returned selection. */
  selectionCurrent?: boolean
  modelRevision?: number
  /** Selection coordinates belong to this host when present. */
  host?: DocumentAuxiliaryBlockHost
  /** The committed on-open cache, captured before observers can close or replace
   * the document. Hosts use selectionCurrent before applying it to a view. */
  fieldCache?: DocumentFieldCacheState
}

export const DOCUMENT_EDITOR_COMMAND_LABELS: Record<Exclude<DocumentCommand['id'], 'doc.command.set-document-permission-rules'>, string> = {
  'doc.command.insert-text': '输入文本',
  'doc.command.delete': '删除内容',
  'doc.command.format-text': '设置字符格式',
  'doc.command.apply-format-painter': '应用格式刷',
  'doc.command.clear-formatting': '清除格式',
  'doc.command.format-paragraph': '设置段落格式',
  'doc.command.change-paragraph-indent': '调整段落缩进',
  'doc.command.set-paragraph-named-style': '设置段落样式',
  'doc.command.set-block-type': '设置段落类型',
  'doc.command.change-list-type': '设置列表样式',
  'doc.command.change-list-nesting-level': '调整列表级别',
  'doc.command.restart-list-numbering': '重新开始编号',
  'doc.command.continue-list-numbering': '继续编号',
  'doc.command.toggle-checklist-item': '切换清单状态',
  'doc.command.insert-page-break': '插入分页符',
  'doc.command.insert-column-break': '插入分栏符',
  'doc.command.insert-section-break': '插入分节符',
  'doc.command.insert-table': '插入表格',
  'doc.command.insert-image': '插入图片',
  'doc.command.insert-divider': '插入分隔线',
  'doc.command.paste-blocks': '粘贴文档内容',
  'doc.command.table-tab': '跳转表格单元格',
  'doc.command.set-section-header-footer': '设置页眉页脚',
  'doc.command.set-section-setup': '设置节格式',
  'doc.command.refresh-header-footer-fields': '更新页眉页脚字段',
  'doc.command.update-selected-fields': '更新选中字段',
  'doc.command.insert-page-field': '插入页面域',
  'doc.command.update-table-of-contents': '更新目录',
  'doc.command.delete-table-of-contents': '删除目录',
  'doc.command.table-insert-row': '插入表格行',
  'doc.command.table-insert-column': '插入表格列',
  'doc.command.table-delete-rows': '删除表格行',
  'doc.command.table-delete-columns': '删除表格列',
  'doc.command.table-delete': '删除表格',
  'doc.command.table-merge-cells': '合并表格单元格',
  'doc.command.table-unmerge-cells': '拆分表格单元格',
  'doc.command.table-move-rows': '移动表格行',
  'doc.command.table-move-columns': '移动表格列',
  'doc.command.table-move': '移动表格',
  'doc.command.table-resize-row': '设置表格行高',
  'doc.command.table-resize-column': '设置表格列宽',
  'doc.command.table-distribute-rows': '平均分布表格行',
  'doc.command.table-distribute-columns': '平均分布表格列',
  'doc.command.table-set-background': '设置表格背景',
  'doc.command.table-set-border': '设置表格边框',
  'doc.command.table-set-border-color': '设置表格边框颜色',
  'doc.command.table-set-border-width': '设置表格边框宽度',
  'doc.command.table-set-border-style': '设置表格边框样式',
  'doc.command.table-set-vertical-align': '设置单元格垂直对齐',
  'doc.command.table-set-text-align': '设置表格文字对齐',
  'doc.command.table-set-align': '设置整表对齐',
  'doc.command.table-set-cell-margin': '设置表格单元格边距',
  'doc.command.table-set-cell-text': '设置表格单元格文本',
  'doc.command.table-set-header-row-count': '设置表格表头行数',
  'doc.command.table-insert-title-row': '插入表格标题行',
  'doc.command.table-set-column-type': '设置表格列类型',
  'doc.command.table-sort': '排序表格',
  'doc.command.table-set-style': '设置表格样式',
  'doc.command.table-set-range-text': '编辑表格选区内容',
}

export function assertDocumentEditorCommandRequest(request: DocumentEditorCommandRequest, commandId: string): void {
  if (!request || typeof request !== 'object' || typeof request.unitId !== 'string' || !request.unitId
    || !request.command || typeof request.command !== 'object' || request.command.id !== commandId
    || !request.selection || typeof request.selection !== 'object') throw new TypeError('文档命令参数无效。')
  assertPoint(request.selection.anchor)
  assertPoint(request.selection.focus)
  if (request.command.host) validateDocumentBlockHost(request.command.host)
  if (request.selection.tableSelection !== undefined) assertDocumentTableViewSelectionShape(request.selection.tableSelection)
  if (request.expectedRevision !== undefined && (!Number.isSafeInteger(request.expectedRevision) || request.expectedRevision < 0)
    || request.expectedHistoryGeneration !== undefined && (!Number.isSafeInteger(request.expectedHistoryGeneration) || request.expectedHistoryGeneration < 1)
    || request.inputGeneration !== undefined && (!Number.isSafeInteger(request.inputGeneration) || request.inputGeneration < 1)) {
    throw new TypeError('文档输入版本无效。')
  }
}

function assertPoint(point: DocumentPoint): void {
  if (!point || typeof point !== 'object' || typeof point.blockId !== 'string' || !point.blockId
    || !Number.isSafeInteger(point.offset) || point.offset < 0
    || point.row !== undefined && (!Number.isSafeInteger(point.row) || point.row < 0)
    || point.col !== undefined && (!Number.isSafeInteger(point.col) || point.col < 0 || point.row === undefined)
    || point.cellBlockId !== undefined && typeof point.cellBlockId !== 'string'
    || point.pageAffinity !== undefined && (!Number.isSafeInteger(point.pageAffinity) || point.pageAffinity < 0)) throw new TypeError('文档选区坐标无效。')
}

function modelPoint(model: DocumentModel, point: DocumentPoint): DocumentPoint {
  const normalized = clampPoint(model, point)
  if (normalized.blockId !== point.blockId || point.row !== undefined && normalized.row !== point.row
    || point.col !== undefined && normalized.col !== point.col
    || point.cellBlockId !== undefined && normalized.cellBlockId !== point.cellBlockId
    || normalized.offset !== point.offset) throw new RangeError('文档选区指向的内容或文字位置已不存在。')
  return normalized
}

export function resolveDocumentEditorSelection(model: DocumentModel, selection: DocumentSelection, host?: DocumentAuxiliaryBlockHost): DocumentSelection {
  if (host) model = documentModelForBlockHost(model, host)
  if (!selection || typeof selection !== 'object') throw new TypeError('文档选区无效。')
  assertPoint(selection.anchor)
  assertPoint(selection.focus)
  return { anchor: modelPoint(model, selection.anchor), focus: modelPoint(model, selection.focus),
    ...(selection.tableSelection ? { tableSelection: resolveDocumentTableViewSelection(model, selection.tableSelection) } : {}) }
}

/** Both UI and facade commands edit a transaction draft with the same logical selection. */
export function applyDocumentEditorCommand(model: DocumentModel, request: DocumentEditorCommandRequest, permissionContext: DocumentPermissionContext): DocumentSelection {
  const host = request.command.host, scoped = host ? documentModelForBlockHost(cloneModel(model), host, true) : model
  const selection = resolveDocumentEditorSelection(scoped, request.selection)
  let codeSelection: DocumentSelection | undefined
  if (!selection.tableSelection && (request.command.id === 'doc.command.insert-text' || request.command.id === 'doc.command.delete')) {
    assertDocumentPermission(model.permissionRules ?? [], permissionContext, 'edit', permissionTargetsForSelection(scoped, selection))
    codeSelection = request.command.id === 'doc.command.insert-text'
      ? editCodeBlockText(scoped, selection, { text: request.command.text, marks: request.command.marks })
      : editCodeBlockText(scoped, selection, { deleteDirection: request.command.direction })
  }
  if (codeSelection && host) {
    const candidate = cloneModel(model)
    applyDocumentBlockHostModel(candidate, host, scoped)
    assertDocumentPermission(model.permissionRules ?? [], permissionContext, 'edit', permissionTargetsForChangedBlocks(model, candidate))
    Object.assign(model, candidate)
  }
  const nextSelection = codeSelection ?? executeDocumentCommand(model, selection, request.command, permissionContext)
  model.wordCount = textTargets(model).reduce((count, target) => count + runText(runsForTarget(model, target)).replace(/\s/g, '').length, 0)
  const tableMutation = request.command.id.startsWith('doc.command.table-') && request.command.id !== 'doc.command.table-tab'
    || !!selection.tableSelection && ['doc.command.insert-text', 'doc.command.delete'].includes(request.command.id)
  const affinity = (point: DocumentPoint): DocumentPoint => !tableMutation && point.pageAffinity === undefined && selection.focus.pageAffinity !== undefined && sameTarget(point, selection.focus)
    ? { ...point, pageAffinity: selection.focus.pageAffinity } : point
  return { ...cloneDocumentSelection(nextSelection), anchor: affinity(nextSelection.anchor), focus: affinity(nextSelection.focus) }
}
