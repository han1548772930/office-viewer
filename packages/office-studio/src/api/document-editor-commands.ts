import type { DocumentModel } from '../domain/model.ts'
import { editCodeBlockText } from '../features/document/code-blocks/code-blocks.ts'
import { executeDocumentCommand, type DocumentCommand } from '../features/document/commands/document-commands.ts'
import { clampPoint, runText, runsForTarget, textTargets, type DocumentPoint, type DocumentSelection } from '../features/document/model/document-model.ts'
import { assertDocumentPermission, permissionTargetsForSelection, type DocumentPermissionContext } from '../features/document/permissions/document-permissions.ts'

export interface DocumentEditorCommandRequest {
  unitId: string
  command: DocumentCommand
  selection: DocumentSelection
  /** Optional facade guard for content changes made by BeforeCommandExecute listeners. */
  expectedRevision?: number
  inputGeneration?: number
}
export interface DocumentEditorCommandResult {
  changed: boolean
  selection: DocumentSelection
  inputChanged?: boolean
  /** False when navigation or another edit has superseded the returned selection. */
  selectionCurrent?: boolean
  modelRevision?: number
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
  'doc.command.insert-page-field': '插入页面域',
  'doc.command.update-table-of-contents': '更新目录',
  'doc.command.delete-table-of-contents': '删除目录',
  'doc.command.table-insert-row': '插入表格行',
  'doc.command.table-insert-column': '插入表格列',
  'doc.command.table-delete-rows': '删除表格行',
  'doc.command.table-delete-columns': '删除表格列',
  'doc.command.table-delete': '删除表格',
}

export function assertDocumentEditorCommandRequest(request: DocumentEditorCommandRequest, commandId: string): void {
  if (!request || typeof request !== 'object' || typeof request.unitId !== 'string' || !request.unitId
    || !request.command || typeof request.command !== 'object' || request.command.id !== commandId
    || !request.selection || typeof request.selection !== 'object') throw new TypeError('文档命令参数无效。')
  assertPoint(request.selection.anchor)
  assertPoint(request.selection.focus)
  if (request.expectedRevision !== undefined && (!Number.isSafeInteger(request.expectedRevision) || request.expectedRevision < 0)
    || request.inputGeneration !== undefined && (!Number.isSafeInteger(request.inputGeneration) || request.inputGeneration < 1)) {
    throw new TypeError('文档输入版本无效。')
  }
}

function assertPoint(point: DocumentPoint): void {
  if (!point || typeof point !== 'object' || typeof point.blockId !== 'string' || !point.blockId
    || !Number.isSafeInteger(point.offset) || point.offset < 0
    || point.row !== undefined && (!Number.isSafeInteger(point.row) || point.row < 0)
    || point.col !== undefined && (!Number.isSafeInteger(point.col) || point.col < 0 || point.row === undefined)
    || point.cellBlockId !== undefined && typeof point.cellBlockId !== 'string') throw new TypeError('文档选区坐标无效。')
}

function modelPoint(model: DocumentModel, point: DocumentPoint): DocumentPoint {
  const normalized = clampPoint(model, point)
  if (normalized.blockId !== point.blockId || point.row !== undefined && normalized.row !== point.row
    || point.col !== undefined && normalized.col !== point.col
    || point.cellBlockId !== undefined && normalized.cellBlockId !== point.cellBlockId
    || normalized.offset !== point.offset) throw new RangeError('文档选区指向的内容或文字位置已不存在。')
  return normalized
}

export function resolveDocumentEditorSelection(model: DocumentModel, selection: DocumentSelection): DocumentSelection {
  if (!selection || typeof selection !== 'object') throw new TypeError('文档选区无效。')
  assertPoint(selection.anchor)
  assertPoint(selection.focus)
  return { anchor: modelPoint(model, selection.anchor), focus: modelPoint(model, selection.focus) }
}

/** Both UI and facade commands edit a transaction draft with the same logical selection. */
export function applyDocumentEditorCommand(model: DocumentModel, request: DocumentEditorCommandRequest, permissionContext: DocumentPermissionContext): DocumentSelection {
  const selection = resolveDocumentEditorSelection(model, request.selection)
  let codeSelection: DocumentSelection | undefined
  if (request.command.id === 'doc.command.insert-text' || request.command.id === 'doc.command.delete') {
    assertDocumentPermission(model.permissionRules ?? [], permissionContext, 'edit', permissionTargetsForSelection(model, selection))
    codeSelection = request.command.id === 'doc.command.insert-text'
      ? editCodeBlockText(model, selection, { text: request.command.text, marks: request.command.marks })
      : editCodeBlockText(model, selection, { deleteDirection: request.command.direction })
  }
  const nextSelection = codeSelection ?? executeDocumentCommand(model, selection, request.command, permissionContext)
  model.wordCount = textTargets(model).reduce((count, target) => count + runText(runsForTarget(model, target)).replace(/\s/g, '').length, 0)
  return nextSelection
}
