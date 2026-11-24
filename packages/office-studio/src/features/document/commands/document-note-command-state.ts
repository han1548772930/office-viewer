import type { DocumentModel } from '../../../domain/model'
import type { DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import { orderedSelection, collapsedSelection, type DocumentSelection } from '../model/document-model'
import { assertDocumentNoteInsertion, documentSupportsNotes } from '../notes/document-note-editing'
import type { DocumentNoteKind } from '../notes/document-note-types'
import { documentNotePermissionTargets } from '../permissions/document-note-permissions'
import { permissionTargetsForSelection, resolveDocumentPermission, type DocumentPermissionContext } from '../permissions/document-permissions'
import { documentSectionIds } from '../services/section-header-footer'

export interface DocumentNoteActionState { available: boolean; reason?: string }
export interface DocumentNoteActionContext { permission: DocumentPermissionContext; host?: DocumentAuxiliaryBlockHost; disabled?: boolean }

export function documentNoteInsertState(model: DocumentModel, selection: DocumentSelection, context: DocumentNoteActionContext): DocumentNoteActionState {
  if (context.disabled) return { available: false, reason: '请先完成当前输入法输入。' }
  if (context.host) return { available: false, reason: '请在正文或单元格文字中插入脚注或尾注引用。' }
  try {
    assertDocumentNoteInsertion(model, selection)
    const point = orderedSelection(model, selection).end
    const permission = resolveDocumentPermission(model.permissionRules ?? [], context.permission, 'edit', permissionTargetsForSelection(model, collapsedSelection(point)))
    return permission.allowed ? { available: true } : { available: false, reason: '当前正文位置没有编辑权限。' }
  } catch (error) {
    if (error instanceof Error) return { available: false, reason: error.message }
    throw error
  }
}

export function documentNoteEditState(model: DocumentModel, kind: DocumentNoteKind, id: string, context: DocumentNoteActionContext): DocumentNoteActionState {
  if (context.disabled) return { available: false, reason: '请先完成当前输入法输入。' }
  if (!(kind === 'footnote' ? model.footnotes : model.endnotes)?.[id]) return { available: false, reason: '此注释已不存在。' }
  const permission = resolveDocumentPermission(model.permissionRules ?? [], context.permission, 'edit', documentNotePermissionTargets(model, kind, id))
  return permission.allowed ? { available: true } : { available: false, reason: '当前注释或引用内容区没有编辑权限。' }
}

export function documentNoteSettingsState(model: DocumentModel, context: DocumentNoteActionContext, sectionId?: string): DocumentNoteActionState {
  if (context.disabled) return { available: false, reason: '请先完成当前输入法输入。' }
  if (!documentSupportsNotes(model)) return { available: false, reason: '脚注和尾注设置适用于传统分页文档。' }
  const permission = resolveDocumentPermission(model.permissionRules ?? [], context.permission, 'edit',
    (sectionId === undefined ? documentSectionIds(model) : [sectionId]).map(id => ({ scope: 'section' as const, id })))
  return permission.allowed ? { available: true } : { available: false, reason: '当前注释设置作用域没有编辑权限。' }
}
