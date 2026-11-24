import type { DocumentModel } from '../../../domain/model'
import type { DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import { documentModelForBlockHost } from '../services/document-block-host'
import { assertDocumentContentControlValueEditable } from '../services/content-controls'
import { permissionTargetsForBlock, resolveDocumentPermission, type DocumentPermissionContext } from '../permissions/document-permissions'
import type { DocumentContentControlValueCommand } from './content-control-command-types'

export interface ContentControlCommandState { available: boolean; reason?: string }

export function contentControlValueCommandState(model: DocumentModel, command: DocumentContentControlValueCommand,
  options: { host?: DocumentAuxiliaryBlockHost; permission?: DocumentPermissionContext; disabled?: boolean } = {}): ContentControlCommandState {
  if (options.disabled) return { available: false, reason: '当前内容区无法编辑。' }
  try {
    const scoped = options.host ? documentModelForBlockHost(model, options.host, true) : model
    const control = scoped.contentControls?.find(control => control.id === command.params.controlId)
    if (!control) return { available: false, reason: '内容控件已删除或已切换内容区。' }
    if (command.id === 'doc.command.set-content-control-checked' && control.type !== 'checkbox'
      || command.id === 'doc.command.select-content-control-item' && !['dropDownList', 'comboBox'].includes(control.type)
      || command.id === 'doc.command.set-content-control-value' && !['text', 'richText', 'date', 'dropDownList', 'comboBox'].includes(control.type)) return { available: false, reason: '此控件不支持该值操作。' }
    const permission = options.permission ?? { documentId: model.id ?? '', userId: 'local-user' }
    const decision = resolveDocumentPermission(model.permissionRules ?? [], permission, 'edit', permissionTargetsForBlock(scoped, control.blockId))
    if (!decision.allowed) return { available: false, reason: '没有编辑此内容控件的权限。' }
    assertDocumentContentControlValueEditable(scoped, control.id)
    return { available: true }
  } catch (reason) { return { available: false, reason: reason instanceof Error ? reason.message : String(reason) } }
}
