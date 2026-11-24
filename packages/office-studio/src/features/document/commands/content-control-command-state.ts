import type { DocumentContentControl, DocumentModel } from '../../../domain/model'
import type { DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import { documentModelForBlockHost } from '../services/document-block-host'
import { assertDocumentContentControlValueEditable, assertDocumentRepeatingSectionItemEditable } from '../services/content-controls'
import { resolveDocumentPermission, type DocumentPermissionContext } from '../permissions/document-permissions'
import { documentContentControlPermissionTargets } from '../permissions/content-control-permissions'
import type { DocumentContentControlRepeatingItemCommand, DocumentContentControlValueCommand } from './content-control-command-types'

export interface ContentControlCommandState { available: boolean; reason?: string }

export interface ContentControlCommandOptions {
  host?: DocumentAuxiliaryBlockHost
  permission?: DocumentPermissionContext
  disabled?: boolean
}

export function contentControlValueCommandState(model: DocumentModel, command: DocumentContentControlValueCommand,
  options: ContentControlCommandOptions = {}): ContentControlCommandState {
  return contentControlEditCommandState(model, command.params.controlId, options, control => {
    if (command.id === 'doc.command.set-content-control-checked' && control.type !== 'checkbox'
      || command.id === 'doc.command.select-content-control-item' && !['dropDownList', 'comboBox'].includes(control.type)
      || command.id === 'doc.command.set-content-control-value' && !['text', 'richText', 'date', 'dropDownList', 'comboBox'].includes(control.type)) return { available: false, reason: '此控件不支持该值操作。' }
    if (command.id === 'doc.command.reset-content-control-placeholder') {
      if (!['text', 'richText', 'date', 'dropDownList', 'comboBox', 'picture'].includes(control.type)
        || !control.placeholder) return { available: false, reason: '此控件没有可恢复的占位文字。' }
      if (control.showingPlaceholder) return { available: false, reason: '此控件已经显示占位文字。' }
    }
    return undefined
  })
}

export function contentControlPictureCommandState(model: DocumentModel, controlId: string,
  options: ContentControlCommandOptions = {}): ContentControlCommandState {
  return contentControlEditCommandState(model, controlId, options, control => control.type === 'picture'
    ? undefined : { available: false, reason: '此控件不支持图片替换。' })
}

export function contentControlRepeatingItemCommandState(model: DocumentModel, command: DocumentContentControlRepeatingItemCommand,
  options: ContentControlCommandOptions = {}): ContentControlCommandState {
  if (options.disabled) return { available: false, reason: '当前内容区无法编辑。' }
  try {
    const scoped = options.host ? documentModelForBlockHost(model, options.host, true) : model
    const control = scoped.contentControls?.find(control => control.id === command.params.controlId)
    if (!control) return { available: false, reason: '内容控件已删除或已切换内容区。' }
    const permission = options.permission ?? { documentId: model.id ?? '', userId: 'local-user' }
    const targets = documentContentControlPermissionTargets(scoped, control, model)
    const decision = resolveDocumentPermission(model.permissionRules ?? [], permission, 'edit', targets)
    if (!decision.allowed) return { available: false, reason: '没有编辑此重复项目的权限。' }
    if (command.id === 'doc.command.duplicate-repeating-section-item'
      && !resolveDocumentPermission(model.permissionRules ?? [], permission, 'copy', targets).allowed) return { available: false, reason: '没有复制此重复项目的权限。' }
    assertDocumentRepeatingSectionItemEditable(scoped, control.id, command.id === 'doc.command.duplicate-repeating-section-item' ? 'duplicate' : 'delete')
    return { available: true }
  } catch (reason) { return { available: false, reason: reason instanceof Error ? reason.message : String(reason) } }
}

function contentControlEditCommandState(model: DocumentModel, controlId: string, options: ContentControlCommandOptions,
  validateType: (control: DocumentContentControl) => ContentControlCommandState | undefined): ContentControlCommandState {
  if (options.disabled) return { available: false, reason: '当前内容区无法编辑。' }
  try {
    const scoped = options.host ? documentModelForBlockHost(model, options.host, true) : model
    const control = scoped.contentControls?.find(control => control.id === controlId)
    if (!control) return { available: false, reason: '内容控件已删除或已切换内容区。' }
    const unsupported = validateType(control)
    if (unsupported) return unsupported
    const permission = options.permission ?? { documentId: model.id ?? '', userId: 'local-user' }
    const decision = resolveDocumentPermission(model.permissionRules ?? [], permission, 'edit', documentContentControlPermissionTargets(scoped, control, model))
    if (!decision.allowed) return { available: false, reason: '没有编辑此内容控件的权限。' }
    assertDocumentContentControlValueEditable(scoped, control.id)
    return { available: true }
  } catch (reason) { return { available: false, reason: reason instanceof Error ? reason.message : String(reason) } }
}
