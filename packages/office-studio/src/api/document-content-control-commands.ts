import type { DocumentModel } from '../domain/model'
import type { OfficeStudio } from './office-studio'
import type { DocumentImage } from '../features/document/model/document-model'
import { documentBlockHostForId, sameDocumentBlockHost, type DocumentAuxiliaryBlockHost } from '../features/document/model/document-block-host'
import { createDocumentContentControl, deleteRepeatingSectionItem, duplicateRepeatingSectionItem, getDocumentContentControl,
  removeDocumentContentControl, replacePictureContentControlImage, resetDocumentContentControlPlaceholder, selectDocumentContentControlItem,
  setDocumentContentControlChecked, setDocumentContentControlValue, updateDocumentContentControl,
  type ContentControlInput, type ContentControlPatch } from '../features/document/services/content-controls'

export interface DocumentContentControlExecutionContext {
  unitId: string
  expectedRevision?: number
  inputGeneration?: number
  /** A present undefined explicitly addresses the body. Omission permits API host resolution. */
  host?: DocumentAuxiliaryBlockHost
}

type ControlInput = DocumentContentControlExecutionContext & { controlId: string }

function assertExecutionContext(api: OfficeStudio, input: DocumentContentControlExecutionContext, controlId?: string): DocumentModel {
  const model = api.getDocumentModel(input.unitId)
  if (!model) throw new Error(`Document not found: ${input.unitId}`)
  const revision = api.getDocumentRevision(input.unitId), session = api.getDocumentInputState(input.unitId)
  if (input.expectedRevision !== undefined && (!Number.isSafeInteger(input.expectedRevision) || input.expectedRevision < 0 || input.expectedRevision !== revision)) {
    throw new Error('内容控件所属文档已更新，请重新选择。')
  }
  if (Object.hasOwn(input, 'inputGeneration') && (!Number.isSafeInteger(input.inputGeneration) || !session?.valid
    || session.modelRevision !== revision || session.generation !== input.inputGeneration)) throw new Error('内容控件输入会话已更改，请重新选择。')
  if (controlId !== undefined && Object.hasOwn(input, 'host')) {
    const control = getDocumentContentControl(model, controlId)
    if (!control || !sameDocumentBlockHost(documentBlockHostForId(model, control.blockId), input.host)) throw new Error('内容控件已删除或已切换内容区。')
    if (Object.hasOwn(input, 'inputGeneration') && (session?.location.kind !== 'body' || !sameDocumentBlockHost(session.location.host, input.host))) {
      throw new Error('内容控件命令与当前输入内容区不一致。')
    }
  }
  return model
}

/** All UI, shortcuts and public control facades share the same guarded history transaction. */
export function registerDocumentContentControlCommands(api: OfficeStudio): void {
  const register = <P extends ControlInput>(id: string, label: string, mutate: (model: DocumentModel, input: P) => void) => {
    api.registerCommand<P, boolean>(id, input => {
      assertExecutionContext(api, input, input.controlId)
      return api.commitDocument(input.unitId, label, model => mutate(model, input))
    })
  }
  api.registerCommand<DocumentContentControlExecutionContext & { control: ContentControlInput }, boolean>('doc.command.add-content-control', input => {
    assertExecutionContext(api, input)
    const session = api.getDocumentInputState(input.unitId)
    const host = Object.hasOwn(input, 'host') ? input.host : session?.valid && session.modelRevision === api.getDocumentRevision(input.unitId)
      && session.location.kind === 'body' && (session.location.selection.anchor.cellBlockId ?? session.location.selection.anchor.blockId) === input.control.blockId
      ? session.location.host : undefined
    return api.commitDocument(input.unitId, 'Add content control', model => {
      const created = createDocumentContentControl(model, input.control, input.control.id, host)
      if (created.type === 'repeatingSection') createDocumentContentControl(model, { type: 'repeatingSectionItem', blockId: created.blockId,
        startOffset: created.startOffset, endOffset: created.endOffset, placement: created.placement, parentId: created.id }, undefined, host)
    })
  })
  register<ControlInput & { patch: ContentControlPatch }>('doc.command.update-content-control', 'Update content control', (model, input) => updateDocumentContentControl(model, input.controlId, input.patch))
  register<ControlInput & { value: string }>('doc.command.set-content-control-value', 'Set content control value', (model, input) => setDocumentContentControlValue(model, input.controlId, input.value))
  register<ControlInput & { checked: boolean }>('doc.command.set-content-control-checked', 'Set content control checkbox', (model, input) => setDocumentContentControlChecked(model, input.controlId, input.checked))
  register<ControlInput>('doc.command.reset-content-control-placeholder', 'Reset content control placeholder', (model, input) => resetDocumentContentControlPlaceholder(model, input.controlId))
  register<ControlInput & { image: DocumentImage }>('doc.command.replace-picture-content-control-image', 'Replace picture content control image', (model, input) => replacePictureContentControlImage(model, input.controlId, input.image))
  register<ControlInput>('doc.command.duplicate-repeating-section-item', 'Duplicate repeating-section item', (model, input) => { duplicateRepeatingSectionItem(model, input.controlId) })
  register<ControlInput>('doc.command.delete-repeating-section-item', 'Delete repeating-section item', (model, input) => deleteRepeatingSectionItem(model, input.controlId))
  register<ControlInput & { value: string }>('doc.command.select-content-control-item', 'Select content control item', (model, input) => selectDocumentContentControlItem(model, input.controlId, input.value))
  register<ControlInput>('doc.command.remove-content-control', 'Remove content control', (model, input) => removeDocumentContentControl(model, input.controlId))
}
