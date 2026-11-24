import type { DocumentModel } from '../domain/model'
import type { OfficeStudio } from './office-studio'
import { cloneModel } from '../domain/serialization'
import type { DocumentImage } from '../features/document/model/document-model'
import { documentBlockHostForId, sameDocumentBlockHost, type DocumentAuxiliaryBlockHost } from '../features/document/model/document-block-host'
import { createDocumentContentControl, getDocumentContentControl,
  removeDocumentContentControl, replacePictureContentControlImage, resetDocumentContentControlPlaceholder, selectDocumentContentControlItem,
  setDocumentContentControlChecked, setDocumentContentControlValue, updateDocumentContentControl,
  type ContentControlInput, type ContentControlPatch } from '../features/document/services/content-controls'
import { synchronizeDocumentContentControlBindings } from '../features/document/services/content-control-bindings'
import { assertDocumentPermission, permissionTargetsForBlock, permissionTargetsForChangedBlocks } from '../features/document/permissions/document-permissions'
import { runText, runsForTarget, textTargets } from '../features/document/model/document-model'
import { documentModelForBlockHost } from '../features/document/services/document-block-host'
import { documentContentControlTableRowBlockIds, resolveDocumentContentControlTableRows } from '../features/document/model/content-control-table-rows'
import { documentContentControlOwnedBlockIds, resolveDocumentContentControlBlocks } from '../features/document/model/content-control-block-range'
import { documentContentControlPermissionTargets } from '../features/document/permissions/content-control-permissions'
import { prepareDocumentRepeatingContentControlMutation } from '../features/document/services/repeating-content-control-mutation'
import { prepareRepeatingBindingStructure } from '../features/document/services/document-custom-xml-repeating'
import { prepareDocumentCustomXmlCandidate } from '../features/document/services/document-custom-xml-candidate'

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
    if (!control || !sameDocumentBlockHost(documentBlockHostForId(model, control.blockRange?.blockIds[0] ?? control.tableRowRange?.tableId ?? control.blockId), input.host)
      || !sameDocumentBlockHost(documentBlockHostForId(model, control.blockId), input.host)) throw new Error('内容控件已删除或已切换内容区。')
    if (Object.hasOwn(input, 'inputGeneration') && (session?.location.kind !== 'body' || !sameDocumentBlockHost(session.location.host, input.host))) {
      throw new Error('内容控件命令与当前输入内容区不一致。')
    }
  }
  return model
}

/** All UI, shortcuts and public control facades share the same guarded history transaction. */
export function registerDocumentContentControlCommands(api: OfficeStudio): void {
  const register = <P extends ControlInput>(id: string, label: string, mutate: (model: DocumentModel, input: P) => void,
    scalar?: (input: P) => string | undefined) => {
    api.registerCommand<P, boolean>(id, input => {
      assertExecutionContext(api, input, input.controlId)
      return api.commitDocument(input.unitId, label, model => {
        const control = getDocumentContentControl(model, input.controlId)
        if (!control) throw new RangeError('内容控件已删除。')
        const permissionContext = { documentId: input.unitId, userId: api.currentUserId }
        const host = documentBlockHostForId(model, control.blockRange?.blockIds[0] ?? control.tableRowRange?.tableId ?? control.blockId)
        const scoped = host ? documentModelForBlockHost(model, host, true) : model
        assertDocumentPermission(model.permissionRules ?? [], permissionContext, 'edit', documentContentControlPermissionTargets(scoped, control, model))
        const before = cloneModel(model)
        mutate(model, input)
        if (scalar) synchronizeDocumentContentControlBindings(before, model, { sourceControlId: input.controlId,
          sourceValue: scalar?.(input), permissionContext })
        assertDocumentPermission(before.permissionRules ?? [], permissionContext, 'edit', permissionTargetsForChangedBlocks(before, model))
        if (JSON.stringify(before) !== JSON.stringify(model)) {
          model.wordCount = textTargets(model).reduce((count, target) => count + runText(runsForTarget(model, target)).replace(/\s/g, '').length, 0)
        }
      })
    })
  }
  api.registerCommand<DocumentContentControlExecutionContext & { control: ContentControlInput }, boolean>('doc.command.add-content-control', input => {
    const source = assertExecutionContext(api, input)
    const session = api.getDocumentInputState(input.unitId)
    const host = Object.hasOwn(input, 'host') ? input.host : session?.valid && session.modelRevision === api.getDocumentRevision(input.unitId)
      && session.location.kind === 'body' && (session.location.selection.anchor.cellBlockId ?? session.location.selection.anchor.blockId) === input.control.blockId
      ? session.location.host : documentBlockHostForId(source, input.control.blockRange?.blockIds[0] ?? input.control.tableRowRange?.tableId ?? input.control.blockId)
    if (!sameDocumentBlockHost(documentBlockHostForId(source, input.control.blockRange?.blockIds[0] ?? input.control.tableRowRange?.tableId ?? input.control.blockId), host)
      || !sameDocumentBlockHost(documentBlockHostForId(source, input.control.blockId), host)) throw new Error('内容控件插入范围已切换内容区。')
    if (Object.hasOwn(input, 'inputGeneration') && (session?.location.kind !== 'body'
      || !sameDocumentBlockHost(session.location.host, host))) throw new Error('内容控件命令与当前输入内容区不一致。')
    return api.commitDocument(input.unitId, 'Add content control', model => {
      const before = cloneModel(model), scoped = host ? documentModelForBlockHost(model, host, true) : model
      const permissionContext = { documentId: input.unitId, userId: api.currentUserId }
      const blockIds = input.control.tableRowRange ? [input.control.tableRowRange.tableId,
        ...documentContentControlTableRowBlockIds(resolveDocumentContentControlTableRows(scoped, input.control.tableRowRange))]
        : input.control.blockRange ? [...documentContentControlOwnedBlockIds(resolveDocumentContentControlBlocks(scoped, input.control.blockRange))] : [input.control.blockId]
      assertDocumentPermission(model.permissionRules ?? [], permissionContext, 'edit', blockIds.flatMap(id => permissionTargetsForBlock(model, id)))
      const created = createDocumentContentControl(model, input.control, input.control.id, host)
      if (created.type === 'repeatingSection') createDocumentContentControl(model, { type: 'repeatingSectionItem', blockId: created.blockId,
        startOffset: created.startOffset, endOffset: created.endOffset, placement: created.placement, parentId: created.id,
        ...(created.tableRowRange ? { tableRowRange: cloneModel(created.tableRowRange) } : {}),
        ...(created.blockRange ? { blockRange: cloneModel(created.blockRange) } : {}) }, undefined, host)
      assertDocumentPermission(before.permissionRules ?? [], permissionContext, 'edit', permissionTargetsForChangedBlocks(before, model))
      model.wordCount = textTargets(model).reduce((count, target) => count + runText(runsForTarget(model, target)).replace(/\s/g, '').length, 0)
    })
  })
  register<ControlInput & { patch: ContentControlPatch }>('doc.command.update-content-control', 'Update content control', (model, input) => updateDocumentContentControl(model, input.controlId, input.patch), () => undefined)
  register<ControlInput & { value: string }>('doc.command.set-content-control-value', 'Set content control value', (model, input) => setDocumentContentControlValue(model, input.controlId, input.value), input => input.value)
  register<ControlInput & { checked: boolean }>('doc.command.set-content-control-checked', 'Set content control checkbox', (model, input) => setDocumentContentControlChecked(model, input.controlId, input.checked), input => input.checked ? 'true' : 'false')
  register<ControlInput>('doc.command.reset-content-control-placeholder', 'Reset content control placeholder', (model, input) => resetDocumentContentControlPlaceholder(model, input.controlId), () => '')
  register<ControlInput & { image: DocumentImage }>('doc.command.replace-picture-content-control-image', 'Replace picture content control image', (model, input) => replacePictureContentControlImage(model, input.controlId, input.image))
  const registerRepeatingItem = (id: string, action: 'duplicate' | 'delete') => api.registerCommand<ControlInput, boolean>(id, input => {
    assertExecutionContext(api, input, input.controlId)
    return api.commitDocument(input.unitId, action === 'duplicate' ? 'Duplicate repeating-section item' : 'Delete repeating-section item', model => {
      const before = cloneModel(model)
      const permissionContext = { documentId: input.unitId, userId: api.currentUserId }
      const plan = prepareDocumentRepeatingContentControlMutation(before, input.controlId, action, permissionContext, input.host)
      const binding = prepareRepeatingBindingStructure(before, plan.model, { sectionId: plan.sectionId, sourceItemId: input.controlId,
        action, resultItemId: plan.controlId, identityMap: plan.identityMap, permissionContext })
      binding?.commit()
      prepareDocumentCustomXmlCandidate(before, plan.model, permissionContext, true)
      assertDocumentPermission(before.permissionRules ?? [], permissionContext, 'edit', permissionTargetsForChangedBlocks(before, plan.model))
      plan.model.wordCount = textTargets(plan.model).reduce((count, target) => count + runText(runsForTarget(plan.model, target)).replace(/\s/g, '').length, 0)
      Object.assign(model, plan.model)
    })
  })
  registerRepeatingItem('doc.command.duplicate-repeating-section-item', 'duplicate')
  registerRepeatingItem('doc.command.delete-repeating-section-item', 'delete')
  register<ControlInput & { value: string }>('doc.command.select-content-control-item', 'Select content control item', (model, input) => selectDocumentContentControlItem(model, input.controlId, input.value), input => input.value)
  register<ControlInput>('doc.command.remove-content-control', 'Remove content control', (model, input) => removeDocumentContentControl(model, input.controlId))
}
