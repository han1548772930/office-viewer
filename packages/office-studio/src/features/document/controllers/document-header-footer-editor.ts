import type { DocumentModel } from '../../../domain/model'
import type { DocumentStoryKey } from '../model/document-story'
import { documentBlockHostId } from '../model/document-block-host'
import { textTargets, runText, runsForTarget } from '../model/document-model'
import { documentTableDescendantIds } from '../model/document-table-selection'
import { documentModelForBlockHost, resolveDocumentBlockHost } from '../services/document-block-host'
import { permissionTargetsForBlock, permissionTargetsForSelection, resolveDocumentPermission,
  type DocumentPermissionContext, type DocumentPermissionTarget } from '../permissions/document-permissions'

export const DOCUMENT_HEADER_FOOTER_COMMANDS = {
  open: 'doc.command.open-header-footer-panel',
  close: 'doc.command.close-header-footer',
  ribbon: 'sidebar.operation.doc-header-footer-panel',
} as const
export const DOCUMENT_HEADER_FOOTER_LOCAL_COMMANDS = {
  open: 'doc.operation.open-header-footer',
  close: 'doc.operation.close-header-footer',
} as const

export type DocumentHeaderFooterEditArea = 'HEADER' | 'FOOTER'
export interface DocumentHeaderFooterRequestContext {
  unitId?: string
  expectedRevision?: number
  expectedHistoryGeneration?: number
  expectedInputGeneration?: number
}
export interface DocumentOpenHeaderFooterRequest extends DocumentHeaderFooterRequestContext { editArea?: DocumentHeaderFooterEditArea }
export interface DocumentCloseHeaderFooterRequest extends DocumentHeaderFooterRequestContext {}
export interface DocumentHeaderFooterRibbonRequest extends DocumentHeaderFooterRequestContext { value: 'open' | 'close' }

export type DocumentHeaderFooterRequest = DocumentHeaderFooterRequestContext & (
  | { action: 'open'; editArea: DocumentHeaderFooterEditArea }
  | { action: 'close' }
  | { action: 'ribbon'; visible: boolean }
)
/** Runtime source and version guards never enter the document snapshot. */
export type DocumentHeaderFooterEditorTarget = DocumentHeaderFooterRequest & {
  unitId: string
  source: DocumentModel
  revision: number
  historyGeneration: number
  inputGeneration?: number
}

export function isDocumentHeaderFooterCommand(commandId: string): boolean {
  return Object.values(DOCUMENT_HEADER_FOOTER_COMMANDS).some(id => id === commandId)
    || Object.values(DOCUMENT_HEADER_FOOTER_LOCAL_COMMANDS).some(id => id === commandId)
}

/** The fixed edit-area values are strings, distinct from local canvas area names. */
export function resolveDocumentHeaderFooterRequest(commandId: string, raw?: unknown): DocumentHeaderFooterRequest {
  const action = commandId === DOCUMENT_HEADER_FOOTER_COMMANDS.open || commandId === DOCUMENT_HEADER_FOOTER_LOCAL_COMMANDS.open ? 'open'
    : commandId === DOCUMENT_HEADER_FOOTER_COMMANDS.close || commandId === DOCUMENT_HEADER_FOOTER_LOCAL_COMMANDS.close ? 'close'
      : commandId === DOCUMENT_HEADER_FOOTER_COMMANDS.ribbon ? 'ribbon' : undefined
  if (!action) throw new TypeError('页眉页脚命令无效。')
  const fields = ['unitId', 'expectedRevision', 'expectedHistoryGeneration', 'expectedInputGeneration',
    ...(action === 'open' ? ['editArea'] : action === 'ribbon' ? ['value'] : [])]
  if (raw !== undefined && (!raw || typeof raw !== 'object' || Array.isArray(raw)
    || Object.keys(raw).some(key => !fields.includes(key)))) throw new TypeError('页眉页脚参数无效。')
  const params: object = raw ?? {}
  const unitId: unknown = Reflect.get(params, 'unitId')
  if (unitId !== undefined && (typeof unitId !== 'string' || !unitId.trim())) throw new TypeError('页眉页脚的文档 ID 无效。')
  const context: DocumentHeaderFooterRequestContext = typeof unitId === 'string' ? { unitId } : {}
  for (const key of ['expectedRevision', 'expectedHistoryGeneration', 'expectedInputGeneration'] as const) {
    const value: unknown = Reflect.get(params, key)
    if (value === undefined) continue
    if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 0) throw new TypeError('页眉页脚的版本参数无效。')
    context[key] = value
  }
  if (action === 'close') return { ...context, action }
  if (action === 'open') {
    const editArea: unknown = Reflect.get(params, 'editArea')
    if (editArea !== undefined && editArea !== 'HEADER' && editArea !== 'FOOTER') throw new TypeError('页眉页脚的编辑区域无效。')
    return { ...context, action, editArea: editArea ?? 'HEADER' }
  }
  const value: unknown = Reflect.get(params, 'value')
  if (value !== 'open' && value !== 'close') throw new TypeError('页眉页脚 Ribbon 的显隐参数无效。')
  return { ...context, action, visible: value === 'open' }
}

/** Reading a linked story also checks the real source, without reading unrelated body paragraphs. */
export function documentHeaderFooterStoryViewState(model: DocumentModel, sectionId: string, story: DocumentStoryKey,
  permission: DocumentPermissionContext): { available: boolean; reason?: string } {
  const host = { kind: 'story' as const, sectionId, story }, view = resolveDocumentBlockHost(model, host)
  const targets = new Map<string, DocumentPermissionTarget>()
  const add = (target: DocumentPermissionTarget) => targets.set(`${target.scope}:${target.id}`, target)
  for (const scope of [host, view.sourceHost]) {
    if (scope.kind !== 'story') continue
    add({ scope: 'section', id: scope.sectionId })
    add({ scope: 'entity', id: documentBlockHostId(scope) })
    add({ scope: 'entity', id: `${scope.sectionId}:${scope.story.toLowerCase().includes('header') ? 'header' : 'footer'}` })
    const scoped = documentModelForBlockHost(model, scope), text = textTargets(scoped), first = text[0], last = text.at(-1)
    for (const id of documentTableDescendantIds(scoped.blocks)) for (const target of permissionTargetsForBlock(scoped, id)) add(target)
    if (first && last) for (const target of permissionTargetsForSelection(scoped, {
      anchor: { ...first, offset: 0 }, focus: { ...last, offset: runText(runsForTarget(scoped, last)).length },
    })) add(target)
  }
  const available = resolveDocumentPermission(model.permissionRules ?? [], permission, 'view', [...targets.values()]).allowed
  return available ? { available } : { available, reason: '当前用户没有查看此页眉或页脚的权限。' }
}
