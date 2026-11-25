import type { OfficeStudio } from './office-studio'
import { FDocument } from './facade'
import { insertDocumentLatexAtRange, type DocumentLatexCreateOptions, type DocumentLatexRange } from './document-latex'
import { documentSegmentIdForHost } from './document-segments'
import { DOCUMENT_LATEX_TEMPLATE_LOCAL_OPERATION, DOCUMENT_LATEX_TEMPLATE_OPERATION } from '../features/document/latex/latex-templates'
import { requestDocumentLatexTemplate } from './document-latex-template-editors'
import { DOCUMENT_LATEX_POPUP_OPERATIONS, assertDocumentLatexPopupRequest, type DocumentLatexPopupOperationRequest } from '../features/document/controllers/document-latex-popup'
import { resolveDocumentLatexPopupTarget } from './document-latex-popup'
import { requestDocumentLatexPopupOpen, requestDocumentLatexPopupClose } from './document-latex-popup-editors'

export const DOCUMENT_LATEX_PUBLIC_COMMANDS = {
  insert: 'docs-latex.command.insert-formula', update: 'docs-latex.command.update-formula',
  remove: 'docs-latex.command.remove-formula', replaceWithText: 'docs-latex.command.replace-formula-with-text',
} as const

interface LatexPublicCommandBase {
  unitId?: string; segmentId?: string; expectedRevision?: number; expectedHistoryGeneration?: number
}
export interface DocumentLatexPublicInsertRequest extends LatexPublicCommandBase, DocumentLatexCreateOptions {
  latex: string; startOffset?: number; endOffset?: number
  selection?: { startOffset: number; endOffset: number; segmentId?: string }
}
export interface DocumentLatexPublicUpdateRequest extends LatexPublicCommandBase {
  rangeId: string; latex: string; properties?: DocumentLatexCreateOptions['properties']
}
export interface DocumentLatexPublicRemoveRequest extends LatexPublicCommandBase { rangeId: string }
export interface DocumentLatexPublicReplaceRequest extends DocumentLatexPublicRemoveRequest { text: string }
export interface DocumentLatexTemplateOperationRequest extends LatexPublicCommandBase { templateId: string }

function commandDocument(api: OfficeStudio, input: LatexPublicCommandBase): FDocument {
  if (!input || typeof input !== 'object' || Array.isArray(input)) throw new TypeError('公式命令参数无效。')
  if (input.segmentId !== undefined && typeof input.segmentId !== 'string'
    || [input.expectedRevision, input.expectedHistoryGeneration].some(value => value !== undefined && (!Number.isSafeInteger(value) || value < 0))) throw new TypeError('公式命令的内容区或版本参数无效。')
  const unitId = input.unitId ?? api.activeDocumentId
  if (typeof unitId !== 'string' || !unitId.trim() || !api.getDocumentModel(unitId)) throw new RangeError('公式命令目标文档不存在。')
  if (input.expectedRevision !== undefined && input.expectedRevision !== api.getDocumentRevision(unitId)
    || input.expectedHistoryGeneration !== undefined && input.expectedHistoryGeneration !== api.editor.history.captureSavepoint(unitId).generation) throw new RangeError('公式命令对应的文档版本或历史已变化。')
  return new FDocument(api, unitId)
}

/** Public aliases dispatch the registered editor command, sharing one history. */
export function registerDocumentLatexCommands(api: OfficeStudio): void {
  for (const id of [DOCUMENT_LATEX_POPUP_OPERATIONS.open, DOCUMENT_LATEX_POPUP_OPERATIONS.localOpen]) {
    api.registerCommand<DocumentLatexPopupOperationRequest | undefined, boolean>(id, input => {
      const request = input ?? {}
      assertDocumentLatexPopupRequest(request)
      const document = commandDocument(api, request)
      return requestDocumentLatexPopupOpen(api, resolveDocumentLatexPopupTarget(document, request))
    })
  }
  for (const id of [DOCUMENT_LATEX_POPUP_OPERATIONS.close, DOCUMENT_LATEX_POPUP_OPERATIONS.localClose]) {
    api.registerCommand<DocumentLatexPopupOperationRequest | undefined, boolean>(id, input => {
      const request = input ?? {}
      assertDocumentLatexPopupRequest(request)
      return requestDocumentLatexPopupClose(api, commandDocument(api, request).id)
    })
  }
  for (const id of [DOCUMENT_LATEX_TEMPLATE_OPERATION, DOCUMENT_LATEX_TEMPLATE_LOCAL_OPERATION]) {
    api.registerCommand<DocumentLatexTemplateOperationRequest, boolean>(id, input =>
      requestDocumentLatexTemplate(api, commandDocument(api, input).id, input.templateId))
  }
  api.registerCommand<DocumentLatexPublicInsertRequest, DocumentLatexRange | null>(DOCUMENT_LATEX_PUBLIC_COMMANDS.insert, input => {
    const document = commandDocument(api, input), segmentId = input.segmentId ?? input.selection?.segmentId ?? ''
    if (input.selection !== undefined && (!input.selection || typeof input.selection !== 'object' || Array.isArray(input.selection)
      || !Number.isSafeInteger(input.selection.startOffset) || !Number.isSafeInteger(input.selection.endOffset)
      || input.selection.segmentId !== undefined && typeof input.selection.segmentId !== 'string')) throw new TypeError('公式命令的选区无效。')
    const start = input.startOffset ?? input.selection?.startOffset, end = input.endOffset ?? input.selection?.endOffset ?? start
    if (input.selection && (input.startOffset !== undefined && input.startOffset !== input.selection.startOffset
      || input.endOffset !== undefined && input.endOffset !== input.selection.endOffset
      || input.segmentId !== undefined && input.selection.segmentId !== undefined && input.segmentId !== input.selection.segmentId)) throw new RangeError('公式命令的显式范围与选区不一致。')
    if (start === undefined && end !== undefined) throw new RangeError('公式命令的结束位置缺少起始位置。')
    if (start === undefined && (input.segmentId !== undefined || input.selection?.segmentId !== undefined)) {
      const state = document.getInputState()
      const activeSegment = state?.valid && state.location.kind === 'body'
        ? state.location.host ? documentSegmentIdForHost(document.getModel(), state.location.host) : '' : undefined
      if (activeSegment !== segmentId) throw new RangeError('公式命令的指定内容区与当前输入选区不一致。')
    }
    const formula = start === undefined ? document.insertLatexAtSelection(input.latex, input)
      : insertDocumentLatexAtRange(document, start, end!, input.latex, { properties: input.properties, mode: input.mode, segmentId })
    return formula?.getRange() ?? null
  })
  api.registerCommand<DocumentLatexPublicUpdateRequest, boolean>(DOCUMENT_LATEX_PUBLIC_COMMANDS.update, input =>
    commandDocument(api, input).getLatexFormula(input.rangeId, input.segmentId)?.update({ latex: input.latex, properties: input.properties }) ?? false)
  api.registerCommand<DocumentLatexPublicRemoveRequest, boolean>(DOCUMENT_LATEX_PUBLIC_COMMANDS.remove, input =>
    commandDocument(api, input).getLatexFormula(input.rangeId, input.segmentId)?.remove() ?? false)
  api.registerCommand<DocumentLatexPublicReplaceRequest, boolean>(DOCUMENT_LATEX_PUBLIC_COMMANDS.replaceWithText, input =>
    commandDocument(api, input).getLatexFormula(input.rangeId, input.segmentId)?.replaceWithText(input.text) ?? false)
}
