import type { FDocument } from './facade'
import { cloneDocumentSelection } from '../features/document/model/document-table-selection'
import { sameDocumentBlockHost } from '../features/document/model/document-block-host'
import { documentLatexSourceForSelection, selectedLatexFormula } from '../features/document/latex/latex-formula'
import { documentLatexSelectionState } from '../features/document/commands/document-latex-commands'
import { assertDocumentLatexPopupRequest, type DocumentLatexPopupOperationRequest, type DocumentLatexPopupTarget } from '../features/document/controllers/document-latex-popup'
import { documentSegmentIdForHost } from './document-segments'
import { documentSelectionAtPublicBodyRange } from './document-public-body'
import { assertDocumentLatexAtomicRange } from './document-latex'

/** Resolve source offsets and hosts before asking a UI owner to open a draft. */
export function resolveDocumentLatexPopupTarget(document: FDocument, input: DocumentLatexPopupOperationRequest): DocumentLatexPopupTarget {
  assertDocumentLatexPopupRequest(input)
  if (input.selection?.rangeType === 'RECT') throw new RangeError('公式编辑器需要文字选区。')
  const source = document.getModel(), state = document.getInputState()
  const modelRevision = document.getModelRevision(), historyGeneration = document.getHistoryGeneration()
  if (input.unitId !== undefined && input.unitId !== document.id) throw new RangeError('公式编辑器的目标文档不一致。')
  if (input.expectedRevision !== undefined && input.expectedRevision !== modelRevision
    || input.expectedHistoryGeneration !== undefined && input.expectedHistoryGeneration !== historyGeneration) throw new RangeError('公式编辑器的内容或历史已变化。')
  if (!state?.valid || state.location.kind !== 'body' || state.modelRevision !== modelRevision) throw new RangeError('公式编辑器需要有效的文档文字输入会话。')
  if (input.expectedInputGeneration !== undefined && input.expectedInputGeneration !== state.generation) throw new RangeError('公式编辑器的输入会话已变化。')
  const activeSegment = state.location.host ? documentSegmentIdForHost(source, state.location.host) : ''
  const segmentId = input.segmentId ?? input.selection?.segmentId ?? activeSegment
  const host = document.getSegmentHost(segmentId), scoped = document.getSegmentModel(segmentId)
  let start = input.selection?.startOffset ?? input.startIndex
  let end = input.selection?.endOffset ?? (input.endIndex === undefined ? undefined : input.endIndex + 1)
  if (input.rangeId !== undefined) {
    const range = document.getLatexFormula(input.rangeId, segmentId)?.getRange()
    if (!range) throw new RangeError('公式编辑目标已被删除或替换。')
    if (start !== undefined && (start !== range.startOffset || end !== range.endOffset)) throw new RangeError('公式身份与显式范围不一致。')
    start = range.startOffset; end = range.endOffset
  }
  let selection = cloneDocumentSelection(state.location.selection)
  if (start !== undefined && end !== undefined) {
    assertDocumentLatexAtomicRange(document.getPublicBodyIndex(segmentId), start, end)
    const resolved = documentSelectionAtPublicBodyRange(source, start, end, host)
    if (!resolved) throw new RangeError('公式编辑目标缺少可编辑的文字范围。')
    selection = resolved
  } else if (!sameDocumentBlockHost(host, state.location.host)) throw new RangeError('指定内容区缺少显式公式或文字范围。')
  if (start !== undefined && input.selection?.direction === 'backward') selection = { anchor: selection.focus, focus: selection.anchor }
  const segmentPage = input.segmentPage ?? input.selection?.segmentPage
  if (segmentPage !== undefined) {
    selection.anchor.pageAffinity = segmentPage; selection.focus.pageAffinity = segmentPage
  }
  const permission = documentLatexSelectionState(scoped, selection, { documentId: document.id, userId: document.getCurrentUserId() })
  if (!permission.available) throw new RangeError(permission.reason ?? '当前不能编辑公式。')
  const formula = selectedLatexFormula(scoped, selection)
  const initialLatex = input.initialLatex ?? formula?.formula.latex ?? documentLatexSourceForSelection(scoped, selection)
  if (initialLatex.length > 65_536) throw new RangeError('公式源文本超过编辑器预算。')
  return { source, unitId: document.id, selection, ...(host ? { host: { ...host } } : {}),
    ...(formula ? { rangeId: formula.rangeId } : {}), initialLatex, modelRevision, historyGeneration, inputGeneration: state.generation }
}
