import type { DocumentModel } from '../../../domain/model'
import { orderedSelection, runText, runsForTarget, sameTarget, textTargets, type DocumentSelection } from '../model/document-model'
import { isDocumentTextBoundary } from '../model/document-story'
import { assertDocumentPermission, permissionTargetsForSelection, type DocumentPermissionContext } from '../permissions/document-permissions'
import { assertTextRangeEditable } from './content-controls'
import { flattenDocumentBody } from './document-facade'

/** Formatting and insertion share the same text boundaries and content-lock decision. */
export function assertDocumentTextSelectionEditable(model: DocumentModel, selection: DocumentSelection): void {
  const targets = textTargets(model)
  for (const point of [selection.anchor, selection.focus]) {
    if (!targets.some(target => sameTarget(target, point)) || !isDocumentTextBoundary(runText(runsForTarget(model, point)), point.offset)) {
      throw new RangeError('字符选区指向的内容或 UTF-16 位置无效。')
    }
  }
  if (!model.contentControls?.length) return
  const { start, end } = orderedSelection(model, selection), body = flattenDocumentBody(model)
  const first = body.segments.find(segment => sameTarget(segment, start)), last = body.segments.find(segment => sameTarget(segment, end))
  if (!first || !last) throw new RangeError('字符选区指向的内容已不存在。')
  assertTextRangeEditable(model, first.startOffset + start.offset, last.startOffset + end.offset)
}

export function documentTextFormattingState(model: DocumentModel, selection: DocumentSelection, context: DocumentPermissionContext): { available: boolean; reason?: string } {
  try {
    assertDocumentPermission(model.permissionRules ?? [], context, 'edit', permissionTargetsForSelection(model, selection))
    assertDocumentTextSelectionEditable(model, selection)
    return { available: true }
  } catch (error) { return { available: false, reason: error instanceof Error ? error.message : String(error) } }
}
