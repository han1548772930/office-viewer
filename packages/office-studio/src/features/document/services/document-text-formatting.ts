import type { DocumentModel } from '../../../domain/model'
import { findDocumentBlock, orderedSelection, runText, runsForTarget, sameTarget, textTargets, type DocumentSelection } from '../model/document-model'
import { isDocumentTextBoundary } from '../model/document-story'
import { assertDocumentPermission, permissionTargetsForSelection, type DocumentPermissionContext } from '../permissions/document-permissions'
import { assertTextRangeEditable } from './content-controls'
import { flattenDocumentBody } from './document-facade'
import { documentTableSelectedTextTargets } from '../model/document-table-selection'
import { DOCUMENT_PARAGRAPH_PAGINATION_FIELDS } from '../model/document-paragraph-pagination'

/** Formatting and insertion share the same text boundaries and content-lock decision. */
export function assertDocumentTextSelectionEditable(model: DocumentModel, selection: DocumentSelection): void {
  if (selection.tableSelection) {
    const selected = documentTableSelectedTextTargets(model, selection.tableSelection)
    if (!selected.length) throw new RangeError('选中的单元格没有可设置字符格式的文本')
    if (model.contentControls?.length) {
      const body = flattenDocumentBody(model)
      for (const target of selected) {
        const segment = body.segments.find(candidate => sameTarget(candidate, target))
        if (!segment) throw new RangeError('表格选区的文本内容已不存在')
        assertTextRangeEditable(model, segment.startOffset, segment.endOffset)
      }
    }
    return
  }
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

/** A generated quote rule changes the entire paragraph, even if the insertion
 * point lies outside a locked span. Check both sides before publishing. */
export function assertDocumentParagraphPaginationChangesEditable(before: DocumentModel, after: DocumentModel): void {
  const originals = new Map(textTargets(before).map(target => [target.cellBlockId ?? target.blockId, target]))
  for (const target of textTargets(after)) {
    const id = target.cellBlockId ?? target.blockId, block = findDocumentBlock(after, id)!, originalTarget = originals.get(id)
    const original = originalTarget && findDocumentBlock(before, id)
    if (original && DOCUMENT_PARAGRAPH_PAGINATION_FIELDS.every(key => original.style[key] === block.style[key])) continue
    if (!original && DOCUMENT_PARAGRAPH_PAGINATION_FIELDS.every(key => block.style[key] === undefined)) continue
    assertDocumentTextSelectionEditable(after, { anchor: { ...target, offset: 0 }, focus: { ...target, offset: runText(block.runs).length } })
    if (original && originalTarget) assertDocumentTextSelectionEditable(before,
      { anchor: { ...originalTarget, offset: 0 }, focus: { ...originalTarget, offset: runText(original.runs).length } })
  }
}
