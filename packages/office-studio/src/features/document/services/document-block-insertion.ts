import type { DocumentBlock, DocumentModel } from '../../../domain/model'
import {
  blockAtPoint, clampPoint, collapsedSelection, createParagraph, findDocumentBlockContainer, runText, setTargetRuns,
  sliceRuns, textTargets, type DocumentPoint, type DocumentSelection,
} from '../model/document-model'
import { documentTextFieldParagraphs, documentTextFieldRanges, transformDocumentTextFieldsForBlockRangeEdit } from '../model/document-text-fields'
import { isDocumentTextBoundary } from '../model/document-story'
import { getDocumentContentControls, transformContentControlsForBlockEdit } from './content-controls'
import { assertDocumentTextSelectionEditable } from './document-text-formatting'
import { transformDocumentCommentAnchorsForBodyRangeEdit } from '../comments/document-comment-anchors'

/** An offset is resolved once to a stable paragraph, before a table is moved. */
export function documentPointAtClampedBodyOffset(model: DocumentModel, requestedOffset: number): DocumentPoint {
  if (!Number.isSafeInteger(requestedOffset)) throw new RangeError('正文插入偏移必须为安全整数')
  const paragraphs = documentTextFieldParagraphs(model), last = paragraphs.at(-1)
  if (!last) throw new RangeError('正文缺少可编辑段落')
  const offset = Math.max(0, Math.min(last.absoluteEnd, requestedOffset))
  const paragraph = paragraphs.find(candidate => offset <= candidate.absoluteEnd) ?? last
  const local = Math.max(0, Math.min(paragraph.text.length, offset - paragraph.absoluteStart))
  if (!isDocumentTextBoundary(paragraph.text, local)) throw new RangeError('正文插入偏移不能切开 UTF-16 字符')
  const target = textTargets(model).find(candidate => (candidate.cellBlockId ?? candidate.blockId) === paragraph.block.id)
  if (!target) throw new RangeError('正文插入目标已不存在')
  return { ...target, offset: local }
}

export interface DocumentBlockInsertionResult {
  selection: DocumentSelection
  transformedFieldIds: ReadonlySet<string>
}

/** Only a detached candidate calls this helper; paragraph text remains model-owned. */
export function insertDocumentBlocksAtPoint(model: DocumentModel, requested: DocumentPoint, blocks: DocumentBlock[], movingBlock?: DocumentBlock): DocumentBlockInsertionResult {
  const point = clampPoint(model, requested), selection = collapsedSelection(point)
  assertDocumentTextSelectionEditable(model, selection)
  const block = blockAtPoint(model, point), container = block && findDocumentBlockContainer(model, block.id)
  if (!block || !container) throw new RangeError('插入块的目标段落已不存在')
  const oldText = runText(block.runs), tail = createParagraph()
  tail.style = { ...block.style }
  tail.runs = sliceRuns(block.runs, point.offset)
  const fieldIds = new Set(documentTextFieldRanges(model).map(range => range.field.id))
  const controls: DocumentModel = { ...model, contentControls: getDocumentContentControls(model) }
  transformContentControlsForBlockEdit(controls, block.id, point.offset, point.offset, 0, {
    blockId: tail.id, firstBreakOffset: point.offset, lastParagraphOffset: point.offset,
  })
  transformDocumentTextFieldsForBlockRangeEdit([block], point.offset, point.offset, '\n', [block, tail], undefined, model)
  model.contentControls = controls.contentControls
  setTargetRuns(model, point, sliceRuns(block.runs, 0, point.offset))
  if (movingBlock) {
    const source = findDocumentBlockContainer(model, movingBlock.id)
    if (!source) throw new RangeError('移动块的源位置已不存在')
    source.splice(source.indexOf(movingBlock), 1)
    if (!documentTextFieldParagraphs({ blocks: source }).length) source.push(createParagraph())
  }
  container.splice(container.indexOf(block) + 1, 0, ...blocks, tail)
  // Map comments on the split paragraph to its prefix/suffix; comments in
  // moved table paragraphs keep their stable IDs and local offsets.
  transformDocumentCommentAnchorsForBodyRangeEdit(model,
    [{ block: { id: block.id }, absoluteStart: 0, absoluteEnd: oldText.length }],
    [
      { block, text: runText(block.runs), absoluteStart: 0, absoluteEnd: point.offset },
      { block: tail, text: runText(tail.runs), absoluteStart: point.offset + 1, absoluteEnd: oldText.length + 1 },
    ], point.offset, point.offset, '\n')
  for (const range of documentTextFieldRanges(model)) fieldIds.add(range.field.id)
  return { selection: collapsedSelection(point.row === undefined ? { blockId: tail.id, offset: 0, pageAffinity: point.pageAffinity }
    : { ...point, cellBlockId: tail.id, offset: 0 }), transformedFieldIds: fieldIds }
}
