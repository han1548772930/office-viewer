import type { DocumentBlock, DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { transformDocumentCommentAnchorsForParagraphEdits, type DocumentCommentParagraphEdit } from '../comments/document-comment-anchors'
import { blockAtPoint, createParagraph, runText, textTargets, type DocumentPoint, type DocumentSelection } from '../model/document-model'
import { isDocumentTextBoundary } from '../model/document-story'
import { documentTextFieldParagraphs, documentTextFieldRanges, documentTextFields, setDocumentTextFields, type DocumentTextField, type DocumentTextFieldParagraph } from '../model/document-text-fields'
import { assertContentControlBlockRemovable, transformContentControlsForBlockEdit, validateDocumentContentControlRanges } from '../services/content-controls'

function fieldSource(field: DocumentTextField): Record<string, unknown> {
  const source = field.metadata?.source
  return source !== null && typeof source === 'object' && !Array.isArray(source) ? source as Record<string, unknown> : {}
}

/** Keep explicit roots and recover the nearest surviving explicit ancestor only. */
export function reconcileDocumentTocFieldParents(previous: DocumentModel, candidate: DocumentModel): void {
  const before = new Map(documentTextFieldRanges(previous).map(range => [range.field.id, range]))
  const ranges = documentTextFieldRanges(candidate), after = new Map(ranges.map(range => [range.field.id, range]))
  const parents = new Map(ranges.map(range => [range.field.id, fieldSource(range.field).parentFieldId]))
  for (const range of ranges) {
    let parent = parents.get(range.field.id)
    if (typeof parent !== 'string') continue
    const visited = new Set([range.field.id])
    while (typeof parent === 'string') {
      if (visited.has(parent)) throw new RangeError('目录变更后的字段父子关系存在循环。')
      visited.add(parent)
      const kept = after.get(parent)
      if (kept && kept.absoluteStart <= range.absoluteStart && kept.absoluteEnd >= range.absoluteEnd) break
      const ancestor = kept ?? before.get(parent), next = ancestor && fieldSource(ancestor.field).parentFieldId
      parent = typeof next === 'string' ? next : null
    }
    parents.set(range.field.id, parent)
  }
  // Equal and empty ranges are valid when the explicit graph is acyclic.
  const settled = new Set<string>()
  for (const id of parents.keys()) {
    const path = new Set<string>()
    let current: unknown = id
    while (typeof current === 'string' && !settled.has(current)) {
      if (path.has(current)) throw new RangeError('目录变更后的字段父子关系存在循环。')
      path.add(current); current = parents.get(current)
    }
    for (const visited of path) settled.add(visited)
  }
  const changed = new Map<string, DocumentTextField>()
  for (const range of ranges) {
    const parent = parents.get(range.field.id)
    if (parent === fieldSource(range.field).parentFieldId) continue
    changed.set(range.field.id, { ...range.field, metadata: { ...range.field.metadata,
      source: { ...fieldSource(range.field), parentFieldId: parent } } })
  }
  for (const paragraph of documentTextFieldParagraphs(candidate)) {
    const fields = documentTextFields(paragraph.block)
    if (fields.some(field => changed.has(field.id))) setDocumentTextFields(paragraph.block, fields.map(field => changed.get(field.id) ?? field))
  }
}

export interface DocumentTocMutationPlan {
  blocks: DocumentModel['blocks']
  contentControls?: DocumentModel['contentControls']
  comments?: DocumentModel['comments']
  selection?: DocumentSelection
  changed: boolean
  /** Includes valid unchanged results; locked/unsupported page ranges are excluded. */
  updatablePageFields?: number
}

/** Field/control definitions share entity ACLs with their other public commands. */
export function documentTocMutationEntityIds(before: DocumentModel, plan: DocumentTocMutationPlan): string[] {
  const after: DocumentModel = { ...before, blocks: plan.blocks, contentControls: plan.contentControls, comments: plan.comments }
  const previous = new Map<string, unknown>(), next = new Map<string, unknown>()
  for (const range of documentTextFieldRanges(before)) previous.set(range.field.id, { blockId: range.block.id, field: range.field })
  for (const range of documentTextFieldRanges(after)) next.set(range.field.id, { blockId: range.block.id, field: range.field })
  const changed = new Set([...new Set([...previous.keys(), ...next.keys()])].filter(id => JSON.stringify(previous.get(id)) !== JSON.stringify(next.get(id))))
  previous.clear(); next.clear()
  for (const control of before.contentControls ?? []) previous.set(control.id, control)
  for (const control of after.contentControls ?? []) next.set(control.id, control)
  for (const id of new Set([...previous.keys(), ...next.keys()])) if (JSON.stringify(previous.get(id)) !== JSON.stringify(next.get(id))) changed.add(id)
  return [...changed]
}

function textEdit(before: string, after: string, blockId: string): DocumentCommentParagraphEdit {
  let startOffset = 0, suffix = 0
  while (startOffset < before.length && startOffset < after.length && before[startOffset] === after[startOffset]) startOffset++
  while (startOffset > 0 && (!isDocumentTextBoundary(before, startOffset) || !isDocumentTextBoundary(after, startOffset))) startOffset--
  while (suffix < before.length - startOffset && suffix < after.length - startOffset
    && before[before.length - suffix - 1] === after[after.length - suffix - 1]) suffix++
  while (suffix > 0 && (!isDocumentTextBoundary(before, before.length - suffix) || !isDocumentTextBoundary(after, after.length - suffix))) suffix--
  return { blockId, startOffset, endOffset: before.length - suffix, insertedText: after.slice(startOffset, after.length - suffix) }
}

/** Heading text and PAGEREF results have separate anchors and formatting. */
function paragraphEdits(before: DocumentBlock, after: DocumentBlock): DocumentCommentParagraphEdit[] {
  const oldText = runText(before.runs), newText = runText(after.runs)
  const oldPages = documentTextFields(before).filter(field => field.field.toUpperCase() === 'PAGEREF' && !field.endBlockId)
  const oldPage = oldPages.length === 1 ? oldPages[0] : undefined
  const newPages = documentTextFields(after).filter(field => field.field.toUpperCase() === 'PAGEREF' && !field.endBlockId)
  const newPage = oldPage && newPages.length === 1 && newPages[0]!.id === oldPage.id ? newPages[0] : undefined
  if (!oldPage || !newPage) return [textEdit(oldText, newText, before.id)]
  const prefix = textEdit(oldText.slice(0, oldPage.startOffset), newText.slice(0, newPage.startOffset), before.id)
  const page = textEdit(oldText.slice(oldPage.startOffset, oldPage.endOffset), newText.slice(newPage.startOffset, newPage.endOffset), before.id)
  const suffix = textEdit(oldText.slice(oldPage.endOffset), newText.slice(newPage.endOffset), before.id)
  return [prefix, { ...page, startOffset: newPage.startOffset + page.startOffset, endOffset: newPage.startOffset + page.endOffset },
    { ...suffix, startOffset: newPage.endOffset + suffix.startOffset, endOffset: newPage.endOffset + suffix.endOffset }]
}

function fieldRangeAfterReplacement(from: number, to: number, start: number, end: number, length: number): { start: number; end: number } {
  const delta = length - (end - start)
  if (start === end) {
    if (start <= from) return { start: from + length, end: to + length }
    return { start: from, end: start < to ? to + length : to }
  }
  if (end <= from) return { start: from + delta, end: to + delta }
  if (start >= to) return { start: from, end: to }
  return { start: from >= start ? start : from, end: to >= end ? to + delta : start + length }
}

function paragraphAt(paragraphs: readonly DocumentTextFieldParagraph[], offset: number): DocumentTextFieldParagraph {
  const paragraph = paragraphs.find(current => offset >= current.absoluteStart && offset <= current.absoluteEnd)
  if (!paragraph) throw new RangeError('目录变更后的字段端点无法映射到段落。')
  return paragraph
}

function mapSelectionPoint(
  model: DocumentModel,
  point: DocumentPoint,
  edits: ReadonlyMap<string, readonly DocumentCommentParagraphEdit[]>,
  paragraphs: ReadonlyMap<string, DocumentTextFieldParagraph>,
): DocumentPoint {
  const changes = edits.get(point.cellBlockId ?? point.blockId)
  if (!changes?.length) return point
  const blockId = changes.at(-1)!.blockId
  const paragraph = paragraphs.get(blockId)
  const target = textTargets(model).find(current => (current.cellBlockId ?? current.blockId) === blockId)
  if (!paragraph || !target) throw new RangeError('目录变更后的选区目标不存在。')
  let offset = point.offset
  for (const edit of changes) offset = (edit.targetOffset ?? 0) + (offset <= edit.startOffset ? offset
    : offset >= edit.endOffset ? offset + edit.insertedText.length - (edit.endOffset - edit.startOffset)
      : edit.startOffset + Math.min(offset - edit.startOffset, edit.insertedText.length))
  if (!isDocumentTextBoundary(paragraph.text, offset)) offset--
  if (!isDocumentTextBoundary(paragraph.text, offset)) throw new RangeError('目录变更后的选区 UTF-16 位置无效。')
  return { ...target, offset }
}

/** Prepare the complete row/field/SDT/comment mutation without touching the source model. */
export function prepareDocumentTocBlockMutation(
  model: DocumentModel,
  sourceBlocks: readonly DocumentBlock[],
  replacementBlocks: readonly DocumentBlock[],
  replacedFieldIds: ReadonlySet<string>,
  selection?: DocumentSelection,
  timestamp = Date.now(),
): DocumentTocMutationPlan {
  if (typeof timestamp !== 'number' || !Number.isFinite(new Date(timestamp).getTime())) throw new TypeError('目录更新时间无效。')
  if (!sourceBlocks.length) throw new RangeError('目录变更缺少源段落。')
  const firstIndex = model.blocks.findIndex(block => block.id === sourceBlocks[0]!.id)
  if (firstIndex < 0 || sourceBlocks.some((block, index) => model.blocks[firstIndex + index]?.id !== block.id)) {
    throw new RangeError('目录段落已经不连续，请重新定位目录。')
  }
  if (selection) for (const point of [selection.anchor, selection.focus]) {
    const block = blockAtPoint(model, point)
    if (!block || !isDocumentTextBoundary(runText(block.runs), point.offset)) throw new RangeError('目录变更选区的 UTF-16 位置无效。')
  }
  const source = documentTextFieldParagraphs(model), ranges = documentTextFieldRanges(model)
  const sourceIds = new Set(sourceBlocks.map(block => block.id))
  const first = source.find(paragraph => paragraph.block.id === sourceBlocks[0]!.id)
  const last = source.find(paragraph => paragraph.block.id === sourceBlocks.at(-1)!.id)
  if (!first || !last || sourceBlocks.some(block => !source.some(paragraph => paragraph.block.id === block.id))) {
    throw new RangeError('目录只能替换连续的文本段落。')
  }
  const deleting = replacementBlocks.length === 0
  const previous = source.filter(paragraph => paragraph.absoluteEnd < first.absoluteStart).at(-1)
  const next = source.find(paragraph => paragraph.absoluteStart > last.absoluteEnd)
  let start = first.absoluteStart, end = last.absoluteEnd
  if (deleting && next) end = next.absoluteStart
  else if (deleting && previous) start = previous.absoluteEnd
  const candidate: DocumentModel = { ...model, blocks: cloneModel(model.blocks),
    contentControls: model.contentControls && cloneModel(model.contentControls), comments: model.comments && cloneModel(model.comments) }
  // Source endpoints remain authoritative until all ranges have been projected.
  for (const paragraph of documentTextFieldParagraphs(candidate)) if (paragraph.block.metadata.documentTextFields !== undefined) setDocumentTextFields(paragraph.block, [])
  const replacement = cloneModel([...replacementBlocks])
  candidate.blocks.splice(firstIndex, sourceBlocks.length, ...replacement)
  if (!documentTextFieldParagraphs(candidate).length) candidate.blocks.splice(firstIndex, 0, createParagraph())
  const destination = documentTextFieldParagraphs(candidate), byId = new Map(destination.map(paragraph => [paragraph.block.id, paragraph]))
  const resultStart = replacement.find(block => documentTextFields(block).some(field => field.field.toUpperCase() === 'TOC')) ?? replacement[0]
  const boundary = resultStart ? byId.get(resultStart.id) : next ? byId.get(next.block.id) : previous ? byId.get(previous.block.id) : destination[0]
  if (!boundary) throw new RangeError('目录变更缺少可编辑的目标段落。')
  const boundaryOffset = deleting && !next && previous ? boundary.text.length : 0
  const edits = new Map<string, DocumentCommentParagraphEdit[]>()
  for (const block of sourceBlocks) {
    const oldText = runText(block.runs), target = byId.get(block.id)
    if (target) {
      const changes = paragraphEdits(block, target.block)
      edits.set(block.id, changes)
      for (const edit of changes) if (edit.startOffset !== edit.endOffset || edit.insertedText.length) {
        transformContentControlsForBlockEdit(candidate, block.id, edit.startOffset, edit.endOffset, edit.insertedText.length, undefined, true)
      }
    } else {
      assertContentControlBlockRemovable(candidate, block.id)
      edits.set(block.id, [{ blockId: boundary.block.id, startOffset: 0, endOffset: oldText.length, insertedText: '', targetOffset: boundaryOffset }])
    }
  }
  validateDocumentContentControlRanges(candidate)
  transformDocumentCommentAnchorsForParagraphEdits(candidate, source, destination, edits, timestamp)
  const plans = new Map(destination.map(paragraph => [paragraph.block.id, [...documentTextFields(paragraph.block)] as DocumentTextField[]]))
  const text = replacement.map(block => runText(block.runs)).join('\n')
  for (const range of ranges) {
    const contained = range.absoluteStart >= start && range.absoluteEnd <= end
      && (deleting || range.absoluteStart > start || range.absoluteEnd < end || sourceIds.has(range.block.id) && range.absoluteStart === range.absoluteEnd)
    if (replacedFieldIds.has(range.field.id) || contained) continue
    const mapped = fieldRangeAfterReplacement(range.absoluteStart, range.absoluteEnd, start, end, text.length)
    const owner = paragraphAt(destination, mapped.start), tail = paragraphAt(destination, mapped.end)
    const { endBlockId: _endBlockId, ...field } = range.field
    plans.get(owner.block.id)!.push({ ...cloneModel(field), startOffset: mapped.start - owner.absoluteStart, endOffset: mapped.end - tail.absoluteStart,
      ...(owner.block.id === tail.block.id ? {} : { endBlockId: tail.block.id }) })
  }
  for (const paragraph of destination) setDocumentTextFields(paragraph.block, plans.get(paragraph.block.id)!, paragraph.text)
  reconcileDocumentTocFieldParents(model, candidate)
  documentTextFieldRanges(candidate)
  const nextSelection = selection && { anchor: mapSelectionPoint(candidate, selection.anchor, edits, byId), focus: mapSelectionPoint(candidate, selection.focus, edits, byId) }
  return { blocks: candidate.blocks, contentControls: candidate.contentControls, comments: candidate.comments,
    changed: JSON.stringify(candidate.blocks) !== JSON.stringify(model.blocks) || JSON.stringify(candidate.contentControls) !== JSON.stringify(model.contentControls)
      || JSON.stringify(candidate.comments) !== JSON.stringify(model.comments), ...(nextSelection ? { selection: nextSelection } : {}) }
}

/** Preserve unchanged objects while publishing an already validated structural candidate. */
export function applyDocumentTocMutation(model: DocumentModel, plan: DocumentTocMutationPlan): void {
  if (!plan.changed) return
  const current = new Map(model.blocks.map(block => [block.id, block]))
  const blocks = plan.blocks.map(next => {
    const block = current.get(next.id)
    if (!block) return next
    for (const key of Object.keys(block)) if (!Object.hasOwn(next, key)) Reflect.deleteProperty(block, key)
    for (const [key, value] of Object.entries(next)) if (JSON.stringify(Reflect.get(block, key)) !== JSON.stringify(value)) Reflect.set(block, key, value)
    return block
  })
  if (blocks.length !== model.blocks.length || blocks.some((block, index) => block !== model.blocks[index])) model.blocks = blocks
  if (JSON.stringify(model.contentControls) !== JSON.stringify(plan.contentControls)) model.contentControls = plan.contentControls
  if (JSON.stringify(model.comments) !== JSON.stringify(plan.comments)) model.comments = plan.comments
}
