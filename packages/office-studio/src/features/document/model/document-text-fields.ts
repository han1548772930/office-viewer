import type { DocumentBlock, DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { documentId, runText } from './document-model'
import { documentBlockChildContainers, isDocumentTextBlock } from './document-structural-blocks'
import { isDocumentTextBoundary } from './document-story'
import { assertDocumentNoteReferenceBoundaryRank, documentNoteReferenceOffsets, remapDocumentNoteReferenceBoundary, type DocumentNoteReferenceOffset } from './document-note-boundary'
import { assertDocumentNoteReferenceOffsets, compareDocumentNoteBoundaries, remapDocumentNoteBoundaryParagraph, type DocumentNoteEditRanks } from './document-note-boundary-edit'
import { documentTextFieldTopologyParents } from './document-text-field-topology'
import { assertDocumentNoteReferenceParagraphIds, documentNoteEmptyReplacementBoundary, documentNoteInsertionBoundary,
  documentNoteReferenceEmptyRemovalRange, documentNoteReferenceEmptyReplacementAffinity, documentNoteReferenceInsertionAffinity } from './document-note-insertion-affinity'

/** A native field result is ordinary Canvas text; offsets exclude file-format delimiters. */
export interface DocumentTextField {
  id: string
  /** Unknown field types remain in the model and are never evaluated as DATE/TIME. */
  field: string
  instruction?: string
  cachedResult?: string
  /** Half-open UTF-16 result start in the paragraph containing this metadata. */
  startOffset: number
  /** UTF-16 result end in endBlockId, or in the owning paragraph when absent. */
  endOffset: number
  /** Cumulative paragraph note count; empty labels retain distinct native boundaries. */
  startNoteReferenceCount?: number
  endNoteReferenceCount?: number
  endBlockId?: string
  locked?: boolean
  containsNestedFields?: boolean
  language?: string
  metadata?: Record<string, unknown>
}

export interface DocumentTextFieldCollection {
  version: 1
  fields: DocumentTextField[]
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function validateFields(fields: readonly DocumentTextField[], text: string, blockId?: string, references: readonly DocumentNoteReferenceOffset[] = []): void {
  const ids = new Set<string>()
  for (const field of fields) {
    if (!isRecord(field) || typeof field.id !== 'string' || !field.id.trim() || ids.has(field.id)
      || typeof field.field !== 'string' || !field.field.trim()) throw new TypeError('正文字段标识或类型无效。')
    ids.add(field.id)
    if (field.endBlockId !== undefined && (typeof field.endBlockId !== 'string' || !field.endBlockId.trim())) throw new TypeError(`正文字段 ${field.id} 的末段标识无效。`)
    if (!isDocumentTextBoundary(text, field.startOffset) || !Number.isSafeInteger(field.endOffset) || field.endOffset < 0
      || (field.endBlockId === undefined || field.endBlockId === blockId) && (!isDocumentTextBoundary(text, field.endOffset) || field.endOffset < field.startOffset)) {
      throw new RangeError(`正文字段 ${field.id} 的 UTF-16 范围无效。`)
    }
    for (const key of ['startNoteReferenceCount', 'endNoteReferenceCount'] as const) {
      if (field[key] !== undefined && (!Number.isSafeInteger(field[key]) || field[key]! < 0)) throw new RangeError(`正文字段 ${field.id} 的累计注释引用序号无效。`)
    }
    const startRank = assertDocumentNoteReferenceBoundaryRank(references, field.startOffset, field.startNoteReferenceCount)
    if (field.endBlockId === undefined || field.endBlockId === blockId) {
      const endRank = assertDocumentNoteReferenceBoundaryRank(references, field.endOffset, field.endNoteReferenceCount)
      if (compareDocumentNoteBoundaries(field.startOffset, startRank, field.endOffset, endRank) > 0) throw new RangeError(`正文字段 ${field.id} 的注释引用端点顺序相反。`)
    }
    for (const key of ['instruction', 'cachedResult', 'language'] as const) {
      if (field[key] !== undefined && typeof field[key] !== 'string') throw new TypeError(`正文字段 ${field.id} 的 ${key} 无效。`)
    }
    for (const key of ['locked', 'containsNestedFields'] as const) {
      if (field[key] !== undefined && typeof field[key] !== 'boolean') throw new TypeError(`正文字段 ${field.id} 的 ${key} 必须为布尔值。`)
    }
    if (field.metadata !== undefined && !isRecord(field.metadata)) throw new TypeError(`正文字段 ${field.id} 的来源无效。`)
  }
  // A paragraph alone cannot order endpoints in another paragraph. The model
  // validator below checks those ranges together with ordinary local fields.
  const sorted = fields.filter(field => field.endBlockId === undefined || field.endBlockId === blockId).map(field => ({ field,
    startRank: assertDocumentNoteReferenceBoundaryRank(references, field.startOffset, field.startNoteReferenceCount),
    endRank: assertDocumentNoteReferenceBoundaryRank(references, field.endOffset, field.endNoteReferenceCount) }))
    .sort((a, b) => compareDocumentNoteBoundaries(a.field.startOffset, a.startRank, b.field.startOffset, b.startRank)
      || compareDocumentNoteBoundaries(b.field.endOffset, b.endRank, a.field.endOffset, a.endRank))
  const enclosing: typeof sorted = []
  for (const candidate of sorted) {
    const field = candidate.field
    while (enclosing.length && compareDocumentNoteBoundaries(enclosing.at(-1)!.field.endOffset, enclosing.at(-1)!.endRank, field.startOffset, candidate.startRank) <= 0) enclosing.pop()
    const parent = enclosing.at(-1)
    if (parent && compareDocumentNoteBoundaries(field.endOffset, candidate.endRank, parent.field.endOffset, parent.endRank) > 0) throw new RangeError(`正文字段 ${field.id} 与 ${parent.field.id} 部分交叉。`)
    if (compareDocumentNoteBoundaries(field.startOffset, candidate.startRank, field.endOffset, candidate.endRank) < 0) enclosing.push(candidate)
  }
}

/** Invalid metadata is reported instead of silently dropping the imported field definition. */
export function documentTextFields(block: DocumentBlock): readonly DocumentTextField[] {
  const collection = block.metadata.documentTextFields
  if (collection === undefined) return []
  if (!isRecord(collection) || collection.version !== 1 || !Array.isArray(collection.fields)) throw new TypeError('正文字段集合版本或结构无效。')
  const fields = collection.fields as DocumentTextField[]
  validateFields(fields, runText(block.runs), block.id, documentNoteReferenceOffsets(block.runs))
  return fields
}

export function setDocumentTextFields(block: DocumentBlock, fields: readonly DocumentTextField[], projectedText = runText(block.runs),
  projectedNoteReferences: readonly DocumentNoteReferenceOffset[] = documentNoteReferenceOffsets(block.runs)): void {
  assertDocumentNoteReferenceOffsets(projectedNoteReferences, projectedText)
  validateFields(fields, projectedText, block.id, projectedNoteReferences)
  if (!fields.length && block.metadata.documentTextFields === undefined) return
  const source = block.metadata.documentTextFields
  block.metadata = { ...block.metadata, documentTextFields: {
    ...(isRecord(source) ? source : {}), version: 1, fields: fields.map(field => ({ ...field })),
  } satisfies DocumentTextFieldCollection }
}

export interface DocumentTextFieldParagraph {
  block: DocumentBlock
  container: DocumentBlock[]
  index: number
  text: string
  /** Computed positions include one synthetic newline between text paragraphs. */
  absoluteStart: number
  absoluteEnd: number
  /** Captured before edits mutate the referenced block's runs. */
  noteReferences: readonly DocumentNoteReferenceOffset[]
}

export interface DocumentTextFieldRange {
  block: DocumentBlock
  field: DocumentTextField
  startParagraph: DocumentTextFieldParagraph
  endParagraph: DocumentTextFieldParagraph
  absoluteStart: number
  absoluteEnd: number
  startNoteReferenceCount: number
  endNoteReferenceCount: number
}

/** Body order includes paragraphs in every nested table; stable IDs remain the stored contract. */
export function documentTextFieldParagraphs(model: Pick<DocumentModel, 'blocks'>): DocumentTextFieldParagraph[] {
  const paragraphs: DocumentTextFieldParagraph[] = []
  let offset = 0
  const visit = (blocks: DocumentBlock[]) => {
    for (const [index, block] of blocks.entries()) {
      if (isDocumentTextBlock(block)) {
        const text = runText(block.runs)
        paragraphs.push({ block, container: blocks, index, text, absoluteStart: offset, absoluteEnd: offset + text.length,
          noteReferences: documentNoteReferenceOffsets(block.runs) })
        offset += text.length + 1
      } else if (block.metadata.documentTextFields !== undefined && documentTextFields(block).length) {
        throw new TypeError(`正文字段起始位置必须属于文本段落：${block.id}`)
      }
      for (const content of documentBlockChildContainers(block)) visit(content)
    }
  }
  visit(model.blocks)
  return paragraphs
}

function rangesForParagraphs(paragraphs: readonly DocumentTextFieldParagraph[], planned?: ReadonlyMap<string, readonly DocumentTextField[]>): DocumentTextFieldRange[] {
  assertDocumentNoteReferenceParagraphIds(paragraphs)
  const byId = new Map<string, DocumentTextFieldParagraph>()
  for (const paragraph of paragraphs) {
    if (byId.has(paragraph.block.id)) throw new TypeError(`正文字段段落标识重复：${paragraph.block.id}`)
    byId.set(paragraph.block.id, paragraph)
  }
  const ids = new Set<string>(), ranges: DocumentTextFieldRange[] = []
  for (const paragraph of paragraphs) {
    const fields = planned?.get(paragraph.block.id) ?? documentTextFields(paragraph.block)
    validateFields(fields, paragraph.text, paragraph.block.id, paragraph.noteReferences)
    for (const field of fields) {
      if (ids.has(field.id)) throw new TypeError(`正文字段标识重复：${field.id}`)
      ids.add(field.id)
      const end = field.endBlockId === undefined ? paragraph : byId.get(field.endBlockId)
      if (!end) throw new RangeError(`正文字段 ${field.id} 的末段不存在：${field.endBlockId}`)
      if (!isDocumentTextBoundary(end.text, field.endOffset)
        || end.absoluteStart + field.endOffset < paragraph.absoluteStart + field.startOffset) throw new RangeError(`正文字段 ${field.id} 的跨段 UTF-16 范围无效。`)
      const startNoteReferenceCount = assertDocumentNoteReferenceBoundaryRank(paragraph.noteReferences, field.startOffset, field.startNoteReferenceCount)
      const endNoteReferenceCount = assertDocumentNoteReferenceBoundaryRank(end.noteReferences, field.endOffset, field.endNoteReferenceCount)
      if (compareDocumentNoteBoundaries(paragraph.absoluteStart + field.startOffset, startNoteReferenceCount, end.absoluteStart + field.endOffset, endNoteReferenceCount) > 0) {
        throw new RangeError(`正文字段 ${field.id} 的跨段注释引用端点顺序相反。`)
      }
      ranges.push({ block: paragraph.block, field, startParagraph: paragraph, endParagraph: end,
        absoluteStart: paragraph.absoluteStart + field.startOffset, absoluteEnd: end.absoluteStart + field.endOffset, startNoteReferenceCount, endNoteReferenceCount })
    }
  }
  documentTextFieldTopologyParents(ranges)
  return ranges
}

/** Validate the full result range rather than interpreting an end offset in the owning paragraph. */
export function documentTextFieldRanges(model: Pick<DocumentModel, 'blocks'>): DocumentTextFieldRange[] {
  return rangesForParagraphs(documentTextFieldParagraphs(model))
}

/** Includes empty collections and validates nested/cross-paragraph fields before a global refresh. */
export function documentTextFieldBlocks(model: Pick<DocumentModel, 'blocks'>): DocumentBlock[] {
  const paragraphs = documentTextFieldParagraphs(model)
  rangesForParagraphs(paragraphs)
  return paragraphs.filter(({ block }) => block.metadata.documentTextFields !== undefined).map(({ block }) => block)
}

type RefreshField = Pick<DocumentTextField, 'locked' | 'containsNestedFields' | 'metadata'>

function nativeFlag(field: RefreshField, key: 'locked' | 'containsNestedFields'): boolean {
  // An explicit false supersedes immutable import provenance after an API edit.
  if (field[key] !== undefined) return field[key] === true
  const source = field.metadata?.source
  const properties = isRecord(source) && isRecord(source.properties) ? source.properties : undefined
  const value = isRecord(source) ? source[key] ?? properties?.[key] : undefined
  return value === true || value === 1 || value === '1' || value === 'true'
}

export function documentFieldRetainsNestedFields(field: RefreshField): boolean {
  const source = field.metadata?.source
  // These records are written back with offsets relative to the current cache.
  // A flag override cannot make their ranges safe after changing that cache.
  return isRecord(source) && Array.isArray(source.nestedFields) && source.nestedFields.length > 0
}

export function documentFieldLocked(field: RefreshField): boolean {
  return nativeFlag(field, 'locked')
}

export function documentFieldAutomaticRefreshBlocked(field: RefreshField): boolean {
  return documentFieldLocked(field) || nativeFlag(field, 'containsNestedFields') || documentFieldRetainsNestedFields(field)
}

export function documentTextFieldContainsNestedRange(field: DocumentTextField, fields: readonly DocumentTextField[], references?: readonly DocumentNoteReferenceOffset[]): boolean {
  const rank = (candidate: DocumentTextField, edge: 'start' | 'end') => references
    ? assertDocumentNoteReferenceBoundaryRank(references, edge === 'start' ? candidate.startOffset : candidate.endOffset,
      edge === 'start' ? candidate.startNoteReferenceCount : candidate.endNoteReferenceCount)
    : (edge === 'start' ? candidate.startNoteReferenceCount : candidate.endNoteReferenceCount) ?? 0
  return nativeFlag(field, 'containsNestedFields') || documentFieldRetainsNestedFields(field) || fields.some(inner => inner.id !== field.id
    && (isRecord(inner.metadata?.source) && inner.metadata.source.parentFieldId === field.id
      || field.endBlockId === inner.endBlockId
        && compareDocumentNoteBoundaries(inner.startOffset, rank(inner, 'start'), field.startOffset, rank(field, 'start')) >= 0
        && compareDocumentNoteBoundaries(inner.endOffset, rank(inner, 'end'), field.endOffset, rank(field, 'end')) <= 0
        && (compareDocumentNoteBoundaries(inner.startOffset, rank(inner, 'start'), field.startOffset, rank(field, 'start')) > 0
          || compareDocumentNoteBoundaries(inner.endOffset, rank(inner, 'end'), field.endOffset, rank(field, 'end')) < 0)))
}

export function documentTextFieldRangeContainsNested(range: DocumentTextFieldRange, ranges: readonly DocumentTextFieldRange[]): boolean {
  return nativeFlag(range.field, 'containsNestedFields') || documentFieldRetainsNestedFields(range.field) || ranges.some(inner => inner.field.id !== range.field.id
    && (isRecord(inner.field.metadata?.source) && inner.field.metadata.source.parentFieldId === range.field.id
      || compareDocumentNoteBoundaries(inner.absoluteStart, inner.startNoteReferenceCount, range.absoluteStart, range.startNoteReferenceCount) >= 0
        && compareDocumentNoteBoundaries(inner.absoluteEnd, inner.endNoteReferenceCount, range.absoluteEnd, range.endNoteReferenceCount) <= 0
        && (compareDocumentNoteBoundaries(inner.absoluteStart, inner.startNoteReferenceCount, range.absoluteStart, range.startNoteReferenceCount) > 0
          || compareDocumentNoteBoundaries(inner.absoluteEnd, inner.endNoteReferenceCount, range.absoluteEnd, range.endNoteReferenceCount) < 0)))
}

function intersectsBoundaries(from: number, fromRank: number, to: number, toRank: number, start: number, startRank: number, end: number, endRank: number): boolean {
  const collapsed = compareDocumentNoteBoundaries(start, startRank, end, endRank) === 0
  return collapsed
    ? compareDocumentNoteBoundaries(start, startRank, from, fromRank) >= 0
      && (compareDocumentNoteBoundaries(start, startRank, to, toRank) < 0
        || compareDocumentNoteBoundaries(from, fromRank, to, toRank) === 0 && compareDocumentNoteBoundaries(start, startRank, from, fromRank) === 0)
    : compareDocumentNoteBoundaries(start, startRank, to, toRank) < 0 && compareDocumentNoteBoundaries(from, fromRank, end, endRank) < 0
}

export function documentTextFieldRangeIntersects(range: DocumentTextFieldRange, start: number, end: number, startRank?: number, endRank?: number): boolean {
  const resolved = (offset: number, rank: number | undefined) => {
    const paragraph = [range.startParagraph, range.endParagraph].find(paragraph => offset >= paragraph.absoluteStart && offset <= paragraph.absoluteEnd)
    return paragraph ? assertDocumentNoteReferenceBoundaryRank(paragraph.noteReferences, offset - paragraph.absoluteStart, rank) : rank ?? 0
  }
  return intersectsBoundaries(range.absoluteStart, range.startNoteReferenceCount, range.absoluteEnd, range.endNoteReferenceCount,
    start, resolved(start, startRank), end, resolved(end, endRank))
}

export function documentTextFieldIntersects(field: DocumentTextField, start: number, end: number,
  startRank?: number, endRank?: number, references?: readonly DocumentNoteReferenceOffset[]): boolean {
  const resolved = (offset: number, rank: number | undefined) => references ? assertDocumentNoteReferenceBoundaryRank(references, offset, rank) : rank ?? 0
  return intersectsBoundaries(field.startOffset, resolved(field.startOffset, field.startNoteReferenceCount), field.endOffset, resolved(field.endOffset, field.endNoteReferenceCount),
    start, resolved(start, startRank), end, resolved(end, endRank))
}

interface FieldRange { startOffset: number; endOffset: number }

/** Editing a result preserves its definition/cache; a larger deletion removes the enclosed field. */
function editedRange(field: FieldRange & { id: string; startNoteReferenceCount: number; endNoteReferenceCount: number },
  start: number, end: number, length: number, startRank: number, endRank: number, preserveFieldId?: string): FieldRange | undefined {
  const from = field.startOffset, to = field.endOffset, delta = length - (end - start)
  const startFrom = compareDocumentNoteBoundaries(start, startRank, from, field.startNoteReferenceCount)
  const startTo = compareDocumentNoteBoundaries(start, startRank, to, field.endNoteReferenceCount)
  const endFrom = compareDocumentNoteBoundaries(end, endRank, from, field.startNoteReferenceCount)
  const endTo = compareDocumentNoteBoundaries(end, endRank, to, field.endNoteReferenceCount)
  if (field.id === preserveFieldId) return { startOffset: from, endOffset: to + delta }
  if (compareDocumentNoteBoundaries(start, startRank, end, endRank) === 0) {
    if (startFrom <= 0) return { startOffset: from + length, endOffset: to + length }
    return { startOffset: from, endOffset: startTo < 0 ? to + length : to }
  }
  if (endFrom <= 0) return { startOffset: from + delta, endOffset: to + delta }
  if (startTo >= 0) return { startOffset: from, endOffset: to }
  if (startFrom <= 0 && endTo >= 0 && (startFrom < 0 || endTo > 0)) return undefined
  return {
    startOffset: startFrom > 0 ? from : endFrom <= 0 ? from + delta : start,
    endOffset: endTo <= 0 ? to + delta : startTo < 0 ? start + length : to,
  }
}

function rebaseParagraphs(paragraphs: readonly DocumentTextFieldParagraph[]): DocumentTextFieldParagraph[] {
  let offset = 0
  return paragraphs.map(paragraph => {
    const next = { ...paragraph, absoluteStart: offset, absoluteEnd: offset + paragraph.text.length }
    offset += paragraph.text.length + 1
    return next
  })
}

function sameFieldEndpoints(before: readonly DocumentTextField[], after: readonly DocumentTextField[]): boolean {
  return before.length === after.length && before.every((field, index) => {
    const next = after[index]!
    return field.id === next.id && field.startOffset === next.startOffset && field.endOffset === next.endOffset && field.endBlockId === next.endBlockId
      && field.startNoteReferenceCount === next.startNoteReferenceCount && field.endNoteReferenceCount === next.endNoteReferenceCount
  })
}

/**
 * Move both endpoints and the owning definition before changing paragraph runs.
 * The full model finds fields which begin outside the edited paragraphs. All
 * endpoint/containment plans are validated before any field collection changes.
 */
export function transformDocumentTextFieldsForBlockRangeEdit(
  inputs: readonly DocumentBlock[], startOffset: number, endOffset: number, text: string,
  outputs: readonly DocumentBlock[] = inputs.slice(0, 1), preserveFieldId?: string,
  model?: Pick<DocumentModel, 'blocks'>,
  projectedNoteReferences?: readonly (readonly DocumentNoteReferenceOffset[])[],
  sourceNoteRanks?: DocumentNoteEditRanks,
): void {
  if (!inputs.length || !outputs.length) throw new RangeError('正文字段编辑缺少段落。')
  const inputParagraphs = documentTextFieldParagraphs({ blocks: [...inputs] })
  if (!inputParagraphs.length) throw new RangeError('正文字段编辑缺少文本段落。')
  const overrides = new Map(inputParagraphs.map(paragraph => [paragraph.block.id, paragraph]))
  const source = rebaseParagraphs(documentTextFieldParagraphs(model ?? { blocks: [...inputs] }).map(paragraph => {
    const override = overrides.get(paragraph.block.id)
    return override ? { ...paragraph, block: override.block, text: override.text, noteReferences: override.noteReferences } : paragraph
  }))
  const ranges = rangesForParagraphs(source)
  if (!ranges.length) return
  const first = source.findIndex(paragraph => paragraph.block.id === inputParagraphs[0]!.block.id)
  const last = source.findIndex(paragraph => paragraph.block.id === inputParagraphs.at(-1)!.block.id)
  if (first < 0 || last < first) throw new RangeError('正文字段编辑的段落映射无效。')
  const firstParagraph = source[first]!, lastParagraph = source[last]!
  if (!isDocumentTextBoundary(firstParagraph.text, startOffset) || !isDocumentTextBoundary(lastParagraph.text, endOffset)) throw new RangeError('正文字段编辑的 UTF-16 位置无效。')
  const start = firstParagraph.absoluteStart + startOffset, end = lastParagraph.absoluteStart + endOffset
  if (end < start) throw new RangeError('正文字段编辑范围倒置。')
  const prefix = firstParagraph.text.slice(0, startOffset), suffix = lastParagraph.text.slice(endOffset)
  const parts = text.split('\n')
  const projectedTexts = outputs.length === 1 ? [prefix + text + suffix] : parts.map((part, index) =>
    `${index === 0 ? prefix : ''}${part}${index === parts.length - 1 ? suffix : ''}`)
  if (projectedTexts.length !== outputs.length) throw new RangeError('正文字段编辑的段落映射无效。')
  if (projectedNoteReferences !== undefined && projectedNoteReferences.length !== outputs.length) throw new RangeError('正文字段编辑的注释引用目的段落数量不一致。')
  const sourceStartRank = assertDocumentNoteReferenceBoundaryRank(firstParagraph.noteReferences, startOffset, sourceNoteRanks?.start)
  const sourceEndRank = assertDocumentNoteReferenceBoundaryRank(lastParagraph.noteReferences, endOffset, sourceNoteRanks?.end)
  const prefixReferences = firstParagraph.noteReferences.filter((reference, index) =>
    compareDocumentNoteBoundaries(reference.endOffset, index + 1, startOffset, sourceStartRank) <= 0).map(reference => ({ ...reference }))
  const suffixBase = projectedTexts.at(-1)!.length - suffix.length
  const suffixReferences = lastParagraph.noteReferences.filter((reference, index) =>
    compareDocumentNoteBoundaries(reference.startOffset, index, endOffset, sourceEndRank) >= 0).map(reference => ({ ...reference,
    startOffset: suffixBase + reference.startOffset - endOffset, endOffset: suffixBase + reference.endOffset - endOffset }))
  // Equal display text does not prove equal future atom identity: replacing a
  // reference label by the same plain text still removes the note entity.
  const actualOutputGeometry = outputs.every((block, index) => !inputs.includes(block) && runText(block.runs) === projectedTexts[index])
  const outputReferences = outputs.map((block, index) => projectedNoteReferences?.[index] ?? (actualOutputGeometry ? documentNoteReferenceOffsets(block.runs)
    : [...(index === 0 ? prefixReferences : []), ...(index + 1 === outputs.length ? suffixReferences : [])]))
  for (const [index, references] of outputReferences.entries()) assertDocumentNoteReferenceOffsets(references, projectedTexts[index]!)
  const projected = rebaseParagraphs([...source.slice(0, first), ...outputs.map((block, index) => ({
    block, container: [...outputs], index, text: projectedTexts[index]!, absoluteStart: 0, absoluteEnd: 0,
    noteReferences: outputReferences[index]!,
  })), ...source.slice(last + 1)])
  const plans = new Map(projected.map(paragraph => [paragraph.block.id, [] as DocumentTextField[]]))
  const targets = projected.map(paragraph => ({ ...paragraph, id: paragraph.block.id }))
  assertDocumentNoteReferenceParagraphIds(targets)
  let editStartRank = sourceStartRank
  let editEndRank = sourceEndRank
  const outputsForEdit = targets.slice(first, first + outputs.length)
  const emptyRemoval = start === end ? documentNoteReferenceEmptyRemovalRange(firstParagraph.noteReferences,
    new Set(outputsForEdit.flatMap(paragraph => paragraph.noteReferences.map(reference => reference.runId))), startOffset) : undefined
  const insertionAffinity = start === end && !emptyRemoval
    ? documentNoteReferenceInsertionAffinity(firstParagraph.noteReferences, outputsForEdit, startOffset, text.length) : undefined
  const replacementAffinity = emptyRemoval
    ? documentNoteReferenceEmptyReplacementAffinity(firstParagraph.noteReferences, outputsForEdit, startOffset, text.length, emptyRemoval) : undefined
  if (sourceNoteRanks && (emptyRemoval && (editStartRank !== emptyRemoval.startRank || editEndRank !== emptyRemoval.endRank)
    || insertionAffinity && (editStartRank !== insertionAffinity.sourceRank || editEndRank !== insertionAffinity.sourceRank))) {
    throw new RangeError('正文字段编辑的源注释引用端点与真实目的内容不一致。')
  }
  if (emptyRemoval) { editStartRank = emptyRemoval.startRank; editEndRank = emptyRemoval.endRank }
  else if (insertionAffinity) editStartRank = editEndRank = insertionAffinity.sourceRank
  if (compareDocumentNoteBoundaries(start, editStartRank, end, editEndRank) > 0) throw new RangeError('正文字段编辑的源注释引用端点倒置。')
  const parents = insertionAffinity ? documentTextFieldTopologyParents(ranges) : []
  const pointAtParentEnd = (index: number): boolean => {
    if (!insertionAffinity || ranges[index]!.startParagraph.block.id !== firstParagraph.block.id
      || ranges[index]!.field.startOffset !== insertionAffinity.sourceOffset || ranges[index]!.startNoteReferenceCount !== insertionAffinity.sourceRank) return false
    let parent = parents[index]
    while (parent !== undefined) {
      const enclosing = ranges[parent]!
      if (enclosing.field.id === preserveFieldId || enclosing.endParagraph.block.id !== firstParagraph.block.id
        || enclosing.field.endOffset !== insertionAffinity.sourceOffset || enclosing.endNoteReferenceCount !== insertionAffinity.sourceRank) return false
      if (compareDocumentNoteBoundaries(enclosing.absoluteStart, enclosing.startNoteReferenceCount, enclosing.absoluteEnd, enclosing.endNoteReferenceCount) < 0) return true
      parent = parents[parent]
    }
    return false
  }
  for (const [index, range] of ranges.entries()) {
    const { field, absoluteStart, absoluteEnd, startParagraph, endParagraph, startNoteReferenceCount, endNoteReferenceCount } = range
    const edited = editedRange({ id: field.id, startOffset: absoluteStart, endOffset: absoluteEnd, startNoteReferenceCount, endNoteReferenceCount },
      start, end, text.length, editStartRank, editEndRank, preserveFieldId)
    if (!edited) continue
    const from = projected.find(paragraph => edited.startOffset >= paragraph.absoluteStart && edited.startOffset <= paragraph.absoluteEnd)
    const to = projected.find(paragraph => edited.endOffset >= paragraph.absoluteStart && edited.endOffset <= paragraph.absoluteEnd)
    if (!from || !to) throw new RangeError(`正文字段 ${field.id} 的编辑后端点无法映射。`)
    let mappedStart = remapDocumentNoteBoundaryParagraph(startParagraph.noteReferences, field.startOffset, field.startNoteReferenceCount, targets,
      { paragraph: targets.find(paragraph => paragraph.block.id === from.block.id)!, offset: edited.startOffset - from.absoluteStart })
    let mappedEnd = remapDocumentNoteBoundaryParagraph(endParagraph.noteReferences, field.endOffset, field.endNoteReferenceCount, targets,
      { paragraph: targets.find(paragraph => paragraph.block.id === to.block.id)!, offset: edited.endOffset - to.absoluteStart })
    const point = compareDocumentNoteBoundaries(absoluteStart, startNoteReferenceCount, absoluteEnd, endNoteReferenceCount) === 0
    const pointBefore = point && field.id !== preserveFieldId && pointAtParentEnd(index)
    if (startParagraph.block.id === firstParagraph.block.id) mappedStart = documentNoteEmptyReplacementBoundary(replacementAffinity,
      field.startOffset, startNoteReferenceCount, 'start') ?? (pointBefore ? insertionAffinity?.before
      : documentNoteInsertionBoundary(insertionAffinity, field.startOffset, startNoteReferenceCount, 'start', point, field.id === preserveFieldId)) ?? mappedStart
    if (endParagraph.block.id === firstParagraph.block.id) mappedEnd = documentNoteEmptyReplacementBoundary(replacementAffinity,
      field.endOffset, endNoteReferenceCount, 'end') ?? (pointBefore ? insertionAffinity?.before
      : documentNoteInsertionBoundary(insertionAffinity, field.endOffset, endNoteReferenceCount, 'end', point, field.id === preserveFieldId)) ?? mappedEnd
    const { endBlockId: _endBlockId, startNoteReferenceCount: _startRank, endNoteReferenceCount: _endRank, ...definition } = field
    plans.get(mappedStart.paragraph.block.id)!.push({ ...definition,
      startOffset: mappedStart.offset, endOffset: mappedEnd.offset,
      ...(mappedStart.rank === undefined ? {} : { startNoteReferenceCount: mappedStart.rank }),
      ...(mappedEnd.rank === undefined ? {} : { endNoteReferenceCount: mappedEnd.rank }),
      ...(mappedStart.paragraph.block.id === mappedEnd.paragraph.block.id ? {} : { endBlockId: mappedEnd.paragraph.block.id }),
    })
  }
  rangesForParagraphs(projected, plans)
  const before = new Map(source.map(paragraph => [paragraph.block.id, documentTextFields(paragraph.block)]))
  const outputIds = new Set(outputs.map(block => block.id))
  for (const paragraph of projected) {
    const plan = plans.get(paragraph.block.id)!
    if (outputIds.has(paragraph.block.id) || !sameFieldEndpoints(before.get(paragraph.block.id) ?? [], plan)) setDocumentTextFields(paragraph.block, plan, paragraph.text, paragraph.noteReferences)
  }
}

export function transformDocumentTextFieldsForBlockEdit(block: DocumentBlock, start: number, end: number, text: string, preserveFieldId?: string, model?: Pick<DocumentModel, 'blocks'>,
  projectedNoteReferences?: readonly DocumentNoteReferenceOffset[], sourceNoteRanks?: DocumentNoteEditRanks): void {
  transformDocumentTextFieldsForBlockRangeEdit([block], start, end, text, [block], preserveFieldId, model,
    projectedNoteReferences === undefined ? undefined : [projectedNoteReferences], sourceNoteRanks)
}

export function remapDocumentTextFieldIds(block: DocumentBlock, blockIds?: ReadonlyMap<string, string>, fieldIds?: ReadonlyMap<string, string>): void {
  const fields = documentTextFields(block)
  const ids = fieldIds ?? new Map(fields.map(field => [field.id, documentId('field')]))
  if (fields.length) setDocumentTextFields(block, fields.map(field => {
    const { endBlockId, ...definition } = field
    const end = endBlockId === undefined ? undefined : blockIds ? blockIds.get(endBlockId) : endBlockId
    if (endBlockId !== undefined && end === undefined) throw new RangeError(`复制范围不包含字段 ${field.id} 的末段。`)
    const id = ids.get(field.id)
    if (id === undefined) throw new RangeError(`复制范围没有字段 ${field.id} 的新身份。`)
    const metadata = cloneModel(field.metadata ?? {}), source = metadata.source
    if (isRecord(source) && source.parentFieldId !== undefined && source.parentFieldId !== null) {
      if (typeof source.parentFieldId !== 'string') throw new TypeError(`字段 ${field.id} 的父身份无效。`)
      const parent = ids.get(source.parentFieldId)
      if (parent === undefined) throw new RangeError(`复制范围不包含字段 ${field.id} 的父字段。`)
      metadata.officeTextFieldCopySource = { version: 1, fieldId: field.id, parentFieldId: source.parentFieldId }
      source.parentFieldId = parent
    }
    return { ...definition, id, metadata, ...(end === undefined || end === block.id ? {} : { endBlockId: end }) }
  }))
}

export function insertDocumentTextFields(block: DocumentBlock, fields: readonly DocumentTextField[], offset: number,
  sourceNoteReferences: readonly DocumentNoteReferenceOffset[] = []): void {
  if (!fields.length) return
  const references = documentNoteReferenceOffsets(block.runs)
  setDocumentTextFields(block, [...documentTextFields(block), ...fields.map(field => {
    const { startNoteReferenceCount, endNoteReferenceCount, ...definition } = cloneModel(field)
    const start = remapDocumentNoteReferenceBoundary(sourceNoteReferences, references, field.startOffset, offset + field.startOffset, startNoteReferenceCount)
    const localEnd = field.endBlockId === undefined || field.endBlockId === block.id
    const end = localEnd ? remapDocumentNoteReferenceBoundary(sourceNoteReferences, references, field.endOffset, offset + field.endOffset, endNoteReferenceCount)
      : { offset: field.endOffset, rank: endNoteReferenceCount }
    return { ...definition, startOffset: start.offset, endOffset: end.offset,
      ...(start.rank === undefined ? {} : { startNoteReferenceCount: start.rank }),
      ...(end.rank === undefined ? {} : { endNoteReferenceCount: end.rank }) }
  })])
}
