import type { DocumentBlock, DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { documentId, runText, tableForBlock } from './document-model'
import { isDocumentTextBoundary } from './document-story'

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

function validateFields(fields: readonly DocumentTextField[], text: string, blockId?: string): void {
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
  const sorted = fields.filter(field => field.endBlockId === undefined || field.endBlockId === blockId).sort((a, b) => a.startOffset - b.startOffset || b.endOffset - a.endOffset)
  const enclosing: DocumentTextField[] = []
  for (const field of sorted) {
    while (enclosing.length && enclosing.at(-1)!.endOffset <= field.startOffset) enclosing.pop()
    const parent = enclosing.at(-1)
    if (parent && field.endOffset > parent.endOffset) throw new RangeError(`正文字段 ${field.id} 与 ${parent.id} 部分交叉。`)
    enclosing.push(field)
  }
}

/** Invalid metadata is reported instead of silently dropping the imported field definition. */
export function documentTextFields(block: DocumentBlock): readonly DocumentTextField[] {
  const collection = block.metadata.documentTextFields
  if (collection === undefined) return []
  if (!isRecord(collection) || collection.version !== 1 || !Array.isArray(collection.fields)) throw new TypeError('正文字段集合版本或结构无效。')
  const fields = collection.fields as DocumentTextField[]
  validateFields(fields, runText(block.runs), block.id)
  return fields
}

export function setDocumentTextFields(block: DocumentBlock, fields: readonly DocumentTextField[], projectedText = runText(block.runs)): void {
  validateFields(fields, projectedText, block.id)
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
}

export interface DocumentTextFieldRange {
  block: DocumentBlock
  field: DocumentTextField
  startParagraph: DocumentTextFieldParagraph
  endParagraph: DocumentTextFieldParagraph
  absoluteStart: number
  absoluteEnd: number
}

const NON_TEXT_BLOCK_TYPES = new Set(['table', 'image', 'divider', 'pageBreak', 'columnBreak', 'sectionBreak'])

/** Body order includes paragraphs in every nested table; stable IDs remain the stored contract. */
export function documentTextFieldParagraphs(model: Pick<DocumentModel, 'blocks'>): DocumentTextFieldParagraph[] {
  const paragraphs: DocumentTextFieldParagraph[] = []
  let offset = 0
  const visit = (blocks: DocumentBlock[]) => {
    for (const [index, block] of blocks.entries()) {
      if (!NON_TEXT_BLOCK_TYPES.has(block.type)) {
        const text = runText(block.runs)
        paragraphs.push({ block, container: blocks, index, text, absoluteStart: offset, absoluteEnd: offset + text.length })
        offset += text.length + 1
      } else if (block.metadata.documentTextFields !== undefined && documentTextFields(block).length) {
        throw new TypeError(`正文字段起始位置必须属于文本段落：${block.id}`)
      }
      for (const row of tableForBlock(block)?.rows ?? []) for (const cell of row.cells) visit(cell.blocks)
    }
  }
  visit(model.blocks)
  return paragraphs
}

function rangesForParagraphs(paragraphs: readonly DocumentTextFieldParagraph[], planned?: ReadonlyMap<string, readonly DocumentTextField[]>): DocumentTextFieldRange[] {
  const byId = new Map<string, DocumentTextFieldParagraph>()
  for (const paragraph of paragraphs) {
    if (byId.has(paragraph.block.id)) throw new TypeError(`正文字段段落标识重复：${paragraph.block.id}`)
    byId.set(paragraph.block.id, paragraph)
  }
  const ids = new Set<string>(), ranges: DocumentTextFieldRange[] = []
  for (const paragraph of paragraphs) {
    const fields = planned?.get(paragraph.block.id) ?? documentTextFields(paragraph.block)
    validateFields(fields, paragraph.text, paragraph.block.id)
    for (const field of fields) {
      if (ids.has(field.id)) throw new TypeError(`正文字段标识重复：${field.id}`)
      ids.add(field.id)
      const end = field.endBlockId === undefined ? paragraph : byId.get(field.endBlockId)
      if (!end) throw new RangeError(`正文字段 ${field.id} 的末段不存在：${field.endBlockId}`)
      if (!isDocumentTextBoundary(end.text, field.endOffset)
        || end.absoluteStart + field.endOffset < paragraph.absoluteStart + field.startOffset) throw new RangeError(`正文字段 ${field.id} 的跨段 UTF-16 范围无效。`)
      ranges.push({ block: paragraph.block, field, startParagraph: paragraph, endParagraph: end,
        absoluteStart: paragraph.absoluteStart + field.startOffset, absoluteEnd: end.absoluteStart + field.endOffset })
    }
  }
  const enclosing: DocumentTextFieldRange[] = []
  for (const range of [...ranges].sort((a, b) => a.absoluteStart - b.absoluteStart || b.absoluteEnd - a.absoluteEnd)) {
    while (enclosing.length && enclosing.at(-1)!.absoluteEnd <= range.absoluteStart) enclosing.pop()
    const parent = enclosing.at(-1)
    if (parent && range.absoluteEnd > parent.absoluteEnd) throw new RangeError(`正文字段 ${range.field.id} 与 ${parent.field.id} 部分交叉。`)
    enclosing.push(range)
  }
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

export function documentFieldLocked(field: RefreshField): boolean {
  return nativeFlag(field, 'locked')
}

export function documentFieldAutomaticRefreshBlocked(field: RefreshField): boolean {
  return documentFieldLocked(field) || nativeFlag(field, 'containsNestedFields')
}

export function documentTextFieldContainsNestedRange(field: DocumentTextField, fields: readonly DocumentTextField[]): boolean {
  return nativeFlag(field, 'containsNestedFields') || fields.some(inner => inner !== field && field.endBlockId === inner.endBlockId
    && inner.startOffset >= field.startOffset && inner.endOffset <= field.endOffset
    && (inner.startOffset > field.startOffset || inner.endOffset < field.endOffset))
}

export function documentTextFieldRangeContainsNested(range: DocumentTextFieldRange, ranges: readonly DocumentTextFieldRange[]): boolean {
  return nativeFlag(range.field, 'containsNestedFields') || ranges.some(inner => inner.field.id !== range.field.id
    && inner.absoluteStart >= range.absoluteStart && inner.absoluteEnd <= range.absoluteEnd
    && (inner.absoluteStart > range.absoluteStart || inner.absoluteEnd < range.absoluteEnd))
}

export function documentTextFieldRangeIntersects(range: DocumentTextFieldRange, start: number, end: number): boolean {
  return start === end
    ? start >= range.absoluteStart && (start < range.absoluteEnd || range.absoluteStart === range.absoluteEnd && start === range.absoluteStart)
    : start < range.absoluteEnd && range.absoluteStart < end
}

export function documentTextFieldIntersects(field: DocumentTextField, start: number, end: number): boolean {
  return start === end
    ? start >= field.startOffset && (start < field.endOffset || field.startOffset === field.endOffset && start === field.startOffset)
    : start < field.endOffset && field.startOffset < end
}

interface FieldRange { startOffset: number; endOffset: number }

/** Editing a result preserves its definition/cache; a larger deletion removes the enclosed field. */
function editedRange(field: FieldRange & { id: string }, start: number, end: number, length: number, preserveFieldId?: string): FieldRange | undefined {
  const from = field.startOffset, to = field.endOffset, delta = length - (end - start)
  if (field.id === preserveFieldId) return { startOffset: from, endOffset: to + delta }
  if (start === end) {
    if (start <= from) return { startOffset: from + length, endOffset: to + length }
    return { startOffset: from, endOffset: start < to ? to + length : to }
  }
  if (end <= from) return { startOffset: from + delta, endOffset: to + delta }
  if (start >= to) return { startOffset: from, endOffset: to }
  if (start <= from && end >= to && (start < from || end > to)) return undefined
  return {
    startOffset: from < start ? from : from >= end ? from + delta : start,
    endOffset: to >= end ? to + delta : to > start ? start + length : to,
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
): void {
  if (!inputs.length || !outputs.length) throw new RangeError('正文字段编辑缺少段落。')
  const inputParagraphs = documentTextFieldParagraphs({ blocks: [...inputs] })
  if (!inputParagraphs.length) throw new RangeError('正文字段编辑缺少文本段落。')
  const overrides = new Map(inputParagraphs.map(paragraph => [paragraph.block.id, paragraph]))
  const source = rebaseParagraphs(documentTextFieldParagraphs(model ?? { blocks: [...inputs] }).map(paragraph => {
    const override = overrides.get(paragraph.block.id)
    return override ? { ...paragraph, block: override.block, text: override.text } : paragraph
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
  const projected = rebaseParagraphs([...source.slice(0, first), ...outputs.map((block, index) => ({
    block, container: [...outputs], index, text: projectedTexts[index]!, absoluteStart: 0, absoluteEnd: 0,
  })), ...source.slice(last + 1)])
  const plans = new Map(projected.map(paragraph => [paragraph.block.id, [] as DocumentTextField[]]))
  for (const { field, absoluteStart, absoluteEnd } of ranges) {
    const edited = editedRange({ id: field.id, startOffset: absoluteStart, endOffset: absoluteEnd }, start, end, text.length, preserveFieldId)
    if (!edited) continue
    const from = projected.find(paragraph => edited.startOffset >= paragraph.absoluteStart && edited.startOffset <= paragraph.absoluteEnd)
    const to = projected.find(paragraph => edited.endOffset >= paragraph.absoluteStart && edited.endOffset <= paragraph.absoluteEnd)
    if (!from || !to) throw new RangeError(`正文字段 ${field.id} 的编辑后端点无法映射。`)
    const { endBlockId: _endBlockId, ...definition } = field
    plans.get(from.block.id)!.push({ ...definition,
      startOffset: edited.startOffset - from.absoluteStart, endOffset: edited.endOffset - to.absoluteStart,
      ...(from.block.id === to.block.id ? {} : { endBlockId: to.block.id }),
    })
  }
  rangesForParagraphs(projected, plans)
  const before = new Map(source.map(paragraph => [paragraph.block.id, documentTextFields(paragraph.block)]))
  const outputIds = new Set(outputs.map(block => block.id))
  for (const paragraph of projected) {
    const plan = plans.get(paragraph.block.id)!
    if (outputIds.has(paragraph.block.id) || !sameFieldEndpoints(before.get(paragraph.block.id) ?? [], plan)) setDocumentTextFields(paragraph.block, plan, paragraph.text)
  }
}

export function transformDocumentTextFieldsForBlockEdit(block: DocumentBlock, start: number, end: number, text: string, preserveFieldId?: string, model?: Pick<DocumentModel, 'blocks'>): void {
  transformDocumentTextFieldsForBlockRangeEdit([block], start, end, text, [block], preserveFieldId, model)
}

export function remapDocumentTextFieldIds(block: DocumentBlock, blockIds?: ReadonlyMap<string, string>): void {
  const fields = documentTextFields(block)
  if (fields.length) setDocumentTextFields(block, fields.map(field => {
    const { endBlockId, ...definition } = field
    const end = endBlockId === undefined ? undefined : blockIds ? blockIds.get(endBlockId) : endBlockId
    if (endBlockId !== undefined && end === undefined) throw new RangeError(`复制范围不包含字段 ${field.id} 的末段。`)
    return { ...definition, id: documentId('field'), ...(end === undefined || end === block.id ? {} : { endBlockId: end }) }
  }))
}

export function insertDocumentTextFields(block: DocumentBlock, fields: readonly DocumentTextField[], offset: number): void {
  if (!fields.length) return
  setDocumentTextFields(block, [...documentTextFields(block), ...fields.map(field => ({
    ...cloneModel(field), startOffset: offset + field.startOffset,
    endOffset: field.endBlockId !== undefined && field.endBlockId !== block.id ? field.endOffset : offset + field.endOffset,
  }))])
}
