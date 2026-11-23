import type { DocumentBlock, DocumentModel } from '../../../domain/model'
import { documentId, runText, tableForBlock } from './document-model'
import { isDocumentTextBoundary } from './document-story'

/** A native field result is ordinary Canvas text; offsets exclude file-format delimiters. */
export interface DocumentTextField {
  id: string
  /** Unknown field types remain in the model and are never evaluated as DATE/TIME. */
  field: string
  instruction?: string
  cachedResult?: string
  /** Half-open UTF-16 result range in the paragraph containing this metadata. */
  startOffset: number
  endOffset: number
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

function validateFields(fields: readonly DocumentTextField[], text: string): void {
  const ids = new Set<string>()
  for (const field of fields) {
    if (!isRecord(field) || typeof field.id !== 'string' || !field.id.trim() || ids.has(field.id)
      || typeof field.field !== 'string' || !field.field.trim()) throw new TypeError('正文字段标识或类型无效。')
    ids.add(field.id)
    if (!isDocumentTextBoundary(text, field.startOffset) || !isDocumentTextBoundary(text, field.endOffset)
      || field.endOffset < field.startOffset) throw new RangeError(`正文字段 ${field.id} 的 UTF-16 范围无效。`)
    for (const key of ['instruction', 'cachedResult', 'language'] as const) {
      if (field[key] !== undefined && typeof field[key] !== 'string') throw new TypeError(`正文字段 ${field.id} 的 ${key} 无效。`)
    }
    for (const key of ['locked', 'containsNestedFields'] as const) {
      if (field[key] !== undefined && typeof field[key] !== 'boolean') throw new TypeError(`正文字段 ${field.id} 的 ${key} 必须为布尔值。`)
    }
    if (field.metadata !== undefined && !isRecord(field.metadata)) throw new TypeError(`正文字段 ${field.id} 的来源无效。`)
  }
  const sorted = [...fields].sort((a, b) => a.startOffset - b.startOffset || b.endOffset - a.endOffset)
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
  validateFields(fields, runText(block.runs))
  return fields
}

export function setDocumentTextFields(block: DocumentBlock, fields: readonly DocumentTextField[], projectedText = runText(block.runs)): void {
  validateFields(fields, projectedText)
  if (!fields.length && block.metadata.documentTextFields === undefined) return
  const source = block.metadata.documentTextFields
  block.metadata = { ...block.metadata, documentTextFields: {
    ...(isRecord(source) ? source : {}), version: 1, fields: fields.map(field => ({ ...field })),
  } satisfies DocumentTextFieldCollection }
}

/** Includes text in nested table cells when the global refresh traverses the document. */
export function documentTextFieldBlocks(model: Pick<DocumentModel, 'blocks'>): DocumentBlock[] {
  const result: DocumentBlock[] = []
  const visit = (blocks: readonly DocumentBlock[]) => {
    for (const block of blocks) {
      if (block.metadata.documentTextFields !== undefined) result.push(block)
      for (const row of tableForBlock(block)?.rows ?? []) for (const cell of row.cells) visit(cell.blocks)
    }
  }
  visit(model.blocks)
  const ids = new Set<string>()
  for (const block of result) for (const field of documentTextFields(block)) {
    if (ids.has(field.id)) throw new TypeError(`正文字段标识重复：${field.id}`)
    ids.add(field.id)
  }
  return result
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

export function documentFieldAutomaticRefreshBlocked(field: RefreshField): boolean {
  return nativeFlag(field, 'locked') || nativeFlag(field, 'containsNestedFields')
}

export function documentTextFieldContainsNestedRange(field: DocumentTextField, fields: readonly DocumentTextField[]): boolean {
  return nativeFlag(field, 'containsNestedFields') || fields.some(inner => inner !== field
    && inner.startOffset >= field.startOffset && inner.endOffset <= field.endOffset
    && (inner.startOffset > field.startOffset || inner.endOffset < field.endOffset))
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

/**
 * Move fields with a paragraph suffix or a merged range before changing runs.
 * All ranges are prepared first, so a paragraph split through an imported field
 * reports the unsupported cross-paragraph result without losing its metadata.
 */
export function transformDocumentTextFieldsForBlockRangeEdit(
  inputs: readonly DocumentBlock[], startOffset: number, endOffset: number, text: string,
  outputs: readonly DocumentBlock[] = inputs.slice(0, 1), preserveFieldId?: string,
): void {
  if (!inputs.length || !outputs.length) throw new RangeError('正文字段编辑缺少段落。')
  const sourceTexts = inputs.map(block => runText(block.runs))
  const fields = inputs.flatMap(block => documentTextFields(block))
  if (!fields.length) return
  if (!isDocumentTextBoundary(sourceTexts[0]!, startOffset) || !isDocumentTextBoundary(sourceTexts.at(-1)!, endOffset)) throw new RangeError('正文字段编辑的 UTF-16 位置无效。')
  const source = sourceTexts.join('\n')
  const end = source.length - sourceTexts.at(-1)!.length + endOffset
  if (end < startOffset) throw new RangeError('正文字段编辑范围倒置。')
  const prefix = sourceTexts[0]!.slice(0, startOffset), suffix = sourceTexts.at(-1)!.slice(endOffset)
  const parts = text.split('\n')
  const projected = outputs.length === 1 ? [prefix + text + suffix] : parts.map((part, index) =>
    `${index === 0 ? prefix : ''}${part}${index === parts.length - 1 ? suffix : ''}`)
  if (projected.length !== outputs.length) throw new RangeError('正文字段编辑的段落映射无效。')
  const plans = outputs.map(() => [] as DocumentTextField[])
  const outputOffsets: number[] = []
  let cursor = 0
  for (const part of projected) { outputOffsets.push(cursor); cursor += part.length + 1 }
  const ids = new Set<string>()
  cursor = 0
  for (const [index, block] of inputs.entries()) {
    for (const field of documentTextFields(block)) {
      if (ids.has(field.id)) throw new TypeError(`正文字段标识重复：${field.id}`)
      ids.add(field.id)
      const range = editedRange({ ...field, startOffset: cursor + field.startOffset, endOffset: cursor + field.endOffset },
        startOffset, end, text.length, preserveFieldId)
      if (!range) continue
      const destination = projected.findIndex((part, output) => range.startOffset >= outputOffsets[output]!
        && range.endOffset <= outputOffsets[output]! + part.length)
      if (destination < 0) throw new RangeError(`段落拆分穿过字段 ${field.id}；此字段定义已保留，请在字段结果之外换段。`)
      plans[destination]!.push({ ...field, startOffset: range.startOffset - outputOffsets[destination]!, endOffset: range.endOffset - outputOffsets[destination]! })
    }
    cursor += sourceTexts[index]!.length + 1
  }
  for (const [index, plan] of plans.entries()) validateFields(plan, projected[index]!)
  for (const [index, block] of outputs.entries()) setDocumentTextFields(block, plans[index]!, projected[index]!)
}

export function transformDocumentTextFieldsForBlockEdit(block: DocumentBlock, start: number, end: number, text: string, preserveFieldId?: string): void {
  transformDocumentTextFieldsForBlockRangeEdit([block], start, end, text, [block], preserveFieldId)
}

export function remapDocumentTextFieldIds(block: DocumentBlock): void {
  const fields = documentTextFields(block)
  if (fields.length) setDocumentTextFields(block, fields.map(field => ({ ...field, id: documentId('field') })))
}

export function insertDocumentTextFields(block: DocumentBlock, fields: readonly DocumentTextField[], offset: number): void {
  if (!fields.length) return
  setDocumentTextFields(block, [...documentTextFields(block), ...fields.map(field => ({
    ...field, startOffset: offset + field.startOffset, endOffset: offset + field.endOffset,
  }))])
}
