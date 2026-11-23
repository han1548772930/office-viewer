import type { DocumentBlock, DocumentModel, DocumentParagraphStyle } from '../../../domain/model'
import { transformContentControlsForBlockEdit, transformContentControlsForBodyRangeEdit } from './content-controls'
import { documentTextFieldParagraphs, transformDocumentTextFieldsForBlockRangeEdit } from '../model/document-text-fields'
import { transformDocumentCommentAnchorsForBodyRangeEdit } from '../comments/document-comment-anchors'
import {
  applyRunPatch,
  compactRuns,
  createParagraph,
  createRun,
  documentId,
  marksAtPoint,
  runText,
  runsForTarget,
  sameTarget,
  setTargetRuns,
  sliceRuns,
  tableForBlock,
  textTargets,
  type DocumentRun,
  type DocumentPoint,
  type TextTarget,
} from '../model/document-model'

/** A body segment with offsets in the flattened body text. */
export interface DocumentBodySegment extends TextTarget {
  index: number
  startOffset: number
  endOffset: number
  text: string
}

/** The body text and its stable mapping back to document targets. */
export interface DocumentBodySnapshot {
  text: string
  segments: DocumentBodySegment[]
}

export interface DocumentParagraphQuery {
  text?: string
  paragraphId?: string
  blockId?: string
}

export interface DocumentParagraphInfo extends DocumentBodySegment {
  paragraphId: string
  type: DocumentBlock['type']
}

/** A half-open range in the flattened main-body text. */
export interface DocumentTextRange {
  startOffset: number
  endOffset: number
  text: string
}

export type DocumentTextStyle = Partial<Pick<DocumentRun, 'bold' | 'italic' | 'underline' | 'strike' | 'verticalAlign' | 'color' | 'fontFamily' | 'fontSize' | 'highlight' | 'href'>>
export type DocumentTextStylePatch = DocumentTextStyle
export interface DocumentParagraphIdentity { paragraphId: string }

export interface DocumentTextStyleRun {
  startOffset: number
  endOffset: number
  textStyle: DocumentTextStyle
}

/**
 * Build the main-body text used by the Docs range helpers.
 *
 * Text targets are separated by one synthetic newline. The separator keeps
 * paragraph boundaries addressable while leaving the persisted model in its
 * existing block/run shape. Table-cell paragraphs participate in document
 * order after their containing table block.
 */
export function flattenDocumentBody(model: DocumentModel): DocumentBodySnapshot {
  const segments: DocumentBodySegment[] = []
  let offset = 0
  const targets = textTargets(model)
  targets.forEach((target, index) => {
    const text = runText(runsForTarget(model, target))
    const startOffset = offset
    const endOffset = startOffset + text.length
    segments.push({ ...target, index, startOffset, endOffset, text })
    offset = endOffset + (index < targets.length - 1 ? 1 : 0)
  })
  return { text: segments.map((segment) => segment.text).join('\n'), segments }
}

export function documentBodyText(model: DocumentModel): string {
  return flattenDocumentBody(model).text
}

/** Find paragraphs by literal text, stable id, or containing block id. */
export function findDocumentParagraphs(
  model: DocumentModel,
  query: string | DocumentParagraphQuery
): DocumentParagraphInfo[] {
  const normalized = typeof query === 'string' ? { text: query } : query
  const body = flattenDocumentBody(model)
  return body.segments.flatMap((segment) => {
    const block = blockForTarget(model, segment)
    if (!block || ['table', 'image', 'divider', 'pageBreak', 'columnBreak', 'sectionBreak'].includes(block.type)) return []
    const paragraphId = segment.cellBlockId ?? segment.blockId
    if (normalized.paragraphId && normalized.paragraphId !== paragraphId) return []
    if (normalized.blockId && normalized.blockId !== segment.blockId) return []
    if (normalized.text && !segment.text.includes(normalized.text)) return []
    return [{ ...segment, paragraphId, type: block.type }]
  })
}

export function findDocumentParagraph(
  model: DocumentModel,
  query: string | DocumentParagraphQuery
): DocumentParagraphInfo | undefined {
  return findDocumentParagraphs(model, query)[0]
}

export function readDocumentRange(model: DocumentModel, startOffset: number, endOffset: number): DocumentTextRange {
  const body = flattenDocumentBody(model)
  assertRange(body.text.length, startOffset, endOffset)
  return { startOffset, endOffset, text: body.text.slice(startOffset, endOffset) }
}

/** Return run styles clipped to a body range; synthetic paragraph breaks have no text style. */
export function getDocumentTextStyleRuns(model: DocumentModel, startOffset: number, endOffset: number): DocumentTextStyleRun[] {
  const body = flattenDocumentBody(model)
  assertRange(body.text.length, startOffset, endOffset)
  const result: DocumentTextStyleRun[] = []
  for (const segment of body.segments) {
    const rangeStart = Math.max(startOffset, segment.startOffset)
    const rangeEnd = Math.min(endOffset, segment.endOffset)
    if (rangeStart >= rangeEnd) continue
    let runOffset = 0
    for (const run of runsForTarget(model, segment)) {
      const runStart = segment.startOffset + runOffset
      const runEnd = runStart + run.text.length
      const clippedStart = Math.max(rangeStart, runStart)
      const clippedEnd = Math.min(rangeEnd, runEnd)
      if (clippedStart < clippedEnd) {
        result.push({
          startOffset: clippedStart,
          endOffset: clippedEnd,
          textStyle: textStyleOf(run),
        })
      }
      runOffset += run.text.length
    }
  }
  return result
}

/** Return style properties shared by every character, including unstyled paragraph breaks. */
export function getCommonDocumentTextStyle(model: DocumentModel, startOffset: number, endOffset: number): DocumentTextStyle {
  const body = flattenDocumentBody(model)
  assertRange(body.text.length, startOffset, endOffset)
  if (startOffset === endOffset) return {}

  const runs = getDocumentTextStyleRuns(model, startOffset, endOffset)
  const segments: DocumentTextStyle[] = []
  let offset = startOffset
  for (const run of runs) {
    if (offset < run.startOffset) segments.push({})
    segments.push(run.textStyle)
    offset = run.endOffset
  }
  if (offset < endOffset) segments.push({})
  if (!segments.length) return {}

  const common: Partial<DocumentTextStyle> = {}
  for (const key of textStyleKeys) {
    const value = segments[0]![key]
    if (value !== undefined && segments.every((segment) => segment[key] === value)) {
      ;(common as Record<keyof DocumentTextStyle, unknown>)[key] = value
    }
  }
  return common
}

/** Apply a text-style patch to characters in a body range, preserving surrounding runs. */
export function setDocumentTextStyle(
  model: DocumentModel,
  startOffset: number,
  endOffset: number,
  patch: DocumentTextStylePatch
): void {
  const body = flattenDocumentBody(model)
  assertRange(body.text.length, startOffset, endOffset)
  const effectivePatch = Object.fromEntries(Object.entries(patch).filter(([, value]) => value !== undefined)) as DocumentTextStylePatch
  validateTextStylePatch(effectivePatch)
  if (startOffset === endOffset || Object.keys(effectivePatch).length === 0) return

  const selectedRuns = getDocumentTextStyleRuns(model, startOffset, endOffset)
  if (!selectedRuns.some(({ textStyle }) => textStyleKeys.some((key) => key in effectivePatch && textStyle[key] !== effectivePatch[key]))) return

  for (const segment of body.segments) {
    const rangeStart = Math.max(startOffset, segment.startOffset)
    const rangeEnd = Math.min(endOffset, segment.endOffset)
    if (rangeStart >= rangeEnd) continue
    const from = rangeStart - segment.startOffset
    const to = rangeEnd - segment.startOffset
    const runs = runsForTarget(model, segment)
    const selected = sliceRuns(runs, from, to).map((run) => applyRunPatch(run, effectivePatch))
    setTargetRuns(model, segment, [
      ...sliceRuns(runs, 0, from),
      ...selected,
      ...sliceRuns(runs, to),
    ])
  }
}

export function getDocumentParagraphStyle(
  model: DocumentModel,
  identity: DocumentParagraphIdentity
): DocumentParagraphStyle | undefined {
  const reference = locateParagraphTarget(model, identity)
  return reference ? { ...reference.block.style } : undefined
}

/** Apply paragraph formatting to a top-level or table-cell paragraph by stable target identity. */
export function setDocumentParagraphStyle(
  model: DocumentModel,
  identity: DocumentParagraphIdentity,
  patch: Partial<DocumentParagraphStyle>
): void {
  const effectivePatch = Object.fromEntries(Object.entries(patch).filter(([, value]) => value !== undefined)) as Partial<DocumentParagraphStyle>
  validateParagraphStylePatch(effectivePatch)
  const reference = locateParagraphTarget(model, identity)
  if (!reference) throw new Error(`Document paragraph with id ${identity.paragraphId} not found`)
  reference.block.style = { ...reference.block.style, ...effectivePatch }
}

/** Insert plain text into a cloned model at a flattened-body offset. */
export function insertDocumentText(model: DocumentModel, index: number, text: string): DocumentModel {
  return replaceDocumentRange(model, index, index, text)
}

/** Replace a half-open flattened-body range in a cloned model. */
export function replaceDocumentRange(
  model: DocumentModel,
  startOffset: number,
  endOffset: number,
  text: string
): DocumentModel {
  const body = flattenDocumentBody(model)
  assertRange(body.text.length, startOffset, endOffset)
  if (startOffset === endOffset && text === '') return structuredCloneDocument(model)

  const next = structuredCloneDocument(model)
  const nextBody = flattenDocumentBody(next)
  if (!nextBody.segments.length) {
    if (text === '') return next
    next.blocks.splice(0, 0, ...text.replace(/\r\n?/g, '\n').split('\n').map((line) => createParagraph(line)))
    updateDocumentStats(next)
    next.modified = true
    return next
  }

  const start = pointAtOffset(nextBody, startOffset)
  const end = pointAtOffset(nextBody, endOffset)
  if (!start || !end) return next

  const startRef = locateTarget(next, start)
  const endRef = locateTarget(next, end)
  if (!startRef || !endRef) return next

  if (startRef.container !== endRef.container) {
    throw new RangeError('A document text range cannot cross independent table cells.')
  }
  if (!sameTarget(start, end) && hasUnaddressableBlocksBetween(startRef, endRef)) {
    throw new RangeError('A document text range cannot cross non-text blocks.')
  }
  const sourceParagraphs = documentTextFieldParagraphs(next)
  const sourceStart = sourceParagraphs.find(paragraph => paragraph.block.id === (start.cellBlockId ?? start.blockId))!
  const sourceEnd = sourceParagraphs.find(paragraph => paragraph.block.id === (end.cellBlockId ?? end.blockId))!
  if (sameTarget(start, end)) replaceWithinBlock(next, start, end, startRef, text)
  else {
    replaceAcrossBlocks(next, start, end, startRef, endRef, text)
    transformContentControlsForBodyRangeEdit(next, sourceParagraphs,
      sourceStart.absoluteStart + start.offset, sourceEnd.absoluteStart + end.offset, text.replace(/\r\n?/g, '\n').length, false)
  }
  transformDocumentCommentAnchorsForBodyRangeEdit(next, sourceParagraphs, documentTextFieldParagraphs(next),
    sourceStart.absoluteStart + start.offset, sourceEnd.absoluteStart + end.offset, text.replace(/\r\n?/g, '\n'))
  updateDocumentStats(next)
  next.modified = true
  return next
}

export function deleteDocumentRange(model: DocumentModel, startOffset: number, endOffset: number): DocumentModel {
  return replaceDocumentRange(model, startOffset, endOffset, '')
}

/** Return a detached, export-safe snapshot without changing the live model. */
export function saveDocumentSnapshot(model: DocumentModel, name?: string): DocumentModel {
  const snapshot = structuredCloneDocument(model)
  snapshot.name = safeDocumentName(name ?? snapshot.name)
  snapshot.modified = false
  updateDocumentStats(snapshot)
  return snapshot
}

/** Sanitize a document title before it is used as a file name or snapshot name. */
export function safeDocumentName(value: string, fallback = 'Untitled document'): string {
  const cleaned = value
    .replace(/[\u0000-\u001f\u007f]/g, '')
    .replace(/[<>:"/\\|?*]/g, '-')
    .replace(/[. ]+$/g, '')
    .trim()
  if (!cleaned || cleaned === '.' || cleaned === '..') return fallback
  const truncated = Array.from(cleaned).slice(0, 180).join('').replace(/[. ]+$/g, '').trim()
  if (!truncated) return fallback
  return /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i.test(truncated)
    ? `document-${truncated}`
    : truncated
}

function structuredCloneDocument(model: DocumentModel): DocumentModel {
  return JSON.parse(JSON.stringify(model)) as DocumentModel
}

const textStyleKeys = ['bold', 'italic', 'underline', 'strike', 'verticalAlign', 'color', 'fontFamily', 'fontSize', 'highlight', 'href'] as const satisfies ReadonlyArray<keyof DocumentTextStyle>

function textStyleOf(run: DocumentRun): DocumentTextStyle {
  const { bold, italic, underline, strike, verticalAlign, color, fontFamily, fontSize, highlight, href } = run
  return { bold, italic, underline, strike, ...(verticalAlign === undefined ? {} : { verticalAlign }), color, fontFamily, fontSize, ...(highlight === undefined ? {} : { highlight }), ...(href === undefined ? {} : { href }) }
}

function validateTextStylePatch(patch: DocumentTextStylePatch): void {
  if (patch.fontSize !== undefined && (!Number.isFinite(patch.fontSize) || patch.fontSize <= 0)) {
    throw new RangeError('Font size must be a positive number')
  }
  if (patch.verticalAlign !== undefined && !['baseline', 'subscript', 'superscript'].includes(patch.verticalAlign)) {
    throw new TypeError(`Unsupported vertical alignment: ${patch.verticalAlign}`)
  }
}

function validateParagraphStylePatch(patch: Partial<DocumentParagraphStyle>): void {
  if (patch.align !== undefined && !['left', 'center', 'right', 'justify'].includes(patch.align)) {
    throw new TypeError(`Unsupported paragraph alignment: ${patch.align}`)
  }
  for (const key of ['indent', 'lineHeight', 'spaceBefore', 'spaceAfter'] as const) {
    const value = patch[key]
    if (value !== undefined && !Number.isFinite(value)) throw new RangeError(`Paragraph ${key} must be finite`)
  }
  if (patch.lineHeight !== undefined && patch.lineHeight <= 0) throw new RangeError('Paragraph line height must be positive')
}

function locateParagraphTarget(model: DocumentModel, identity: DocumentParagraphIdentity): TargetReference | undefined {
  const target = textTargets(model).find((candidate) => (candidate.cellBlockId ?? candidate.blockId) === identity.paragraphId)
  if (!target) return undefined
  const reference = locateTarget(model, { ...target, offset: 0 })
  if (!reference || ['table', 'image', 'divider', 'pageBreak', 'columnBreak', 'sectionBreak'].includes(reference.block.type)) return undefined
  return reference
}

function assertRange(length: number, startOffset: number, endOffset: number): void {
  if (!Number.isInteger(startOffset) || !Number.isInteger(endOffset) || startOffset < 0 || endOffset < startOffset || endOffset > length) {
    throw new RangeError(`Invalid document text range [${startOffset}, ${endOffset}) for body length ${length}.`)
  }
}

function pointAtOffset(body: DocumentBodySnapshot, offset: number): DocumentPoint | undefined {
  if (!body.segments.length) return undefined
  for (const segment of body.segments) {
    if (offset <= segment.endOffset) return { ...segment, offset: offset - segment.startOffset }
    if (offset < segment.startOffset) return { ...segment, offset: 0 }
  }
  const last = body.segments.at(-1)!
  return { ...last, offset: last.text.length }
}

interface TargetReference {
  block: DocumentBlock
  container: DocumentBlock[]
  index: number
}

function hasUnaddressableBlocksBetween(start: TargetReference, end: TargetReference): boolean {
  return start.container
    .slice(start.index + 1, end.index)
    .some((block) => !['paragraph', 'heading', 'bullet', 'numbered', 'checklist', 'quote'].includes(block.type))
}

function locateTarget(model: DocumentModel, point: DocumentPoint): TargetReference | undefined {
  const block = model.blocks.find((candidate) => candidate.id === point.blockId)
  if (!block) return undefined
  if (point.row == null) {
    const index = model.blocks.findIndex((candidate) => candidate.id === point.blockId)
    return index < 0 ? undefined : { block, container: model.blocks, index }
  }
  const cell = tableForBlock(block)?.rows[point.row]?.cells[point.col ?? 0]
  if (!cell) return undefined
  const targetId = point.cellBlockId ?? cell.blocks[0]?.id
  const index = cell.blocks.findIndex((candidate) => candidate.id === targetId)
  const target = cell.blocks[index]
  return target && index >= 0 ? { block: target, container: cell.blocks, index } : undefined
}

function blockForTarget(model: DocumentModel, target: TextTarget): DocumentBlock | undefined {
  const block = model.blocks.find((candidate) => candidate.id === target.blockId)
  if (target.row == null || !block) return block
  const cell = tableForBlock(block)?.rows[target.row]?.cells[target.col ?? 0]
  return cell?.blocks.find((candidate) => candidate.id === target.cellBlockId)
}

function replaceAcrossBlocks(
  model: DocumentModel,
  start: DocumentPoint,
  end: DocumentPoint,
  startRef: TargetReference,
  endRef: TargetReference,
  text: string
): void {
  const startRuns = runsForTarget(model, start)
  const endRuns = runsForTarget(model, end)
  const prefix = sliceRuns(startRuns, 0, start.offset)
  const suffix = sliceRuns(endRuns, end.offset)
  const inherited = marksAtPoint(model, start)
  const lines = text.replace(/\r\n?/g, '\n').split('\n')
  const first = startRef.block

  const additions = lines.slice(1).map((line, index) => {
    const block = paragraphAfter(first, inherited)
    const isLast = index === lines.length - 2
    block.runs = compactRuns([...(line ? [createRun(line, inherited)] : []), ...(isLast ? suffix : [])])
    return block
  })

  const removeCount = endRef.index - startRef.index
  transformDocumentTextFieldsForBlockRangeEdit(startRef.container.slice(startRef.index, endRef.index + 1), start.offset, end.offset, lines.join('\n'), [first, ...additions], undefined, model)
  setTargetRuns(model, start, [
    ...prefix,
    ...(lines[0] ? [createRun(lines[0], inherited)] : []),
    ...(lines.length === 1 ? suffix : []),
  ])
  if (removeCount > 0) startRef.container.splice(startRef.index + 1, removeCount, ...additions)
  else if (additions.length) startRef.container.splice(startRef.index + 1, 0, ...additions)
}

function replaceWithinBlock(
  model: DocumentModel,
  start: DocumentPoint,
  end: DocumentPoint,
  reference: TargetReference,
  text: string
): void {
  const runs = runsForTarget(model, start)
  const prefix = sliceRuns(runs, 0, start.offset)
  const suffix = sliceRuns(runs, end.offset)
  const inherited = marksAtPoint(model, start)
  const lines = text.replace(/\r\n?/g, '\n').split('\n')
  const additions = lines.slice(1).map((line, index) => {
    const block = paragraphAfter(reference.block, inherited)
    const isLast = index === lines.length - 2
    block.runs = compactRuns([...(line ? [createRun(line, inherited)] : []), ...(isLast ? suffix : [])])
    return block
  })
  const normalizedText = lines.join('\n')
  transformContentControlsForBlockEdit(model, start.cellBlockId ?? start.blockId, start.offset, end.offset, normalizedText.length,
    additions.length ? {
      blockId: additions.at(-1)!.id,
      firstBreakOffset: start.offset + normalizedText.indexOf('\n'),
      lastParagraphOffset: start.offset + normalizedText.lastIndexOf('\n') + 1,
    } : undefined)
  transformDocumentTextFieldsForBlockRangeEdit([reference.block], start.offset, end.offset, normalizedText, [reference.block, ...additions], undefined, model)
  setTargetRuns(model, start, [
    ...prefix,
    ...(lines[0] ? [createRun(lines[0], inherited)] : []),
    ...(lines.length === 1 ? suffix : []),
  ])
  reference.container.splice(reference.index + 1, 0, ...additions)
}

function paragraphAfter(source: DocumentBlock, marks: ReturnType<typeof marksAtPoint>): DocumentBlock {
  const block = createParagraph('', marks)
  block.style = { ...source.style }
  if (['bullet', 'numbered', 'checklist', 'quote'].includes(source.type)) {
    block.type = source.type
    block.level = source.level
    if (source.type === 'checklist') block.checklist = { listId: source.checklist?.listId ?? documentId('checklist'), checked: false }
  }
  return block
}

function updateDocumentStats(model: DocumentModel): void {
  model.wordCount = textTargets(model).reduce((count, target) => count + runText(runsForTarget(model, target)).replace(/\s/g, '').length, 0)
}
