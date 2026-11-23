import type { DocumentBlock, DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import type { DocumentSelection } from '../model/document-model'
import { blockAtPoint, createParagraph, createRun, documentId, orderedSelection, runText, sameTarget, sliceRuns, textTargets, type DocumentPoint, type DocumentRun } from '../model/document-model'
import { documentDefaultStyle } from '../model/document-default-style'
import { documentRunStyleAt, splitDocumentRunRange } from '../model/document-run-range'
import { isDocumentTextBoundary, remapDocumentStoryFieldRuns } from '../model/document-story'
import { documentFieldLocked, documentTextFieldParagraphs, documentTextFieldRangeContainsNested, documentTextFieldRangeIntersects, documentTextFieldRanges, documentTextFields, setDocumentTextFields, transformDocumentTextFieldsForBlockEdit, type DocumentTextFieldParagraph, type DocumentTextFieldRange } from '../model/document-text-fields'
import { transformContentControlsForBlockEdit } from '../services/content-controls'
import { transformDocumentCommentAnchorsForTextEdit } from '../comments/thread-comments'

export type DocumentTocFormat = 'fromTemplate' | 'classic' | 'modern' | 'simple'
export type DocumentTocTabLeader = 'none' | 'dots' | 'dashes' | 'underline'
export type DocumentTocUpdateMode = 'pageNumbersOnly' | 'all'

export interface DocumentTocOptions {
  levels: number
  showPageNumbers: boolean
  rightAlignPageNumbers: boolean
  tabLeader: DocumentTocTabLeader
  format: DocumentTocFormat
}

export const DEFAULT_DOCUMENT_TOC_OPTIONS: DocumentTocOptions = {
  levels: 3,
  showPageNumbers: true,
  rightAlignPageNumbers: true,
  tabLeader: 'dots',
  format: 'fromTemplate',
}

export interface DocumentTocHeading {
  id: string
  text: string
  level: number
}

export type DocumentTocPageNumbers = Readonly<Record<string, number | undefined>>

export interface DocumentTocPageNumberUpdatePlan {
  blocks: DocumentModel['blocks']
  contentControls?: DocumentModel['contentControls']
  comments?: DocumentModel['comments']
  selection?: DocumentSelection
  changed: boolean
}

interface DocumentTocMetadata {
  id: string
  role: 'title' | 'entry'
  headingId?: string
  headingText?: string
  titleText?: string
  level?: number
  pageNumber?: number
  options?: DocumentTocOptions
}

const TOC_METADATA_KEY = 'documentToc'

function normalizeTocOptions(options: Partial<DocumentTocOptions>): DocumentTocOptions {
  const showPageNumbers = options.showPageNumbers ?? true
  return {
    levels: typeof options.levels === 'number' && Number.isFinite(options.levels) ? Math.min(9, Math.max(1, Math.trunc(options.levels))) : 3,
    showPageNumbers,
    rightAlignPageNumbers: showPageNumbers && (options.rightAlignPageNumbers ?? true),
    tabLeader: options.tabLeader === 'none' || options.tabLeader === 'dashes' || options.tabLeader === 'underline' ? options.tabLeader : 'dots',
    format: options.format === 'classic' || options.format === 'modern' || options.format === 'simple' ? options.format : 'fromTemplate',
  }
}

function tocMetadata(block: DocumentBlock): DocumentTocMetadata | undefined {
  const value = block.metadata[TOC_METADATA_KEY]
  if (!value || typeof value !== 'object') return undefined
  const metadata = value as Partial<DocumentTocMetadata>
  return typeof metadata.id === 'string' && (metadata.role === 'title' || metadata.role === 'entry')
    ? metadata as DocumentTocMetadata
    : undefined
}

function displayText(heading: DocumentTocHeading, options: DocumentTocOptions, pageNumber: number | undefined): string {
  if (!options.showPageNumbers || pageNumber == null) return heading.text
  return options.rightAlignPageNumbers ? `${heading.text}\t${pageNumber}` : `${heading.text} ${pageNumber}`
}

function createTocBlock(
  tocId: string,
  role: DocumentTocMetadata['role'],
  text: string,
  options: DocumentTocOptions,
  details: Pick<DocumentTocMetadata, 'headingId' | 'headingText' | 'titleText' | 'level' | 'pageNumber'> = {},
  id = documentId('toc'),
): DocumentBlock {
  const block = createParagraph()
  block.id = id
  block.runs = [createRun(text, role === 'title' ? { bold: true, fontSize: options.format === 'modern' ? 18 : 16 } : {})]
  block.style = {
    ...block.style,
    indent: role === 'entry' ? Math.max(0, (details.level ?? 1) - 1) * (options.format === 'classic' ? 14 : options.format === 'simple' ? 12 : options.format === 'modern' ? 20 : 18) : 0,
    spaceAfter: options.format === 'modern' ? 6 : 2,
    lineHeight: options.format === 'simple' ? 1.2 : 1.35,
  }
  block.metadata[TOC_METADATA_KEY] = {
    id: tocId,
    role,
    ...details,
    options: { ...options },
  } satisfies DocumentTocMetadata
  return block
}

export function collectDocumentTocHeadings(model: DocumentModel, levels: number): DocumentTocHeading[] {
  const maxLevel = normalizeTocOptions({ levels }).levels
  const seen = new Set<string>()
  return model.blocks.flatMap((block) => {
    if (block.type !== 'heading' || !block.id || seen.has(block.id) || !Number.isInteger(block.level) || block.level < 1 || block.level > maxLevel) return []
    const text = runText(block.runs).replace(/[\r\n\t]/g, ' ').trim()
    const listLabel = block.metadata.listLabel
    const prefix = typeof listLabel === 'string' ? listLabel.trim() : ''
    if (text) seen.add(block.id)
    return text ? [{ id: block.id, text: prefix ? `${prefix} ${text}` : text, level: block.level }] : []
  })
}

export function buildDocumentTocBlocks(
  model: DocumentModel,
  options: DocumentTocOptions,
  pageNumbers: DocumentTocPageNumbers = {},
  tocId = documentId('document-toc'),
  existingBlocks: readonly DocumentBlock[] = [],
  titleText = 'Contents',
  includeTitle = true,
): DocumentBlock[] {
  options = normalizeTocOptions(options)
  const headings = collectDocumentTocHeadings(model, options.levels)
  if (!headings.length) return []
  const existingIds = new Map(existingBlocks.flatMap((block) => {
    const metadata = tocMetadata(block)
    return metadata?.role === 'entry' && metadata.headingId ? [[metadata.headingId, block.id] as const] : []
  }))
  const oldTitleBlock = existingBlocks.find((block) => tocMetadata(block)?.role === 'title')
  const titleBlocks = includeTitle
    ? [createTocBlock(tocId, 'title', titleText, options, { titleText }, oldTitleBlock?.id)]
    : []
  const entries = headings.map((heading) => {
    const pageNumber = pageNumbers[heading.id]
    return createTocBlock(tocId, 'entry', displayText(heading, options, pageNumber), options, {
      headingId: heading.id,
      headingText: heading.text,
      level: heading.level,
      ...(pageNumber == null ? {} : { pageNumber }),
    }, existingIds.get(heading.id))
  })
  return [...titleBlocks, ...entries]
}

export function updateDocumentToc(
  model: DocumentModel,
  tocId: string,
  mode: DocumentTocUpdateMode,
  pageNumbers: DocumentTocPageNumbers = {},
): boolean {
  const tocBlocks = model.blocks.filter((block) => tocMetadata(block)?.id === tocId)
  if (!tocBlocks.length) return false

  if (mode === 'pageNumbersOnly') {
    const plan = prepareDocumentTocPageNumberUpdate(model, tocId, pageNumbers)
    if (!plan) return false
    applyDocumentTocPageNumberUpdate(model, plan)
    return true
  }

  const titleBlock = tocBlocks.find((block) => tocMetadata(block)?.role === 'title')
  const titleMetadata = titleBlock && tocMetadata(titleBlock)
  const entryMetadata = tocBlocks.map(tocMetadata).find((metadata) => metadata?.role === 'entry')
  const options = titleMetadata?.options ?? entryMetadata?.options
  if (!options) return false
  const replacement = buildDocumentTocBlocks(model, options, pageNumbers, tocId, tocBlocks, titleMetadata?.titleText ?? 'Contents', Boolean(titleMetadata))
  if (!replacement.length) return false
  const firstIndex = model.blocks.findIndex((block) => tocMetadata(block)?.id === tocId)
  model.blocks = model.blocks.filter((block) => tocMetadata(block)?.id !== tocId)
  model.blocks.splice(firstIndex, 0, ...replacement)
  return true
}

function fieldResultText(range: DocumentTextFieldRange, paragraphs: readonly DocumentTextFieldParagraph[]): string {
  return paragraphs.filter(paragraph => paragraph.absoluteStart >= range.startParagraph.absoluteStart
    && paragraph.absoluteStart <= range.endParagraph.absoluteStart).map(paragraph => paragraph.text.slice(
    paragraph.block.id === range.block.id ? range.field.startOffset : 0,
    paragraph.block.id === range.endParagraph.block.id ? range.field.endOffset : undefined,
  )).join('\n')
}

function mapTocPagePoint(point: DocumentPoint, blockId: string, start: number, end: number, length: number): DocumentPoint {
  if ((point.cellBlockId ?? point.blockId) !== blockId || point.offset <= start) return point
  return { ...point, offset: point.offset >= end ? point.offset + length - (end - start) : start + Math.min(point.offset - start, length) }
}

/** Resolve every page before publishing, preserving native field definitions and ordinary anchors. */
export function prepareDocumentTocPageNumberUpdate(
  model: DocumentModel,
  tocId: string,
  pageNumbers: DocumentTocPageNumbers = {},
  selection?: DocumentSelection,
  timestamp = Date.now(),
): DocumentTocPageNumberUpdatePlan | undefined {
  if (typeof timestamp !== 'number' || !Number.isFinite(new Date(timestamp).getTime())) throw new TypeError('目录更新时间无效。')
  const tocBlocks = model.blocks.filter(block => tocMetadata(block)?.id === tocId)
  if (!tocBlocks.length) return undefined
  if (selection) for (const point of [selection.anchor, selection.focus]) {
    const block = blockAtPoint(model, point)
    if (!block || !isDocumentTextBoundary(runText(block.runs), point.offset)) throw new RangeError('目录更新选区的 UTF-16 位置无效。')
  }
  const paragraphs = documentTextFieldParagraphs(model), ranges = documentTextFieldRanges(model)
  const tocIds = new Set(tocBlocks.map(block => block.id))
  const outerFields = ranges.filter(range => range.field.field.toUpperCase() === 'TOC'
    && paragraphs.some(paragraph => tocIds.has(paragraph.block.id) && paragraph.absoluteStart >= range.startParagraph.absoluteStart
      && paragraph.absoluteStart <= range.endParagraph.absoluteStart))
  if (outerFields.some(range => documentFieldLocked(range.field))) return undefined
  const replacements: Array<{ blockId: string; metadata: DocumentTocMetadata; pageNumber: number; start: number; end: number; text: string; fieldId?: string }> = []
  for (const block of tocBlocks) {
    const metadata = tocMetadata(block)
    if (metadata?.role !== 'entry' || !metadata.headingId || !metadata.options?.showPageNumbers) continue
    const pageFields = ranges.filter(range => range.block.id === block.id && range.field.field.toUpperCase() === 'PAGEREF')
    // Multiple or cross-paragraph references need a target mapping beyond a single local TOC row.
    if (pageFields.length > 1) continue
    const pageField = pageFields[0]
    if (pageField && (documentFieldLocked(pageField.field) || documentTextFieldRangeContainsNested(pageField, ranges)
      || pageField.endParagraph.block.id !== block.id)) continue
    const text = runText(block.runs), separator = metadata.options.rightAlignPageNumbers ? '\t' : ' '
    const separatorIndex = metadata.pageNumber == null ? -1 : text.lastIndexOf(separator)
    const start = pageField?.field.startOffset ?? (separatorIndex >= 0 ? separatorIndex + 1 : text.length)
    const end = pageField?.field.endOffset ?? text.length
    const paragraph = paragraphs.find(paragraph => paragraph.block.id === block.id)
    if (!paragraph) throw new TypeError('目录页码必须属于文本段落。')
    const absoluteStart = paragraph.absoluteStart + start, absoluteEnd = paragraph.absoluteStart + end
    const overlapping = ranges.filter(range => documentTextFieldRangeIntersects(range, absoluteStart, absoluteEnd))
    if (overlapping.some(range => documentFieldLocked(range.field)
      || range.field.field.toUpperCase() !== 'TOC' && range.field.id !== pageField?.field.id)) continue
    const pageNumber = pageNumbers[metadata.headingId]
    if (pageNumber == null || !Number.isSafeInteger(pageNumber) || pageNumber < 1) return undefined
    const pageText = String(pageNumber), result = !pageField && separatorIndex < 0 ? `${separator}${pageText}` : pageText
    if (text.slice(start, end) === result && metadata.pageNumber === pageNumber
      && (!pageField || pageField.field.cachedResult === pageText)) continue
    replacements.push({ blockId: block.id, metadata, pageNumber, start, end, text: result, fieldId: pageField?.field.id })
  }
  const repairsOuterCache = outerFields.some(range => range.field.cachedResult !== fieldResultText(range, paragraphs))
  if (!replacements.length && !repairsOuterCache) return { blocks: model.blocks, changed: false, ...(selection ? { selection } : {}) }
  const candidate: DocumentModel = { ...model, blocks: cloneModel(model.blocks),
    contentControls: model.contentControls && cloneModel(model.contentControls), comments: model.comments && cloneModel(model.comments) }
  let nextSelection = selection && cloneModel(selection)
  for (const replacement of replacements) {
    const block = candidate.blocks.find(block => block.id === replacement.blockId)!
    const textChanged = runText(block.runs).slice(replacement.start, replacement.end) !== replacement.text
    if (textChanged) {
      const runs = block.runs as DocumentRun[], split = splitDocumentRunRange(runs, replacement.start, replacement.end)
      const cache = sliceRuns(runs, replacement.start, replacement.end)
      const inherited = documentRunStyleAt(runs, replacement.start, replacement.end) ?? documentDefaultStyle(candidate).text
      const result = remapDocumentStoryFieldRuns(cache.length ? cache : [createRun('', inherited)], replacement.text, documentDefaultStyle(candidate).text)
      transformContentControlsForBlockEdit(candidate, block.id, replacement.start, replacement.end, replacement.text.length, undefined, true)
      transformDocumentTextFieldsForBlockEdit(block, replacement.start, replacement.end, replacement.text, replacement.fieldId, candidate)
      block.runs = [...split.before, ...result, ...split.after]
      delete block.metadata.html
      transformDocumentCommentAnchorsForTextEdit(candidate, block.id, replacement.start, replacement.end, replacement.text, timestamp)
      if (nextSelection) nextSelection = {
        anchor: mapTocPagePoint(nextSelection.anchor, block.id, replacement.start, replacement.end, replacement.text.length),
        focus: mapTocPagePoint(nextSelection.focus, block.id, replacement.start, replacement.end, replacement.text.length),
      }
    }
    block.metadata[TOC_METADATA_KEY] = { ...replacement.metadata, pageNumber: replacement.pageNumber }
    if (replacement.fieldId) setDocumentTextFields(block, documentTextFields(block).map(field => field.id === replacement.fieldId
      ? { ...field, cachedResult: String(replacement.pageNumber) } : field))
  }
  const currentParagraphs = documentTextFieldParagraphs(candidate)
  for (const outer of outerFields) {
    const range = documentTextFieldRanges(candidate).find(range => range.field.id === outer.field.id)!
    const cachedResult = fieldResultText(range, currentParagraphs)
    if (range.field.cachedResult !== cachedResult) setDocumentTextFields(range.block, documentTextFields(range.block).map(field => field.id === range.field.id
      ? { ...field, cachedResult } : field))
  }
  documentTextFieldRanges(candidate)
  return { blocks: candidate.blocks, contentControls: candidate.contentControls, comments: candidate.comments, changed: true,
    ...(nextSelection ? { selection: nextSelection } : {}) }
}

/** Apply an already validated plan; unchanged run objects stay intact for cache-only repairs. */
export function applyDocumentTocPageNumberUpdate(model: DocumentModel, plan: DocumentTocPageNumberUpdatePlan): void {
  if (!plan.changed) return
  const prepared = new Map(plan.blocks.map(block => [block.id, block]))
  for (const block of model.blocks) {
    const next = prepared.get(block.id)!
    if (JSON.stringify(block.runs) !== JSON.stringify(next.runs)) block.runs = next.runs
    if (JSON.stringify(block.metadata) !== JSON.stringify(next.metadata)) block.metadata = next.metadata
  }
  if (JSON.stringify(model.contentControls) !== JSON.stringify(plan.contentControls)) model.contentControls = plan.contentControls
  if (JSON.stringify(model.comments) !== JSON.stringify(plan.comments)) model.comments = plan.comments
}

export function deleteDocumentToc(model: DocumentModel, tocId: string): boolean {
  const next = model.blocks.filter((block) => tocMetadata(block)?.id !== tocId)
  if (next.length === model.blocks.length) return false
  model.blocks = next
  return true
}

export function documentTocIdAtBlock(block: DocumentBlock | undefined): string | undefined {
  return block ? tocMetadata(block)?.id : undefined
}

export function documentTocHeadingTarget(model: DocumentModel, block: DocumentBlock | undefined): string | undefined {
  const metadata = block && tocMetadata(block)
  const headingId = metadata?.role === 'entry' ? metadata.headingId : undefined
  if (!headingId) return undefined
  return model.blocks.some(candidate => candidate.id === headingId && candidate.type === 'heading') ? headingId : undefined
}

export function documentTocOptionsAtBlock(block: DocumentBlock | undefined): DocumentTocOptions | undefined {
  const options = block && tocMetadata(block)?.options
  return options && normalizeTocOptions(options)
}

export function documentTocIdForSelection(model: DocumentModel, selection: DocumentSelection): string | undefined {
  if (selection.anchor.row != null || selection.focus.row != null) return undefined
  const { start, end } = orderedSelection(model, selection)
  const targets = textTargets(model)
  const first = targets.findIndex((target) => sameTarget(target, start))
  const last = targets.findIndex((target) => sameTarget(target, end))
  if (first < 0 || last < first) return undefined
  const endExclusive = last > first && end.offset === 0 ? last : last + 1
  const selectedBlocks = targets.slice(first, endExclusive).map((target) => model.blocks.find((block) => block.id === target.blockId))
  const tocIds = new Set(selectedBlocks.map(documentTocIdAtBlock))
  if (tocIds.size !== 1 || tocIds.has(undefined)) return undefined
  return [...tocIds][0]
}

export function documentTocPageNumbers(headings: readonly { id: string; page: number }[]): DocumentTocPageNumbers {
  return Object.fromEntries(headings.flatMap(({ id, page }) => Number.isInteger(page) && page >= 0 ? [[id, page + 1]] : []))
}

export function supportsDocumentTableOfContents(model: DocumentModel): boolean {
  const source = model.metadata?.source
  const documentStyle = model.metadata?.documentStyle
  const sourceStyle = source && typeof source === 'object' && !Array.isArray(source)
    ? (source as Record<string, unknown>).documentStyle
    : undefined
  const flavor = model.metadata?.documentFlavor
    ?? (documentStyle && typeof documentStyle === 'object' && !Array.isArray(documentStyle) ? (documentStyle as Record<string, unknown>).documentFlavor : undefined)
    ?? (sourceStyle && typeof sourceStyle === 'object' && !Array.isArray(sourceStyle) ? (sourceStyle as Record<string, unknown>).documentFlavor : undefined)
  return flavor == null || flavor === 0 || flavor === 1 || (typeof flavor === 'string' && flavor.toLowerCase() === 'traditional')
}
