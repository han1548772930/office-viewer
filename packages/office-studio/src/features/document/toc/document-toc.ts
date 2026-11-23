import type { DocumentBlock, DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import type { DocumentSelection } from '../model/document-model'
import { blockAtPoint, createParagraph, createRun, documentId, orderedSelection, runText, sameTarget, sliceRuns, textTargets, type DocumentPoint, type DocumentRun } from '../model/document-model'
import { documentDefaultStyle } from '../model/document-default-style'
import { documentRunStyleAt, splitDocumentRunRange } from '../model/document-run-range'
import { isDocumentTextBoundary, remapDocumentStoryFieldRuns } from '../model/document-story'
import { documentFieldLocked, documentFieldRetainsNestedFields, documentTextFieldParagraphs, documentTextFieldRangeContainsNested, documentTextFieldRangeIntersects, documentTextFieldRanges, documentTextFields, setDocumentTextFields, transformDocumentTextFieldsForBlockEdit, type DocumentTextField, type DocumentTextFieldParagraph, type DocumentTextFieldRange } from '../model/document-text-fields'
import { transformContentControlsForBlockEdit, validateDocumentContentControlRanges } from '../services/content-controls'
import { transformDocumentCommentAnchorsForTextEdit } from '../comments/document-comment-anchors'
import { applyDocumentTocMutation, prepareDocumentTocBlockMutation, reconcileDocumentTocFieldParents, type DocumentTocMutationPlan } from './document-toc-edit'
export { applyDocumentTocMutation, type DocumentTocMutationPlan } from './document-toc-edit'

export type DocumentTocFormat = 'fromTemplate' | 'classic' | 'modern' | 'simple'
export type DocumentTocTabLeader = 'none' | 'dots' | 'dashes' | 'underline'
export type DocumentTocUpdateMode = 'pageNumbersOnly' | 'all'

export interface DocumentTocOptions {
  levels: number
  /** Native TOC \\o ranges can start above outline level one. */
  minimumLevel?: number
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

export type DocumentTocPageNumberUpdatePlan = DocumentTocMutationPlan

export interface DocumentNativeTocMetadata {
  fieldId: string
  instruction?: string
  hyperlinks: boolean
  fullUpdateUnavailableReason?: string
  pageUpdateUnavailableReason?: string
}

export interface DocumentTocMetadata {
  id: string
  role: 'title' | 'entry'
  headingId?: string
  headingText?: string
  titleText?: string
  level?: number
  pageNumber?: number
  options?: DocumentTocOptions
  native?: DocumentNativeTocMetadata
}

const TOC_METADATA_KEY = 'documentToc'

function normalizeTocOptions(options: Partial<DocumentTocOptions>): DocumentTocOptions {
  const showPageNumbers = options.showPageNumbers ?? true
  const minimumLevel = typeof options.minimumLevel === 'number' && Number.isFinite(options.minimumLevel)
    ? Math.min(9, Math.max(1, Math.trunc(options.minimumLevel))) : 1
  return {
    levels: typeof options.levels === 'number' && Number.isFinite(options.levels) ? Math.min(9, Math.max(minimumLevel, Math.trunc(options.levels))) : Math.max(minimumLevel, 3),
    ...(minimumLevel > 1 ? { minimumLevel } : {}),
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

export function buildDocumentTocInstruction(options: DocumentTocOptions): string {
  const normalized = normalizeTocOptions(options)
  return `TOC ${[`\\o "${normalized.minimumLevel ?? 1}-${normalized.levels}"`, '\\h', '\\z', '\\u',
    ...(!normalized.showPageNumbers ? ['\\n'] : []),
    ...(normalized.showPageNumbers && !normalized.rightAlignPageNumbers ? ['\\p " "'] : [])].join(' ')}`
}

function pageReferenceInstruction(headingId: string): string {
  return `PAGEREF "${headingId.replaceAll('"', '\\"')}" \\h`
}

function pageFieldWithParent(field: DocumentTextField, parentFieldId: string): DocumentTextField {
  const source = field.metadata?.source
  return { ...field, metadata: { ...field.metadata, source: {
    ...(source !== null && typeof source === 'object' && !Array.isArray(source) ? source : {}),
    parentFieldId, nestedInInstruction: false,
  } } }
}

/** Generated results belong to the final restored TOC, including an outside owner. */
function assignGeneratedTocPageParents(model: DocumentModel, pageIds: ReadonlySet<string>, outerIds?: ReadonlySet<string>): void {
  if (!pageIds.size) return
  const ranges = documentTextFieldRanges(model)
  const outers = ranges.filter(range => range.field.field.toUpperCase() === 'TOC' && (!outerIds || outerIds.has(range.field.id)))
  for (const page of ranges.filter(range => pageIds.has(range.field.id))) {
    const parent = outers.filter(range => range.absoluteStart <= page.absoluteStart && range.absoluteEnd >= page.absoluteEnd)
      .sort((left, right) => left.absoluteEnd - left.absoluteStart - (right.absoluteEnd - right.absoluteStart))[0]
    if (!parent) continue
    setDocumentTextFields(page.block, documentTextFields(page.block).map(field => field.id === page.field.id ? pageFieldWithParent(field, parent.field.id) : field))
    setDocumentTextFields(parent.block, documentTextFields(parent.block).map(field => field.id === parent.field.id ? { ...field, containsNestedFields: true } : field))
  }
}

function installDocumentTocFields(blocks: DocumentBlock[], options: DocumentTocOptions, existingBlocks: readonly DocumentBlock[]): void {
  const entries = blocks.filter(block => tocMetadata(block)?.role === 'entry')
  if (!entries.length) return
  const oldOuterOwner = existingBlocks.find(block => documentTextFields(block).some(field => field.field.toUpperCase() === 'TOC'))
  const oldOuter = oldOuterOwner && documentTextFields(oldOuterOwner).find(field => field.field.toUpperCase() === 'TOC')
  const first = oldOuterOwner && tocMetadata(oldOuterOwner)?.role === 'title' ? blocks[0]! : entries[0]!
  const last = entries.at(-1)!
  const { endBlockId: _oldEndBlockId, ...template } = oldOuter ?? {}
  const outer = { ...cloneModel(template), id: oldOuter?.id ?? documentId('field'), field: 'TOC',
    instruction: oldOuter?.instruction ?? buildDocumentTocInstruction(options), startOffset: 0, endOffset: runText(last.runs).length,
    cachedResult: blocks.slice(blocks.indexOf(first)).map(block => runText(block.runs)).join('\n'),
    containsNestedFields: options.showPageNumbers && entries.some(block => tocMetadata(block)?.pageNumber != null),
    ...(first.id === last.id ? {} : { endBlockId: last.id }) }
  for (const block of blocks) setDocumentTextFields(block, [])
  setDocumentTextFields(first, [outer])
  for (const block of entries) {
    const metadata = tocMetadata(block)!
    if (!options.showPageNumbers || metadata.pageNumber == null || !metadata.headingId) continue
    const previous = existingBlocks.find(old => tocMetadata(old)?.headingId === metadata.headingId)
    const oldPage = previous && documentTextFields(previous).find(field => field.field.toUpperCase() === 'PAGEREF'
      && !documentFieldRetainsNestedFields(field) && field.containsNestedFields !== true)
    const { endBlockId: _pageEndBlockId, ...pageTemplate } = oldPage ?? {}
    const end = runText(block.runs).length
    setDocumentTextFields(block, [...documentTextFields(block), pageFieldWithParent({ ...cloneModel(pageTemplate), id: oldPage?.id ?? documentId('field'), field: 'PAGEREF',
      instruction: oldPage?.instruction ?? pageReferenceInstruction(metadata.headingId),
      startOffset: end - String(metadata.pageNumber).length, endOffset: end, cachedResult: String(metadata.pageNumber), containsNestedFields: false }, outer.id)])
  }
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
    indent: 0,
    indentStart: role === 'entry' ? Math.max(0, (details.level ?? 1) - 1) * (options.format === 'classic' ? 14 : options.format === 'simple' ? 12 : options.format === 'modern' ? 20 : 18) : 0,
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

export function collectDocumentTocHeadings(model: DocumentModel, levels: number, minimumLevel = 1): DocumentTocHeading[] {
  const normalized = normalizeTocOptions({ levels, minimumLevel }), maxLevel = normalized.levels, minLevel = normalized.minimumLevel ?? 1
  const seen = new Set<string>()
  return model.blocks.flatMap((block) => {
    if (block.type !== 'heading' || tocMetadata(block) || !block.id || seen.has(block.id) || !Number.isInteger(block.level) || block.level < minLevel || block.level > maxLevel) return []
    const text = runText(block.runs).replace(/[\r\n\t]/g, ' ').trim()
    const listLabel = block.metadata.listLabel
    const prefix = typeof listLabel === 'string' ? listLabel.trim() : ''
    if (text) seen.add(block.id)
    return text ? [{ id: block.id, text: prefix ? `${prefix} ${text}` : text, level: block.level }] : []
  })
}

function remapTocEntryRuns(previous: DocumentBlock, text: string, options: DocumentTocOptions, pageNumber: number | undefined, model: DocumentModel): DocumentRun[] {
  const pages = documentTextFields(previous).filter(field => field.field.toUpperCase() === 'PAGEREF' && !field.endBlockId)
  const page = pages.length === 1 ? pages[0] : undefined
  if (!page) return remapDocumentStoryFieldRuns(previous.runs as DocumentRun[], text, documentDefaultStyle(model).text)
  const defaults = documentDefaultStyle(model).text, runs = previous.runs as DocumentRun[]
  const prefix = sliceRuns(runs, 0, page.startOffset)
  if (!options.showPageNumbers || pageNumber == null) return remapDocumentStoryFieldRuns(prefix, text, defaults)
  const digits = String(pageNumber), start = text.length - digits.length
  const cached = sliceRuns(runs, page.startOffset, page.endOffset)
  const inherited = documentRunStyleAt(runs, page.startOffset, page.endOffset) ?? defaults
  return [...remapDocumentStoryFieldRuns(prefix, text.slice(0, start), defaults),
    ...remapDocumentStoryFieldRuns(cached.length ? cached : [createRun('', inherited)], digits, defaults)]
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
  const headings = collectDocumentTocHeadings(model, options.levels, options.minimumLevel)
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
  const blocks = [...titleBlocks, ...entries]
  for (const block of blocks) {
    const previous = existingBlocks.find(old => old.id === block.id)
    if (!previous) continue
    const text = runText(block.runs)
    block.metadata = { ...cloneModel(previous.metadata), ...block.metadata }
    block.style = { ...cloneModel(previous.style), ...block.style }
    block.runs = previous.runs.length === 1 ? [{ ...cloneModel(previous.runs[0]!), text }]
      : tocMetadata(block)?.role === 'entry' ? remapTocEntryRuns(previous, text, options, tocMetadata(block)?.pageNumber, model)
        : remapDocumentStoryFieldRuns(previous.runs as DocumentRun[], text, documentDefaultStyle(model).text)
    if (text !== runText(previous.runs)) delete block.metadata.html
  }
  installDocumentTocFields(blocks, options, existingBlocks)
  return blocks
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

  const plan = prepareDocumentTocFullUpdate(model, tocId, pageNumbers)
  if (!plan) return false
  applyDocumentTocMutation(model, plan)
  return true
}

function fieldResultText(range: DocumentTextFieldRange, paragraphs: readonly DocumentTextFieldParagraph[]): string {
  return paragraphs.filter(paragraph => paragraph.absoluteStart >= range.startParagraph.absoluteStart
    && paragraph.absoluteStart <= range.endParagraph.absoluteStart).map(paragraph => paragraph.text.slice(
    paragraph.block.id === range.block.id ? range.field.startOffset : 0,
    paragraph.block.id === range.endParagraph.block.id ? range.field.endOffset : undefined,
  )).join('\n')
}

function tocOuterFields(model: DocumentModel, blocks: readonly DocumentBlock[]): DocumentTextFieldRange[] {
  const paragraphs = documentTextFieldParagraphs(model), ids = new Set(blocks.map(block => block.id))
  return documentTextFieldRanges(model).filter(range => range.field.field.toUpperCase() === 'TOC'
    && paragraphs.some(paragraph => ids.has(paragraph.block.id) && paragraph.absoluteStart >= range.startParagraph.absoluteStart
      && paragraph.absoluteStart <= range.endParagraph.absoluteStart))
}

function tocRetainsPartialNativeResult(blocks: readonly DocumentBlock[], fields: readonly DocumentTextFieldRange[]): boolean {
  const ids = new Set(blocks.map(block => block.id))
  return fields.some(range => ids.has(range.block.id) && range.field.startOffset !== 0
    || ids.has(range.endParagraph.block.id) && range.field.endOffset !== range.endParagraph.text.length)
}

export type DocumentTocAction = DocumentTocUpdateMode | 'delete'

/** Explain the same source constraints used by the detached mutation planners. */
export function documentTocActionUnavailableReason(
  model: DocumentModel, tocId: string, action: DocumentTocAction, pageNumbers: DocumentTocPageNumbers = {},
): string | undefined {
  const blocks = model.blocks.filter(block => tocMetadata(block)?.id === tocId)
  if (!blocks.length) return '目标目录已不存在，请重新选择。'
  if (!supportsDocumentTableOfContents(model)) return '当前文档模式尚未接入正文目录。'
  const outer = tocOuterFields(model, blocks)
  if (outer.some(range => documentFieldLocked(range.field))) return '目录字段已锁定。'
  for (const metadata of blocks.map(tocMetadata)) {
    if (!metadata?.native) continue
    const source = outer.find(range => range.field.id === metadata.native!.fieldId)
    if (!source || source.block.id !== blocks[0]!.id || source.endParagraph.block.id !== blocks.at(-1)!.id
      || source.field.startOffset !== 0 || source.field.endOffset !== source.endParagraph.text.length) return '原生目录的字段范围已改变，请重新解析目录后再操作。'
    if (action !== 'delete' && metadata.native.instruction !== undefined && metadata.native.instruction !== (source.field.instruction ?? '')) return '原生目录指令已改变，请重新解析选项和标题目标后再更新。'
  }
  if (action !== 'delete' && outer.some(range => documentFieldRetainsNestedFields(range.field))) return '原生目录的相对嵌套字段范围尚未转换，当前不能安全更新。'
  if (action === 'pageNumbersOnly') return blocks.map(tocMetadata).find(metadata => metadata?.native?.pageUpdateUnavailableReason)?.native?.pageUpdateUnavailableReason
  const ids = new Set(blocks.map(block => block.id)), paragraphs = documentTextFieldParagraphs(model)
  const first = paragraphs.find(paragraph => paragraph.block.id === blocks[0]!.id)
  const last = paragraphs.find(paragraph => paragraph.block.id === blocks.at(-1)!.id)
  if (first && last && documentTextFieldRanges(model).some(range => documentFieldLocked(range.field)
    && (documentTextFieldRangeIntersects(range, first.absoluteStart, last.absoluteEnd) || ids.has(range.block.id)))) return '目录结果中包含锁定字段。'
  if (tocRetainsPartialNativeResult(blocks, outer)) return '原生目录结果只占部分段落，该范围的结构变更尚未接入。'
  if (blocks.some(block => block.metadata.nativeFieldSource !== undefined)) return '原生目录来源尚未映射为可安全重建的字段范围。'
  if (blocks.some(block => block.runs.some(run => run.metadata?.note || run.metadata?.noteReference))) return '目录中的脚注或尾注引用尚未接入目录结构事务。'
  if (action === 'delete') return undefined
  const nativeReason = blocks.map(tocMetadata).find(metadata => metadata?.native?.fullUpdateUnavailableReason)?.native?.fullUpdateUnavailableReason
  if (nativeReason) return nativeReason
  if (blocks.some(block => block.runs.some(run => run.metadata?.latexFormula))) return '目录中的公式实体尚未接入完整重建事务。'
  const options = blocks.map(tocMetadata).find(metadata => metadata?.options)?.options
  if (!options) return '目录选项已缺失，请重新定位目录。'
  const headings = collectDocumentTocHeadings(model, options.levels, options.minimumLevel)
  if (!headings.length) return '没有可加入目录的标题。'
  if (options.showPageNumbers && headings.some(heading => !Number.isSafeInteger(pageNumbers[heading.id]) || pageNumbers[heading.id]! < 1)) return '部分标题的布局页码尚未就绪。'
  return undefined
}

/** Rebuild the field result and all generated children in one detached structural candidate. */
export function prepareDocumentTocFullUpdate(
  model: DocumentModel,
  tocId: string,
  pageNumbers: DocumentTocPageNumbers = {},
  selection?: DocumentSelection,
  timestamp = Date.now(),
): DocumentTocMutationPlan | undefined {
  const blocks = model.blocks.filter(block => tocMetadata(block)?.id === tocId)
  if (!blocks.length) return undefined
  if (documentTocActionUnavailableReason(model, tocId, 'all', pageNumbers)) return undefined
  const outerFields = tocOuterFields(model, blocks)
  if (outerFields.some(range => documentFieldLocked(range.field) || documentFieldRetainsNestedFields(range.field))
    || tocRetainsPartialNativeResult(blocks, outerFields) || blocks.some(block => block.metadata.nativeFieldSource !== undefined
      || block.runs.some(run => run.metadata?.note || run.metadata?.noteReference || run.metadata?.latexFormula))) return undefined
  const title = blocks.find(block => tocMetadata(block)?.role === 'title'), titleMetadata = title && tocMetadata(title)
  const options = titleMetadata?.options ?? blocks.map(tocMetadata).find(metadata => metadata?.role === 'entry')?.options
  if (!options) return undefined
  const headings = collectDocumentTocHeadings(model, options.levels, options.minimumLevel)
  if (!headings.length || options.showPageNumbers && headings.some(heading => !Number.isSafeInteger(pageNumbers[heading.id]) || pageNumbers[heading.id]! < 1)) return undefined
  const replacement = buildDocumentTocBlocks(model, options, pageNumbers, tocId, blocks, titleMetadata?.titleText ?? 'Contents', Boolean(title))
  const plan = prepareDocumentTocBlockMutation(model, blocks, replacement, new Set(outerFields.map(range => range.field.id)), selection, timestamp)
  const candidate: DocumentModel = { ...model, blocks: plan.blocks, contentControls: plan.contentControls, comments: plan.comments }
  if (outerFields.length) {
    const ids = new Set(replacement.map(block => block.id))
    const sourceIds = new Set(blocks.map(block => block.id))
    for (const block of candidate.blocks) if (ids.has(block.id)) setDocumentTextFields(block, documentTextFields(block).filter(field => field.field.toUpperCase() !== 'TOC'))
    const paragraphs = documentTextFieldParagraphs(candidate), byId = new Map(paragraphs.map(paragraph => [paragraph.block.id, paragraph]))
    const entries = candidate.blocks.filter(block => tocMetadata(block)?.id === tocId && tocMetadata(block)?.role === 'entry')
    const firstEntry = entries[0]!, lastEntry = entries.at(-1)!
    for (const outer of outerFields) {
      const ownerId = sourceIds.has(outer.block.id) ? tocMetadata(outer.block)?.role === 'title' && title ? title.id : firstEntry.id : outer.block.id
      const tailId = sourceIds.has(outer.endParagraph.block.id) ? lastEntry.id : outer.endParagraph.block.id
      const owner = byId.get(ownerId), tail = byId.get(tailId)
      if (!owner || !tail) throw new RangeError('目录重建后的原生外层字段端点不存在。')
      const { endBlockId: _oldEndBlockId, ...field } = outer.field
      setDocumentTextFields(owner.block, [{ ...cloneModel(field), startOffset: sourceIds.has(outer.block.id) ? 0 : outer.field.startOffset,
        endOffset: sourceIds.has(outer.endParagraph.block.id) ? tail.text.length : outer.field.endOffset,
        containsNestedFields: options.showPageNumbers, ...(ownerId === tailId ? {} : { endBlockId: tailId }) }, ...documentTextFields(owner.block)])
      const mapped = documentTextFieldRanges(candidate).find(range => range.field.id === outer.field.id)!
      setDocumentTextFields(owner.block, documentTextFields(owner.block).map(current => current.id === mapped.field.id
        ? { ...current, cachedResult: fieldResultText(mapped, paragraphs) } : current))
    }
  }
  assignGeneratedTocPageParents(candidate, new Set(replacement.flatMap(block => documentTextFields(block)).filter(field => field.field.toUpperCase() === 'PAGEREF').map(field => field.id)),
    outerFields.length ? new Set(outerFields.map(range => range.field.id)) : undefined)
  reconcileDocumentTocFieldParents(model, candidate)
  documentTextFieldRanges(candidate)
  return { ...plan, blocks: candidate.blocks, changed: JSON.stringify(candidate.blocks) !== JSON.stringify(model.blocks)
    || JSON.stringify(candidate.contentControls) !== JSON.stringify(model.contentControls) || JSON.stringify(candidate.comments) !== JSON.stringify(model.comments) }
}

/** Delete the complete directory while retaining surrounding fields and a resolvable input target. */
export function prepareDocumentTocDeletion(
  model: DocumentModel,
  tocId: string,
  selection?: DocumentSelection,
  timestamp = Date.now(),
): DocumentTocMutationPlan | undefined {
  const blocks = model.blocks.filter(block => tocMetadata(block)?.id === tocId)
  if (!blocks.length) return undefined
  if (documentTocActionUnavailableReason(model, tocId, 'delete')) return undefined
  const outerFields = tocOuterFields(model, blocks)
  if (outerFields.some(range => documentFieldLocked(range.field)) || tocRetainsPartialNativeResult(blocks, outerFields)
    || blocks.some(block => block.metadata.nativeFieldSource !== undefined)) return undefined
  return prepareDocumentTocBlockMutation(model, blocks, [], new Set(outerFields.map(range => range.field.id)), selection, timestamp)
}

function mapTocPagePoint(point: DocumentPoint, blockId: string, start: number, end: number, length: number): DocumentPoint {
  if ((point.cellBlockId ?? point.blockId) !== blockId || point.offset <= start) return point
  return { ...point, offset: point.offset >= end ? point.offset + length - (end - start) : start + Math.min(point.offset - start, length) }
}

function tocPageRangeContainsInlineEntity(block: DocumentBlock, start: number, end: number): boolean {
  let offset = 0
  return block.runs.some(run => {
    const from = offset
    offset += run.text.length
    const entity = run.metadata?.note || run.metadata?.noteReference || run.metadata?.latexFormula
    return Boolean(entity && (run.text.length ? from < end && offset > start : from > start && from < end))
  })
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
  if (documentTocActionUnavailableReason(model, tocId, 'pageNumbersOnly', pageNumbers)) return undefined
  if (selection) for (const point of [selection.anchor, selection.focus]) {
    const block = blockAtPoint(model, point)
    if (!block || !isDocumentTextBoundary(runText(block.runs), point.offset)) throw new RangeError('目录更新选区的 UTF-16 位置无效。')
  }
  const paragraphs = documentTextFieldParagraphs(model), ranges = documentTextFieldRanges(model)
  const tocIds = new Set(tocBlocks.map(block => block.id))
  const outerFields = ranges.filter(range => range.field.field.toUpperCase() === 'TOC'
    && paragraphs.some(paragraph => tocIds.has(paragraph.block.id) && paragraph.absoluteStart >= range.startParagraph.absoluteStart
      && paragraph.absoluteStart <= range.endParagraph.absoluteStart))
  if (outerFields.some(range => documentFieldLocked(range.field) || documentFieldRetainsNestedFields(range.field))) return undefined
  const replacements: Array<{ blockId: string; metadata: DocumentTocMetadata; pageNumber: number; start: number; end: number; text: string; fieldId: string; createField: boolean }> = []
  let updatablePageFields = 0
  for (const block of tocBlocks) {
    const metadata = tocMetadata(block)
    if (metadata?.role !== 'entry') continue
    const pageFields = ranges.filter(range => range.block.id === block.id && range.field.field.toUpperCase() === 'PAGEREF')
    if (metadata.native ? !pageFields.length : !metadata.options?.showPageNumbers) continue
    // Multiple or cross-paragraph references need a target mapping beyond a single local TOC row.
    if (pageFields.length > 1) continue
    const pageField = pageFields[0]
    if (pageField && (documentFieldLocked(pageField.field) || documentTextFieldRangeContainsNested(pageField, ranges)
      || pageField.endParagraph.block.id !== block.id)) continue
    if (!metadata.headingId) { if (metadata.native) return undefined; continue }
    const text = runText(block.runs), separator = metadata.options?.rightAlignPageNumbers === false ? ' ' : '\t'
    const separatorIndex = metadata.pageNumber == null ? -1 : text.lastIndexOf(separator)
    const start = pageField?.field.startOffset ?? (separatorIndex >= 0 ? separatorIndex + 1 : text.length)
    const end = pageField?.field.endOffset ?? text.length
    if (tocPageRangeContainsInlineEntity(block, start, end)) continue
    const paragraph = paragraphs.find(paragraph => paragraph.block.id === block.id)
    if (!paragraph) throw new TypeError('目录页码必须属于文本段落。')
    const absoluteStart = paragraph.absoluteStart + start, absoluteEnd = paragraph.absoluteStart + end
    const overlapping = ranges.filter(range => documentTextFieldRangeIntersects(range, absoluteStart, absoluteEnd))
    if (overlapping.some(range => documentFieldLocked(range.field)
      || range.field.field.toUpperCase() !== 'TOC' && range.field.id !== pageField?.field.id)) continue
    const pageNumber = pageNumbers[metadata.headingId]
    if (pageNumber == null || !Number.isSafeInteger(pageNumber) || pageNumber < 1) return undefined
    updatablePageFields++
    const pageText = String(pageNumber), result = !pageField && separatorIndex < 0 ? `${separator}${pageText}` : pageText
    if (text.slice(start, end) === result && metadata.pageNumber === pageNumber
      && pageField?.field.cachedResult === pageText) continue
    replacements.push({ blockId: block.id, metadata, pageNumber, start, end, text: result,
      fieldId: pageField?.field.id ?? documentId('field'), createField: !pageField })
  }
  const repairsOuterCache = outerFields.some(range => range.field.cachedResult !== fieldResultText(range, paragraphs))
  if (!replacements.length && !repairsOuterCache) return { blocks: model.blocks, changed: false, updatablePageFields, ...(selection ? { selection } : {}) }
  const candidate: DocumentModel = { ...model, blocks: cloneModel(model.blocks),
    contentControls: model.contentControls && cloneModel(model.contentControls), comments: model.comments && cloneModel(model.comments) }
  let nextSelection = selection && cloneModel(selection)
  for (const replacement of replacements) {
    const block = candidate.blocks.find(block => block.id === replacement.blockId)!
    const textChanged = runText(block.runs).slice(replacement.start, replacement.end) !== replacement.text
    if (textChanged) {
      const appendedOuterEnds = replacement.start === replacement.end ? documentTextFieldRanges(candidate).filter(range => range.field.field.toUpperCase() === 'TOC'
        && range.endParagraph.block.id === block.id && range.field.endOffset === replacement.start) : []
      const runs = block.runs as DocumentRun[], split = splitDocumentRunRange(runs, replacement.start, replacement.end)
      const cache = sliceRuns(runs, replacement.start, replacement.end)
      const inherited = documentRunStyleAt(runs, replacement.start, replacement.end) ?? documentDefaultStyle(candidate).text
      const result = remapDocumentStoryFieldRuns(cache.length ? cache : [createRun('', inherited)], replacement.text, documentDefaultStyle(candidate).text)
      transformContentControlsForBlockEdit(candidate, block.id, replacement.start, replacement.end, replacement.text.length, undefined, true)
      transformDocumentTextFieldsForBlockEdit(block, replacement.start, replacement.end, replacement.text, replacement.fieldId, candidate)
      block.runs = [...split.before, ...result, ...split.after]
      delete block.metadata.html
      for (const outer of appendedOuterEnds) {
        const current = documentTextFieldRanges(candidate).find(range => range.field.id === outer.field.id)!
        setDocumentTextFields(current.block, documentTextFields(current.block).map(field => field.id === outer.field.id ? { ...field,
          startOffset: outer.block.id === block.id && outer.field.startOffset === replacement.start ? replacement.start : field.startOffset,
          endOffset: replacement.start + replacement.text.length } : field))
      }
      transformDocumentCommentAnchorsForTextEdit(candidate, block.id, replacement.start, replacement.end, replacement.text, timestamp)
      if (nextSelection) nextSelection = {
        anchor: mapTocPagePoint(nextSelection.anchor, block.id, replacement.start, replacement.end, replacement.text.length),
        focus: mapTocPagePoint(nextSelection.focus, block.id, replacement.start, replacement.end, replacement.text.length),
      }
    }
    block.metadata[TOC_METADATA_KEY] = { ...replacement.metadata, pageNumber: replacement.pageNumber }
    if (replacement.createField) setDocumentTextFields(block, [...documentTextFields(block), { id: replacement.fieldId, field: 'PAGEREF',
      instruction: pageReferenceInstruction(replacement.metadata.headingId!), cachedResult: String(replacement.pageNumber),
      startOffset: replacement.start + replacement.text.length - String(replacement.pageNumber).length,
      endOffset: replacement.start + replacement.text.length }])
    else setDocumentTextFields(block, documentTextFields(block).map(field => field.id === replacement.fieldId
      ? { ...field, cachedResult: String(replacement.pageNumber) } : field))
  }
  if (!outerFields.length && replacements.length) {
    const entries = candidate.blocks.filter(block => tocMetadata(block)?.id === tocId && tocMetadata(block)?.role === 'entry')
    const first = entries[0], last = entries.at(-1), options = first && tocMetadata(first)?.options
    if (first && last && options) setDocumentTextFields(first, [{ id: documentId('field'), field: 'TOC', instruction: buildDocumentTocInstruction(options),
      cachedResult: entries.map(block => runText(block.runs)).join('\n'), startOffset: 0, endOffset: runText(last.runs).length,
      ...(first.id === last.id ? {} : { endBlockId: last.id }) }, ...documentTextFields(first)])
  }
  assignGeneratedTocPageParents(candidate, new Set(replacements.filter(replacement => replacement.createField).map(replacement => replacement.fieldId)))
  const currentParagraphs = documentTextFieldParagraphs(candidate)
  for (const outer of outerFields) {
    const range = documentTextFieldRanges(candidate).find(range => range.field.id === outer.field.id)!
    const cachedResult = fieldResultText(range, currentParagraphs)
    if (range.field.cachedResult !== cachedResult) setDocumentTextFields(range.block, documentTextFields(range.block).map(field => field.id === range.field.id
      ? { ...field, cachedResult } : field))
  }
  reconcileDocumentTocFieldParents(model, candidate)
  documentTextFieldRanges(candidate)
  validateDocumentContentControlRanges(candidate)
  return { blocks: candidate.blocks, contentControls: candidate.contentControls, comments: candidate.comments, changed: true, updatablePageFields,
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
  const plan = prepareDocumentTocDeletion(model, tocId)
  if (!plan) return false
  applyDocumentTocMutation(model, plan)
  return true
}

export function documentTocIdAtBlock(block: DocumentBlock | undefined): string | undefined {
  return block ? tocMetadata(block)?.id : undefined
}

export function documentTocHeadingTarget(model: DocumentModel, block: DocumentBlock | undefined): string | undefined {
  const metadata = block && tocMetadata(block)
  const headingId = metadata?.role === 'entry' && metadata.native?.hyperlinks !== false ? metadata.headingId : undefined
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

/** Resolve section page labels; an unavailable layout must not fabricate physical labels. */
export function documentTocPageNumbers(headings: readonly { id: string; page: number }[], pages?: readonly { pageNumber: number; isParityBlank?: boolean }[]): DocumentTocPageNumbers {
  return Object.fromEntries(headings.flatMap(({ id, page }) => {
    if (!Number.isInteger(page) || page < 0) return []
    if (pages === undefined) return [[id, page + 1]]
    const layoutPage = pages[page]
    return layoutPage && !layoutPage.isParityBlank && Number.isInteger(layoutPage.pageNumber) && layoutPage.pageNumber > 0
      ? [[id, layoutPage.pageNumber]] : []
  }))
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
