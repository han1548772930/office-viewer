import type { DocumentBlock, DocumentModel } from '../../../domain/model'
import type { DocumentSelection } from '../model/document-model'
import { createParagraph, createRun, documentId, orderedSelection, runText, sameTarget, textTargets } from '../model/document-model'

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
  const maxLevel = Math.min(9, Math.max(1, Math.floor(levels)))
  return model.blocks.flatMap((block) => {
    if (block.type !== 'heading' || !block.id || block.level < 1 || block.level > maxLevel) return []
    const text = runText(block.runs).replace(/[\r\n\t]/g, ' ').trim()
    return text ? [{ id: block.id, text, level: block.level }] : []
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
    for (const block of tocBlocks) {
      const metadata = tocMetadata(block)
      if (metadata?.role !== 'entry' || !metadata.headingId || !metadata.options) continue
      const pageNumber = pageNumbers[metadata.headingId]
      if (pageNumber == null) continue
      const heading = { id: metadata.headingId, text: metadata.headingText ?? runText(block.runs), level: metadata.level ?? 1 }
      block.runs = [createRun(displayText(heading, metadata.options, pageNumber))]
      block.metadata[TOC_METADATA_KEY] = { ...metadata, pageNumber }
    }
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

export function deleteDocumentToc(model: DocumentModel, tocId: string): boolean {
  const next = model.blocks.filter((block) => tocMetadata(block)?.id !== tocId)
  if (next.length === model.blocks.length) return false
  model.blocks = next
  return true
}

export function documentTocIdAtBlock(block: DocumentBlock | undefined): string | undefined {
  return block ? tocMetadata(block)?.id : undefined
}

export function documentTocOptionsAtBlock(block: DocumentBlock | undefined): DocumentTocOptions | undefined {
  const options = block && tocMetadata(block)?.options
  return options && { ...options }
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
