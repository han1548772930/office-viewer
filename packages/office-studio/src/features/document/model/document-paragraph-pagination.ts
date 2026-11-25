import type { DocumentBlock, DocumentModel, DocumentParagraphStyle } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { isDocumentStructuredBlock, isDocumentTextBlock, visitDocumentBlocks } from './document-structural-blocks'

export const DOCUMENT_PARAGRAPH_PAGINATION_FIELDS = ['keepLines', 'keepNext', 'widowControl', 'pageBreakBefore'] as const
export type DocumentParagraphPaginationField = typeof DOCUMENT_PARAGRAPH_PAGINATION_FIELDS[number]
export type DocumentParagraphPagination = Required<Pick<DocumentParagraphStyle, DocumentParagraphPaginationField>>
const QUOTE_FIELDS = ['keepLines', 'keepNext', 'widowControl'] as const
type QuoteField = typeof QUOTE_FIELDS[number]

function record(value: unknown): Record<string, unknown> | undefined {
  return value !== null && typeof value === 'object' && !Array.isArray(value) ? value as Record<string, unknown> : undefined
}

/** Public 0 means unspecified. Older project snapshots used 0 for traditional;
 * retain that local meaning only when no public mode provenance is present. */
export function documentUsesTraditionalPagination(model: Pick<DocumentModel, 'metadata'>): boolean {
  const metadata = model.metadata ?? {}, publicStyle = record(record(record(metadata.univerDocumentSource)?.snapshot)?.documentStyle)
    ?? record(metadata.documentStyle) ?? record(record(metadata.source)?.documentStyle)
  const flavor = metadata.documentFlavor ?? publicStyle?.documentFlavor ?? metadata.mode
  if (flavor === 1 || flavor === 'traditional') return true
  if (flavor === undefined) return true
  return flavor === 0 && publicStyle?.documentFlavor === undefined && metadata.documentFlavorContract !== 'univer'
}

/** Resolve only optional author rules; layout positions never enter the model. */
export function documentParagraphPagination(style: Partial<DocumentParagraphStyle>, model?: Pick<DocumentModel, 'defaultStyle'>): DocumentParagraphPagination {
  const defaults = model?.defaultStyle?.paragraph
  return { keepLines: style.keepLines ?? defaults?.keepLines ?? false, keepNext: style.keepNext ?? defaults?.keepNext ?? false,
    widowControl: style.widowControl ?? defaults?.widowControl ?? false, pageBreakBefore: style.pageBreakBefore ?? defaults?.pageBreakBefore ?? false }
}

function publicStyle(block: DocumentBlock): Record<string, unknown> | undefined {
  return record(record(record(block.metadata.univerBodyParagraphSource)?.paragraph)?.paragraphStyle)
}
interface GeneratedQuoteField {
  generated: boolean
  original?: boolean
  originalDirect?: unknown
  originalBaseline?: boolean
  originalAuthored: boolean
}
interface QuotePaginationSource {
  version: 1
  quoteId: string
  fields: Partial<Record<QuoteField, GeneratedQuoteField>>
  overrides: QuoteField[]
}
function quoteSource(block: DocumentBlock): QuotePaginationSource | undefined {
  const source = record(block.metadata.documentQuotePaginationSource)
  if (!source) return undefined
  const fields = record(source.fields)
  if (source.version !== 1 || typeof source.quoteId !== 'string' || !fields || !Array.isArray(source.overrides)
    || source.overrides.some(key => !QUOTE_FIELDS.includes(key as QuoteField))) throw new RangeError('引用分页来源无效。')
  const parsed: QuotePaginationSource['fields'] = {}
  for (const key of QUOTE_FIELDS) {
    const value = record(fields[key])
    if (!value) { if (fields[key] !== undefined) throw new RangeError('引用分页属性来源无效。'); continue }
    if (typeof value.generated !== 'boolean' || typeof value.originalAuthored !== 'boolean'
      || value.original !== undefined && typeof value.original !== 'boolean'
      || value.originalBaseline !== undefined && typeof value.originalBaseline !== 'boolean'
      || value.originalDirect !== undefined && value.originalDirect !== null
        && ![true, false, 0, 1].includes(value.originalDirect as boolean | number)) throw new RangeError('引用分页属性来源无效。')
    parsed[key] = { generated: value.generated, originalAuthored: value.originalAuthored,
      ...(value.original === undefined ? {} : { original: value.original as boolean }),
      ...(value.originalBaseline === undefined ? {} : { originalBaseline: value.originalBaseline as boolean }),
      ...(value.originalDirect === undefined ? {} : { originalDirect: cloneModel(value.originalDirect) }) }
  }
  return { version: 1, quoteId: source.quoteId, fields: parsed, overrides: [...source.overrides] as QuoteField[] }
}

function markPublicOverride(block: DocumentBlock, key: DocumentParagraphPaginationField, authored: boolean): void {
  const resolution = record(block.metadata.univerParagraphPaginationResolution)
  if (resolution?.version !== 1) return
  const keys = new Set(Array.isArray(resolution.authored) ? resolution.authored.filter(value => typeof value === 'string') : [])
  if (authored) keys.add(key)
  else keys.delete(key)
  resolution.authored = [...keys]
}

/** Same-value API writes are still explicit author formatting. They also detach
 * a generated quote rule so a later Enter/delete does not undo the author edit. */
export function markDocumentParagraphPaginationAuthored(block: DocumentBlock, patch: Partial<DocumentParagraphStyle>): void {
  const source = quoteSource(block)
  for (const key of DOCUMENT_PARAGRAPH_PAGINATION_FIELDS) {
    if (!Object.hasOwn(patch, key)) continue
    markPublicOverride(block, key, true)
    if (key !== 'pageBreakBefore' && source) {
      delete source.fields[key]
      if (!source.overrides.includes(key)) source.overrides.push(key)
    }
  }
  if (source) block.metadata.documentQuotePaginationSource = source
}

/** Explicit breaks and nested object hosts bound adjacency; their own rules
 * stay in their own container. A quote never captures the following body. */
function quoteParagraphGroups(quote: DocumentBlock): DocumentBlock[][] {
  const groups: DocumentBlock[][] = []
  let current: DocumentBlock[] = []
  for (const child of quote.children ?? []) {
    if (isDocumentTextBlock(child) && !isDocumentStructuredBlock(child)) current.push(child)
    else if (current.length) { groups.push(current); current = [] }
  }
  if (current.length) groups.push(current)
  return groups
}

function applyQuoteGroup(quote: DocumentBlock, paragraphs: DocumentBlock[]): void {
  for (const [index, paragraph] of paragraphs.entries()) {
    const source: QuotePaginationSource = quoteSource(paragraph) ?? { version: 1, quoteId: quote.id, fields: {}, overrides: [] }
    // A different surviving wrapper owns its provenance; do not overwrite it.
    if (source.quoteId !== quote.id) continue
    const resolution = record(paragraph.metadata.univerParagraphPaginationResolution), baseline = record(resolution?.baseline)
    const authored = Array.isArray(resolution?.authored) ? resolution.authored : []
    for (const key of QUOTE_FIELDS) {
      const generated = key === 'keepNext' ? index < paragraphs.length - 1 : true
      if (source.overrides.includes(key)) continue
      const previous = source.fields[key]
      if (previous && paragraph.style[key] !== previous.generated) {
        delete source.fields[key]; source.overrides.push(key); continue
      }
      if (!previous) source.fields[key] = { generated, originalAuthored: authored.includes(key),
        ...(paragraph.style[key] === undefined ? {} : { original: paragraph.style[key] }),
        ...(baseline?.[key] === undefined ? {} : { originalBaseline: baseline[key] as boolean }),
        ...(publicStyle(paragraph)?.[key] === undefined ? {} : { originalDirect: cloneModel(publicStyle(paragraph)![key]) }) }
      else previous.generated = generated
      paragraph.style[key] = generated
      markPublicOverride(paragraph, key, true)
    }
    paragraph.metadata.documentQuotePaginationSource = source
  }
}

export function initializeDocumentQuotePagination(quote: DocumentBlock): void {
  if (quote.type !== 'quote' || !quote.children) return
  quote.metadata.documentQuotePaginationManaged = 1
  for (const paragraphs of quoteParagraphGroups(quote)) applyQuoteGroup(quote, paragraphs)
}

function restoreQuoteParagraph(paragraph: DocumentBlock, quoteId: string): void {
  const source = quoteSource(paragraph)
  if (!source || source.quoteId !== quoteId) return
  const resolution = record(paragraph.metadata.univerParagraphPaginationResolution), baseline = record(resolution?.baseline), raw = publicStyle(paragraph)
  for (const key of QUOTE_FIELDS) {
    const previous = source.fields[key]
    if (!previous || paragraph.style[key] !== previous.generated) continue
    if (previous.original === undefined) delete paragraph.style[key]
    else paragraph.style[key] = previous.original
    if (raw) {
      if (previous.originalDirect === undefined) delete raw[key]
      else raw[key] = cloneModel(previous.originalDirect)
    }
    if (baseline) {
      if (previous.originalBaseline === undefined) delete baseline[key]
      else baseline[key] = previous.originalBaseline
    }
    markPublicOverride(paragraph, key, previous.originalAuthored)
  }
  delete paragraph.metadata.documentQuotePaginationSource
}

/** Restore only verifiable generated values, including after a public-body
 * save/reopen. Independent author changes remain on unwrapped paragraphs. */
export function clearDocumentQuotePagination(quote: DocumentBlock): void {
  visitDocumentBlocks([quote], child => restoreQuoteParagraph(child, quote.id))
  delete quote.metadata.documentQuotePaginationManaged
}

export function hasManagedDocumentQuotePagination(model: Pick<DocumentModel, 'blocks'>): boolean {
  let managed = false
  visitDocumentBlocks(model.blocks, block => { if (block.type === 'quote' && block.metadata.documentQuotePaginationManaged === 1) managed = true })
  return managed
}

/** Reconcile in the same content transaction, after text splits/merges or
 * structure edits. Imported unverified quote formatting is never rewritten. */
export function reconcileDocumentQuotePagination(model: Pick<DocumentModel, 'blocks'>): void {
  const owners = new Map<string, { quote: DocumentBlock; groups: DocumentBlock[][]; paragraphs: Set<DocumentBlock> }>()
  visitDocumentBlocks(model.blocks, block => {
    if (block.type === 'quote' && block.children && block.metadata.documentQuotePaginationManaged === 1) {
      const groups = quoteParagraphGroups(block)
      owners.set(block.id, { quote: block, groups, paragraphs: new Set(groups.flat()) })
    }
  })
  visitDocumentBlocks(model.blocks, block => {
    const source = quoteSource(block), owner = source && owners.get(source.quoteId)
    if (source && (!owner || !owner.paragraphs.has(block))) restoreQuoteParagraph(block, source.quoteId)
  })
  for (const owner of owners.values()) for (const paragraphs of owner.groups) applyQuoteGroup(owner.quote, paragraphs)
}
