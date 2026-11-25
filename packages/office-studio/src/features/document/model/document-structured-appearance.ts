import type { DocumentBlock } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { documentTextColor } from './document-model'
import { visitDocumentBlocks } from './document-structural-blocks'

export interface DocumentQuoteStyle { lineColor: string; textColor: string }

function record(value: unknown): Record<string, unknown> | undefined {
  return value !== null && typeof value === 'object' && !Array.isArray(value) ? value as Record<string, unknown> : undefined
}
function publicParagraphSource(block: DocumentBlock): Record<string, unknown> | undefined {
  const source = record(block.metadata.univerBodyParagraphSource)
  return record(source?.paragraph) ?? source
}
interface QuoteBorderSource { version: 1; quoteId: string; generated: Record<string, unknown>; original?: unknown }
function quoteBorderSource(block: DocumentBlock): QuoteBorderSource | undefined {
  const value = record(block.metadata.documentQuoteBorderSource), generated = record(value?.generated)
  return value?.version === 1 && typeof value.quoteId === 'string' && generated
    ? { version: 1, quoteId: value.quoteId, generated, ...(value.original === undefined ? {} : { original: value.original }) } : undefined
}

/** Export records the pre-quote rule so removing this container can restore author formatting. */
export function documentQuoteParagraphBorderSource(block: DocumentBlock, quoteId: string, generated: Record<string, unknown>): QuoteBorderSource {
  const border = record(publicParagraphSource(block)?.paragraphStyle)?.borderLeft, prior = quoteBorderSource(block)
  const original = prior && JSON.stringify(border) === JSON.stringify(prior.generated) ? prior.original : border
  return { version: 1, quoteId, generated: cloneModel(generated), ...(original === undefined ? {} : { original: cloneModel(original) }) }
}

/** Only rules with matching generated provenance are removed; unverified imported rules survive. */
export function clearDocumentQuoteParagraphBorders(block: DocumentBlock): void {
  visitDocumentBlocks([block], child => {
    const provenance = quoteBorderSource(child)
    if (!provenance || provenance.quoteId !== block.id) return
    const style = record(publicParagraphSource(child)?.paragraphStyle)
    if (style && JSON.stringify(style.borderLeft) === JSON.stringify(provenance.generated)) {
      if (provenance.original === undefined) delete style.borderLeft
      else style.borderLeft = cloneModel(provenance.original)
    }
    delete child.metadata.documentQuoteBorderSource
  })
}

export function documentStructuredTextColor(block: DocumentBlock): string {
  let color: string | undefined
  visitDocumentBlocks([block], child => {
    if (color === undefined) color = child.runs.find(run => run.text.length > 0)?.color
  })
  return documentTextColor(color ?? 'auto')
}

/** Local edits take precedence over imported paragraph-rule provenance. */
export function documentQuoteLineColor(block: DocumentBlock): string {
  const local = block.metadata.quoteStyle
  if (local && typeof local === 'object' && typeof Reflect.get(local, 'lineColor') === 'string') return Reflect.get(local, 'lineColor')
  let color: string | undefined
  visitDocumentBlocks([block], child => {
    if (color !== undefined) return
    const source = publicParagraphSource(child)
    if (!source) return
    const style = source.paragraphStyle, border = style && typeof style === 'object' && Reflect.get(style, 'borderLeft')
    const value = border && typeof border === 'object' && Reflect.get(border, 'color')
    const rgb = value && typeof value === 'object' ? Reflect.get(value, 'rgb') : value
    if (typeof rgb === 'string') color = rgb
  })
  return color ?? '#8796aa'
}

export function documentQuoteStyle(block: DocumentBlock): DocumentQuoteStyle {
  return { lineColor: documentQuoteLineColor(block), textColor: documentStructuredTextColor(block) }
}
