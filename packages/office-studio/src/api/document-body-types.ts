import type { DocumentBlock, DocumentContentControl, DocumentDefaultStyle, DocumentStoryContent } from '../domain/model'
import type { UniverBodyRecord } from './document-body-style'

/** Public v1.6.0 bodies use UTF-16 stream coordinates, independently of Canvas geometry. */
export interface UniverDocumentTextRun extends UniverBodyRecord { st: number; ed: number; ts?: UniverBodyRecord; sId?: string }
export interface UniverDocumentParagraph extends UniverBodyRecord { startIndex: number; paragraphId?: string; paragraphStyle?: UniverBodyRecord }
export interface UniverDocumentCustomTable extends UniverBodyRecord { startIndex: number; endIndex: number; tableId: string }
export interface UniverDocumentCustomBlock extends UniverBodyRecord { startIndex: number; blockId: string }
export interface UniverDocumentCustomRange extends UniverBodyRecord { startIndex: number; endIndex: number; rangeId: string; rangeType: number; properties?: UniverBodyRecord }
export interface UniverDocumentBody extends UniverBodyRecord {
  dataStream: string
  textRuns?: UniverDocumentTextRun[]
  paragraphs?: UniverDocumentParagraph[]
  tables?: UniverDocumentCustomTable[]
  customBlocks?: UniverDocumentCustomBlock[]
  docxRawCustomBlocks?: UniverDocumentCustomBlock[]
  customRanges?: UniverDocumentCustomRange[]
  sectionBreaks?: UniverBodyRecord[]
  renderedPageBreaks?: number[]
}
export interface UniverDocumentBodyResources {
  tableSource?: Record<string, UniverBodyRecord>
  drawings?: Record<string, UniverBodyRecord>
  drawingsOrder?: string[]
  lists?: Record<string, UniverBodyRecord>
}
export interface UniverDocumentBodyData extends UniverDocumentBodyResources { body: UniverDocumentBody }

// core/src/docs/data-model/types.ts, not the older table-token interface comment.
export const DOCUMENT_BODY_TOKENS = {
  paragraph: '\r', section: '\n', customBlock: '\b', pageBreak: '\f', columnBreak: '\v',
  tableStart: '\x1a', rowStart: '\x1b', cellStart: '\x1c', cellEnd: '\x1d', rowEnd: '\x0e', tableEnd: '\x0f',
  rangeStart: '\x1f', rangeEnd: '\x1e', blockStart: '\x10', blockEnd: '\x11',
  groupStart: '\x12', columnStart: '\x13', columnEnd: '\x14', groupEnd: '\x15',
} as const

export interface DocumentBodyTextSpan { sourceStart: number; sourceEnd: number; targetStart: number; targetEnd: number }
export interface DocumentBodyParagraphPosition {
  block: DocumentBlock
  start: number
  end: number
  spans: DocumentBodyTextSpan[]
}
export interface DocumentBodyDecodeResult {
  content: DocumentStoryContent
  contentControls: DocumentContentControl[]
  diagnostics: string[]
}
export interface DocumentBodyEncodeResult {
  data: UniverDocumentBodyData
  positions: DocumentBodyParagraphPosition[]
}
export interface DocumentBodyDecodeOptions {
  defaults: DocumentDefaultStyle
  allowSections?: boolean
  styles?: Record<string, UniverBodyRecord>
}

export function documentBodyArray(value: unknown, label: string): UniverBodyRecord[] {
  if (value === undefined || value === null) return []
  if (!Array.isArray(value) || value.some(item => !item || typeof item !== 'object' || Array.isArray(item))) throw new TypeError(`${label} 必须为对象数组。`)
  return value as UniverBodyRecord[]
}
export function documentBodyOffset(value: unknown, length: number, label: string): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 0 || value > length) throw new RangeError(`${label} 的 UTF-16 位置无效。`)
  return value
}
export function documentBodyIdentity(value: unknown, label: string): string {
  if (typeof value !== 'string' || !value.trim() || value.length > 4096) throw new TypeError(`${label} 标识无效。`)
  return value
}

/** A public stream boundary may collapse around nonprinting range delimiters. */
export function documentBodyLocalOffset(position: DocumentBodyParagraphPosition, source: number, affinity: 'start' | 'end' = 'start'): number {
  for (const span of position.spans) {
    if (source < span.sourceStart) return span.targetStart
    if (source <= span.sourceEnd) {
      const sourceLength = span.sourceEnd - span.sourceStart, targetLength = span.targetEnd - span.targetStart
      if (sourceLength === targetLength) return span.targetStart + Math.max(0, source - span.sourceStart)
      if (source === span.sourceStart) return span.targetStart
      if (source === span.sourceEnd) return span.targetEnd
      return affinity === 'start' ? span.targetStart : span.targetEnd
    }
  }
  return position.spans.at(-1)?.targetEnd ?? 0
}
