import { assertDocumentNoteReferenceBoundaryRank, documentNoteReferenceOffsets } from '../features/document/model/document-note-boundary'
import type { DocumentBodyParagraphPosition, DocumentBodyTextSpan } from './document-body-types'

/** Public streams retain one atom per note even when its display label is empty. */
export function documentBodyNoteReferenceCount(position: DocumentBodyParagraphPosition, streamOffset: number): number | undefined {
  const references = position.spans.filter(span => span.noteReference)
  return references.length ? references.filter(span => span.sourceEnd <= streamOffset).length : undefined
}

/** Writing spans use display offsets as source and stream offsets as target. */
export function documentBodyNoteStreamOffset(
  block: DocumentBodyParagraphPosition['block'], spans: readonly DocumentBodyTextSpan[], offset: number, fallback: number, rank?: number,
): number {
  const references = documentNoteReferenceOffsets(block.runs)
  const count = assertDocumentNoteReferenceBoundaryRank(references, offset, rank)
  const tokens = spans.filter(span => span.noteReference)
  const next = tokens[count], previous = tokens[count - 1]
  if (next?.sourceStart === offset) return next.targetStart
  if (previous?.sourceEnd === offset) return previous.targetEnd
  return fallback
}
