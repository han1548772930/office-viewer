import type { InlineRun } from '../../../domain/model'
import type { DocumentImage } from '../model/document-model'

export const DOCUMENT_IMAGE_ATOM = '\uFFFC'
export interface DocumentImageAtom { version: 1; imageId: string; runId?: string; image: DocumentImage }
const record = (value: unknown): value is Record<string, unknown> => !!value && typeof value === 'object' && !Array.isArray(value)
const imageRecord = (value: unknown): value is DocumentImage => record(value) && typeof value.src === 'string'
  && typeof value.alt === 'string' && typeof value.width === 'number' && Number.isFinite(value.width) && value.width > 0
  && typeof value.height === 'number' && Number.isFinite(value.height) && value.height > 0

/** The source run owns its image identity; a copied run receives its own identity. */
export function documentImageAtom(run: Pick<InlineRun, 'id' | 'text' | 'metadata'>): DocumentImageAtom | undefined {
  const value = run.metadata?.imageAtom
  if (run.text !== DOCUMENT_IMAGE_ATOM || !record(value) || value.version !== 1 || typeof value.imageId !== 'string' || !value.imageId
    || value.runId !== undefined && typeof value.runId !== 'string' || !imageRecord(value.image)) return undefined
  return { version: 1, imageId: value.runId !== undefined && value.runId !== run.id ? run.id : value.imageId,
    ...(typeof value.runId === 'string' ? { runId: value.runId } : {}), image: value.image }
}

export function setDocumentImageAtom(run: InlineRun, imageId: string, image: DocumentImage): void {
  const previous = run.metadata?.imageAtom
  run.text = DOCUMENT_IMAGE_ATOM
  run.metadata = { ...run.metadata, imageAtom: { ...(record(previous) ? previous : {}), version: 1, imageId, runId: run.id, image } satisfies DocumentImageAtom }
}

/** Text-style inheritance must never turn a typed character into another embedded object. */
export function clearDocumentImageAtomStyleMetadata(metadata: Record<string, unknown>): void {
  delete metadata.imageAtom
  delete metadata.univerBodyCustomBlock
  delete metadata.univerBodyDrawingSource
  delete metadata.univerBodyRawCustomBlock
}
