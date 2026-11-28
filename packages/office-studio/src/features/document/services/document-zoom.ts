import type { DocumentModel } from '../../../domain/model'
import { equivalentModelData } from '../../../platform/commands/model-equivalence'

/** The public Docs contract stores a ratio; Ribbon and Footer display percent. */
export const DOCUMENT_ZOOM_SHORTCUTS = [50, 80, 100, 130, 150, 170, 200, 400] as const
export const DOCUMENT_DEFAULT_ZOOM_RATIO = 1

export function validateDocumentZoomRatio(value: unknown): number {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < 0.1 || value > 4) {
    throw new RangeError('文档缩放比例须为 0.1 到 4 之间的有限数。')
  }
  return value
}

export function documentEffectiveZoomRatio(model: Pick<DocumentModel, 'metadata'>): number {
  const value = model.metadata?.documentZoomRatio
  return value === undefined || value === null ? DOCUMENT_DEFAULT_ZOOM_RATIO : validateDocumentZoomRatio(value)
}

/** View settings must not advance the content revision or invalidate text input. */
export function equivalentDocumentContent(left: DocumentModel, right: DocumentModel): boolean {
  const content = (model: DocumentModel): DocumentModel => {
    const metadata = { ...model.metadata }
    delete metadata.documentZoomRatio
    return { ...model, metadata: Object.keys(metadata).length ? metadata : undefined }
  }
  return equivalentModelData(content(left), content(right), ['modified'])
}

/** Undo/redo restores content while retaining the user's current document view. */
export function preserveDocumentZoomRatio(target: DocumentModel, source: DocumentModel): void {
  if (source.metadata?.documentZoomRatio !== undefined) {
    target.metadata = { ...target.metadata, documentZoomRatio: documentEffectiveZoomRatio(source) }
  } else if (target.metadata) delete target.metadata.documentZoomRatio
}
