import type { FDocument } from './facade'
import type { DocumentModel } from '../domain/model'
import { sameDocumentBlockHost, type DocumentAuxiliaryBlockHost } from '../features/document/model/document-block-host'
import { createTextMeasurer, layoutDocument, type DocumentLayout } from '../features/document/layout/document-layout'
import { documentLayoutForBlockHost } from '../features/document/layout/document-auxiliary-layout'
import { layoutDocumentBlockHostInspection } from '../features/document/layout/document-block-host-inspection'
import type { LayoutMeasurer } from '../features/document/layout/document-layout-types'
import { documentImageEntries } from '../features/document/images/document-image-model'

export interface DocumentImageLayoutRead { source: DocumentModel; host?: DocumentAuxiliaryBlockHost; layout: DocumentLayout }
export interface DocumentImageGeometryReadOptions { preferredPage?: number; measurer?: LayoutMeasurer }
const owners = new WeakMap<DocumentModel, { read: () => DocumentImageLayoutRead | undefined }>()

/** A live editor can lend its current derived geometry; replacement owners have independent disposal. */
export function bindDocumentImageLayout(source: DocumentModel, read: () => DocumentImageLayoutRead | undefined): { dispose(): void } {
  if (!source || typeof source !== 'object' || typeof read !== 'function') throw new TypeError('图片排版宿主无效。')
  const owner = { read }
  owners.set(source, owner)
  return { dispose() { if (owners.get(source) === owner) owners.delete(source) } }
}
export function readDocumentImageLayout(document: FDocument, segmentId = '', options: DocumentImageGeometryReadOptions = {}): DocumentLayout {
  if (!options || typeof options !== 'object' || Array.isArray(options)
    || Object.keys(options).some(key => key !== 'preferredPage' && key !== 'measurer')
    || options.preferredPage !== undefined && (!Number.isSafeInteger(options.preferredPage) || options.preferredPage < 0)
    || options.measurer !== undefined && (!options.measurer || typeof options.measurer.measure !== 'function')) throw new TypeError('图片排版读取参数无效。')
  const source = document.getModel(), host = document.getSegmentHost(segmentId), borrowed = owners.get(source)?.read()
  if (borrowed?.source === source && sameDocumentBlockHost(borrowed.host, host)) return borrowed.layout
  // Reuse the authoritative document layout implementation for a detached Facade call.
  const measurer = options.measurer ?? createTextMeasurer(), layout = documentLayoutForBlockHost(layoutDocument(source, measurer), host)
  if (!host || documentImageEntries(document.getSegmentModel(segmentId)).every(entry => layout.images.some(image => image.blockId === entry.imageId))) return layout
  return layoutDocumentBlockHostInspection(source, host, measurer)
}
