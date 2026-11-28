import type { OfficeStudio } from './office-studio'
import { DOCUMENT_DEFAULT_ZOOM_RATIO, validateDocumentZoomRatio } from '../features/document/services/document-zoom'

export const DOCUMENT_ZOOM_COMMANDS = {
  SetRatio: 'doc.operation.set-zoom-ratio',
  SetRatioCommand: 'doc.command.set-zoom-ratio',
} as const

export interface DocumentZoomOperationParams { unitId: string; zoomRatio: number }
export interface DocumentZoomCommandParams { documentId?: string; zoomRatio?: number }

export function resolveDocumentZoomCommand(commandId: string, raw: unknown, activeDocumentId: string): DocumentZoomOperationParams {
  const operation = commandId === DOCUMENT_ZOOM_COMMANDS.SetRatio
  if (raw !== undefined && (raw === null || typeof raw !== 'object' || Array.isArray(raw))) throw new TypeError('文档缩放参数必须是对象。')
  if (operation && !raw) throw new TypeError('文档缩放操作需要文档 ID 和比例。')
  const params = raw as DocumentZoomCommandParams & Partial<DocumentZoomOperationParams> | undefined
  const unitId = operation ? params?.unitId : params?.documentId ?? activeDocumentId
  if (typeof unitId !== 'string' || !unitId.trim()) throw new TypeError('文档缩放的文档 ID 无效。')
  const zoomRatio = validateDocumentZoomRatio(operation ? params?.zoomRatio : params?.zoomRatio ?? DOCUMENT_DEFAULT_ZOOM_RATIO)
  return { unitId, zoomRatio }
}

/** Both public entry points publish the same view operation, outside content history. */
export function registerDocumentZoomCommands(api: OfficeStudio, publish: (params: DocumentZoomOperationParams) => boolean): void {
  for (const commandId of Object.values(DOCUMENT_ZOOM_COMMANDS)) {
    api.registerCommand<unknown, boolean>(commandId, raw => publish(resolveDocumentZoomCommand(commandId, raw, api.activeDocumentId)))
  }
}
