import type { OfficeStudio } from './office-studio'
import { FDocument } from './facade'
import { getDocumentImage, insertDocumentImage, documentImageLegalSelection, type DocumentImageInsertOptions } from './document-images'
import { requestDocumentImageEditorOpen, type DocumentImageEditorMode, type DocumentImageEditorTarget } from './document-image-editors'
import { documentSegmentIdForHost } from './document-segments'
import type { DocumentImageObjectCommand } from '../features/document/commands/document-image-commands'
import { normalizeDocumentImagePosition, documentImageCropFromSourceRect, DOCUMENT_IMAGE_WRAPPING_STYLES, type DocumentImageSourceRect,
  type DocumentImagePosition, type DocumentImageWrappingStyle, type DocumentImagePatch } from '../features/document/images/document-image-model'
import { cloneModel } from '../domain/serialization'

export const DOCUMENT_IMAGE_PUBLIC_COMMANDS = {
  insert: 'doc.command.insert-float-image', transform: 'doc.command.update-drawing-doc-transform',
  wrapping: 'doc.command.update-doc-drawing-wrapping-style', arrange: 'doc.command.set-drawing-arrange', remove: 'doc.command.remove-drawing',
  edit: 'doc.operation.edit-doc-image', crop: 'doc.operation.crop-image', replace: 'doc.operation.replace-image',
} as const
export interface DocumentImageCommandBase { unitId?: string; subUnitId?: string; segmentId?: string; expectedRevision?: number; expectedHistoryGeneration?: number }
export interface DocumentImageOperationRequest extends DocumentImageCommandBase { imageId?: string; drawingId?: string; paragraphMenuPlacement?: 'below' }
export interface DocumentImagePublicInsertRequest extends DocumentImageCommandBase, Partial<DocumentImageInsertOptions> {}
export type DocumentImageTransformItem = { drawingId: string } & (
  { key: 'size'; value: { width: number; height: number } } | { key: 'angle'; value: number }
  | { key: 'flipX' | 'flipY'; value: boolean }
  | { key: 'positionH' | 'positionV'; value: DocumentImagePosition } | { key: 'srcRect'; value: DocumentImageSourceRect | null }
)
export interface DocumentImagePublicTransformRequest extends DocumentImageCommandBase { drawings: DocumentImageTransformItem[] }
export interface DocumentImagePublicWrappingRequest extends DocumentImageCommandBase { drawings: Array<{ drawingId: string }>; wrappingStyle: DocumentImageWrappingStyle }
export interface DocumentImagePublicArrangeRequest extends DocumentImageCommandBase { drawingIds: string[]; arrangeType?: 0 | 1 | 2 | 3; zOrder?: number }
export interface DocumentImagePublicRemoveRequest extends DocumentImageCommandBase { drawings: Array<{ drawingId: string; unitId?: string; subUnitId?: string; drawingType?: number }> }

function commandDocument(api: OfficeStudio, input: DocumentImageCommandBase): FDocument {
  if (!input || typeof input !== 'object' || Array.isArray(input)
    || [input.segmentId, input.subUnitId].some(value => value !== undefined && typeof value !== 'string')
    || [input.expectedRevision, input.expectedHistoryGeneration].some(value => value !== undefined && (!Number.isSafeInteger(value) || value < 0))) throw new TypeError('图片命令的内容区或版本参数无效。')
  const unitId = input.unitId ?? api.activeDocumentId
  if (typeof unitId !== 'string' || !api.getDocumentModel(unitId)) throw new RangeError('图片命令的目标文档不存在。')
  if (input.expectedRevision !== undefined && input.expectedRevision !== api.getDocumentRevision(unitId)
    || input.expectedHistoryGeneration !== undefined && input.expectedHistoryGeneration !== api.editor.history.captureSavepoint(unitId).generation) throw new RangeError('图片命令的内容版本已变化。')
  return new FDocument(api, unitId)
}
function segment(document: FDocument, input: DocumentImageCommandBase): string {
  if (input.segmentId !== undefined && input.subUnitId !== undefined && input.subUnitId !== document.id && input.subUnitId !== input.segmentId) throw new TypeError('图片命令的内容区身份冲突。')
  return input.segmentId ?? (input.subUnitId && input.subUnitId !== document.id ? input.subUnitId : '')
}
function drawingIds(values: readonly string[]): string[] {
  if (!Array.isArray(values) || !values.length || values.length > 1000 || values.some(value => typeof value !== 'string' || !value.trim())
    || new Set(values).size !== values.length) throw new TypeError('图片命令需要不重复的有效对象身份。')
  return [...values]
}
function executeBatch(document: FDocument, segmentId: string, commands: DocumentImageObjectCommand[]): boolean {
  const host = document.getSegmentHost(segmentId)
  return document.executeCommand({ id: 'doc.command.batch-images', commands, ...(host ? { host } : {}) }, documentImageLegalSelection(document, segmentId)).changed
}
export function resolveDocumentImageEditorTarget(document: FDocument, request: DocumentImageOperationRequest, mode: DocumentImageEditorMode): DocumentImageEditorTarget {
  if (request.segmentId !== undefined || request.subUnitId !== undefined) segment(document, request)
  const state = document.getInputState(), segmentId = request.segmentId ?? (request.subUnitId && request.subUnitId !== document.id ? request.subUnitId : undefined)
    ?? (state?.valid && state.location.kind === 'body' && state.location.host ? documentSegmentIdForHost(document.getModel(), state.location.host) : '')
  const imageId = request.imageId ?? request.drawingId
  if (request.imageId !== undefined && request.drawingId !== undefined && request.imageId !== request.drawingId) throw new TypeError('图片编辑目标身份冲突。')
  if (imageId !== undefined && !getDocumentImage(document, imageId, segmentId)) throw new RangeError('图片不属于指定内容区。')
  if (request.paragraphMenuPlacement !== undefined && request.paragraphMenuPlacement !== 'below') throw new TypeError('图片段落插入位置无效。')
  return { unitId: document.id, source: document.getModel(), revision: document.getModelRevision(), historyGeneration: document.getHistoryGeneration(),
    inputGeneration: state?.generation, segmentId, selection: cloneModel(documentImageLegalSelection(document, segmentId)), host: document.getSegmentHost(segmentId),
    imageId, mode, paragraphMenuPlacement: request.paragraphMenuPlacement }
}

export function registerDocumentImageCommands(api: OfficeStudio): void {
  api.registerCommand<DocumentImagePublicInsertRequest | undefined, Promise<boolean>>(DOCUMENT_IMAGE_PUBLIC_COMMANDS.insert, async input => {
    const request = input ?? {}, document = commandDocument(api, request)
    const segmentId = request.segmentId !== undefined || request.subUnitId !== undefined ? segment(document, request) : undefined
    return request.source === undefined ? requestDocumentImageEditorOpen(api, resolveDocumentImageEditorTarget(document, request, 'insert'))
      : !!await insertDocumentImage(document, { ...request, segmentId, source: request.source })
  })
  for (const [id, mode] of [[DOCUMENT_IMAGE_PUBLIC_COMMANDS.edit, 'edit'], [DOCUMENT_IMAGE_PUBLIC_COMMANDS.crop, 'crop'], [DOCUMENT_IMAGE_PUBLIC_COMMANDS.replace, 'replace']] as const) {
    api.registerCommand<DocumentImageOperationRequest | undefined, boolean>(id, input => {
      const request = input ?? {}, document = commandDocument(api, request)
      return requestDocumentImageEditorOpen(api, resolveDocumentImageEditorTarget(document, request, mode))
    })
  }
  api.registerCommand<DocumentImagePublicTransformRequest, boolean>(DOCUMENT_IMAGE_PUBLIC_COMMANDS.transform, input => {
    const document = commandDocument(api, input), segmentId = segment(document, input)
    if (!Array.isArray(input.drawings) || !input.drawings.length || input.drawings.length > 1000) throw new TypeError('图片变换列表无效。')
    const patches = new Map<string, DocumentImagePatch>()
    const sourceRects = new Map<string, DocumentImageSourceRect | null>()
    for (const item of input.drawings) {
      if (!item || typeof item.drawingId !== 'string' || !getDocumentImage(document, item.drawingId, segmentId)) throw new RangeError('图片变换目标不存在。')
      let patch: DocumentImagePatch
      if (item.key === 'size') {
        if (!item.value || typeof item.value !== 'object' || Array.isArray(item.value)) throw new TypeError('图片尺寸参数无效。')
        patch = { width: item.value.width, height: item.value.height }
      }
      else if (item.key === 'angle') patch = { angle: item.value }
      else if (item.key === 'flipX' || item.key === 'flipY') {
        if (typeof item.value !== 'boolean') throw new TypeError('图片翻转值必须为布尔值。')
        patch = { [item.key]: item.value }
      }
      else if (item.key === 'positionH' || item.key === 'positionV') patch = { [item.key]: normalizeDocumentImagePosition(item.value, item.key === 'positionV') }
      else if (item.key === 'srcRect') { sourceRects.set(item.drawingId, item.value); patch = {} }
      else throw new TypeError('该变换字段尚未实现图片语义。')
      patches.set(item.drawingId, { ...patches.get(item.drawingId), ...patch })
    }
    // A crop notification may contain size after srcRect; conversion uses its final visible frame.
    for (const [imageId, sourceRect] of sourceRects) {
      const patch = patches.get(imageId)!, image = getDocumentImage(document, imageId, segmentId)!.getImage()!
      patch.crop = documentImageCropFromSourceRect(sourceRect, patch.width ?? image.width, patch.height ?? image.height)
    }
    return executeBatch(document, segmentId, [...patches].map(([imageId, patch]) => ({ id: 'doc.command.update-image', imageId, patch })))
  })
  api.registerCommand<DocumentImagePublicWrappingRequest, boolean>(DOCUMENT_IMAGE_PUBLIC_COMMANDS.wrapping, input => {
    const document = commandDocument(api, input), segmentId = segment(document, input)
    if (!DOCUMENT_IMAGE_WRAPPING_STYLES.includes(input.wrappingStyle) || !Array.isArray(input.drawings)) throw new TypeError('图片环绕参数无效。')
    return executeBatch(document, segmentId, drawingIds(input.drawings.map(item => item.drawingId)).map(imageId => ({ id: 'doc.command.update-image', imageId, patch: { wrappingStyle: input.wrappingStyle } })))
  })
  api.registerCommand<DocumentImagePublicRemoveRequest, boolean>(DOCUMENT_IMAGE_PUBLIC_COMMANDS.remove, input => {
    const document = commandDocument(api, input), segmentId = segment(document, input)
    if (!Array.isArray(input.drawings) || input.drawings.some(item => !item || item.unitId !== undefined && item.unitId !== document.id
      || item.drawingType !== undefined && item.drawingType !== 0 || item.subUnitId !== undefined && item.subUnitId !== document.id && item.subUnitId !== segmentId)) throw new TypeError('删除图片的目标文档或类型无效。')
    return executeBatch(document, segmentId, drawingIds(input.drawings.map(item => item.drawingId)).map(imageId => ({ id: 'doc.command.remove-image', imageId })))
  })
  api.registerCommand<DocumentImagePublicArrangeRequest, boolean>(DOCUMENT_IMAGE_PUBLIC_COMMANDS.arrange, input => {
    const document = commandDocument(api, input), segmentId = segment(document, input), ids = drawingIds(input.drawingIds)
    if (ids.some(id => !getDocumentImage(document, id, segmentId))) throw new RangeError('图片排列目标不属于指定内容区。')
    const host = document.getSegmentHost(segmentId), selection = documentImageLegalSelection(document, segmentId)
    if (input.zOrder !== undefined) {
      if (input.arrangeType !== undefined || typeof input.zOrder !== 'number' || ids.length !== 1) throw new TypeError('图片绝对层级参数无效。')
      return document.executeCommand({ id: 'doc.command.arrange-images', imageIds: ids, zOrder: input.zOrder, ...(host ? { host } : {}) }, selection).changed
    }
    // Univer core ArrangeTypeEnum: forward=0, backward=1, front=2, back=3.
    const arrange = (['forward', 'backward', 'front', 'back'] as const)[input.arrangeType ?? -1]
    if (!arrange) throw new TypeError('图片排列动作无效。')
    return document.executeCommand({ id: 'doc.command.arrange-images', imageIds: ids, arrange, ...(host ? { host } : {}) }, selection).changed
  })
}
