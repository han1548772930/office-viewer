import type { FDocument } from './facade'
import { cloneModel } from '../domain/serialization'
import type { SheetImageResource } from '../domain/sheet-images'
import { prepareSheetImageFiles, prepareSheetImageResources, prepareSheetImageUrl, decodeSheetImageSource } from '../features/sheet-images/sheet-image-io'
import { sheetImageResourceFromDataUrl } from '../features/sheet-images/sheet-image-source'
import { normalizeSheetImageResource } from '../features/sheet-images/sheet-image-model'
import { documentId, textTargets, clampPoint, type DocumentSelection, type DocumentImage } from '../features/document/model/document-model'
import { sameDocumentBlockHost } from '../features/document/model/document-block-host'
import { documentImageEntries, documentImageEntryMetadata, normalizeDocumentImage, type DocumentImageEntry, type DocumentImagePatch,
  type DocumentImagePosition, type DocumentImageWrappingStyle, type DocumentImageArrange } from '../features/document/images/document-image-model'
import type { DocumentImageCommand } from '../features/document/commands/document-image-commands'
import { assertDocumentPermission, permissionTargetsForImage } from '../features/document/permissions/document-permissions'
import { documentSelectionAtPublicBodyRange } from './document-public-body'
import { documentSegmentIdForHost } from './document-segments'
import { documentImageToDrawing } from './document-image-body'
import { bodyRecord } from './document-body-style'
import { trackDocumentImageTask } from './document-image-editors'

export interface DocumentImageTextRange { startOffset: number; endOffset: number; segmentId?: string }
export interface DocumentImageInsertOptions {
  source: string | File | SheetImageResource; imageSourceType?: 'URL' | 'UUID' | 'BASE64'; width?: number; height?: number
  angle?: number; flipX?: boolean; flipY?: boolean; positionH?: DocumentImagePosition; positionV?: DocumentImagePosition; wrappingStyle?: DocumentImageWrappingStyle
  alt?: string; keepAspectRatio?: boolean; textRange?: DocumentImageTextRange; segmentId?: string
  paragraphMenuPlacement?: 'below'; signal?: AbortSignal; current?: () => boolean
}
export interface DocumentImageData extends Record<string, unknown> {
  drawingId: string; drawingType: 0; source: string; imageSourceType: 'URL' | 'BASE64'
  description: string; docTransform: { size: { width: number; height: number }; angle: number; flipX?: boolean; flipY?: boolean; positionH?: DocumentImagePosition; positionV?: DocumentImagePosition }
}

export function documentImageTargets(document: FDocument, segmentId = ''): DocumentImageEntry[] {
  const model = document.getSegmentModel(segmentId), entries = documentImageEntries(model)
  for (const entry of entries) assertDocumentPermission(model.permissionRules ?? [], { documentId: document.id, userId: document.getCurrentUserId() }, 'view', permissionTargetsForImage(model, entry.imageId))
  return entries
}
export function documentImageLegalSelection(document: FDocument, segmentId = ''): DocumentSelection {
  const model = document.getSegmentModel(segmentId), state = document.getInputState(), host = document.getSegmentHost(segmentId)
  if (state?.valid && state.location.kind === 'body' && sameDocumentBlockHost(host, state.location.host)) {
    return { anchor: clampPoint(model, state.location.selection.anchor), focus: clampPoint(model, state.location.selection.focus) }
  }
  const target = textTargets(model)[0]
  if (!target) throw new RangeError('图片操作缺少合法文字定位点。')
  return { anchor: { ...target, offset: 0 }, focus: { ...target, offset: 0 } }
}
export function documentImageFromResource(resource: SheetImageResource, options: Pick<DocumentImageInsertOptions, 'width' | 'height' | 'angle' | 'flipX' | 'flipY' | 'positionH' | 'positionV' | 'wrappingStyle' | 'alt' | 'keepAspectRatio'> = {}, originalSource?: string): DocumentImage {
  const width = options.width ?? (options.height === undefined ? resource.width : options.height * resource.width / resource.height)
  const height = options.height ?? (options.width === undefined ? resource.height : options.width * resource.height / resource.width)
  return normalizeDocumentImage({ src: resource.source, alt: options.alt ?? resource.name, width, height, angle: options.angle ?? 0, flipX: options.flipX, flipY: options.flipY,
    positionH: options.positionH, positionV: options.positionV, wrappingStyle: options.wrappingStyle ?? 'inline', keepAspectRatio: options.keepAspectRatio ?? true,
    resourceId: resource.id, intrinsicWidth: resource.width, intrinsicHeight: resource.height, mimeType: resource.mimeType, byteLength: resource.byteLength, originalSource })
}
export async function prepareDocumentImageSource(source: DocumentImageInsertOptions['source'], signal?: AbortSignal, current?: () => boolean): Promise<SheetImageResource> {
  const assertCurrent = () => { if (signal?.aborted || current?.() === false) throw new DOMException('图片读取对应的文档或选区已变化。', 'AbortError') }
  assertCurrent()
  if (typeof File !== 'undefined' && source instanceof File) {
    const prepared = (await prepareSheetImageFiles([source], { signal, current }))[0]
    assertCurrent()
    if (!prepared) throw new Error('图片资源读取未返回结果。')
    return prepared
  }
  if (typeof source === 'string' && /^https?:\/\//i.test(source)) {
    const prepared = await prepareSheetImageUrl(source, { signal })
    assertCurrent()
    return prepared
  }
  const resource = typeof source === 'string' ? sheetImageResourceFromDataUrl(source) : normalizeSheetImageResource(source)
  const prepared = (await prepareSheetImageResources([resource], { signal, current }))[0]!
  const decoded = await decodeSheetImageSource(prepared.source, signal)
  try {
    assertCurrent()
    if (decoded.naturalWidth !== prepared.width || decoded.naturalHeight !== prepared.height) throw new RangeError('图片声明尺寸与实际解码不一致。')
    return prepared
  } finally { decoded.src = '' }
}

/** Each mutation resolves the live object and uses the document's existing content history. */
export class FDocumentImage {
  constructor(private readonly document: FDocument, private readonly imageId: string, private readonly segmentId = '') {}
  getUnitId(): string { return this.document.id }
  getId(): string { return this.imageId }
  getSegmentId(): string { return this.segmentId }
  private source(): DocumentImageEntry | undefined { return documentImageTargets(this.document, this.segmentId).find(entry => entry.imageId === this.imageId) }
  getSource(): string | undefined { const image = this.source()?.image; return image?.originalSource ?? image?.src }
  getSourceType(): 'URL' | 'BASE64' | undefined { const source = this.getSource(); return source === undefined ? undefined : /^data:/i.test(source) ? 'BASE64' : 'URL' }
  getSize(): { width: number; height: number } | null { const image = this.source()?.image; return image ? { width: image.width, height: image.height } : null }
  getAngle(): number | undefined { const image = this.source()?.image; return image ? image.angle ?? 0 : undefined }
  getFlipX(): boolean | undefined { const image = this.source()?.image; return image ? image.flipX ?? false : undefined }
  getFlipY(): boolean | undefined { const image = this.source()?.image; return image ? image.flipY ?? false : undefined }
  getPositionH(): DocumentImagePosition | null { const value = this.source()?.image.positionH; return value ? cloneModel(value) : null }
  getPositionV(): DocumentImagePosition | null { const value = this.source()?.image.positionV; return value ? cloneModel(value) : null }
  getImage(): DocumentImage | null { const image = this.source()?.image; return image ? cloneModel(image) : null }
  getImageData(): DocumentImageData | null {
    const entry = this.source()
    if (!entry) return null
    const image = entry.image, raw = documentImageToDrawing(image, bodyRecord(documentImageEntryMetadata(entry).univerBodyDrawingSource))
    return cloneModel({ ...raw, drawingId: this.imageId, drawingType: 0, source: image.src, imageSourceType: /^data:/i.test(image.src) ? 'BASE64' : 'URL', description: image.alt,
      docTransform: { ...bodyRecord(raw.docTransform), size: { ...bodyRecord(bodyRecord(raw.docTransform).size), width: image.width, height: image.height },
        angle: image.angle ?? 0,
        ...(image.positionH ? { positionH: { ...bodyRecord(bodyRecord(raw.docTransform).positionH), ...cloneModel(image.positionH) } } : {}),
        ...(image.positionV ? { positionV: { ...bodyRecord(bodyRecord(raw.docTransform).positionV), ...cloneModel(image.positionV) } } : {}) } })
  }
  private execute(command: DocumentImageCommand): boolean {
    if (!this.source()) return false
    const host = this.document.getSegmentHost(this.segmentId)
    return this.document.executeCommand({ ...command, ...(host ? { host } : {}) }, documentImageLegalSelection(this.document, this.segmentId)).changed
  }
  update(patch: DocumentImagePatch): boolean {
    if (patch && typeof patch === 'object' && 'src' in patch) throw new TypeError('替换图片来源必须经过异步资源读取。')
    return this.execute({ id: 'doc.command.update-image', imageId: this.imageId, patch })
  }
  setSize(width: number, height: number): boolean { return this.update({ width, height }) }
  setRotate(angle: number): boolean { return this.update({ angle }) }
  setFlipX(flipX: boolean): boolean { return this.update({ flipX }) }
  setFlipY(flipY: boolean): boolean { return this.update({ flipY }) }
  setPositionH(positionH: DocumentImagePosition): boolean { return this.update({ positionH }) }
  setPositionV(positionV: DocumentImagePosition): boolean { return this.update({ positionV }) }
  setWrappingStyle(wrappingStyle: DocumentImageWrappingStyle): boolean { return this.update({ wrappingStyle }) }
  setCrop(crop: NonNullable<DocumentImage['crop']>): boolean { return this.update({ crop }) }
  arrange(arrange: DocumentImageArrange): boolean { return this.execute({ id: 'doc.command.arrange-image', imageId: this.imageId, arrange }) }
  setForward(): boolean { return this.arrange('forward') }
  setBackward(): boolean { return this.arrange('backward') }
  setFront(): boolean { return this.arrange('front') }
  setBack(): boolean { return this.arrange('back') }
  resetSize(): boolean { return this.execute({ id: 'doc.command.reset-image-size', imageId: this.imageId }) }
  remove(): boolean { return this.execute({ id: 'doc.command.remove-image', imageId: this.imageId }) }
  async replace(source: DocumentImageInsertOptions['source'], options: Pick<DocumentImageInsertOptions, 'signal' | 'current' | 'alt'> = {}): Promise<boolean> {
    if (!this.source()) return false
    const capture = captureImageTask(this.document, options), original = typeof source === 'string' ? source : undefined
    try {
      const resource = await prepareDocumentImageSource(source, capture.signal, capture.current)
      capture.assertCurrent()
      return this.execute({ id: 'doc.command.replace-image', imageId: this.imageId, image: documentImageFromResource(resource, { alt: options.alt }, original) })
    } finally { capture.dispose() }
  }
}
export function getDocumentImages(document: FDocument, segmentId = ''): FDocumentImage[] { return documentImageTargets(document, segmentId).map(entry => new FDocumentImage(document, entry.imageId, segmentId)) }
export function getDocumentImage(document: FDocument, imageId: string, segmentId = ''): FDocumentImage | null { return getDocumentImages(document, segmentId).find(image => image.getId() === imageId) ?? null }

function captureImageTask(document: FDocument, options: Pick<DocumentImageInsertOptions, 'signal' | 'current'>) {
  const source = document.getModel(), revision = document.getModelRevision(), history = document.getHistoryGeneration(), generation = document.getInputState()?.generation
  const controller = new AbortController(), cancel = () => controller.abort(), release = trackDocumentImageTask(source, controller)
  options.signal?.addEventListener('abort', cancel, { once: true }); if (options.signal?.aborted) controller.abort()
  const current = () => {
    try { return !controller.signal.aborted && options.current?.() !== false && document.getModel() === source && document.getModelRevision() === revision
      && document.getHistoryGeneration() === history && document.getInputState()?.generation === generation } catch { return false }
  }
  return { signal: controller.signal, current, assertCurrent: () => { if (!current()) throw new DOMException('图片读取对应的文档或选区已变化。', 'AbortError') },
    dispose: () => { options.signal?.removeEventListener('abort', cancel); release() } }
}
export async function insertDocumentImage(document: FDocument, options: DocumentImageInsertOptions): Promise<FDocumentImage | null> {
  if (!options || typeof options !== 'object' || Array.isArray(options)) throw new TypeError('插入图片参数无效。')
  if (options.imageSourceType === 'UUID') throw new RangeError('UUID 图片需要本项目媒体解析服务；当前尚未接入。')
  if (options.imageSourceType !== undefined && !['URL', 'BASE64'].includes(options.imageSourceType)) throw new TypeError('图片来源类型无效。')
  if (options.imageSourceType === 'URL' && (typeof options.source !== 'string' || !/^https?:\/\//i.test(options.source))
    || options.imageSourceType === 'BASE64' && (typeof options.source !== 'string' || !/^data:image\//i.test(options.source))) throw new TypeError('图片来源与声明的来源类型不一致。')
  if (options.segmentId !== undefined && typeof options.segmentId !== 'string'
    || options.current !== undefined && typeof options.current !== 'function'
    || options.paragraphMenuPlacement !== undefined && options.paragraphMenuPlacement !== 'below') throw new TypeError('图片插入上下文无效。')
  const range = options.textRange
  if (range !== undefined && (!range || typeof range !== 'object' || Array.isArray(range)
    || !Number.isSafeInteger(range.startOffset) || !Number.isSafeInteger(range.endOffset) || range.startOffset < 0 || range.endOffset < range.startOffset
    || range.segmentId !== undefined && typeof range.segmentId !== 'string'
    || range.segmentId !== undefined && options.segmentId !== undefined && range.segmentId !== options.segmentId)) throw new TypeError('图片插入的文字范围或内容区冲突。')
  const state = document.getInputState(), segmentId = options.segmentId ?? options.textRange?.segmentId
    ?? (state?.valid && state.location.kind === 'body' && state.location.host ? documentSegmentIdForHost(document.getModel(), state.location.host) : '')
  const host = document.getSegmentHost(segmentId)
  const selection = options.textRange ? documentSelectionAtPublicBodyRange(document.getModel(), options.textRange.startOffset, options.textRange.endOffset, host)
    : documentImageLegalSelection(document, segmentId)
  if (!selection) return null
  const capture = captureImageTask(document, options), id = documentId('image')
  try {
    const resource = await prepareDocumentImageSource(options.source, capture.signal, capture.current)
    capture.assertCurrent()
    const result = document.executeCommand({ id: 'doc.command.insert-image', imageId: id, image: documentImageFromResource(resource, options, typeof options.source === 'string' ? options.source : undefined),
      paragraphMenuPlacement: options.paragraphMenuPlacement, ...(host ? { host } : {}) }, selection)
    return result.changed ? getDocumentImage(document, id, segmentId) : null
  } finally { capture.dispose() }
}
