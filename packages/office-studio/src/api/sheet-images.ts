import type { OfficeStudio } from './office-studio.ts'
import type { SheetDrawingAnchorType, SheetFloatingImage, SheetImagePlacement, SheetImagePlacementInput, SheetImageResource, SheetImageSourceRect } from '../domain/sheet-images.ts'
import type { SpreadsheetSheet, WorkbookModel } from '../domain/model.ts'
import { cloneModel } from '../domain/serialization.ts'
import { canViewProtectedRanges } from '../features/sheet-protection/range-protection.ts'
import { SHEET_IMAGE_COMMANDS } from '../features/sheet-images/sheet-image-commands.ts'
import { sheetImageClipboardBounds } from '../features/sheet-images/sheet-image-clipboard.ts'
import { sheetClipboardObjectAxes, sheetClipboardObjectRange } from '../features/sheet-workbench/sheet-clipboard-object-geometry.ts'
import { applySheetImagePlacement, sheetImageAnchorAt, sheetImageFrame, sheetImagePlacement, sheetImageWithFrame, type SheetFloatingImagePatch } from '../features/sheet-images/sheet-image-placement.ts'
import { prepareSheetImageResources, prepareSheetImageUrl } from '../features/sheet-images/sheet-image-io.ts'
import { normalizeSheetFloatingImage, sheetImageId, sheetImageIndex, sheetImageNumber, sheetImageRecord, sheetImageText } from '../features/sheet-images/sheet-image-model.ts'
import { sheetImageResourceFromDataUrl } from '../features/sheet-images/sheet-image-source.ts'
import { SheetImageSourceNeedsPreparationError, sheetFloatingImageFromPublicDrawing, sheetFloatingImageToPublicDrawing,
  type SheetImageSourceType, type UniverSheetImageDrawing } from '../features/sheet-images/sheet-image-public.ts'

export type PreparedSheetImage = SheetFloatingImage & UniverSheetImageDrawing
export interface SheetImageAsyncOptions { signal?: AbortSignal }
export function canViewSheetImage(api: OfficeStudio, workbook: WorkbookModel, sheet: SpreadsheetSheet, image: SheetFloatingImage): boolean {
  const range = sheetClipboardObjectRange(workbook, sheet.id, sheetImageClipboardBounds(workbook, sheet, image))
  return canViewProtectedRanges(sheet, [range], api.currentUserId)
}
function target(api: OfficeStudio, unitId: string, sheetId: string) {
  const workbook = api.getWorkbookModel(unitId), sheet = workbook?.sheets.find(item => item.id === sheetId)
  if (!workbook || !sheet) throw new Error('找不到图片所在的工作表。')
  api.assertCanViewWorksheet(unitId, sheetId)
  return { workbook, sheet }
}
export function preparedSheetImages(api: OfficeStudio, unitId: string, sheetId: string, values: readonly (SheetFloatingImage | UniverSheetImageDrawing)[]): SheetFloatingImage[] {
  if (!Array.isArray(values)) throw new TypeError('图片参数必须是数组。')
  const { workbook, sheet } = target(api, unitId, sheetId)
  return values.map(value => {
    const raw = sheetImageRecord(value)
    if (raw.unitId !== undefined && raw.unitId !== unitId || raw.subUnitId !== undefined && raw.subUnitId !== sheetId) throw new Error('图片必须属于目标工作表。')
    return sheetFloatingImageFromPublicDrawing(raw, workbook, sheet)
  })
}

/** The builder prepares detached image data. Only insert/update commands can publish it. */
export class FOverGridImageBuilder {
  private id = crypto.randomUUID()
  private resource?: SheetImageResource
  private source = ''
  private sourceType: SheetImageSourceType = 'URL'
  private row = 0
  private column = 0
  private rowOffset = 0
  private columnOffset = 0
  private width?: number
  private height?: number
  private opacity = 1
  private title = ''
  private altText = ''
  private keepRatio = true
  private anchorType: SheetDrawingAnchorType = '0'
  private placement?: SheetImagePlacementInput
  private explicitPlacement = false
  private crop?: SheetImageSourceRect
  private rotation = 0
  private flipX = false
  private flipY = false
  private metadata?: Record<string, unknown>
  constructor(private readonly api: OfficeStudio, private unitId: string, private sheetId: string) {}
  setImage(value: SheetFloatingImage | UniverSheetImageDrawing): this {
    const raw = sheetImageRecord(value), { workbook, sheet } = target(this.api, this.unitId, this.sheetId)
    this.placement = undefined; this.explicitPlacement = false; this.crop = undefined; this.rotation = 0; this.flipX = false; this.flipY = false
    this.row = 0; this.column = 0; this.rowOffset = 0; this.columnOffset = 0; this.width = undefined; this.height = undefined; this.metadata = undefined
    if (raw.resource === undefined && typeof raw.source === 'string' && !/^data:/i.test(raw.source) && raw.officeSheetImage === undefined) {
      if (raw.drawingType !== 0) throw new TypeError('目标绘图不是图片。')
      this.id = sheetImageId(raw.drawingId); this.setSource(raw.source, raw.imageSourceType === 'UUID' ? 'UUID' : 'URL')
      const transform = sheetImageRecord(raw.transform ?? {}), position = sheetImageRecord(raw.axisAlignSheetTransform ?? raw.sheetTransform ?? {}), from = sheetImageRecord(position.from ?? {})
      this.setRow(sheetImageIndex(from.row ?? 0, 'row')); this.setColumn(sheetImageIndex(from.column ?? 0, 'column'))
      this.setRowOffset(sheetImageNumber(from.rowOffset ?? from.offsetY ?? 0, '图片行偏移', -Number.MAX_SAFE_INTEGER))
      this.setColumnOffset(sheetImageNumber(from.columnOffset ?? from.offsetX ?? 0, '图片列偏移', -Number.MAX_SAFE_INTEGER))
      if (typeof transform.width === 'number') this.setWidth(transform.width)
      if (typeof transform.height === 'number') this.setHeight(transform.height)
      this.setRotate(sheetImageNumber(transform.angle ?? position.angle ?? 0, '图片旋转角度', -Number.MAX_SAFE_INTEGER))
      this.setAnchorType((raw.anchorType ?? '0') as SheetDrawingAnchorType)
      if (raw.srcRect !== undefined) {
        const crop = sheetImageRecord(raw.srcRect, '图片裁剪')
        this.setCropTop(sheetImageNumber(crop.top ?? 0, '顶部裁剪', -100_000, 100_000)); this.setCropLeft(sheetImageNumber(crop.left ?? 0, '左侧裁剪', -100_000, 100_000))
        this.setCropBottom(sheetImageNumber(crop.bottom ?? 0, '底部裁剪', -100_000, 100_000)); this.setCropRight(sheetImageNumber(crop.right ?? 0, '右侧裁剪', -100_000, 100_000))
      }
      this.setFlipX((transform.flipX ?? position.flipX ?? false) as boolean); this.setFlipY((transform.flipY ?? position.flipY ?? false) as boolean)
      this.setOpacity(sheetImageNumber(raw.opacity ?? 1, '图片不透明度', 0, 1)); this.setKeepRatio((raw.keepRatio ?? true) as boolean)
      this.setTitle(sheetImageText(raw.title, '图片标题')); this.setAltText(sheetImageText(raw.description, '替代文本'))
      if (this.anchorType === '1') this.placement = { kind: '1', from: { row: this.row, column: this.column, rowOffset: this.rowOffset, columnOffset: this.columnOffset },
        to: sheetImageRecord(position.to, '图片结束锚点') as unknown as Extract<SheetImagePlacement, { kind: '1' }>['to'] }
      else if (this.anchorType === '2' && this.width && this.height) this.placement = { kind: '2', left: sheetImageNumber(transform.left, '图片水平位置', -Number.MAX_SAFE_INTEGER),
        top: sheetImageNumber(transform.top, '图片垂直位置', -Number.MAX_SAFE_INTEGER), width: this.width, height: this.height }
      if (this.anchorType === '2' && !this.placement) {
        this.row = 0; this.column = 0; this.columnOffset = sheetImageNumber(transform.left, '图片水平位置', -Number.MAX_SAFE_INTEGER)
        this.rowOffset = sheetImageNumber(transform.top, '图片垂直位置', -Number.MAX_SAFE_INTEGER)
      }
      return this
    }
    const image = sheetFloatingImageFromPublicDrawing(raw, workbook, sheet)
    const axes = sheetClipboardObjectAxes(workbook, sheet), frame = sheetImageFrame(image, axes), anchor = sheetImageAnchorAt(axes, frame.left, frame.top)
    this.id = image.id; this.resource = cloneModel(image.resource); this.source = image.resource.source
    this.sourceType = /^data:/i.test(this.source) ? 'BASE64' : 'URL'
    this.row = anchor.row; this.column = anchor.column; this.rowOffset = anchor.offsetY; this.columnOffset = anchor.offsetX
    this.width = frame.width; this.height = frame.height; this.opacity = image.opacity; this.keepRatio = image.keepRatio
    this.anchorType = image.anchorType ?? '0'; this.placement = sheetImagePlacement(image, axes); this.rotation = image.rotation ?? 0
    this.flipX = image.flipX ?? false; this.flipY = image.flipY ?? false
    this.crop = image.crop && { top: image.crop.top * frame.height / image.size.height, bottom: image.crop.bottom * frame.height / image.size.height,
      left: image.crop.left * frame.width / image.size.width, right: image.crop.right * frame.width / image.size.width }
    this.title = image.title; this.altText = image.altText; this.metadata = image.metadata ? cloneModel(image.metadata) : undefined
    return this
  }
  setSource(source: string, sourceType: SheetImageSourceType = 'URL'): this {
    if (typeof source !== 'string' || !source || !['URL', 'BASE64', 'UUID'].includes(sourceType)) throw new TypeError('图片来源无效。')
    this.source = source; this.sourceType = sourceType; this.resource = undefined; return this
  }
  getSource(): string { return this.source }
  getSourceType(): SheetImageSourceType { return this.sourceType }
  private clearInferredPlacement(): void { if (!this.explicitPlacement) this.placement = undefined }
  setColumn(column: number): this { this.column = sheetImageIndex(column, 'column'); this.clearInferredPlacement(); return this }
  setRow(row: number): this { this.row = sheetImageIndex(row, 'row'); this.clearInferredPlacement(); return this }
  setColumnOffset(offset: number): this { this.columnOffset = sheetImageNumber(offset, '图片列偏移', -Number.MAX_SAFE_INTEGER); this.clearInferredPlacement(); return this }
  setRowOffset(offset: number): this { this.rowOffset = sheetImageNumber(offset, '图片行偏移', -Number.MAX_SAFE_INTEGER); this.clearInferredPlacement(); return this }
  setWidth(width: number): this { const value = sheetImageNumber(width, '图片宽度', 0, 100_000); this.width = value || undefined; this.clearInferredPlacement(); return this }
  setHeight(height: number): this { const value = sheetImageNumber(height, '图片高度', 0, 100_000); this.height = value || undefined; this.clearInferredPlacement(); return this }
  setOpacity(opacity: number): this { this.opacity = sheetImageNumber(opacity, '图片不透明度', 0, 1); return this }
  getOpacity(): number { return this.opacity }
  setTitle(title: string): this { this.title = sheetImageText(title, '图片标题'); return this }
  setAltText(altText: string): this { this.altText = sheetImageText(altText, '替代文本'); return this }
  setKeepRatio(keepRatio: boolean): this { if (typeof keepRatio !== 'boolean') throw new TypeError('纵横比选项无效。'); this.keepRatio = keepRatio; return this }
  setAnchorType(kind: SheetDrawingAnchorType): this { if (!['0', '1', '2'].includes(kind)) throw new TypeError('图片锚点类型无效。'); this.anchorType = kind; this.clearInferredPlacement(); return this }
  setPlacement(placement: SheetImagePlacementInput): this { sheetImageRecord(placement, '图片放置'); this.placement = cloneModel(placement); this.explicitPlacement = true; return this }
  private setCropEdge(edge: keyof SheetImageSourceRect, value: number): this {
    const crop = this.crop ?? { top: 0, left: 0, bottom: 0, right: 0 }
    this.crop = { ...crop, [edge]: sheetImageNumber(value, '图片裁剪', -100_000, 100_000) }; return this
  }
  setCropTop(value: number): this { return this.setCropEdge('top', value) }
  setCropLeft(value: number): this { return this.setCropEdge('left', value) }
  setCropBottom(value: number): this { return this.setCropEdge('bottom', value) }
  setCropRight(value: number): this { return this.setCropEdge('right', value) }
  setRotate(value: number): this { this.rotation = sheetImageNumber(value, '图片旋转角度', -Number.MAX_SAFE_INTEGER); this.clearInferredPlacement(); return this }
  setFlipX(value: boolean): this { if (typeof value !== 'boolean') throw new TypeError('图片翻转参数无效。'); this.flipX = value; return this }
  setFlipY(value: boolean): this { if (typeof value !== 'boolean') throw new TypeError('图片翻转参数无效。'); this.flipY = value; return this }
  setUnitId(unitId: string): this { this.unitId = sheetImageId(unitId); return this }
  setSubUnitId(sheetId: string): this { this.sheetId = sheetImageId(sheetId); return this }
  async buildAsync(options: SheetImageAsyncOptions = {}): Promise<PreparedSheetImage> {
    const config = { unitId: this.unitId, sheetId: this.sheetId, id: this.id, source: this.source, sourceType: this.sourceType,
      resource: this.resource && cloneModel(this.resource), row: this.row, column: this.column, rowOffset: this.rowOffset,
      columnOffset: this.columnOffset, width: this.width, height: this.height, opacity: this.opacity, keepRatio: this.keepRatio,
      title: this.title, altText: this.altText, metadata: this.metadata && cloneModel(this.metadata), anchorType: this.anchorType,
      placement: this.placement && cloneModel(this.placement), crop: this.crop && cloneModel(this.crop), rotation: this.rotation, flipX: this.flipX, flipY: this.flipY }
    const { workbook, sheet } = target(this.api, config.unitId, config.sheetId), revision = this.api.getWorkbookRevision(config.unitId)
    const controller = new AbortController(), cancel = () => controller.abort()
    const current = () => !controller.signal.aborted && !this.api.isDisposed && this.api.getWorkbookModel(config.unitId) === workbook
      && this.api.getWorkbookRevision(config.unitId) === revision
    const listeners = [this.api.events.addEvent('WorkbookDisposed', event => { if (event.workbookId === config.unitId) cancel() }),
      this.api.events.addEvent('LifeCycleChanged', event => { if (event.stage === 'disposed') cancel() })]
    options.signal?.addEventListener('abort', cancel, { once: true })
    if (options.signal?.aborted) cancel()
    try {
      if (!current()) throw new DOMException('图片准备已取消。', 'AbortError')
      if (config.sourceType === 'UUID' && !config.resource) throw new Error('UUID 图片需要先绑定实际媒体资源。')
      const resource = config.resource ? (await prepareSheetImageResources([config.resource], { signal: controller.signal, current }))[0]!
        : await prepareSheetImageUrl(config.source, { signal: controller.signal })
      if (!current()) throw new DOMException('图片目标已更改或准备已取消。', 'AbortError')
      this.api.assertCanViewWorksheet(config.unitId, config.sheetId)
      let image = normalizeSheetFloatingImage({ id: config.id, resource, anchor: { row: config.row, column: config.column, offsetX: config.columnOffset, offsetY: config.rowOffset },
        size: { width: config.width ?? resource.width, height: config.height ?? resource.height }, opacity: config.opacity, keepRatio: config.keepRatio,
        title: config.title || resource.name, altText: config.altText, metadata: config.metadata })
      const axes = sheetClipboardObjectAxes(workbook, sheet), frame = { left: axes.columns.offset(config.column) + config.columnOffset,
        top: axes.rows.offset(config.row) + config.rowOffset, ...image.size }
      image = config.placement ? applySheetImagePlacement({ ...image, rotation: config.rotation }, config.placement, axes)
        : sheetImageWithFrame({ ...image, rotation: config.rotation }, frame, axes, config.anchorType)
      image = normalizeSheetFloatingImage({ ...image, crop: config.crop, flipX: config.flipX, flipY: config.flipY })
      if (!canViewSheetImage(this.api, workbook, sheet, image)) throw new Error('没有查看目标图片范围的权限。')
      return { ...image, ...sheetFloatingImageToPublicDrawing(workbook, sheet, image) }
    } finally {
      options.signal?.removeEventListener('abort', cancel)
      for (const listener of listeners) listener.dispose()
    }
  }
}

/** Every read resolves the current authoritative model; wrappers do not retain a second drawing snapshot. */
export class FOverGridImage {
  constructor(private readonly api: OfficeStudio, private readonly unitId: string, private readonly sheetId: string, private readonly id: string) {}
  getId(): string { return this.id }
  getUnitId(): string { return this.unitId }
  getSubUnitId(): string { return this.sheetId }
  getType(): 0 { return 0 }
  getOpacity(): number { return this.model().opacity }
  getSource(): string { return this.model().resource.source }
  getPlacement(): SheetImagePlacement { const { workbook, sheet } = target(this.api, this.unitId, this.sheetId); return sheetImagePlacement(this.model(), sheetClipboardObjectAxes(workbook, sheet)) }
  getAnchorType(): SheetDrawingAnchorType { return this.model().anchorType ?? '0' }
  getRotate(): number { return this.model().rotation ?? 0 }
  getCrop(): SheetImageSourceRect { const image = this.model(), { workbook, sheet } = target(this.api, this.unitId, this.sheetId), frame = sheetImageFrame(image, sheetClipboardObjectAxes(workbook, sheet)), crop = image.crop ?? { top: 0, left: 0, bottom: 0, right: 0 };
    return { top: crop.top * frame.height / image.size.height, bottom: crop.bottom * frame.height / image.size.height,
      left: crop.left * frame.width / image.size.width, right: crop.right * frame.width / image.size.width } }
  getSnapshot(): PreparedSheetImage { const { workbook, sheet } = target(this.api, this.unitId, this.sheetId), image = this.model(); return { ...cloneModel(image), ...sheetFloatingImageToPublicDrawing(workbook, sheet, image) } }
  toBuilder(): FOverGridImageBuilder { return new FOverGridImageBuilder(this.api, this.unitId, this.sheetId).setImage(this.model()) }
  setOpacity(opacity: number): boolean { return this.update({ opacity }) }
  setTitle(title: string): boolean { return this.update({ title }) }
  setAltText(altText: string): boolean { return this.update({ altText }) }
  setPlacement(placement: SheetImagePlacementInput): boolean { return this.update({ placement }) }
  setAnchorType(anchorType: SheetDrawingAnchorType): boolean { return this.update({ anchorType }) }
  setCrop(top?: number, left?: number, bottom?: number, right?: number): boolean {
    const image = this.model(), crop = this.getCrop(), { workbook, sheet } = target(this.api, this.unitId, this.sheetId), frame = sheetImageFrame(image, sheetClipboardObjectAxes(workbook, sheet))
    return this.update({ size: { width: frame.width, height: frame.height }, crop: { top: top ?? crop.top, left: left ?? crop.left, bottom: bottom ?? crop.bottom, right: right ?? crop.right } })
  }
  setRotate(rotation: number): boolean { return this.update({ rotation }) }
  setFlipX(flipX: boolean): boolean { return this.update({ flipX }) }
  setFlipY(flipY: boolean): boolean { return this.update({ flipY }) }
  setForward(): boolean { return this.arrange(0) }
  setBackward(): boolean { return this.arrange(1) }
  setFront(): boolean { return this.arrange(2) }
  setBack(): boolean { return this.arrange(3) }
  setZOrder(zOrder: number): boolean { this.model(); return this.api.executeCommandSync(SHEET_IMAGE_COMMANDS.arrange, { unitId: this.unitId, subUnitId: this.sheetId, drawingIds: [this.id], zOrder }) }
  private arrange(arrangeType: 0 | 1 | 2 | 3): boolean { this.model(); return this.api.executeCommandSync(SHEET_IMAGE_COMMANDS.arrange, { unitId: this.unitId, subUnitId: this.sheetId, drawingIds: [this.id], arrangeType }) }
  copy(): boolean { const image = this.model(); return this.api.executeCommandSync<unknown, { accepted: boolean }>('sheet.command.copy-image-object', {
    unitId: this.unitId, sheetId: this.sheetId, imageIds: [this.id], range: { startRow: image.anchor.row, endRow: image.anchor.row,
      startCol: image.anchor.column, endCol: image.anchor.column } }).accepted }
  cut(): boolean { const image = this.model(); return this.api.executeCommandSync<unknown, { accepted: boolean }>('sheet.command.cut-image-object', {
    unitId: this.unitId, sheetId: this.sheetId, imageIds: [this.id], range: { startRow: image.anchor.row, endRow: image.anchor.row,
      startCol: image.anchor.column, endCol: image.anchor.column } }).accepted }
  remove(): boolean { this.model(); return this.api.executeCommandSync(SHEET_IMAGE_COMMANDS.remove, { unitId: this.unitId, sheetId: this.sheetId, imageId: this.id }) }
  setSource(source: string, _sourceType: SheetImageSourceType = 'URL'): boolean {
    if (!/^data:/i.test(source)) throw new SheetImageSourceNeedsPreparationError()
    return this.update({ resource: sheetImageResourceFromDataUrl(source, this.model().resource.name) })
  }
  setSourceAsync(source: string | File, options: SheetImageAsyncOptions = {}): Promise<boolean> {
    this.model()
    return this.api.executeCommand(SHEET_IMAGE_COMMANDS.replace, { ...options, unitId: this.unitId, sheetId: this.sheetId, imageId: this.id,
      ...(typeof source === 'string' ? { source } : { files: [source] }) })
  }
  async setPositionAsync(row: number, column: number, rowOffset = 0, columnOffset = 0): Promise<boolean> {
    return this.update({ anchor: { row, column, offsetX: columnOffset, offsetY: rowOffset } })
  }
  async setSizeAsync(width: number, height: number): Promise<boolean> { return this.update({ size: { width, height } }) }
  private update(patch: SheetFloatingImagePatch): boolean {
    this.model()
    return this.api.executeCommandSync(SHEET_IMAGE_COMMANDS.update, { unitId: this.unitId, sheetId: this.sheetId, imageId: this.id, patch })
  }
  private model(): SheetFloatingImage {
    const { workbook, sheet } = target(this.api, this.unitId, this.sheetId), image = sheet.images?.find(item => item.id === this.id)
    if (!image) throw new Error('图片已删除。')
    if (!canViewSheetImage(this.api, workbook, sheet, image)) throw new Error('没有查看目标图片的权限。')
    return image
  }
}

export async function insertWorksheetImage(api: OfficeStudio, unitId: string, sheetId: string, source: string | File,
  column = 0, row = 0, offsetX = 0, offsetY = 0, options: SheetImageAsyncOptions = {}): Promise<boolean> {
  target(api, unitId, sheetId)
  api.assertCanEditWorksheet(unitId, sheetId)
  return api.executeCommand(SHEET_IMAGE_COMMANDS.insertFloat, { ...options, unitId, sheetId, row, column, offsetX, offsetY, fitToViewport: false,
    ...(typeof source === 'string' ? { source } : { files: [source] }) })
}
