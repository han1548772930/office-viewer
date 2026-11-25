import type { SpreadsheetCell, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import type { SheetBackgroundImage, SheetCellImage, SheetFloatingImage, SheetImageAnchor, SheetImageResource } from '../../domain/sheet-images.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH } from '../../rendering/spreadsheet/layout.ts'
import { sheetStyleToPublicData, type SheetPublicStyleData } from '../sheet-workbench/sheet-public-style.ts'
import { exportSheetRangeThemeResource, SHEET_RANGE_THEME_MODEL_PLUGIN } from '../sheet-workbench/sheet-range-theme-model.ts'
import { sheetImageAnchorAt, sheetImageAxisSwapped, sheetImageFrame, sheetImageMajorBounds, sheetImageProjectAnchor, sheetImagePublicMarker, sheetImageWithFrame } from './sheet-image-placement.ts'
import { sheetClipboardBoundsDrawing, sheetClipboardObjectAxes } from '../sheet-workbench/sheet-clipboard-object-geometry.ts'
import { normalizeSheetBackgroundImage, normalizeSheetCellImage, normalizeSheetFloatingImage, SHEET_IMAGE_MAX_COUNT, sheetImageIndex, sheetImageNumber, sheetImageRecord } from './sheet-image-model.ts'
import { sheetImageResourceFromDataUrl } from './sheet-image-source.ts'
import { ensureSheetImageBodyDrawing, removeSheetImageBodyDrawing } from './sheet-image-public-body.ts'

export const SHEET_DRAWING_PLUGIN = 'SHEET_DRAWING_PLUGIN'
export const SheetImageSourceTypes = { URL: 'URL', UUID: 'UUID', BASE64: 'BASE64' } as const
export type SheetImageSourceType = typeof SheetImageSourceTypes[keyof typeof SheetImageSourceTypes]
export interface UniverSheetImageDrawing extends Record<string, unknown> {
  drawingId: string
  drawingType: number
  source?: string
  imageSourceType?: SheetImageSourceType
  transform?: Record<string, unknown> & { left?: number; top?: number; width?: number; height?: number }
  title?: string
  description?: string
  officeSheetImage?: SheetCellImage | SheetFloatingImage
}
export interface UniverSheetImageBody extends Record<string, unknown> {
  id?: string
  body?: Record<string, unknown> & { dataStream: string; customBlocks?: Array<{ startIndex: number; blockId: string; blockType?: number }> }
  drawings?: Record<string, UniverSheetImageDrawing>
  drawingsOrder?: string[]
}
export interface UniverWorksheetBackgroundImage extends Record<string, unknown> {
  source: string
  imageSourceType: SheetImageSourceType
  scaleX?: number
  scaleY?: number
  officeSheetImageResource?: SheetImageResource
}
export interface UniverSheetCellData extends Record<string, unknown> {
  v?: SpreadsheetCell['value']
  f?: string | null
  s?: SheetPublicStyleData | string | null
  t?: number | string | null
  p?: UniverSheetImageBody | null
}
export interface UniverWorksheetData extends Record<string, unknown> {
  id: string
  name: string
  cellData: Record<string, Record<string, UniverSheetCellData>>
  backgroundImage?: UniverWorksheetBackgroundImage
}
export interface UniverWorkbookData extends Record<string, unknown> {
  id: string
  name: string
  sheetOrder: string[]
  sheets: Record<string, UniverWorksheetData>
  resources?: Array<{ name: string; data: string }>
}
const object = (value: unknown): Record<string, unknown> => value && typeof value === 'object' && !Array.isArray(value) ? value as Record<string, unknown> : {}
const sourceType = (source: string): SheetImageSourceType => /^data:/i.test(source) ? 'BASE64' : 'URL'
function preservedDrawing(drawing: Record<string, unknown>): Record<string, unknown> {
  const result = cloneModel(drawing)
  delete result.officeSheetImage
  return result
}

export class SheetImageSourceNeedsPreparationError extends Error {
  constructor() { super('远程图片需要先读取实际尺寸和字节，请使用 createWorkbookAsync 或 prepareWorkbookSheetImageInput。'); this.name = 'SheetImageSourceNeedsPreparationError' }
}
function imageResource(drawing: Record<string, unknown>, fallbackId: string): SheetImageResource {
  const extension = object(drawing.officeSheetImage), extendedResource = object(extension.resource ?? drawing.officeSheetImageResource ?? drawing.resource)
  const source = typeof drawing.source === 'string' ? drawing.source : extendedResource.source
  if (typeof source !== 'string') throw new TypeError('公开图片缺少可读取的 source。')
  if (/^data:/i.test(source)) return { ...sheetImageResourceFromDataUrl(source, typeof drawing.title === 'string' ? drawing.title : 'image', fallbackId),
    ...(Object.keys(extendedResource).length ? { id: typeof extendedResource.id === 'string' ? extendedResource.id : fallbackId, metadata: cloneModel(object(extendedResource.metadata)) } : {}) }
  if (extendedResource.source === source) return normalizeSheetCellImage({ id: fallbackId, resource: extendedResource, title: '', altText: '' }).resource
  throw new SheetImageSourceNeedsPreparationError()
}
export function sheetCellImageFromPublicBody(value: unknown): SheetCellImage | undefined {
  if (value === null || value === undefined) return undefined
  const body = sheetImageRecord(value, '单元格图片正文'), drawings = object(body.drawings)
  const order = Array.isArray(body.drawingsOrder) ? body.drawingsOrder : Object.keys(drawings)
  for (const id of order) {
    if (typeof id !== 'string') throw new TypeError('单元格绘图顺序无效。')
    const drawing = object(drawings[id])
    if (drawing.drawingType !== 0) continue
    const extension = object(drawing.officeSheetImage), image = normalizeSheetCellImage({ ...extension, id,
      resource: imageResource(drawing, id), title: drawing.title ?? extension.title ?? '', altText: drawing.description ?? extension.altText ?? '',
      metadata: { ...object(extension.metadata), publicDrawing: preservedDrawing(drawing) } })
    return image
  }
  return undefined
}
/** Preserve the complete body even when the current cell image occupies only one drawing slot. */
export function applySheetPublicCellBody(cell: SpreadsheetCell, value: unknown): void {
  if (value === undefined) return
  if (value === null) {
    delete cell.image
    if (cell.metadata) { delete cell.metadata.publicBody; delete cell.metadata.sheetImagePublicDrawingId }
    return
  }
  const body = sheetImageRecord(value, '单元格正文'), image = sheetCellImageFromPublicBody(body)
  const preserved = cloneModel(body)
  if (preserved.drawings && typeof preserved.drawings === 'object') preserved.drawings = Object.fromEntries(
    Object.entries(preserved.drawings).map(([id, drawing]) => [id, preservedDrawing(object(drawing))]))
  cell.metadata = { ...cell.metadata, publicBody: preserved }
  if (image) { cell.image = image; cell.metadata.sheetImagePublicDrawingId = image.id }
  else { delete cell.image; delete cell.metadata.sheetImagePublicDrawingId }
  const stream = object(body.body).dataStream
  if (cell.value === null && typeof stream === 'string') {
    const text = stream.replace(/\r\n$/, '').replace(/\u0008/g, '')
    if (text) { cell.value = text; cell.display = text; cell.dataType = 'text' }
  }
}
export function sheetBackgroundImageFromPublic(value: unknown): SheetBackgroundImage | undefined {
  if (value === undefined || value === null) return undefined
  const input = sheetImageRecord(value, '工作表背景图')
  if (input.resource !== undefined) return normalizeSheetBackgroundImage(input)
  if (typeof input.source === 'string' && !/^data:/i.test(input.source) && input.officeSheetImageResource === undefined) {
    return normalizeSheetBackgroundImage({ ...input, imageSourceType: input.imageSourceType ?? 'URL' })
  }
  const resource = imageResource(input, crypto.randomUUID())
  const saved = cloneModel(input)
  delete saved.officeSheetImageResource; delete saved.resource
  return normalizeSheetBackgroundImage({ resource, scaleX: input.scaleX ?? resource.scaleX ?? 1, scaleY: input.scaleY ?? resource.scaleY ?? 1,
    metadata: { publicBackgroundImage: saved } })
}
export function normalizeWorkbookSheetImages(workbook: WorkbookModel): void {
  for (const sheet of workbook.sheets) {
    if (sheet.images !== undefined) {
      if (!Array.isArray(sheet.images) || sheet.images.length > SHEET_IMAGE_MAX_COUNT) throw new RangeError('工作表浮动图片数量无效。')
      sheet.images = sheet.images.map(normalizeSheetFloatingImage)
      if (new Set(sheet.images.map(image => image.id)).size !== sheet.images.length) throw new TypeError('工作表浮动图片 ID 重复。')
    }
    if (sheet.backgroundImage) sheet.backgroundImage = normalizeSheetBackgroundImage(sheet.backgroundImage)
    for (const cell of Object.values(sheet.cells)) if (cell.image) cell.image = normalizeSheetCellImage(cell.image)
  }
}
export function sheetFloatingImageFromPublicDrawing(value: unknown, workbook: WorkbookModel, sheet: SpreadsheetSheet): SheetFloatingImage {
  const drawing = sheetImageRecord(value, '工作表图片'), internal = drawing.resource !== undefined && drawing.anchor !== undefined
  const publicFields = ['drawingType', 'transform', 'sheetTransform', 'axisAlignSheetTransform', 'source', 'srcRect']
  if (internal && !publicFields.some(key => Object.hasOwn(drawing, key))) return normalizeSheetFloatingImage(drawing)
  if (drawing.drawingType !== 0 && !(internal && drawing.drawingType === undefined)) throw new TypeError('目标绘图不是图片。')
  const id = typeof drawing.drawingId === 'string' ? drawing.drawingId : drawing.id
  if (typeof id !== 'string') throw new TypeError('公开图片缺少 drawingId。')
  // Explicit public fields take precedence over the local extension in prepared or edited mixed snapshots.
  const extension = { ...(internal ? drawing : {}), ...object(drawing.officeSheetImage) },
    position = object(drawing.axisAlignSheetTransform ?? drawing.sheetTransform), transform = object(drawing.transform),
    resource = imageResource(drawing, id), axes = sheetClipboardObjectAxes(workbook, sheet),
    kind = drawing.anchorType ?? extension.anchorType ?? '0',
    rotation = sheetImageNumber(transform.angle ?? position.angle ?? extension.rotation ?? 0, '图片旋转角度', -Number.MAX_SAFE_INTEGER)
  const marker = (value: unknown): SheetImageAnchor | undefined => {
    const raw = object(value)
    if (!Object.keys(raw).length) return undefined
    return { row: sheetImageIndex(raw.row, 'row'), column: sheetImageIndex(raw.column, 'column'),
      offsetX: sheetImageNumber(raw.columnOffset ?? raw.offsetX ?? 0, '图片列偏移', -Number.MAX_SAFE_INTEGER),
      offsetY: sheetImageNumber(raw.rowOffset ?? raw.offsetY ?? 0, '图片行偏移', -Number.MAX_SAFE_INTEGER) }
  }
  let from = marker(position.from), to = marker(position.to)
  const start = from && sheetImageProjectAnchor(axes, from), finish = to && sheetImageProjectAnchor(axes, to),
    swapped = drawing.axisAlignSheetTransform !== undefined && sheetImageAxisSwapped(rotation),
    markerWidth = start && finish ? Math.max(0, finish.left - start.left) : undefined,
    markerHeight = start && finish ? Math.max(0, finish.top - start.top) : undefined
  const cachedExtent = (value: unknown, label: string): number | undefined =>
    value === undefined ? undefined : sheetImageNumber(value, label, 0, 100_000) || undefined
  const extensionSize = object(extension.size),
    authoredWidth = extensionSize.width === undefined ? undefined : sheetImageNumber(extensionSize.width, '图片作者宽度', 1, 100_000),
    authoredHeight = extensionSize.height === undefined ? undefined : sheetImageNumber(extensionSize.height, '图片作者高度', 1, 100_000),
    width = cachedExtent(transform.width, '图片宽度') ?? ((swapped ? markerHeight : markerWidth) || authoredWidth || resource.width),
    height = cachedExtent(transform.height, '图片高度') ?? ((swapped ? markerWidth : markerHeight) || authoredHeight || resource.height),
    absolute = object(extension.absolutePosition),
    markerLeft = start && start.left + (kind === '2' && swapped ? (height - width) / 2 : 0),
    markerTop = start && start.top + (kind === '2' && swapped ? (width - height) / 2 : 0),
    rawLeft = transform.left ?? markerLeft ?? absolute.left, rawTop = transform.top ?? markerTop ?? absolute.top,
    left = rawLeft === undefined ? undefined : sheetImageNumber(rawLeft, '图片水平位置', -Number.MAX_SAFE_INTEGER),
    top = rawTop === undefined ? undefined : sheetImageNumber(rawTop, '图片垂直位置', -Number.MAX_SAFE_INTEGER)
  const bounds = left !== undefined && top !== undefined ? { left, top, width, height } : undefined
  if (!from && bounds) {
    const major = sheetImageMajorBounds(bounds, rotation); from = sheetImageAnchorAt(axes, major.left, major.top)
    if (!to) to = sheetImageAnchorAt(axes, major.left + major.width, major.top + major.height)
  }
  const anchor = from ?? extension.anchor
  if (anchor === undefined) throw new TypeError('公开浮动图片缺少有效位置。')
  let crop = Object.hasOwn(drawing, 'srcRect') ? drawing.srcRect ?? undefined : extension.crop
  if (!Object.hasOwn(drawing, 'srcRect') && crop && authoredWidth !== undefined && authoredHeight !== undefined) {
    const authored = sheetImageRecord(crop, '图片裁剪'),
      offset = (edge: 'left' | 'right' | 'top' | 'bottom') => sheetImageNumber(authored[edge], '图片裁剪偏移', -100_000, 100_000)
    crop = { left: offset('left') * width / authoredWidth, right: offset('right') * width / authoredWidth,
      top: offset('top') * height / authoredHeight, bottom: offset('bottom') * height / authoredHeight }
  }
  const image = normalizeSheetFloatingImage({ ...extension, id, resource, anchor, size: { width, height }, anchorType: kind,
    endAnchor: kind === '1' ? to ?? extension.endAnchor : undefined,
    absolutePosition: kind === '2' ? { left, top } : undefined, crop, rotation,
    flipX: transform.flipX ?? position.flipX ?? extension.flipX, flipY: transform.flipY ?? position.flipY ?? extension.flipY,
    keepRatio: drawing.keepRatio ?? extension.keepRatio ?? true, opacity: drawing.opacity ?? extension.opacity ?? 1,
    title: drawing.title ?? extension.title ?? '', altText: drawing.description ?? extension.altText ?? '',
    metadata: { ...object(extension.metadata), publicDrawing: preservedDrawing(drawing) } })
  // Legacy markers describe the unrotated frame. Canonical markers describe its nearest principal-axis rectangle.
  const collapsedMarkers = kind === '1' && start && finish && (finish.left <= start.left || finish.top <= start.top)
  if (drawing.axisAlignSheetTransform === undefined && bounds && !collapsedMarkers) return sheetImageWithFrame(image, bounds, axes)
  return image
}
export function importWorkbookSheetImageResources(workbook: WorkbookModel, resources: unknown): void {
  if (resources === undefined) return
  if (!Array.isArray(resources)) throw new TypeError('工作簿资源必须是数组。')
  workbook.metadata = { ...workbook.metadata, sheetImagePublicResources: cloneModel(resources) }
  for (const entry of resources) {
    const plugin = object(entry)
    if (plugin.name !== SHEET_DRAWING_PLUGIN) continue
    if (typeof plugin.data !== 'string') throw new TypeError('公开绘图资源必须是 JSON 文本。')
    const units: unknown = plugin.data ? JSON.parse(plugin.data) : {}
    for (const [sheetId, raw] of Object.entries(sheetImageRecord(units, '公开绘图资源'))) {
      const sheet = workbook.sheets.find(item => item.id === sheetId)
      if (!sheet) continue
      const table = object(raw), drawings = object(table.data), order = Array.isArray(table.order) ? table.order : Object.keys(drawings)
      const seen = new Set<string>(), imported: SheetFloatingImage[] = []
      for (const id of [...order, ...Object.keys(drawings)]) {
        if (typeof id !== 'string' || seen.has(id)) continue
        seen.add(id)
        const drawing = object(drawings[id])
        if (drawing.drawingType !== 0) continue
        const image = sheetFloatingImageFromPublicDrawing({ ...drawing, drawingId: id }, workbook, sheet)
        imported.push(image)
      }
      if (imported.length) {
        const incomingIds = new Set(imported.map(image => image.id))
        sheet.images = [...(sheet.images ?? []).filter(image => !incomingIds.has(image.id)), ...imported]
        sheet.metadata = { ...sheet.metadata, sheetImagePublicDrawingIds: [...incomingIds] }
      }
    }
  }
}
export function sheetBackgroundImageToPublic(image: SheetBackgroundImage): UniverWorksheetBackgroundImage {
  if (!image.resource) return { ...cloneModel(object(image.metadata?.publicBackgroundImage)), source: image.source,
    imageSourceType: image.imageSourceType, scaleX: image.scaleX, scaleY: image.scaleY }
  return { ...cloneModel(object(image.metadata?.publicBackgroundImage)), source: image.resource.source, imageSourceType: sourceType(image.resource.source),
    scaleX: image.scaleX, scaleY: image.scaleY, officeSheetImageResource: cloneModel(image.resource) }
}
function publicDrawing(image: SheetCellImage | SheetFloatingImage, unitId: string, subUnitId: string): UniverSheetImageDrawing {
  return { ...cloneModel(object(image.metadata?.publicDrawing)), unitId, subUnitId, drawingId: image.id, drawingType: 0,
    source: image.resource.source, imageSourceType: sourceType(image.resource.source), title: image.title, description: image.altText,
    officeSheetImage: cloneModel(image) }
}
export function sheetFloatingImageToPublicDrawing(workbook: WorkbookModel, sheet: SpreadsheetSheet, image: SheetFloatingImage): UniverSheetImageDrawing {
  const axes = sheetClipboardObjectAxes(workbook, sheet), frame = sheetImageFrame(image, axes), major = sheetImageMajorBounds(frame, image.rotation)
  const positions = sheetClipboardBoundsDrawing(workbook, sheet.id, { ...major, left: major.left + GRID_HEADER_WIDTH, top: major.top + GRID_HEADER_HEIGHT })
  const legacy = sheetClipboardBoundsDrawing(workbook, sheet.id, { ...frame, left: frame.left + GRID_HEADER_WIDTH, top: frame.top + GRID_HEADER_HEIGHT })
  const position = { from: sheetImagePublicMarker(image.anchorType === '2' ? positions.from : image.anchor),
    to: sheetImagePublicMarker(image.anchorType === '1' && image.endAnchor ? image.endAnchor : positions.to) }
  const crop = image.crop && { top: image.crop.top * frame.height / image.size.height, bottom: image.crop.bottom * frame.height / image.size.height,
    left: image.crop.left * frame.width / image.size.width, right: image.crop.right * frame.width / image.size.width }
  return { ...publicDrawing(image, workbook.id!, sheet.id), anchorType: image.anchorType ?? '0',
    sheetTransform: { from: sheetImagePublicMarker(legacy.from), to: sheetImagePublicMarker(legacy.to), angle: image.rotation ?? 0, flipX: image.flipX ?? false, flipY: image.flipY ?? false },
    axisAlignSheetTransform: position, srcRect: crop,
    transform: { ...object(image.metadata?.publicDrawing && object(image.metadata.publicDrawing).transform),
      left: frame.left, top: frame.top, width: frame.width, height: frame.height, angle: image.rotation ?? 0, flipX: image.flipX ?? false, flipY: image.flipY ?? false },
    keepRatio: image.keepRatio, opacity: image.opacity }
}
export function sheetCellImageToPublicBody(cell: SpreadsheetCell): UniverSheetImageBody | undefined {
  const saved = object(cell.metadata?.publicBody)
  if (!cell.image && !Object.keys(saved).length) return undefined
  const body = cloneModel(saved) as UniverSheetImageBody, priorId = cell.metadata?.sheetImagePublicDrawingId
  const drawings = body.drawings ??= {}, order = body.drawingsOrder ??= []
  if (typeof priorId === 'string' && priorId !== cell.image?.id) {
    delete drawings[priorId]; body.drawingsOrder = order.filter(id => id !== priorId)
    if (!cell.image && body.body) removeSheetImageBodyDrawing(body.body, priorId)
  }
  if (!cell.image) return body
  const image = cell.image, id = typeof body.id === 'string' ? body.id : `sheet-cell-image-${image.id}`
  body.id = id
  drawings[image.id] = { ...publicDrawing(image, id, id), layoutType: 0, wrapText: 0, behindDoc: 0,
    transform: { ...object(drawings[image.id]?.transform), left: 0, top: 0, width: image.resource.width, height: image.resource.height },
    docTransform: { size: { width: image.resource.width, height: image.resource.height }, positionH: { posOffset: 0 }, positionV: { posOffset: 0 } } }
  if (!body.drawingsOrder!.includes(image.id)) body.drawingsOrder!.push(image.id)
  if (!body.body) body.body = { dataStream: '\u0008\r\n', paragraphs: [{ startIndex: 1 }], sectionBreaks: [{ startIndex: 2 }],
    customBlocks: [{ startIndex: 0, blockId: image.id, blockType: 0 }] }
  else ensureSheetImageBodyDrawing(body.body, image.id, typeof priorId === 'string' ? priorId : undefined)
  return body
}
export function sheetCellDataWithPublicImage(cell: SpreadsheetCell): SpreadsheetCell & UniverSheetCellData {
  const result = { ...cloneModel(cell), v: cell.value, f: cell.formula, s: sheetStyleToPublicData(cell.style),
    t: cell.dataType === 'number' ? 2 : cell.dataType === 'boolean' ? 3 : cell.dataType === 'force_string' ? 4 : 1 } as SpreadsheetCell & UniverSheetCellData
  const body = sheetCellImageToPublicBody(cell)
  if (body) result.p = body
  return result
}
/** A public snapshot is detached. The current media model replaces previously imported image slots. */
export function workbookWithPublicSheetImages(workbook: WorkbookModel): UniverWorkbookData {
  const snapshot = cloneModel(workbook) as unknown as UniverWorkbookData, sheets: UniverWorkbookData['sheets'] = {}
  const resources = Array.isArray(workbook.metadata?.sheetImagePublicResources) ? cloneModel(workbook.metadata.sheetImagePublicResources) as Array<{ name: string; data: string }> : []
  const prior = resources.find(resource => resource.name === SHEET_DRAWING_PLUGIN)
  const drawingMap = prior?.data ? sheetImageRecord(JSON.parse(prior.data), '公开绘图资源') : {}
  for (const sheet of workbook.sheets) {
    const data: UniverWorksheetData = { ...cloneModel(sheet), id: sheet.id, name: sheet.name, cellData: {} }
    delete data.cells; delete data.images; delete data.backgroundImage; delete data.rangeThemeRules
    for (const cell of Object.values(sheet.cells)) (data.cellData[cell.row] ??= {})[cell.col] = sheetCellDataWithPublicImage(cell)
    if (sheet.backgroundImage) data.backgroundImage = sheetBackgroundImageToPublic(sheet.backgroundImage)
    sheets[sheet.id] = data
    const table = object(drawingMap[sheet.id]), previous = object(table.data), drawings = { ...previous }, importedIds = sheet.metadata?.sheetImagePublicDrawingIds
    for (const [id, drawing] of Object.entries(previous)) if (object(drawing).drawingType === 0) delete drawings[id]
    for (const id of Array.isArray(importedIds) ? importedIds : []) if (typeof id === 'string') delete drawings[id]
    const imageIds = (sheet.images ?? []).map(image => image.id), order: string[] = []
    let imageIndex = 0
    for (const id of Array.isArray(table.order) ? table.order : Object.keys(previous)) {
      if (typeof id !== 'string') continue
      const next = object(previous[id]).drawingType === 0 ? imageIds[imageIndex++] : Object.hasOwn(drawings, id) ? id : undefined
      if (next && !order.includes(next)) order.push(next)
    }
    for (const image of sheet.images ?? []) {
      if (Object.hasOwn(drawings, image.id) && object(drawings[image.id]).drawingType !== 0) throw new TypeError('图片 ID 与保留的非图片绘图冲突。')
      drawings[image.id] = sheetFloatingImageToPublicDrawing(workbook, sheet, image)
      if (!order.includes(image.id)) order.push(image.id)
    }
    if (Object.keys(drawings).length || Object.keys(table).length) drawingMap[sheet.id] = { ...table, data: drawings, order }
  }
  for (const sheetId of Object.keys(drawingMap)) if (!workbook.sheets.some(sheet => sheet.id === sheetId)) delete drawingMap[sheetId]
  const outputResources = resources.filter(resource => resource.name !== SHEET_DRAWING_PLUGIN && resource.name !== SHEET_RANGE_THEME_MODEL_PLUGIN)
  const rangeThemes = exportSheetRangeThemeResource(workbook, resources.find(resource => resource.name === SHEET_RANGE_THEME_MODEL_PLUGIN))
  if (rangeThemes) outputResources.push(rangeThemes)
  if (Object.keys(drawingMap).length) outputResources.push({ name: SHEET_DRAWING_PLUGIN, data: JSON.stringify(drawingMap) })
  delete snapshot.kind; delete snapshot.rangeThemes
  snapshot.id = workbook.id!; snapshot.name = workbook.name; snapshot.sheetOrder = workbook.sheets.map(sheet => sheet.id); snapshot.sheets = sheets
  snapshot.resources = outputResources
  return snapshot
}
