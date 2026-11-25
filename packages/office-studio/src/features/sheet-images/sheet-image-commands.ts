import type { CellRange, SpreadsheetSheet } from '../../domain/model.ts'
import { createCell } from '../../domain/model.ts'
import type { SheetBackgroundImage, SheetCellImage, SheetFloatingImage, SheetImageResource } from '../../domain/sheet-images.ts'
import { cloneModel } from '../../domain/serialization.ts'
import type { SheetImageSaveConfig } from './sheet-image-save.ts'
import type { SheetImageCropType } from './sheet-image-crop.ts'
import { alignSheetImageSelection, SHEET_IMAGE_ALIGN_OPERATIONS, sheetImageAlignTypeForOperation, type SheetImageAlignType } from './sheet-image-selection.ts'
import { patchSheetFloatingImage, sheetImageWithFrame, type SheetFloatingImagePatch, type SheetImageAxes } from './sheet-image-placement.ts'
import { createSheetFloatingImage, normalizeSheetBackgroundImage, normalizeSheetCellImage, normalizeSheetFloatingImage,
  normalizeSheetImageResource, sheetImageId, sheetImageIndex, sheetImageRecord, SHEET_IMAGE_MAX_BATCH_BYTES, SHEET_IMAGE_MAX_COUNT } from './sheet-image-model.ts'

export const SHEET_IMAGE_COMMANDS = {
  insertFloat: 'sheet.command.insert-float-image', insertCell: 'sheet.command.insert-cell-image',
  insertPrepared: 'sheet.command.insert-sheet-image-drawings', updateBatch: 'sheet.command.set-sheet-image-drawings',
  replace: 'sheet.command.replace-sheet-image-source',
  update: 'sheet.command.set-sheet-image', remove: 'sheet.command.remove-sheet-images',
  addBackground: 'sheet.command.add-worksheet-background-image', deleteBackground: 'sheet.command.delete-worksheet-background-image',
  setBackground: 'sheet.command.set-worksheet-background-image', saveCell: 'sheet.command.save-cell-images',
  save: 'sheet.command.save-sheet-image',
  edit: 'sheet.operation.open-sheet-image-editor',
  arrange: 'sheet.command.set-drawing-arrange',
  cropOpen: 'sheet.operation.open-image-crop', cropClose: 'sheet.operation.close-image-crop', cropAuto: 'sheet.operation.Auto-image-crop',
} as const
export const SHEET_IMAGE_COMMAND_IDS: readonly string[] = [...Object.values(SHEET_IMAGE_COMMANDS), ...Object.values(SHEET_IMAGE_ALIGN_OPERATIONS)]
export type SheetImageCommandId = typeof SHEET_IMAGE_COMMANDS[keyof typeof SHEET_IMAGE_COMMANDS] | typeof SHEET_IMAGE_ALIGN_OPERATIONS[keyof typeof SHEET_IMAGE_ALIGN_OPERATIONS]
export interface SheetImageCommandInput {
  unitId?: string
  sheetId?: string
  subUnitId?: string
  row?: number
  column?: number
  offsetX?: number
  offsetY?: number
  fitToViewport?: boolean
  resources?: readonly SheetImageResource[]
  files?: readonly File[] | null
  source?: string
  name?: string
  images?: readonly SheetFloatingImage[]
  imageId?: string
  drawingId?: string
  cropType?: SheetImageCropType
  isCancel?: boolean
  isAuto?: boolean
  imageIds?: readonly string[]
  patch?: SheetFloatingImagePatch
  drawingIds?: readonly string[]
  drawings?: readonly { unitId?: string; subUnitId?: string; drawingId: string }[]
  alignType?: SheetImageAlignType
  arrangeType?: 0 | 1 | 2 | 3
  zOrder?: number
  backgroundImage?: SheetBackgroundImage | null
  expectedRevision?: number
  expectedHistoryGeneration?: number
  host?: { unitId: string; sheetId: string }
  ranges?: readonly CellRange[]
  saveConfig?: SheetImageSaveConfig
  showSaveDialog?: boolean
  signal?: AbortSignal
}
export type SheetImageTargetSheet = Pick<SpreadsheetSheet, 'id' | 'images' | 'backgroundImage'>
export interface SheetImagePlan {
  images?: SheetFloatingImage[]
  backgroundImage?: SheetBackgroundImage
  cell?: { row: number; column: number; image?: SheetCellImage; replaceContent?: boolean }
  affected: CellRange[]
  changed: boolean
  writesImages: boolean
  writesBackground: boolean
}
export function sheetImageCommandInput(value: unknown): SheetImageCommandInput {
  if (value === undefined) return {}
  const input = sheetImageRecord(value)
  for (const key of ['unitId', 'sheetId', 'subUnitId'] as const) if (input[key] !== undefined) sheetImageId(input[key])
  for (const key of ['expectedRevision', 'expectedHistoryGeneration'] as const) if (input[key] !== undefined &&
    (typeof input[key] !== 'number' || !Number.isSafeInteger(input[key]) || input[key] < 0)) throw new TypeError('图片事务版本无效。')
  if (input.row !== undefined) sheetImageIndex(input.row, 'row')
  if (input.column !== undefined) sheetImageIndex(input.column, 'column')
  if (input.offsetX !== undefined && (typeof input.offsetX !== 'number' || !Number.isFinite(input.offsetX) || input.offsetX < 0)
    || input.offsetY !== undefined && (typeof input.offsetY !== 'number' || !Number.isFinite(input.offsetY) || input.offsetY < 0)) throw new TypeError('图片偏移无效。')
  if (input.fitToViewport !== undefined && typeof input.fitToViewport !== 'boolean') throw new TypeError('图片适应大小选项无效。')
  if (input.cropType !== undefined && !['0', '1', '2', '3', '4', '5', '6', '7', '8', '9'].includes(input.cropType as string)) throw new TypeError('图片裁剪比例无效。')
  for (const key of ['isCancel', 'isAuto'] as const) if (input[key] !== undefined && typeof input[key] !== 'boolean') throw new TypeError('图片裁剪关闭参数无效。')
  if (input.drawingId !== undefined) sheetImageId(input.drawingId)
  if (input.imageId !== undefined) sheetImageId(input.imageId)
  if (input.alignType !== undefined && !['0', '1', '2', '3', '4', '5', '6', '7', '8'].includes(input.alignType as string)) throw new TypeError('图片对齐方式无效。')
  if (input.source !== undefined && typeof input.source !== 'string' || input.name !== undefined && typeof input.name !== 'string') throw new TypeError('图片地址或文件名无效。')
  if (input.resources !== undefined && !Array.isArray(input.resources) || input.images !== undefined && !Array.isArray(input.images)
    || input.files !== undefined && input.files !== null && !Array.isArray(input.files) || input.imageIds !== undefined && !Array.isArray(input.imageIds)
    || input.drawingIds !== undefined && !Array.isArray(input.drawingIds) || input.drawings !== undefined && !Array.isArray(input.drawings)) throw new TypeError('图片批次必须是数组。')
  if (input.arrangeType !== undefined && ![0, 1, 2, 3].includes(input.arrangeType as number)) throw new TypeError('图片排列方式无效。')
  if (input.zOrder !== undefined && typeof input.zOrder !== 'number') throw new TypeError('图片层级必须是数字。')
  if (input.host !== undefined) { const host = sheetImageRecord(input.host, '图片宿主'); sheetImageId(host.unitId); sheetImageId(host.sheetId) }
  if (input.ranges !== undefined && (!Array.isArray(input.ranges) || input.ranges.length > 1024)) throw new TypeError('图片范围必须是有效数组。')
  if (input.showSaveDialog !== undefined && typeof input.showSaveDialog !== 'boolean') throw new TypeError('图片保存对话框参数无效。')
  const result = { ...input } as SheetImageCommandInput
  if (result.imageId === undefined && result.drawingId !== undefined) result.imageId = result.drawingId
  if (result.resources) result.resources = result.resources.map(resource => cloneModel(resource))
  if (result.images) result.images = result.images.map(image => cloneModel(image))
  if (result.files) result.files = [...result.files]
  if (result.imageIds) result.imageIds = [...result.imageIds]
  if (result.drawingIds) result.drawingIds = [...result.drawingIds]
  if (result.drawings) result.drawings = result.drawings.map(value => {
    const drawing = sheetImageRecord(value, '图片对象')
    if (drawing.unitId !== undefined) sheetImageId(drawing.unitId)
    if (drawing.subUnitId !== undefined) sheetImageId(drawing.subUnitId)
    return { unitId: drawing.unitId as string | undefined, subUnitId: drawing.subUnitId as string | undefined, drawingId: sheetImageId(drawing.drawingId) }
  })
  if (result.patch) result.patch = cloneModel(result.patch)
  if (result.backgroundImage) result.backgroundImage = cloneModel(result.backgroundImage)
  if (result.host) result.host = { ...result.host }
  if (result.ranges) result.ranges = result.ranges.map(range => {
    const point = sheetImageRecord(range, '图片范围')
    const startRow = sheetImageIndex(point.startRow, 'row'), endRow = sheetImageIndex(point.endRow, 'row'), startCol = sheetImageIndex(point.startCol, 'column'), endCol = sheetImageIndex(point.endCol, 'column')
    if (endRow < startRow || endCol < startCol) throw new RangeError('图片范围起止坐标无效。')
    return { startRow, endRow, startCol, endCol }
  })
  if (result.saveConfig) result.saveConfig = cloneModel(result.saveConfig)
  return result
}
const cellRange = (row: number, column: number): CellRange => ({ startRow: row, endRow: row, startCol: column, endCol: column })
const imageRange = (image: SheetFloatingImage): CellRange => cellRange(image.anchor.row, image.anchor.column)

/** All media changes are prepared before publication; no resource is registered ahead of content. */
export function planSheetImageCommand(sheet: SheetImageTargetSheet, id: SheetImageCommandId, input: SheetImageCommandInput,
  anchor = { row: 0, column: 0 }, previousCellImage?: SheetCellImage, axes?: SheetImageAxes): SheetImagePlan {
  const result: SheetImagePlan = { affected: [], changed: false, writesImages: false, writesBackground: false }
  const row = sheetImageIndex(input.row ?? anchor.row, 'row'), column = sheetImageIndex(input.column ?? anchor.column, 'column')
  const alignType = sheetImageAlignTypeForOperation(id, input.alignType)
  if (alignType !== undefined) {
    const requested = input.drawingIds ?? input.imageIds ?? input.drawings?.map(drawing => drawing.drawingId) ?? (input.imageId ? [input.imageId] : []),
      ids = new Set(requested.map(sheetImageId))
    if (alignType === '0' || ids.size < 2) return result
    if (ids.size !== requested.length || input.drawings?.some(drawing => drawing.subUnitId !== undefined && drawing.subUnitId !== sheet.id)
      || [...ids].some(id => !sheet.images?.some(image => image.id === id))) throw new TypeError('对齐的图片重复、不存在或不在同一工作表。')
    if (!axes) throw new TypeError('图片对齐需要目标工作表尺寸索引。')
    const selected = (sheet.images ?? []).filter(image => ids.has(image.id)), aligned = alignSheetImageSelection(selected, axes, alignType),
      replacements = new Map(aligned.map(image => [image.id, image])), changed = aligned.some((image, index) => JSON.stringify(image) !== JSON.stringify(selected[index]))
    return { ...result, images: (sheet.images ?? []).map(image => cloneModel(replacements.get(image.id) ?? image)), changed, writesImages: changed,
      affected: selected.map(imageRange) }
  }
  if (id === SHEET_IMAGE_COMMANDS.insertFloat || id === SHEET_IMAGE_COMMANDS.insertPrepared || id === SHEET_IMAGE_COMMANDS.insertCell || id === SHEET_IMAGE_COMMANDS.addBackground) {
    const resources = (input.resources ?? []).map(normalizeSheetImageResource)
    if (resources.reduce((sum, resource) => sum + resource.byteLength, 0) > SHEET_IMAGE_MAX_BATCH_BYTES) throw new RangeError('本次图片总大小超过 32 MiB。')
    if (id !== SHEET_IMAGE_COMMANDS.insertFloat && id !== SHEET_IMAGE_COMMANDS.insertPrepared && resources.length > 1) throw new RangeError('单元格图片和背景图一次只能选择一张。')
    if (id === SHEET_IMAGE_COMMANDS.insertFloat || id === SHEET_IMAGE_COMMANDS.insertPrepared) {
      const added = input.images ? input.images.map(normalizeSheetFloatingImage) : resources.map((resource, index) => {
        const image = createSheetFloatingImage(resource, row, column, index)
        if (input.fitToViewport === false) image.size = { width: resource.width, height: resource.height }
        if (input.offsetX !== undefined) image.anchor.offsetX = input.offsetX + index * 12
        if (input.offsetY !== undefined) image.anchor.offsetY = input.offsetY + index * 12
        if (axes) {
          const left = axes.columns.offset(image.anchor.column) + image.anchor.offsetX, top = axes.rows.offset(image.anchor.row) + image.anchor.offsetY
          return sheetImageWithFrame(image, { left, top, ...image.size }, axes)
        }
        return normalizeSheetFloatingImage(image)
      })
      if (!added.length) return result
      if (added.reduce((sum, image) => sum + image.resource.byteLength, 0) > SHEET_IMAGE_MAX_BATCH_BYTES) throw new RangeError('本次图片总大小超过 32 MiB。')
      const images = [...cloneModel(sheet.images ?? []), ...added]
      if (images.length > SHEET_IMAGE_MAX_COUNT) throw new RangeError('当前工作表图片数量超过 1000 张。')
      if (new Set(images.map(image => image.id)).size !== images.length) throw new Error('图片 ID 已存在。')
      return { ...result, images, affected: added.map(imageRange), changed: true, writesImages: true }
    }
    const resource = resources[0]
    if (!resource) return result
    if (id === SHEET_IMAGE_COMMANDS.insertCell) {
      const image = normalizeSheetCellImage({ id: crypto.randomUUID(), resource, altText: '', title: resource.name })
      return { ...result, cell: { row, column, image, replaceContent: true }, affected: [cellRange(row, column)], changed: true }
    }
    const backgroundImage = normalizeSheetBackgroundImage({ resource, scaleX: resource.scaleX ?? 1, scaleY: resource.scaleY ?? 1 })
    return { ...result, backgroundImage, writesBackground: true, changed: JSON.stringify(backgroundImage) !== JSON.stringify(sheet.backgroundImage) }
  }
  if (id === SHEET_IMAGE_COMMANDS.setBackground || id === SHEET_IMAGE_COMMANDS.deleteBackground) {
    const backgroundImage = id === SHEET_IMAGE_COMMANDS.deleteBackground || input.backgroundImage === null ? undefined : normalizeSheetBackgroundImage(input.backgroundImage)
    return { ...result, backgroundImage, writesBackground: true, changed: JSON.stringify(backgroundImage) !== JSON.stringify(sheet.backgroundImage) }
  }
  if (id === SHEET_IMAGE_COMMANDS.remove) {
    const ids = new Set((input.imageIds ?? (input.imageId ? [input.imageId] : [])).map(sheetImageId))
    const removed = (sheet.images ?? []).filter(image => ids.has(image.id))
    const removesCell = Boolean(previousCellImage && ids.has(previousCellImage.id))
    return { ...result, images: (sheet.images ?? []).filter(image => !ids.has(image.id)).map(cloneModel),
      cell: removesCell ? { row, column } : undefined,
      affected: [...removed.map(imageRange), ...(removesCell ? [cellRange(row, column)] : [])], changed: removed.length > 0 || removesCell, writesImages: removed.length > 0 }
  }
  if (id === SHEET_IMAGE_COMMANDS.arrange) {
    const requested = input.drawingIds ?? input.imageIds ?? (input.imageId ? [input.imageId] : [])
    const ids = new Set(requested.map(sheetImageId)), images = cloneModel(sheet.images ?? [])
    if (!ids.size) return result
    if (input.zOrder !== undefined && ids.size !== 1) return result
    if (ids.size !== requested.length || [...ids].some(imageId => !images.some(image => image.id === imageId))) throw new TypeError('排列的图片 ID 重复或不存在。')
    if (input.zOrder === undefined && input.arrangeType === undefined || input.zOrder !== undefined && input.arrangeType !== undefined) throw new TypeError('请选择一种图片排列方式或目标层级。')
    let ordered = [...images]
    if (input.zOrder !== undefined || input.arrangeType === 2 || input.arrangeType === 3) {
      const selected = images.filter(image => ids.has(image.id)), remaining = images.filter(image => !ids.has(image.id))
      const index = input.arrangeType === 2 ? remaining.length : input.arrangeType === 3 ? 0
        : Math.max(0, Math.min(Number.isNaN(input.zOrder) ? 0 : Math.trunc(input.zOrder!), remaining.length))
      ordered = [...remaining.slice(0, index), ...selected, ...remaining.slice(index)]
    } else if (input.arrangeType === 0) {
      for (let index = ordered.length - 2; index >= 0; index--) if (ids.has(ordered[index]!.id) && !ids.has(ordered[index + 1]!.id))
        [ordered[index], ordered[index + 1]] = [ordered[index + 1]!, ordered[index]!]
    } else {
      for (let index = 1; index < ordered.length; index++) if (ids.has(ordered[index]!.id) && !ids.has(ordered[index - 1]!.id))
        [ordered[index], ordered[index - 1]] = [ordered[index - 1]!, ordered[index]!]
    }
    const changed = ordered.some((image, index) => image.id !== images[index]?.id)
    return { ...result, images: ordered, changed, writesImages: changed, affected: images.filter(image => ids.has(image.id)).map(imageRange) }
  }
  if (id === SHEET_IMAGE_COMMANDS.updateBatch) {
    const updates = (input.images ?? []).map(normalizeSheetFloatingImage), ids = new Set(updates.map(image => image.id))
    if (ids.size !== updates.length || updates.some(image => !sheet.images?.some(current => current.id === image.id))) throw new TypeError('更新的图片 ID 重复或不存在。')
    if (updates.reduce((sum, image) => sum + image.resource.byteLength, 0) > SHEET_IMAGE_MAX_BATCH_BYTES) throw new RangeError('本次图片总大小超过 32 MiB。')
    const replacements = new Map(updates.map(image => [image.id, image])), changed = updates.filter(image => JSON.stringify(image) !== JSON.stringify(sheet.images!.find(current => current.id === image.id)))
    return { ...result, images: (sheet.images ?? []).map(image => cloneModel(replacements.get(image.id) ?? image)), changed: changed.length > 0, writesImages: changed.length > 0,
      affected: changed.flatMap(image => [imageRange(sheet.images!.find(current => current.id === image.id)!), imageRange(image)]) }
  }
  if (id === SHEET_IMAGE_COMMANDS.update) {
    const imageId = sheetImageId(input.imageId), current = sheet.images?.find(image => image.id === imageId)
    const patch = sheetImageRecord(input.patch, '图片更新')
    if (!current) {
      if (!previousCellImage || previousCellImage.id !== imageId) throw new Error('找不到目标图片。')
      if (['anchor', 'size', 'opacity', 'keepRatio', 'placement', 'anchorType', 'endAnchor', 'absolutePosition', 'crop', 'rotation', 'flipX', 'flipY'].some(key => patch[key] !== undefined)) throw new TypeError('单元格图片不接受浮动位置或尺寸。')
      const image = normalizeSheetCellImage({ ...previousCellImage, ...patch, id: imageId })
      return { ...result, cell: { row, column, image }, affected: [cellRange(row, column)], changed: JSON.stringify(image) !== JSON.stringify(previousCellImage) }
    }
    const next = patchSheetFloatingImage(current, patch as SheetFloatingImagePatch, axes)
    return { ...result, images: sheet.images!.map(image => image.id === imageId ? next : cloneModel(image)),
      affected: [imageRange(current), imageRange(next)], changed: JSON.stringify(current) !== JSON.stringify(next), writesImages: true }
  }
  return result
}
export function applySheetImagePlan(sheet: SpreadsheetSheet, plan: SheetImagePlan): void {
  if (plan.writesImages) sheet.images = cloneModel(plan.images ?? [])
  if (plan.writesBackground) { if (plan.backgroundImage) sheet.backgroundImage = cloneModel(plan.backgroundImage); else delete sheet.backgroundImage }
  if (plan.cell) {
    const { row, column, image } = plan.cell, cell = cloneModel(sheet.cells[`${row}:${column}`] ?? createCell(row, column, null))
    if (plan.cell.replaceContent) {
      cell.value = null; cell.formula = null; cell.display = ''; cell.dataType = 'text'; cell.cachedDataType = null
      delete cell.richText; delete cell.sparkline; cell.hyperlink = null
      if (cell.metadata) { delete cell.metadata.formulaError; delete cell.metadata.publicBody; delete cell.metadata.sheetImagePublicDrawingId }
    }
    if (image) cell.image = cloneModel(image)
    else delete cell.image
    sheet.cells[`${row}:${column}`] = cell
  }
}
