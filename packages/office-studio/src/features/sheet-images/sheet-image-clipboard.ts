import type { CellRange, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import type { SheetFloatingImage } from '../../domain/sheet-images.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH } from '../../rendering/spreadsheet/layout.ts'
import type { SheetClipboardSnapshot, SheetClipboardLocation } from '../sheet-workbench/sheet-clipboard-types.ts'
import type { SheetClipboardTargetGeometry } from '../sheet-workbench/sheet-clipboard-target.ts'
import type { SheetPasteMode } from '../sheet-workbench/sheet-paste.ts'
import type { SheetClipboardMove } from '../sheet-workbench/sheet-clipboard-references.ts'
import { sheetClipboardBoundsContain, sheetClipboardObjectAxes, sheetClipboardObjectRange,
  sheetClipboardRangeBounds, type SheetClipboardObjectBounds } from '../sheet-workbench/sheet-clipboard-object-geometry.ts'
import { normalizeSheetFloatingImage, SHEET_IMAGE_MAX_BATCH_BYTES, SHEET_IMAGE_MAX_COUNT } from './sheet-image-model.ts'
import { sheetImageFootprint, sheetImageFrame, sheetImageRotatedBounds, sheetImageWithFrame } from './sheet-image-placement.ts'

export interface SheetImageClipboardEntry { image: SheetFloatingImage; bounds: SheetClipboardObjectBounds; frame?: SheetClipboardObjectBounds }
export function sheetImageClipboardBounds(workbook: WorkbookModel, sheet: SpreadsheetSheet, image: SheetFloatingImage): SheetClipboardObjectBounds {
  const frame = sheetImageFootprint(image, sheetClipboardObjectAxes(workbook, sheet))
  return { ...frame, left: frame.left + GRID_HEADER_WIDTH, top: frame.top + GRID_HEADER_HEIGHT }
}
export function sheetImageClipboardFrame(workbook: WorkbookModel, sheet: SpreadsheetSheet, image: SheetFloatingImage): SheetClipboardObjectBounds {
  const frame = sheetImageFrame(image, sheetClipboardObjectAxes(workbook, sheet))
  return { ...frame, left: frame.left + GRID_HEADER_WIDTH, top: frame.top + GRID_HEADER_HEIGHT }
}
export function sheetImageClipboardEntryFrame(entry: SheetImageClipboardEntry): SheetClipboardObjectBounds {
  return entry.frame ?? { left: entry.bounds.left + (entry.bounds.width - entry.image.size.width) / 2,
    top: entry.bounds.top + (entry.bounds.height - entry.image.size.height) / 2, width: entry.image.size.width, height: entry.image.size.height }
}
/** Capability queries inspect geometry only; media is fully validated when the clipboard is captured or actually pasted. */
export function assertSheetImageClipboardEntryGeometry(entry: SheetImageClipboardEntry): void {
  if (!entry || typeof entry !== 'object' || !entry.bounds || ![entry.bounds.left, entry.bounds.top, entry.bounds.width, entry.bounds.height].every(Number.isFinite)
    || entry.bounds.width <= 0 || entry.bounds.height <= 0) throw new RangeError('图片剪贴板位置或尺寸无效。')
  const image = entry.image
  if (!image || typeof image !== 'object' || !image.size || ![image.size.width, image.size.height, image.rotation ?? 0].every(Number.isFinite)
    || image.size.width <= 0 || image.size.height <= 0) throw new TypeError('图片剪贴板的显示框无效。')
  const frame = sheetImageClipboardEntryFrame(entry)
  if (![frame.left, frame.top, frame.width, frame.height].every(Number.isFinite) || frame.width !== image.size.width || frame.height !== image.size.height) throw new TypeError('图片剪贴板的显示框无效。')
  const bounds = sheetImageRotatedBounds(frame, image.rotation)
  if ((['left', 'top', 'width', 'height'] as const).some(key => Math.abs(bounds[key] - entry.bounds[key]) > Math.max(1e-6, Math.abs(bounds[key]) * 1e-12))) throw new TypeError('图片剪贴板范围与实际旋转几何不一致。')
}
export function assertSheetImageClipboardEntry(entry: SheetImageClipboardEntry): void {
  if (!entry || typeof entry !== 'object') throw new TypeError('图片剪贴板条目无效。')
  assertSheetImageClipboardEntryGeometry({ ...entry, image: normalizeSheetFloatingImage(entry.image) })
}
function fullySelected(rows: readonly number[], first: number, last: number): boolean {
  let low = 0, high = rows.length
  while (low < high) { const middle = (low + high) >>> 1; if (rows[middle]! < first) low = middle + 1; else high = middle }
  return rows[low] === first && rows[low + last - first] === last
}
/** Copy only drawings wholly contained in the selected physical rectangle and visible row projection. */
export function captureSheetClipboardImages(workbook: WorkbookModel, sheet: SpreadsheetSheet, range: CellRange, rows: readonly number[]): SheetImageClipboardEntry[] {
  const selectedBounds = sheetClipboardRangeBounds(workbook, sheet, range)
  return (sheet.images ?? []).flatMap(image => {
    const bounds = sheetImageClipboardBounds(workbook, sheet, image)
    if (!sheetClipboardBoundsContain(selectedBounds, bounds)) return []
    const footprint = sheetClipboardObjectRange(workbook, sheet.id, bounds)
    if (!fullySelected(rows, footprint.startRow, footprint.endRow)) return []
    const frame = sheetImageClipboardFrame(workbook, sheet, image), axes = sheetClipboardObjectAxes(workbook, sheet)
    return [{ image: sheetImageWithFrame(image, { ...frame, left: frame.left - GRID_HEADER_WIDTH, top: frame.top - GRID_HEADER_HEIGHT }, axes), bounds, frame }]
  })
}
export function assertSheetClipboardImages(snapshot: SheetClipboardSnapshot): void {
  if (snapshot.images === undefined) return
  if (!Array.isArray(snapshot.images) || snapshot.images.length > SHEET_IMAGE_MAX_COUNT || !snapshot.source
    || snapshot.formulaOnly || snapshot.externalFormat || !snapshot.sourceOrigin || !snapshot.sourceBounds) throw new TypeError('图片剪贴板缺少有效来源。')
  const ids = new Set<string>()
  let bytes = 0
  for (const entry of snapshot.images) {
    assertSheetImageClipboardEntry(entry)
    if (!sheetClipboardBoundsContain(snapshot.sourceBounds, entry.bounds)) throw new RangeError('图片剪贴板位置或尺寸无效。')
    const image = normalizeSheetFloatingImage(entry.image)
    if (ids.has(image.id)) throw new TypeError('图片剪贴板身份无效。')
    ids.add(image.id); bytes += image.resource.byteLength
  }
  if (bytes > SHEET_IMAGE_MAX_BATCH_BYTES) throw new RangeError('图片剪贴板超过 32 MiB 批次预算。')
}
/** Drawings are sheet content; they join the existing detached paste candidate and its single history transaction. */
export function pasteSheetClipboardImages(original: WorkbookModel, candidate: WorkbookModel, target: SheetClipboardLocation,
  snapshot: SheetClipboardSnapshot, geometry: SheetClipboardTargetGeometry, mode: SheetPasteMode, move?: SheetClipboardMove): void {
  if ((mode !== 'all' && mode !== 'exceptBorders') || snapshot.formulaOnly || snapshot.externalFormat || geometry.contentOnlyMerge || !snapshot.images?.length) return
  assertSheetClipboardImages(snapshot)
  const targetSheet = candidate.sheets.find(sheet => sheet.id === target.sheetId)!
  // Every sheet written here must be detached, including a source with no copied cells.
  const detached = (sheetId: string) => {
    let sheet = candidate.sheets.find(item => item.id === sheetId)!
    if (sheet === original.sheets.find(item => item.id === sheetId)) {
      sheet = cloneModel(sheet); candidate.sheets[candidate.sheets.findIndex(item => item.id === sheetId)] = sheet
    }
    return sheet
  }
  const destination = detached(targetSheet.id), movedIds = new Set(snapshot.images.map(entry => entry.image.id))
  if (move) {
    const source = detached(move.sourceSheetId)
    source.images = (source.images ?? []).filter(image => !movedIds.has(image.id))
  }
  destination.images ??= []
  const origin = snapshot.sourceOrigin!, existingIds = new Set(destination.images.map(image => image.id))
  for (let rowTile = 0; rowTile < geometry.rows.length; rowTile += snapshot.rows.length) {
    for (let colTile = 0; colTile < geometry.columns.length; colTile += snapshot.columns.length) {
      const tile = { startRow: geometry.rows[rowTile]!, endRow: geometry.rows[Math.min(geometry.rows.length - 1, rowTile + snapshot.rows.length - 1)]!,
        startCol: geometry.columns[colTile]!, endCol: geometry.columns[Math.min(geometry.columns.length - 1, colTile + snapshot.columns.length - 1)]! }
      const targetBounds = sheetClipboardRangeBounds(candidate, destination, tile)
      for (const entry of snapshot.images) {
        const bounds = { ...entry.bounds, left: entry.bounds.left + targetBounds.left - origin.left, top: entry.bounds.top + targetBounds.top - origin.top }
        const sourceFrame = sheetImageClipboardEntryFrame(entry)
        const image = sheetImageWithFrame({ ...cloneModel(entry.image), id: move ? entry.image.id : crypto.randomUUID() },
          { ...sourceFrame, left: sourceFrame.left + targetBounds.left - origin.left - GRID_HEADER_WIDTH,
            top: sourceFrame.top + targetBounds.top - origin.top - GRID_HEADER_HEIGHT }, sheetClipboardObjectAxes(candidate, destination))
        if (existingIds.has(image.id)) throw new Error('粘贴后的图片 ID 已存在。')
        existingIds.add(image.id); destination.images.push(image)
        const footprint = sheetClipboardObjectRange(candidate, destination.id, bounds)
        if (destination.rowCount !== undefined) destination.rowCount = Math.max(destination.rowCount, footprint.endRow + 1)
        if (destination.columnCount !== undefined) destination.columnCount = Math.max(destination.columnCount, footprint.endCol + 1)
      }
      if (destination.images.length > SHEET_IMAGE_MAX_COUNT) throw new RangeError('粘贴后图片数量超过 1000 张。')
    }
  }
}
