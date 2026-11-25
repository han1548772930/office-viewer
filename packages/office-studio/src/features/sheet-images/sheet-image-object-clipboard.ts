import type { CellRange, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import type { SheetFloatingImage } from '../../domain/sheet-images.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { assertSheetClipboardBytes } from '../sheet-workbench/sheet-clipboard-data.ts'
import { sheetClipboardObjectAxes, sheetClipboardObjectRange, sheetClipboardRangeBounds } from '../sheet-workbench/sheet-clipboard-object-geometry.ts'
import { assertSheetImageClipboardEntry, assertSheetImageClipboardEntryGeometry, sheetImageClipboardBounds, sheetImageClipboardEntryFrame, sheetImageClipboardFrame, type SheetImageClipboardEntry } from './sheet-image-clipboard.ts'
import { normalizeSheetFloatingImage, sheetImageId, sheetImageNumber, sheetImageRecord, SHEET_IMAGE_MAX_BATCH_BYTES, SHEET_IMAGE_MAX_COUNT } from './sheet-image-model.ts'
import { GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH } from '../../rendering/spreadsheet/layout.ts'
import { sheetImageWithFrame } from './sheet-image-placement.ts'

export interface SheetImageObjectClipboard {
  version: 1
  id: string
  source: { unitId: string; sheetId: string; userId: string }
  copyType: 'copy' | 'cut'
  origin: { left: number; top: number }
  entries: SheetImageClipboardEntry[]
}
export const SHEET_IMAGE_OBJECT_COMMANDS = { copy: 'sheet.command.copy-image-object', cut: 'sheet.command.cut-image-object', paste: 'sheet.command.paste-image-object' } as const
const escapeHtml = (text: string) => text.replace(/[&<>"']/g, character => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[character]!)

export function assertSheetImageObjectClipboard(value: unknown): asserts value is SheetImageObjectClipboard {
  const input = sheetImageRecord(value, '图片对象剪贴板'), source = sheetImageRecord(input.source, '图片剪贴板来源'), origin = sheetImageRecord(input.origin, '图片剪贴板原点')
  if (input.version !== 1 || input.copyType !== 'copy' && input.copyType !== 'cut') throw new TypeError('图片剪贴板版本或类型无效。')
  sheetImageId(input.id); sheetImageId(source.unitId); sheetImageId(source.sheetId); sheetImageId(source.userId)
  sheetImageNumber(origin.left, '图片剪贴板水平原点', -Number.MAX_SAFE_INTEGER); sheetImageNumber(origin.top, '图片剪贴板垂直原点', -Number.MAX_SAFE_INTEGER)
  if (!Array.isArray(input.entries) || !input.entries.length || input.entries.length > SHEET_IMAGE_MAX_COUNT) throw new RangeError('图片对象剪贴板批次数量无效。')
  const ids = new Set<string>(); let bytes = 0
  for (const raw of input.entries) {
    const entry = sheetImageRecord(raw, '图片剪贴板条目') as unknown as SheetImageClipboardEntry
    assertSheetImageClipboardEntry(entry)
    if (ids.has(entry.image.id)) throw new TypeError('图片剪贴板 ID 重复。')
    ids.add(entry.image.id); bytes += entry.image.resource.byteLength
  }
  if (bytes > SHEET_IMAGE_MAX_BATCH_BYTES) throw new RangeError('图片对象剪贴板超过 32 MiB。')
  const entries = input.entries as SheetImageClipboardEntry[]
  if (origin.left !== Math.min(...entries.map(entry => entry.bounds.left)) || origin.top !== Math.min(...entries.map(entry => entry.bounds.top))) throw new TypeError('图片剪贴板原点与对象范围不一致。')
  assertSheetClipboardBytes(JSON.stringify(input))
}

export function captureSheetImageObjects(workbook: WorkbookModel, sheet: SpreadsheetSheet, imageIds: readonly string[], userId: string, copyType: 'copy' | 'cut' = 'copy'): SheetImageObjectClipboard {
  if (!imageIds.length || imageIds.length > SHEET_IMAGE_MAX_COUNT || new Set(imageIds).size !== imageIds.length) throw new RangeError('图片对象选区无效。')
  const entries = imageIds.map(id => {
    const image = sheet.images?.find(image => image.id === id)
    if (!image) throw new Error('找不到选中的图片。')
    const frame = sheetImageClipboardFrame(workbook, sheet, image)
    return { image: sheetImageWithFrame(normalizeSheetFloatingImage(image), { ...frame, left: frame.left - GRID_HEADER_WIDTH,
      top: frame.top - GRID_HEADER_HEIGHT }, sheetClipboardObjectAxes(workbook, sheet)), bounds: sheetImageClipboardBounds(workbook, sheet, image), frame }
  })
  if (entries.reduce((bytes, entry) => bytes + entry.image.resource.byteLength, 0) > SHEET_IMAGE_MAX_BATCH_BYTES) throw new RangeError('图片对象剪贴板超过 32 MiB。')
  const result: SheetImageObjectClipboard = { version: 1, id: crypto.randomUUID(), source: { unitId: workbook.id!, sheetId: sheet.id, userId }, copyType,
    origin: { left: Math.min(...entries.map(entry => entry.bounds.left)), top: Math.min(...entries.map(entry => entry.bounds.top)) }, entries }
  assertSheetImageObjectClipboard(result)
  return result
}
export function sheetImageObjectClipboardContent(clipboard: SheetImageObjectClipboard): { plain: string; html: string } {
  assertSheetImageObjectClipboard(clipboard)
  const html = `<div data-office-image-clipboard-id="${escapeHtml(clipboard.id)}">${clipboard.entries.map(({ image }) =>
    `<img src="${escapeHtml(image.resource.source)}" width="${image.size.width}" height="${image.size.height}" alt="${escapeHtml(image.altText)}" title="${escapeHtml(image.title)}">`).join('')}</div>`
  assertSheetClipboardBytes(html)
  return { plain: '', html }
}
export function sheetImageObjectClipboardIdFromHtml(html: string): string | undefined {
  const match = /\bdata-office-image-clipboard-id\s*=\s*(?:"([A-Za-z0-9._:-]+)"|'([A-Za-z0-9._:-]+)')/i.exec(html)
  return match?.[1] ?? match?.[2]
}
export function sheetImageObjectPasteRanges(workbook: WorkbookModel, sheet: SpreadsheetSheet, target: CellRange, clipboard: SheetImageObjectClipboard): CellRange[] {
  if (!clipboard || clipboard.version !== 1 || !Array.isArray(clipboard.entries) || !clipboard.entries.length || clipboard.entries.length > SHEET_IMAGE_MAX_COUNT
    || !clipboard.origin || ![clipboard.origin.left, clipboard.origin.top].every(Number.isFinite)) throw new TypeError('图片剪贴板几何无效。')
  clipboard.entries.forEach(assertSheetImageClipboardEntryGeometry)
  const origin = sheetClipboardRangeBounds(workbook, sheet, target)
  return clipboard.entries.map(entry => sheetClipboardObjectRange(workbook, sheet.id, { ...entry.bounds,
    left: origin.left + entry.bounds.left - clipboard.origin.left, top: origin.top + entry.bounds.top - clipboard.origin.top }))
}
/** Pasting creates fresh identities and preserves relative positions, including a clipboard produced by immediate cut. */
export function planPasteSheetImageObjects(workbook: WorkbookModel, sheet: SpreadsheetSheet, target: CellRange, clipboard: SheetImageObjectClipboard): { images: SheetFloatingImage[]; ranges: CellRange[] } {
  assertSheetImageObjectClipboard(clipboard)
  const origin = sheetClipboardRangeBounds(workbook, sheet, target)
  const images = clipboard.entries.map(entry => {
    const frame = sheetImageClipboardEntryFrame(entry)
    return sheetImageWithFrame({ ...cloneModel(entry.image), id: crypto.randomUUID() }, { ...frame,
      left: origin.left + frame.left - clipboard.origin.left - GRID_HEADER_WIDTH,
      top: origin.top + frame.top - clipboard.origin.top - GRID_HEADER_HEIGHT }, sheetClipboardObjectAxes(workbook, sheet))
  })
  if ((sheet.images?.length ?? 0) + images.length > SHEET_IMAGE_MAX_COUNT) throw new RangeError('粘贴后图片数量超过 1000 张。')
  return { images, ranges: sheetImageObjectPasteRanges(workbook, sheet, target, clipboard) }
}
