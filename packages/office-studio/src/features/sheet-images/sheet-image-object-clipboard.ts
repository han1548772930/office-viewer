import type { CellRange, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import type { SheetFloatingImage } from '../../domain/sheet-images.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { assertSheetClipboardBytes } from '../sheet-workbench/sheet-clipboard-data.ts'
import { sheetClipboardObjectAxes, sheetClipboardObjectRange, sheetClipboardRangeBounds } from '../sheet-workbench/sheet-clipboard-object-geometry.ts'
import { assertSheetImageClipboardEntry, assertSheetImageClipboardEntryGeometry, captureSheetImageClipboardEntry, sheetImageClipboardBounds,
  translateSheetImageClipboardEntry, type SheetImageClipboardEntry } from './sheet-image-clipboard.ts'
import { sheetImageId, sheetImageNumber, sheetImageRecord, SHEET_IMAGE_MAX_BATCH_BYTES, SHEET_IMAGE_MAX_COUNT } from './sheet-image-model.ts'

export interface SheetImageObjectClipboard {
  version: 1
  id: string
  source: { unitId: string; sheetId: string; userId: string }
  copyType: 'copy' | 'cut'
  origin: { left: number; top: number }
  entries: SheetImageClipboardEntry[]
}
export const SHEET_IMAGE_OBJECT_COMMANDS = { copy: 'sheet.command.copy-image-object', cut: 'sheet.command.cut-image-object', paste: 'sheet.command.paste-image-object' } as const
const htmlEscapes: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }
const htmlEntities: Record<string, string> = { '&amp;': '&', '&lt;': '<', '&gt;': '>', '&quot;': '"', '&#39;': "'" }
const escapeHtml = (text: string) => text.replace(/[&<>"']/g, character => htmlEscapes[character]!)

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
    return captureSheetImageClipboardEntry(workbook, sheet, image)
  })
  if (entries.reduce((bytes, entry) => bytes + entry.image.resource.byteLength, 0) > SHEET_IMAGE_MAX_BATCH_BYTES) throw new RangeError('图片对象剪贴板超过 32 MiB。')
  const result: SheetImageObjectClipboard = { version: 1, id: crypto.randomUUID(), source: { unitId: workbook.id!, sheetId: sheet.id, userId }, copyType,
    origin: { left: Math.min(...entries.map(entry => entry.bounds.left)), top: Math.min(...entries.map(entry => entry.bounds.top)) }, entries }
  assertSheetImageObjectClipboard(result)
  return result
}
export function sheetImageObjectClipboardContent(clipboard: SheetImageObjectClipboard): { plain: string; html: string } {
  assertSheetImageObjectClipboard(clipboard)
  // HTML carries the actual author snapshot, so a new Studio can preserve crop/flip/affine after the broker is disposed.
  // Media sources already live in IMG.src; do not duplicate their base64 bytes in the author descriptor.
  const descriptor = { ...clipboard, entries: clipboard.entries.map(entry => {
    const { source: _source, ...resource } = entry.image.resource
    return { ...entry, image: { ...entry.image, resource } }
  }) }
  const encoded = encodeURIComponent(JSON.stringify(descriptor))
  const html = `<div data-office-image-clipboard-id="${escapeHtml(clipboard.id)}" data-office-image-clipboard-model="${encoded}">${clipboard.entries.map(({ image }, index) =>
    `<img data-office-image-index="${index}" src="${escapeHtml(image.resource.source)}" width="${image.size.width}" height="${image.size.height}" alt="${escapeHtml(image.altText)}" title="${escapeHtml(image.title)}">`).join('')}</div>`
  assertSheetClipboardBytes(html)
  return { plain: '', html }
}
export function sheetImageObjectClipboardIdFromHtml(html: string): string | undefined {
  const match = /\bdata-office-image-clipboard-id\s*=\s*(?:"([A-Za-z0-9._:-]+)"|'([A-Za-z0-9._:-]+)')/i.exec(html)
  return match?.[1] ?? match?.[2]
}
/** Decode only this versioned internal flavour. Foreign IMG content stays on the ordinary external image path. */
export function sheetImageObjectClipboardFromHtml(html: string): SheetImageObjectClipboard | undefined {
  assertSheetClipboardBytes(html)
  const match = /\bdata-office-image-clipboard-model\s*=\s*(?:"([^"]*)"|'([^']*)')/i.exec(html)
  if (!match) return undefined
  let parsed: unknown
  try { parsed = JSON.parse(decodeURIComponent(match[1] ?? match[2]!)) }
  catch { throw new TypeError('图片剪贴板作者数据无法读取。') }
  const input = sheetImageRecord(parsed, '图片剪贴板作者数据')
  if (!Array.isArray(input.entries) || !input.entries.length || input.entries.length > SHEET_IMAGE_MAX_COUNT) throw new TypeError('图片剪贴板作者条目无效。')
  const attribute = (tag: string, name: string): string | undefined => {
    const found = new RegExp(`(?:^|\\s)${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`, 'i').exec(tag)
    return found ? found[1] ?? found[2] : undefined
  }
  const sources = new Map<number, string>()
  for (const tag of html.match(/<img\b[^>]*>/gi) ?? []) {
    const rawIndex = attribute(tag, 'data-office-image-index')
    if (rawIndex === undefined) continue
    const index = Number(rawIndex), source = attribute(tag, 'src')
    if (!/^\d+$/.test(rawIndex) || !Number.isInteger(index) || index >= input.entries.length || source === undefined || sources.has(index)) throw new TypeError('图片剪贴板媒体身份无效。')
    // Only the entities emitted by escapeHtml are decoded, exactly once.
    sources.set(index, source.replace(/&(amp|lt|gt|quot|#39);/g, entity => htmlEntities[entity]!))
  }
  input.entries.forEach((raw, index) => {
    const entry = sheetImageRecord(raw, '图片剪贴板条目'), image = sheetImageRecord(entry.image, '图片剪贴板图片'), resource = sheetImageRecord(image.resource, '图片剪贴板资源')
    if (resource.source === undefined) {
      const source = sources.get(index)
      if (!source) throw new TypeError('图片剪贴板缺少实际媒体。')
      resource.source = source
    }
  })
  assertSheetImageObjectClipboard(parsed)
  if (sheetImageObjectClipboardIdFromHtml(html) !== parsed.id) throw new TypeError('图片剪贴板标识与作者数据不一致。')
  return cloneModel(parsed)
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
  const axes = sheetClipboardObjectAxes(workbook, sheet), images = clipboard.entries.map(entry => translateSheetImageClipboardEntry(entry,
    axes, origin.left - clipboard.origin.left, origin.top - clipboard.origin.top, crypto.randomUUID()))
  if ((sheet.images?.length ?? 0) + images.length > SHEET_IMAGE_MAX_COUNT) throw new RangeError('粘贴后图片数量超过 1000 张。')
  const ranges = images.map(image => sheetClipboardObjectRange(workbook, sheet.id, sheetImageClipboardBounds(workbook, sheet, image)))
  return { images, ranges }
}
