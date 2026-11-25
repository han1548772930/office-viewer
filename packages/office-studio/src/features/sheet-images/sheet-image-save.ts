import type { CellRange, SpreadsheetSheet } from '../../domain/model.ts'
import type { SheetCellImage, SheetFloatingImage, SheetImageResource } from '../../domain/sheet-images.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { formatProtectionRange } from '../sheet-protection/range-protection.ts'
import { sheetImageIndex, sheetImageRecord, SHEET_IMAGE_MAX_COUNT } from './sheet-image-model.ts'

export interface SaveCellImagesOptions { useCellAddress?: boolean; useColumnIndex?: number }
export interface SheetImageSaveConfig { fileNameParts: Array<'cellAddress' | 'columnValue'>; columnIndex?: number }
export interface SheetImageSaveEntry { row: number; column: number; cellAddress: string; image: SheetCellImage | SheetFloatingImage }
export interface SheetImageSaveRequest { entries: SheetImageSaveEntry[]; columns: Array<{ index: number; label: string }>; columnValues: Record<string, string>; ranges: CellRange[]; target: string }
const inside = (row: number, column: number, range: CellRange) => row >= range.startRow && row <= range.endRow && column >= range.startCol && column <= range.endCol
const rangeAt = (row: number, column: number): CellRange => ({ startRow: row, endRow: row, startCol: column, endCol: column })
export function sheetImageSaveConfig(options: SaveCellImagesOptions = {}): SheetImageSaveConfig {
  if (options.useCellAddress !== undefined && typeof options.useCellAddress !== 'boolean') throw new TypeError('图片文件名选项无效。')
  if (options.useColumnIndex !== undefined) sheetImageIndex(options.useColumnIndex, 'column')
  const fileNameParts: SheetImageSaveConfig['fileNameParts'] = []
  if (options.useCellAddress !== false) fileNameParts.push('cellAddress')
  if (options.useColumnIndex !== undefined) fileNameParts.push('columnValue')
  if (!fileNameParts.length) fileNameParts.push('cellAddress')
  return { fileNameParts, columnIndex: options.useColumnIndex }
}
export function validateSheetImageSaveConfig(value: unknown): SheetImageSaveConfig {
  const input = sheetImageRecord(value, '保存图片')
  if (!Array.isArray(input.fileNameParts) || !input.fileNameParts.length || input.fileNameParts.some(part => part !== 'cellAddress' && part !== 'columnValue')) throw new TypeError('请选择图片文件名的组成方式。')
  if (input.fileNameParts.includes('columnValue')) sheetImageIndex(input.columnIndex, 'column')
  return { fileNameParts: [...new Set(input.fileNameParts)] as SheetImageSaveConfig['fileNameParts'], columnIndex: input.columnIndex as number | undefined }
}
export function collectSheetCellImages(sheet: Pick<SpreadsheetSheet, 'cells'>, ranges: readonly CellRange[]): SheetImageSaveEntry[] {
  const entries = Object.values(sheet.cells).filter(cell => cell.image && ranges.some(range => inside(cell.row, cell.col, range)))
    .sort((left, right) => left.row - right.row || left.col - right.col)
    .map(cell => ({ row: cell.row, column: cell.col, cellAddress: formatProtectionRange(rangeAt(cell.row, cell.col)), image: cloneModel(cell.image!) }))
  if (entries.length > SHEET_IMAGE_MAX_COUNT) throw new RangeError('一次最多保存 1000 张图片。')
  return entries
}
export function createSheetImageSaveRequest(sheet: SpreadsheetSheet, ranges: readonly CellRange[], entries = collectSheetCellImages(sheet, ranges), visible: (row: number, column: number) => boolean = () => true): SheetImageSaveRequest {
  const minRow = Math.min(...ranges.map(range => range.startRow)), maxRow = Math.max(...ranges.map(range => range.endRow))
  const columnValues: Record<string, string> = {}, columns = new Set<number>()
  for (const cell of Object.values(sheet.cells)) {
    if (cell.row < minRow || cell.row > maxRow || !visible(cell.row, cell.col)) continue
    const value = cell.richText?.map(run => run.text).join('') || (cell.value === null ? '' : String(cell.value))
    if (!value.trim()) continue
    if (!ranges.some(range => cell.col >= range.startCol && cell.col <= range.endCol)) columns.add(cell.col)
    columnValues[`${cell.row}:${cell.col}`] = value
  }
  return { entries: cloneModel(entries), ranges: cloneModel([...ranges]), columnValues,
    columns: [...columns].sort((a, b) => a - b).map(index => ({ index, label: formatProtectionRange(rangeAt(0, index)).replace(/1$/, '') })),
    target: ranges.map(formatProtectionRange).join(', ') }
}
export function sheetImageSaveReadRanges(request: SheetImageSaveRequest, config: SheetImageSaveConfig): CellRange[] {
  return request.entries.flatMap(entry => [rangeAt(entry.row, entry.column), ...(config.fileNameParts.includes('columnValue') ? [rangeAt(entry.row, config.columnIndex!)] : [])])
}
function sanitize(value: string): string {
  const result = value.replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_').trim().replace(/[. ]+$/, '').slice(0, 180)
  return /^(?:con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i.test(result) ? `_${result}` : result
}
export function sheetImageExtension(resource: SheetImageResource): string { return resource.mimeType === 'image/jpeg' ? 'jpg' : resource.mimeType.slice(6) }
export function sheetImageSaveNames(request: SheetImageSaveRequest, value: SheetImageSaveConfig): string[] {
  const config = validateSheetImageSaveConfig(value), used = new Set<string>()
  return request.entries.map(entry => {
    const parts = config.fileNameParts.flatMap(part => {
      const text = part === 'cellAddress' ? entry.cellAddress : request.columnValues[`${entry.row}:${config.columnIndex}`] ?? ''
      const safe = sanitize(text); return safe ? [safe] : []
    })
    const base = parts.join('_') || entry.cellAddress, extension = sheetImageExtension(entry.image.resource)
    let count = 0, name = `${base}.${extension}`
    while (used.has(name.toLowerCase())) name = `${base}_${++count}.${extension}`
    used.add(name.toLowerCase()); return name
  })
}
