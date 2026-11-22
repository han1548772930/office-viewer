import type { CellRange, SpreadsheetSheet } from '../../domain/model.ts'
import { createCell } from '../../domain/model.ts'

export const SET_SHEET_HYPERLINK_COMMAND = 'sheet.command.set-cell-hyperlink'
export const REMOVE_SHEET_HYPERLINK_COMMAND = 'sheet.command.remove-cell-hyperlink'

export interface SheetHyperlinkDraft {
  target: string
  display: string
}

export type SheetHyperlinkTarget =
  | { type: 'external'; url: string }
  | { type: 'internal'; sheetId: string; sheetName: string; range: CellRange }

export interface SheetHyperlinkWorkbook {
  sheets: ReadonlyArray<Pick<SpreadsheetSheet, 'id' | 'name' | 'rowCount' | 'columnCount'>>
}

export function normalizeExternalSheetHyperlink(value: string): string | undefined {
  const input = value.trim()
  if (!input) return undefined

  try {
    const target = /^[a-z][a-z\d+.-]*:/i.test(input)
      ? input
      : /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(input) ? `mailto:${input}` : `http://${input}`
    if (/\s/.test(target)) return undefined
    const url = new URL(target)
    if (url.protocol === 'http:' || url.protocol === 'https:') {
      return url.hostname && !url.username && !url.password ? url.href : undefined
    }
    if (url.protocol === 'mailto:' && url.pathname) return url.href
  } catch {
    return undefined
  }

  return undefined
}

export function normalizeSheetHyperlinkTarget(value: string, workbook: SheetHyperlinkWorkbook): string | undefined {
  const external = normalizeExternalSheetHyperlink(value)
  if (external) return external
  const internal = parseSheetHyperlinkTarget(value, workbook)
  if (!internal || internal.type !== 'internal') return undefined
  const range = internal.range
  const address = `${columnLabel(range.startCol)}${range.startRow + 1}${range.startRow === range.endRow && range.startCol === range.endCol ? '' : `:${columnLabel(range.endCol)}${range.endRow + 1}`}`
  return `#gid=${encodeURIComponent(internal.sheetId)}&range=${encodeURIComponent(address)}`
}

export function parseSheetHyperlinkTarget(value: string, workbook: SheetHyperlinkWorkbook): SheetHyperlinkTarget | undefined {
  const external = normalizeExternalSheetHyperlink(value)
  if (external) return { type: 'external', url: external }
  if (!value.startsWith('#')) return undefined

  const query = value.slice(1)
  if (query.includes('=')) {
    const params = new URLSearchParams(query)
    const sheetId = params.get('gid')
    const sheet = workbook.sheets.find(candidate => candidate.id === sheetId)
    if (!sheet) return undefined
    const range = params.get('range') ? parseHyperlinkRange(params.get('range')!) : { startRow: 0, endRow: 0, startCol: 0, endCol: 0 }
    if (!range || !isRangeInsideSheet(range, sheet)) return undefined
    return { type: 'internal', sheetId: sheet.id, sheetName: sheet.name, range }
  }

  const match = /^#(?:(?:'((?:[^']|'')+)')|([^!]+))!(.+)$/.exec(value)
  if (!match) return undefined
  const sheetName = (match[1]?.replace(/''/g, "'") ?? match[2] ?? '').trim()
  const sheet = workbook.sheets.find(candidate => candidate.name.toLocaleLowerCase() === sheetName.toLocaleLowerCase())
  const range = parseHyperlinkRange(match[3]!)
  if (!sheet || !range || !isRangeInsideSheet(range, sheet)) return undefined
  return { type: 'internal', sheetId: sheet.id, sheetName: sheet.name, range }
}

export function canSetSheetHyperlink(sheet: Pick<SpreadsheetSheet, 'dataValidations'>, row: number, col: number, formula?: string | null): boolean {
  if (formula) return false
  return !sheet.dataValidations?.some(rule => ['checkbox', 'list', 'listMultiple'].includes(rule.type)
    && rule.ranges.some(range => row >= range.startRow && row <= range.endRow && col >= range.startCol && col <= range.endCol))
}

export function setSheetCellHyperlink(workbook: SheetHyperlinkWorkbook, sheet: SpreadsheetSheet, row: number, col: number, draft: SheetHyperlinkDraft): void {
  assertHyperlinkPosition(sheet, row, col)
  const cell = sheet.cells[`${row}:${col}`] ?? createCell(row, col)
  if (!canSetSheetHyperlink(sheet, row, col, cell.formula)) throw new Error('公式、复选框和下拉验证单元格不支持插入此类超链接。')
  const hyperlink = normalizeSheetHyperlinkTarget(draft.target, workbook)
  if (!hyperlink) throw new Error('请输入有效的外部链接，或指向当前工作簿中有效工作表和单元格的内部链接。')
  const value = draft.display || cell.display || String(cell.value ?? '') || hyperlink
  sheet.cells[`${row}:${col}`] = { ...cell, value, display: value, hyperlink }
}

function parseHyperlinkRange(value: string): CellRange | undefined {
  const match = /^\$?([A-Z]{1,3})\$?([1-9]\d*)(?::\$?([A-Z]{1,3})\$?([1-9]\d*))?$/i.exec(value.trim())
  if (!match) return undefined
  const startCol = columnIndex(match[1]!)
  const startRow = Number(match[2]) - 1
  const endCol = match[3] ? columnIndex(match[3]) : startCol
  const endRow = match[4] ? Number(match[4]) - 1 : startRow
  if (startCol == null || endCol == null || endRow < startRow || endCol < startCol) return undefined
  return { startRow, endRow, startCol, endCol }
}

function isRangeInsideSheet(range: CellRange, sheet: Pick<SpreadsheetSheet, 'rowCount' | 'columnCount'>): boolean {
  return range.endRow < (sheet.rowCount ?? 1_048_576) && range.endCol < (sheet.columnCount ?? 16_384)
}

function columnIndex(label: string): number | undefined {
  let index = 0
  for (const character of label.toUpperCase()) index = index * 26 + character.charCodeAt(0) - 64
  return index > 0 && index <= 16_384 ? index - 1 : undefined
}

function columnLabel(index: number): string {
  let label = ''
  for (let cursor = index; cursor >= 0; cursor = Math.floor(cursor / 26) - 1) label = String.fromCharCode(cursor % 26 + 65) + label
  return label
}

export function removeSheetCellHyperlink(sheet: SpreadsheetSheet, row: number, col: number): void {
  assertHyperlinkPosition(sheet, row, col)
  const cell = sheet.cells[`${row}:${col}`]
  if (cell?.formula) throw new Error('公式单元格的链接请通过公式编辑。')
  if (cell?.hyperlink) sheet.cells[`${row}:${col}`] = { ...cell, hyperlink: null }
}

function assertHyperlinkPosition(sheet: Pick<SpreadsheetSheet, 'rowCount' | 'columnCount'>, row: number, col: number): void {
  if (!Number.isInteger(row) || !Number.isInteger(col) || row < 0 || col < 0
    || row >= (sheet.rowCount ?? 1_048_576) || col >= (sheet.columnCount ?? 16_384)) throw new Error('超链接目标单元格超出工作表范围。')
}
