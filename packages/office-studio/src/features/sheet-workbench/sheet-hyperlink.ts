import type { SpreadsheetSheet } from '../../domain/model.ts'
import { createCell } from '../../domain/model.ts'

export const SET_SHEET_HYPERLINK_COMMAND = 'sheet.command.set-cell-hyperlink'
export const REMOVE_SHEET_HYPERLINK_COMMAND = 'sheet.command.remove-cell-hyperlink'

export interface SheetHyperlinkDraft {
  target: string
  display: string
}

export function normalizeExternalSheetHyperlink(value: string): string | undefined {
  const input = value.trim()
  if (!input) return undefined

  try {
    const target = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(input)
      ? `mailto:${input}`
      : /^[a-z][a-z\d+.-]*:/i.test(input) ? input : `http://${input}`
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

export function canSetSheetHyperlink(sheet: Pick<SpreadsheetSheet, 'dataValidations'>, row: number, col: number, formula?: string | null): boolean {
  if (formula) return false
  return !sheet.dataValidations?.some(rule => ['checkbox', 'list', 'listMultiple'].includes(rule.type)
    && rule.ranges.some(range => row >= range.startRow && row <= range.endRow && col >= range.startCol && col <= range.endCol))
}

export function setSheetCellHyperlink(sheet: SpreadsheetSheet, row: number, col: number, draft: SheetHyperlinkDraft): void {
  assertHyperlinkPosition(sheet, row, col)
  const cell = sheet.cells[`${row}:${col}`] ?? createCell(row, col)
  if (!canSetSheetHyperlink(sheet, row, col, cell.formula)) throw new Error('公式、复选框和下拉验证单元格不支持插入此类超链接。')
  const hyperlink = normalizeExternalSheetHyperlink(draft.target)
  if (!hyperlink) throw new Error('请输入有效的 HTTP、HTTPS 或 mailto 链接。')
  const value = draft.display || cell.display || String(cell.value ?? '') || hyperlink
  sheet.cells[`${row}:${col}`] = { ...cell, value, display: value, hyperlink }
}

export function removeSheetCellHyperlink(sheet: SpreadsheetSheet, row: number, col: number): void {
  assertHyperlinkPosition(sheet, row, col)
  const cell = sheet.cells[`${row}:${col}`]
  if (cell?.formula) throw new Error('公式单元格的链接请通过公式编辑。')
  if (cell?.hyperlink) sheet.cells[`${row}:${col}`] = { ...cell, hyperlink: null }
}

function assertHyperlinkPosition(sheet: SpreadsheetSheet, row: number, col: number): void {
  if (!Number.isInteger(row) || !Number.isInteger(col) || row < 0 || col < 0
    || row >= (sheet.rowCount ?? 1_048_576) || col >= (sheet.columnCount ?? 16_384)) throw new Error('超链接目标单元格超出工作表范围。')
}
