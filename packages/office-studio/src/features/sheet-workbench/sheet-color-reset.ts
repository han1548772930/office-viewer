import type { CellRange, SpreadsheetSheet } from '../../domain/model.ts'

export const RESET_TEXT_COLOR = 'sheet.command.reset-text-color'
export const RESET_BACKGROUND_COLOR = 'sheet.command.reset-background-color'
export type ResetColorCommand = typeof RESET_TEXT_COLOR | typeof RESET_BACKGROUND_COLOR

export function resetSheetColors(sheet: SpreadsheetSheet, ranges: readonly CellRange[], command: ResetColorCommand) {
  for (const cell of Object.values(sheet.cells)) {
    if (!ranges.some(r => cell.row >= r.startRow && cell.row <= r.endRow && cell.col >= r.startCol && cell.col <= r.endCol)) continue
    if (command === RESET_TEXT_COLOR) cell.style.font.color = null
    else cell.style.fill = null
  }
}
