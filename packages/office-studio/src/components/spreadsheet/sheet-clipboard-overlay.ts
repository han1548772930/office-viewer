import type { CellRange } from '../../domain/model.ts'
import type { SheetPasteMode } from '../../features/sheet-workbench/sheet-paste.ts'
import { GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH, type GridRange, type GridViewport, type SpreadsheetGeometry } from '../../rendering/spreadsheet/layout.ts'

export interface SheetClipboardPasteOptionsView { range: CellRange; mode: SheetPasteMode }

export const SHEET_CLIPBOARD_PASTE_OPTIONS = [
  { mode: 'all', label: '全部粘贴', translationKey: 'sheet.context-menu.paste' },
  { mode: 'values', label: '仅粘贴值', translationKey: 'sheet.context-menu.paste-values' },
  { mode: 'formats', label: '仅粘贴格式', translationKey: 'sheet.context-menu.paste-formats' },
  { mode: 'columnWidths', label: '仅粘贴列宽', translationKey: 'sheet.context-menu.paste-column-widths' },
  { mode: 'exceptBorders', label: '保留目标边框粘贴', translationKey: 'sheet.context-menu.paste-except-borders' },
  { mode: 'formulas', label: '仅粘贴公式', translationKey: 'sheet.context-menu.paste-formulas' },
] as const satisfies readonly { mode: SheetPasteMode; label: string; translationKey: string }[]

export function sheetClipboardGridRange(range: CellRange): GridRange {
  return { startRow: range.startRow, endRow: range.endRow, startColumn: range.startCol, endColumn: range.endCol }
}

/** Anchor only to the visible endpoint in its actual frozen/scrolling pane, using CSS-pixel geometry. */
export function sheetClipboardOverlayPosition(geometry: SpreadsheetGeometry, viewport: GridViewport, range?: CellRange): { left: number; top: number } | undefined {
  const width = 44, height = 24, margin = 4
  if (!range || !Number.isSafeInteger(range.endRow) || !Number.isSafeInteger(range.endCol)
    || range.endRow < 0 || range.endCol < 0 || range.endRow >= geometry.rows.count || range.endCol >= geometry.columns.count
    || ![viewport.width, viewport.height, viewport.scrollLeft, viewport.scrollTop].every(Number.isFinite)
    || !Number.isFinite(viewport.zoom) || viewport.zoom <= 0
    || viewport.width < GRID_HEADER_WIDTH + width + margin * 2 || viewport.height < GRID_HEADER_HEIGHT + height + margin * 2) return undefined
  const pane = geometry.panes(viewport).find(candidate => candidate.clip.width > 0 && candidate.clip.height > 0
    && candidate.rows.includes(range.endRow) && candidate.columns.includes(range.endCol))
  if (!pane) return undefined
  const rect = geometry.cellRect(range.endRow, range.endCol, viewport, false), clip = pane.clip
  if (rect.width <= 0 || rect.height <= 0 || rect.x >= clip.x + clip.width || rect.x + rect.width <= clip.x
    || rect.y >= clip.y + clip.height || rect.y + rect.height <= clip.y) return undefined
  return {
    left: Math.max(GRID_HEADER_WIDTH + margin, Math.min(rect.x + rect.width + margin, viewport.width - width - margin)),
    top: Math.max(GRID_HEADER_HEIGHT + margin, Math.min(rect.y + rect.height + margin, viewport.height - height - margin)),
  }
}
