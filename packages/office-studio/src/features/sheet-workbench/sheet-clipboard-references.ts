import type { CellRange, SpreadsheetCell, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { mapFormulaReferenceAreas, type FormulaReferenceArea } from '../formula/formula-reference-areas.ts'
import { rewriteFormulaImportCell } from '../formula/formula-import-context.ts'
import { rewriteWorkbookFormulaDefinedNames } from '../formula/formula-defined-names.ts'
import { parseSheetHyperlinkTarget } from './sheet-hyperlink.ts'
import { sheetRangeContainsCell } from './sheet-range-command.ts'

export interface SheetClipboardMove {
  sourceSheetId: string
  targetSheetId: string
  sourceSheetName: string
  targetSheetName: string
  sourceRange: CellRange
  targetRange: CellRange
}

const sameName = (first: string | undefined, second: string | undefined): boolean => first?.toLowerCase() === second?.toLowerCase()
const qualifier = (name: string): string => /^[\p{L}_\\][\p{L}\p{N}_.\\]*$/u.test(name) ? `${name}!` : `'${name.replace(/'/g, "''")}'!`

function whollyInside(area: FormulaReferenceArea, range: CellRange): boolean {
  return sheetRangeContainsCell(range, Math.min(area.startRow, area.endRow), Math.min(area.startColumn, area.endColumn))
    && sheetRangeContainsCell(range, Math.max(area.startRow, area.endRow), Math.max(area.startColumn, area.endColumn))
}

/** A cut preserves referents; only references wholly inside the source rectangle move, as in fixed Univer. */
export function moveSheetClipboardFormulaReferences(formula: string, beforeOwner: string | undefined, afterOwner: string | undefined, move: SheetClipboardMove): string {
  return mapFormulaReferenceAreas(formula, area => {
    if (area.external || area.threeDimensional) return { ...area }
    const owner = area.sheetName ?? beforeOwner, endOwner = area.endSheetName ?? owner
    if (!sameName(owner, endOwner)) return { ...area }
    const relocated = sameName(owner, move.sourceSheetName) && whollyInside(area, move.sourceRange)
    const nextOwner = relocated ? move.targetSheetName : owner
    const rowDelta = move.targetRange.startRow - move.sourceRange.startRow, columnDelta = move.targetRange.startCol - move.sourceRange.startCol
    const next = relocated ? { ...area, startRow: area.startRow + rowDelta, endRow: area.endRow + rowDelta,
      startColumn: area.startColumn + columnDelta, endColumn: area.endColumn + columnDelta } : { ...area }
    if (nextOwner && (relocated && !sameName(nextOwner, owner) || !area.qualifier && !sameName(nextOwner, afterOwner))) {
      next.qualifier = area.qualifier || !sameName(nextOwner, afterOwner) ? qualifier(nextOwner) : ''
      next.sheetName = next.qualifier ? nextOwner : undefined
      if (area.endQualifier) { next.endQualifier = qualifier(nextOwner); next.endSheetName = nextOwner }
    }
    return next
  })
}

export function moveSheetClipboardHyperlink(value: string, original: WorkbookModel, move: SheetClipboardMove): string {
  const target = parseSheetHyperlinkTarget(value, original)
  if (target?.type !== 'internal' || target.sheetId !== move.sourceSheetId
    || !sheetRangeContainsCell(move.sourceRange, target.range.startRow, target.range.startCol)
    || !sheetRangeContainsCell(move.sourceRange, target.range.endRow, target.range.endCol)) return value
  const rowDelta = move.targetRange.startRow - move.sourceRange.startRow, columnDelta = move.targetRange.startCol - move.sourceRange.startCol
  const params = new URLSearchParams({ gid: move.targetSheetId })
  const address = (row: number, col: number): string => {
    let letters = ''
    for (let index = col; index >= 0; index = Math.floor(index / 26) - 1) letters = String.fromCharCode(65 + index % 26) + letters
    return `${letters}${row + 1}`
  }
  const first = address(target.range.startRow + rowDelta, target.range.startCol + columnDelta)
  const last = address(target.range.endRow + rowDelta, target.range.endCol + columnDelta)
  params.set('range', first === last ? first : `${first}:${last}`)
  return `#${params}`
}

export function moveSheetClipboardCellReferences(cell: SpreadsheetCell, beforeOwner: string, afterOwner: string, original: WorkbookModel, move: SheetClipboardMove): SpreadsheetCell {
  const next = cloneModel(cell)
  if (next.formula) rewriteFormulaImportCell(next, moveSheetClipboardFormulaReferences(next.formula, beforeOwner, afterOwner, move))
  if (next.sparkline) next.sparkline.sourceRange = moveSheetClipboardFormulaReferences(next.sparkline.sourceRange, beforeOwner, afterOwner, move)
  if (next.hyperlink) next.hyperlink = moveSheetClipboardHyperlink(next.hyperlink, original, move)
  if (next.richText) next.richText = next.richText.map(run => typeof run.hyperlink === 'string'
    ? { ...run, hyperlink: moveSheetClipboardHyperlink(run.hyperlink, original, move) } : run)
  return next
}

export function moveSheetClipboardDefinedNames(metadata: Record<string, unknown> | undefined, original: WorkbookModel, move: SheetClipboardMove,
  ownerSheetId?: string): Record<string, unknown> | undefined {
  if (!metadata) return undefined
  const next = cloneModel(metadata), scoped: WorkbookModel = { ...original,
    metadata: ownerSheetId ? undefined : next,
    sheets: original.sheets.map(sheet => ({ ...sheet, metadata: sheet.id === ownerSheetId ? next : undefined })) }
  rewriteWorkbookFormulaDefinedNames(scoped, (formula, owner) => moveSheetClipboardFormulaReferences(formula, owner, owner, move))
  return next
}

export function sheetClipboardMoveOwner(sheet: SpreadsheetSheet, cell: SpreadsheetCell, move: SheetClipboardMove): string {
  return sheet.id === move.targetSheetId && sheetRangeContainsCell(move.targetRange, cell.row, cell.col) ? move.sourceSheetName : sheet.name
}
