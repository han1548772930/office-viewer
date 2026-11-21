import type { ExcelSheetView, ExcelWorkbookView } from '../../services/excelViewAdapter.ts'
import type { GridRange } from '../../rendering/spreadsheet/layout.ts'
import { cellKey, parseCellKey } from './sheet-workbench.ts'
import { transformFormulaReferencesForStructureEdit, type FormulaStructureEdit } from '../formula/formula-references.ts'

export type SheetCellMoveDirection = 'right' | 'down' | 'left' | 'up'

export type SheetCellMoveResult =
  | { status: 'applied'; sheet: ExcelSheetView }
  | { status: 'confirm-merge'; mergeRanges: string[] }
  | { status: 'invalid'; reason: string }

const MAX_SHEET_ROWS = 1_048_576
const MAX_SHEET_COLUMNS = 16_384

export function moveSheetCellRange(
  sheet: ExcelSheetView,
  range: GridRange,
  direction: SheetCellMoveDirection,
  dismantlePartialMerges = false,
): SheetCellMoveResult {
  if (!isValidRange(range, sheet)) return { status: 'invalid', reason: '选区超出工作表范围。' }

  const edit = structureEdit(range, direction)
  const { height, width } = rangeDimensions(range)
  if (direction === 'right' && sheet.columnCount + width > MAX_SHEET_COLUMNS) return { status: 'invalid', reason: '插入后会超出工作表的最大列数。' }
  if (direction === 'down' && sheet.rowCount + height > MAX_SHEET_ROWS) return { status: 'invalid', reason: '插入后会超出工作表的最大行数。' }
  const mergeResult = transformMerges(sheet.merges ?? [], edit)
  if (mergeResult.partial.length && !dismantlePartialMerges) {
    return { status: 'confirm-merge', mergeRanges: mergeResult.partial }
  }

  const cells: Record<string, ExcelSheetView['cells'][string]> = {}
  for (const [key, cell] of Object.entries(sheet.cells)) {
    const position = parseCellKey(key)
    if (!position) continue
    const moved = transformPosition(position.row, position.col, edit)
    if (moved.deleted) continue
    cells[cellKey(moved.row, moved.column)] = cell
  }

  if (direction === 'right' || direction === 'down') {
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      for (let column = range.startColumn; column <= range.endColumn; column += 1) {
        const original = sheet.cells[cellKey(row, column)]
        if (!original || !hasCellFormatting(original)) continue
        const blank = blankFormattedCell(original)
        cells[cellKey(row, column)] = blank
      }
    }
  }

  const next = { ...sheet, cells, merges: mergeResult.merged }
  if (direction === 'right') next.columnCount += width
  if (direction === 'down') next.rowCount += height
  if (direction === 'left' || direction === 'up') {
  }
  return { status: 'applied', sheet: next }
}

export function transformWorkbookReferencesForSheetEdit(
  workbook: ExcelWorkbookView,
  sheetId: string,
  edit: FormulaStructureEdit,
): ExcelWorkbookView {
  const target = workbook.sheets.find((sheet) => sheet.id === sheetId)
  if (!target) return workbook
  const sheets = workbook.sheets.map((sheet) => ({
    ...sheet,
    cells: Object.fromEntries(Object.entries(sheet.cells).map(([key, cell]) => [key, cell.formula
      ? { ...cell, formula: transformFormulaReferencesForStructureEdit(cell.formula, edit, sheet.name, target.name) }
      : cell])),
  }))
  const metadata = workbook.metadata ? structuredClone(workbook.metadata) : undefined
  if (metadata) {
    metadata.definedNames = transformDefinedNames(metadata.definedNames, edit, target.name, target.name)
    metadata.localDefinedNames = transformDefinedNames(metadata.localDefinedNames, edit, target.name, target.name, true)
  }
  return { ...workbook, sheets, metadata }
}

export function structureEdit(range: GridRange, direction: SheetCellMoveDirection): FormulaStructureEdit {
  const { height, width } = rangeDimensions(range)
  const isVertical = direction === 'down' || direction === 'up'
  return {
    kind: direction === 'right' || direction === 'down' ? 'insert' : 'delete',
    axis: isVertical ? 'row' : 'column',
    start: isVertical ? range.startRow : range.startColumn,
    count: isVertical ? height : width,
    bandStart: isVertical ? range.startColumn : range.startRow,
    bandEnd: isVertical ? range.endColumn : range.endRow,
  }
}

function isValidRange(range: GridRange, sheet: ExcelSheetView): boolean {
  return [range.startRow, range.endRow, range.startColumn, range.endColumn].every(Number.isSafeInteger)
    && range.startRow >= 0 && range.startColumn >= 0
    && range.endRow >= range.startRow && range.endColumn >= range.startColumn
    && range.endRow < sheet.rowCount && range.endColumn < sheet.columnCount
}

function rangeDimensions(range: GridRange): { height: number; width: number } {
  return { height: range.endRow - range.startRow + 1, width: range.endColumn - range.startColumn + 1 }
}

function transformPosition(row: number, column: number, edit: FormulaStructureEdit): { row: number; column: number; deleted: boolean } {
  const band = edit.axis === 'row' ? column : row
  const index = edit.axis === 'row' ? row : column
  if (band < edit.bandStart || band > edit.bandEnd || index < edit.start) return { row, column, deleted: false }
  if (edit.kind === 'insert') return setPositionAxis(row, column, edit.axis, index + edit.count)
  if (index < edit.start + edit.count) return { row, column, deleted: true }
  return setPositionAxis(row, column, edit.axis, index - edit.count)
}

function setPositionAxis(row: number, column: number, axis: FormulaStructureEdit['axis'], value: number) {
  return axis === 'row' ? { row: value, column, deleted: false } : { row, column: value, deleted: false }
}

function transformMerges(merges: string[], edit: FormulaStructureEdit): { merged: string[]; partial: string[] } {
  const merged: string[] = []
  const partial: string[] = []
  for (const merge of merges) {
    const [startKey, endKey = startKey] = merge.split(':')
    const start = parseCellKey(startKey ?? '')
    const end = parseCellKey(endKey ?? '')
    if (!start || !end) continue
    const rows = [start.row, end.row].sort((left, right) => left - right)
    const columns = [start.col, end.col].sort((left, right) => left - right)
    const bandStart = edit.axis === 'row' ? columns[0]! : rows[0]!
    const bandEnd = edit.axis === 'row' ? columns[1]! : rows[1]!
    const axisStart = edit.axis === 'row' ? rows[0]! : columns[0]!
    const axisEnd = edit.axis === 'row' ? rows[1]! : columns[1]!
    if (bandEnd < edit.bandStart || bandStart > edit.bandEnd || axisEnd < edit.start) {
      merged.push(merge)
      continue
    }
    const fullyInBand = bandStart >= edit.bandStart && bandEnd <= edit.bandEnd
    const startsAfterInsertion = edit.kind === 'insert' && axisStart >= edit.start
    const beforeInsertion = edit.kind === 'insert' && axisEnd < edit.start
    const fullyDeleted = edit.kind === 'delete' && axisStart >= edit.start && axisEnd < edit.start + edit.count
    const afterDeletion = edit.kind === 'delete' && axisStart >= edit.start + edit.count
    if (fullyInBand && (beforeInsertion || fullyDeleted)) {
      if (fullyDeleted) continue
      merged.push(merge)
      continue
    }
    if (fullyInBand && (startsAfterInsertion || afterDeletion)) {
      const delta = edit.kind === 'insert' ? edit.count : -edit.count
      const nextStart = setPositionAxis(start.row, start.col, edit.axis, (edit.axis === 'row' ? start.row : start.col) + delta)
      const nextEnd = setPositionAxis(end.row, end.col, edit.axis, (edit.axis === 'row' ? end.row : end.col) + delta)
      merged.push(`${cellKey(nextStart.row, nextStart.column)}:${cellKey(nextEnd.row, nextEnd.column)}`)
      continue
    }
    if (!fullyInBand || (edit.kind === 'insert' && axisStart < edit.start && axisEnd >= edit.start)
      || (edit.kind === 'delete' && axisStart < edit.start + edit.count && axisEnd >= edit.start)) {
      partial.push(merge)
      continue
    }
    merged.push(merge)
  }
  return { merged, partial }
}

function hasCellFormatting(cell: ExcelSheetView['cells'][string]): boolean {
  return cell.bold !== undefined || cell.italic !== undefined || cell.underline !== undefined || cell.strike !== undefined
    || cell.align !== undefined || cell.vertical !== undefined || cell.background !== undefined || cell.color !== undefined
    || cell.fontSize !== undefined || cell.fontFamily !== undefined || cell.numberFormat !== undefined || cell.rawNumberFormat !== undefined
    || cell.wrap !== undefined || cell.border !== undefined || cell.rotation !== undefined || cell.rotationVertical !== undefined
    || cell.indent !== undefined || cell.shrinkToFit !== undefined
}

function blankFormattedCell(cell: ExcelSheetView['cells'][string]): ExcelSheetView['cells'][string] {
  const { value: _value, formula: _formula, cachedValue: _cachedValue, rawValue: _rawValue, displayValue: _displayValue, sparkline: _sparkline, ...formatting } = cell
  return { ...formatting, value: '', rawValue: null, displayValue: '' }
}

function transformDefinedNames(value: unknown, edit: FormulaStructureEdit, formulaSheetName: string, targetSheetName: string, local = false): unknown {
  if (!Array.isArray(value)) return value
  return value.map((entry) => {
    if (!entry || typeof entry !== 'object' || !('refersTo' in entry) || typeof entry.refersTo !== 'string') return entry
    if (local && (!('sheet' in entry) || typeof entry.sheet !== 'string' || entry.sheet.toLowerCase() !== targetSheetName.toLowerCase())) return entry
    return { ...entry, refersTo: transformFormulaReferencesForStructureEdit(entry.refersTo, edit, formulaSheetName, targetSheetName) }
  })
}
