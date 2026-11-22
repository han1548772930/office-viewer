import type { ExcelCellView, ExcelSheetView } from '../../services/excelViewAdapter.ts'
import type { GridRange } from '../../rendering/spreadsheet/layout.ts'
import { translateFormulaReferences } from '../formula/formula-references.ts'
import { sheetCellNumberValue } from './sheet-number-presentation.ts'

export interface SheetFillCellPatch {
  row: number
  column: number
  cell: ExcelCellView
}

export interface SheetFillPlan {
  range: GridRange
  cellPatches: SheetFillCellPatch[]
}

export type SheetFillApplyType = 'COPY' | 'SERIES' | 'ONLY_FORMAT' | 'NO_FORMAT'

export interface SheetFillOptions {
  applyType?: SheetFillApplyType
  targetCells?: readonly (readonly (ExcelCellView | undefined)[])[]
}

export function detectSheetFillDownEndRow(source: GridRange, sheet: Pick<ExcelSheetView, 'cells' | 'rowCount' | 'columnCount'>): number | undefined {
  const firstAdjacentRow = source.endRow + 1
  if (firstAdjacentRow >= sheet.rowCount) return undefined
  const leftColumn = source.startColumn > 0 ? source.startColumn - 1 : undefined
  const rightColumn = source.endColumn + 1 < sheet.columnCount ? source.endColumn + 1 : undefined
  const referenceColumn = leftColumn !== undefined && hasCellContent(sheet, leftColumn, firstAdjacentRow)
    ? leftColumn
    : rightColumn !== undefined && hasCellContent(sheet, rightColumn, firstAdjacentRow)
      ? rightColumn
      : undefined
  if (referenceColumn === undefined) return undefined

  let detectedEndRow = firstAdjacentRow
  while (detectedEndRow + 1 < sheet.rowCount && hasCellContent(sheet, referenceColumn, detectedEndRow + 1)) {
    detectedEndRow += 1
  }
  for (let row = firstAdjacentRow; row <= detectedEndRow; row += 1) {
    for (let column = source.startColumn; column <= source.endColumn; column += 1) {
      if (hasCellContent(sheet, column, row)) return row - 1 > source.endRow ? row - 1 : undefined
    }
  }
  return detectedEndRow > source.endRow ? detectedEndRow : undefined
}

function hasCellContent(sheet: Pick<ExcelSheetView, 'cells'>, column: number, row: number): boolean {
  const address = `${columnName(column)}${row + 1}`
  const cell = sheet.cells[address]
  return Boolean(cell?.formula || cell?.value)
}

function columnName(index: number): string {
  let result = ''
  for (let value = index; value >= 0; value = Math.floor(value / 26) - 1) result = String.fromCharCode(value % 26 + 65) + result
  return result
}

export function createSheetFillPlan(
  source: GridRange,
  target: { row: number; column: number },
  sourceCells: readonly (readonly (ExcelCellView | undefined)[])[],
  options: SheetFillOptions = {},
): SheetFillPlan {
  const range = {
    startRow: Math.min(source.startRow, target.row),
    endRow: Math.max(source.endRow, target.row),
    startColumn: Math.min(source.startColumn, target.column),
    endColumn: Math.max(source.endColumn, target.column),
  }
  const sourceHeight = source.endRow - source.startRow + 1
  const sourceWidth = source.endColumn - source.startColumn + 1
  const applyType = options.applyType ?? 'SERIES'
  const verticalSeries = sourceWidth === 1 && target.column === source.startColumn
  const horizontalSeries = sourceHeight === 1 && target.row === source.startRow
  const seriesStep = applyType === 'SERIES'
    ? inferNumericStep(sourceCells, verticalSeries ? 'vertical' : horizontalSeries ? 'horizontal' : undefined)
    : undefined
  const cellPatches: SheetFillCellPatch[] = []

  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      if (row >= source.startRow && row <= source.endRow && column >= source.startColumn && column <= source.endColumn) continue
      const sourceRow = source.startRow + ((row - source.startRow) % sourceHeight + sourceHeight) % sourceHeight
      const sourceColumn = source.startColumn + ((column - source.startColumn) % sourceWidth + sourceWidth) % sourceWidth
      const cell = sourceCells[sourceRow - source.startRow]?.[sourceColumn - source.startColumn]
      if (!cell) continue
      const existing = options.targetCells?.[row - range.startRow]?.[column - range.startColumn]
      const copy = applyType === 'ONLY_FORMAT'
        ? mergeCellFormat(existing, cell)
        : applyType === 'NO_FORMAT'
          ? mergeCellContent(existing, cell)
          : cloneCell(cell)
      if (applyType !== 'ONLY_FORMAT' && copy.formula) {
        translateFormula(copy, row - sourceRow, column - sourceColumn)
      } else if (applyType === 'SERIES' && seriesStep !== undefined && (verticalSeries || horizontalSeries)) {
        const index = verticalSeries ? row - source.startRow : column - source.startColumn
        const generated = seriesValue(sourceCells, verticalSeries ? 'vertical' : 'horizontal', index, seriesStep)
        if (generated !== undefined) setLiteralValue(copy, generated)
      }
      cellPatches.push({ row, column, cell: copy })
    }
  }

  return { range, cellPatches }
}

/** Generates a formula-free number series for one-dimensional numeric input. */
function inferNumericStep(cells: readonly (readonly (ExcelCellView | undefined)[])[], axis?: 'vertical' | 'horizontal'): number | undefined {
  if (!axis || !cells.length) return undefined
  const values = axis === 'vertical' ? cells.map(row => numericValue(row[0])) : cells[0]!.map(numericValue)
  if (values.some(value => value === undefined)) return undefined
  if (values.length === 1) return 1
  const step = values[values.length - 1]! - values[values.length - 2]!
  if (values.slice(1).some((value, index) => value! - values[index]! !== step)) return undefined
  return step
}

function numericValue(cell: ExcelCellView | undefined): number | undefined {
  if (!cell || cell.formula) return undefined
  const value = sheetCellNumberValue(cell)
  if (typeof value === 'number' && Number.isFinite(value)) return value
  return undefined
}

function seriesValue(
  cells: readonly (readonly (ExcelCellView | undefined)[])[],
  axis: 'vertical' | 'horizontal',
  index: number,
  step: number,
): number | undefined {
  const seeds = axis === 'vertical' ? cells.map(row => numericValue(row[0])) : cells[0]!.map(numericValue)
  const first = seeds[0]
  return first === undefined ? undefined : first + step * index
}

function setLiteralValue(cell: ExcelCellView, value: number): void {
  cell.value = String(value)
  cell.rawValue = value
  cell.formula = undefined
  cell.dataType = 'number'
  cell.cachedDataType = null
  cell.cachedValue = undefined
  cell.displayValue = undefined
}

function translateFormula(cell: ExcelCellView, rowOffset: number, columnOffset: number): void {
  cell.formula = translateFormulaReferences(cell.formula!, rowOffset, columnOffset)
  cell.value = ''
  cell.cachedValue = undefined
  cell.rawValue = null
  cell.dataType = 'formula'
  cell.cachedDataType = null
  cell.displayValue = undefined
}

const CELL_FORMAT_KEYS = [
  'bold', 'italic', 'underline', 'strike', 'align', 'vertical', 'background', 'color', 'fontSize',
  'numberFormat', 'rawNumberFormat', 'fontFamily', 'wrap', 'wrapStrategy', 'border', 'rotation', 'rotationVertical', 'indent', 'shrinkToFit',
] as const

function mergeCellFormat(target: ExcelCellView | undefined, source: ExcelCellView): ExcelCellView {
  const result = cloneCell(target ?? { value: '' })
  for (const key of CELL_FORMAT_KEYS) {
    if (source[key] !== undefined) Object.assign(result, { [key]: source[key] })
  }
  return result
}

function mergeCellContent(target: ExcelCellView | undefined, source: ExcelCellView): ExcelCellView {
  const result = cloneCell(source)
  if (target) {
    for (const key of CELL_FORMAT_KEYS) {
      if (target[key] !== undefined) Object.assign(result, { [key]: target[key] })
      else delete result[key]
    }
  } else {
    for (const key of CELL_FORMAT_KEYS) delete result[key]
  }
  return result
}

function cloneCell(cell: ExcelCellView): ExcelCellView {
  return {
    ...cell,
    border: cell.border ? Object.fromEntries(Object.entries(cell.border).map(([side, value]) => [
      side,
      typeof value === 'object' && value !== null ? { ...value } : value,
    ])) : undefined,
  }
}
