import type { ExcelCellView } from '../../services/excelViewAdapter.ts'
import { translateFormulaReferences } from '../formula/formula-references.ts'
import { clearSheetCellFormat } from './sheet-toolbar-format.ts'

export type SheetPasteMode = 'all' | 'values' | 'formats' | 'columnWidths' | 'exceptBorders'

export interface SheetPasteCellPatch {
  rowOffset: number
  columnOffset: number
  cell: ExcelCellView
}

export interface SheetPasteWidthPatch {
  column: number
  width: number
}

export interface SheetPasteRequest {
  mode: SheetPasteMode
  sourceCells: readonly (readonly (ExcelCellView | undefined)[])[]
  targetCells: readonly (readonly (ExcelCellView | undefined)[])[]
  sourceStartRow: number
  sourceStartColumn: number
  targetStartRow: number
  targetStartColumn: number
  sourceColumnWidths?: Readonly<Record<number, number>>
  sourceDefaultColumnWidth?: number
}

export interface SheetPastePlan {
  cellPatches: SheetPasteCellPatch[]
  columnWidthPatches: SheetPasteWidthPatch[]
}

const FORMAT_KEYS = [
  'bold', 'italic', 'underline', 'strike', 'align', 'vertical', 'background', 'color',
  'fontSize', 'fontFamily', 'numberFormat', 'rawNumberFormat', 'wrap', 'border', 'rotation',
  'rotationVertical', 'indent', 'shrinkToFit',
] as const satisfies readonly (keyof ExcelCellView)[]

export function createSheetPastePlan(request: SheetPasteRequest): SheetPastePlan {
  const { mode, sourceCells } = request
  if (mode === 'columnWidths') {
    const columnCount = sourceCells.reduce((count, row) => Math.max(count, row.length), 0)
    const widths = request.sourceColumnWidths ?? {}
    const defaultWidth = request.sourceDefaultColumnWidth ?? 96
    return {
      cellPatches: [],
      columnWidthPatches: Array.from({ length: columnCount }, (_, offset) => ({
        column: request.targetStartColumn + offset,
        width: widths[request.sourceStartColumn + offset] ?? defaultWidth,
      })),
    }
  }

  const cellPatches: SheetPasteCellPatch[] = []
  const rowCount = sourceCells.length
  for (let rowOffset = 0; rowOffset < rowCount; rowOffset += 1) {
    const sourceRow = sourceCells[rowOffset] ?? []
    const targetRow = request.targetCells[rowOffset] ?? []
    for (let columnOffset = 0; columnOffset < sourceRow.length; columnOffset += 1) {
      const source = sourceRow[columnOffset]
      const target = targetRow[columnOffset] ?? { value: '' }
      cellPatches.push({
        rowOffset,
        columnOffset,
        cell: pasteCell(source, target, mode, request.targetStartRow - request.sourceStartRow, request.targetStartColumn - request.sourceStartColumn),
      })
    }
  }
  return { cellPatches, columnWidthPatches: [] }
}

function pasteCell(
  source: ExcelCellView | undefined,
  target: ExcelCellView,
  mode: Exclude<SheetPasteMode, 'columnWidths'>,
  rowDelta: number,
  columnDelta: number,
): ExcelCellView {
  if (mode === 'values') return pasteValue(source, target)
  if (mode === 'formats') return pasteFormat(source, target)
  const copied = source ? cloneCell(source) : clearSheetCellFormat({ value: '' })
  if (copied.formula) {
    copied.formula = translateFormulaReferences(copied.formula, rowDelta, columnDelta)
    copied.value = ''
    copied.cachedValue = undefined
    copied.rawValue = undefined
    copied.displayValue = undefined
  }
  if (mode === 'exceptBorders') copied.border = cloneBorder(target.border)
  return copied
}

function pasteValue(source: ExcelCellView | undefined, target: ExcelCellView): ExcelCellView {
  const value = source?.formula
    ? source.cachedValue ?? source.rawValue ?? ''
    : source?.rawValue ?? source?.value ?? ''
  const pastedValue = value === null || value === undefined ? '' : String(value)
  const { formula: _formula, cachedValue: _cachedValue, rawValue: _rawValue, displayValue: _displayValue, ...preserved } = target
  return {
    ...preserved,
    value: pastedValue,
    rawValue: source?.formula ? source.rawValue ?? source.cachedValue : source?.rawValue,
    displayValue: undefined,
  }
}

function pasteFormat(source: ExcelCellView | undefined, target: ExcelCellView): ExcelCellView {
  const defaults = clearSheetCellFormat({ value: '' })
  const sourceStyle = source ?? defaults
  const formatted = { ...defaults }
  for (const key of FORMAT_KEYS) {
    if (key === 'border') {
      if (sourceStyle.border !== undefined) formatted.border = cloneBorder(sourceStyle.border)
    } else {
      const value = sourceStyle[key]
      if (value !== undefined) Object.assign(formatted, { [key]: value })
    }
  }
  return {
    ...formatted,
    value: target.value,
    formula: target.formula,
    cachedValue: target.cachedValue,
    rawValue: target.rawValue,
    displayValue: target.displayValue,
  }
}

function cloneCell(cell: ExcelCellView): ExcelCellView {
  return { ...cell, border: cloneBorder(cell.border) }
}

function cloneBorder(border: Record<string, unknown> | undefined): Record<string, unknown> | undefined {
  if (!border) return undefined
  return Object.fromEntries(Object.entries(border).map(([side, value]) => [
    side,
    typeof value === 'object' && value !== null ? { ...value } : value,
  ]))
}
