import { createCell, defaultCellStyle, type CellScalar, type SpreadsheetCell, type WorkbookModel } from '../../domain/model.ts'
import { inferSheetCellValueType, normalizeSheetCellValueType } from '../../domain/sheet-cell-value.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { applyExcelCellView, toExcelCellView } from '../../services/excelViewAdapter.ts'
import { translateFormulaReferenceAreas } from '../formula/formula-reference-areas.ts'
import { editSheetCellText } from './sheet-cell-edit.ts'
import { formatSheetValue, isSheetTextNumberFormat } from './sheet-number-format.ts'
import type { SheetPasteMode } from './sheet-paste.ts'

export interface SheetClipboardCellOptions {
  rowDelta: number
  columnDelta: number
  externalFormat?: 'text' | 'html'
  date1904: WorkbookModel['date1904']
  region?: string
  preserveFormulaReferences?: boolean
  contentOnly?: boolean
}

function presentation(value: CellScalar, cell: SpreadsheetCell, options: SheetClipboardCellOptions): string {
  return formatSheetValue(value, cell.style.numberFormat, options.date1904, options.region) ?? (value === null ? '' : String(value))
}

function invalidateContentMetadata(cell: SpreadsheetCell): void {
  if (cell.metadata) delete cell.metadata.formulaError
}

function externalCell(source: SpreadsheetCell, target: SpreadsheetCell, options: SheetClipboardCellOptions): SpreadsheetCell {
  if (source.formula || typeof source.value !== 'string' || source.richText?.length
    || normalizeSheetCellValueType(source.dataType) === 'force_string') return cloneModel(source)
  const base = options.externalFormat === 'text' ? cloneModel(target) : cloneModel(source)
  if (options.externalFormat === 'html' && source.style.numberFormat === 'General' && isSheetTextNumberFormat(target.style.numberFormat)) {
    base.style.numberFormat = target.style.numberFormat
  }
  const parsed = editSheetCellText(toExcelCellView(base), source.value, options.date1904, options.region)
  return applyExcelCellView({ row: source.row, col: source.col }, { ...parsed, richText: null, hyperlink: source.hyperlink ?? null,
    metadata: source.metadata ? cloneModel(source.metadata) : null }, undefined, { date1904: options.date1904, numberFormatLocale: options.region })
}

/** Domain cells keep extensions and explicit clearing independent of the UI view's undefined fallbacks. */
export function createSheetClipboardPastedCell(
  originalSource: SpreadsheetCell | undefined, originalTarget: SpreadsheetCell | undefined,
  position: { row: number; col: number }, mode: Exclude<SheetPasteMode, 'columnWidths'>, options: SheetClipboardCellOptions,
): SpreadsheetCell | undefined {
  const target = originalTarget ? cloneModel(originalTarget) : createCell(position.row, position.col, null)
  const source = originalSource && options.externalFormat ? externalCell(originalSource, target, options) : originalSource
  if (mode === 'formulas') {
    if (!source?.formula) return originalTarget ? cloneModel(originalTarget) : undefined
    target.formula = options.preserveFormulaReferences || options.externalFormat ? source.formula
      : translateFormulaReferenceAreas(source.formula, options.rowDelta, options.columnDelta)
    target.value = null; target.display = target.formula; target.dataType = 'formula'; target.cachedDataType = null
    delete target.richText; delete target.sparkline; invalidateContentMetadata(target)
    return { ...target, id: `${position.row}:${position.col}`, ...position }
  }
  if (mode === 'values') {
    const value = source?.formula && options.externalFormat ? source.formula : source?.value ?? null
    target.value = value; target.formula = null; target.cachedDataType = null
    target.dataType = source?.formula && !options.externalFormat ? source.cachedDataType ?? inferSheetCellValueType(value)
      : source?.formula ? 'text' : source?.dataType ?? inferSheetCellValueType(value)
    if (normalizeSheetCellValueType(target.dataType) === 'formula') target.dataType = inferSheetCellValueType(value)
    target.display = presentation(value, target, options)
    if (source?.richText && !source.formula) target.richText = cloneModel(source.richText)
    else delete target.richText
    delete target.sparkline; invalidateContentMetadata(target)
    return originalTarget || source ? { ...target, id: `${position.row}:${position.col}`, ...position } : undefined
  }
  if (mode === 'formats') {
    if (!source && !originalTarget) return undefined
    target.style = source ? cloneModel(source.style) : defaultCellStyle()
    if (!target.formula) target.display = presentation(target.value, target, options)
    return { ...target, id: `${position.row}:${position.col}`, ...position }
  }
  if (options.contentOnly) {
    const content = createSheetClipboardPastedCell(source, target, position, source?.formula ? 'formulas' : 'values', options)
    if (content) content.hyperlink = source?.hyperlink ?? null
    return content
  }
  if (!source && (mode === 'all' || originalTarget?.style.border === undefined)) return undefined
  const cell = source ? cloneModel(source) : createCell(position.row, position.col, null)
  if (options.externalFormat === 'text') cell.style = cloneModel(target.style)
  if (mode === 'exceptBorders') {
    if (originalTarget?.style.border !== undefined) cell.style.border = cloneModel(originalTarget.style.border)
    else delete cell.style.border
  }
  if (cell.formula && !options.preserveFormulaReferences) {
    cell.formula = options.externalFormat ? cell.formula : translateFormulaReferenceAreas(cell.formula, options.rowDelta, options.columnDelta)
    cell.value = null; cell.display = cell.formula; cell.dataType = 'formula'; cell.cachedDataType = null
    delete cell.richText; invalidateContentMetadata(cell)
  } else if (!cell.formula) cell.display = presentation(cell.value, cell, options)
  if (cell.sparkline && !options.preserveFormulaReferences) cell.sparkline.sourceRange = options.externalFormat ? cell.sparkline.sourceRange
    : translateFormulaReferenceAreas(cell.sparkline.sourceRange, options.rowDelta, options.columnDelta)
  return { ...cell, id: `${position.row}:${position.col}`, ...position }
}
