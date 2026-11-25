import { createCell, defaultCellStyle, type CellScalar, type CellStyle, type SpreadsheetCell, type WorkbookModel } from '../../domain/model.ts'
import { inferSheetCellValueType, normalizeSheetCellValueType } from '../../domain/sheet-cell-value.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { applyExcelCellView, toExcelCellView } from '../../services/excelViewAdapter.ts'
import { translateFormulaReferenceAreas } from '../formula/formula-reference-areas.ts'
import { editSheetCellText } from './sheet-cell-edit.ts'
import { formatSheetValue, isSheetTextNumberFormat } from './sheet-number-format.ts'
import type { SheetPasteMode } from './sheet-paste.ts'
import { explicitSheetCellStyle, isDefaultSheetCellStyle } from './sheet-style-region.ts'
import { copySheetCellImage } from '../sheet-images/sheet-image-model.ts'

export interface SheetClipboardCellOptions {
  rowDelta: number
  columnDelta: number
  externalFormat?: 'text' | 'html'
  date1904: WorkbookModel['date1904']
  region?: string
  preserveFormulaReferences?: boolean
  /** A cut moves the author sources and range rule together; a copied effective theme must not become direct formatting. */
  preserveRangeThemeSource?: boolean
  contentOnly?: boolean
  sourceIsVirtual?: boolean
  targetStyle?: CellStyle
  targetHasInheritedStyle?: boolean
}

function presentation(value: CellScalar, cell: SpreadsheetCell, options: SheetClipboardCellOptions): string {
  return formatSheetValue(value, cell.style.numberFormat, options.date1904, options.region) ?? (value === null ? '' : String(value))
}

function invalidateContentMetadata(cell: SpreadsheetCell): void {
  if (cell.metadata) { delete cell.metadata.formulaError; delete cell.metadata.publicBody; delete cell.metadata.sheetImagePublicDrawingId }
}

function externalCell(source: SpreadsheetCell, target: SpreadsheetCell, options: SheetClipboardCellOptions): SpreadsheetCell {
  if (source.image || source.formula || typeof source.value !== 'string' || source.richText?.length
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
  if (options.targetStyle) target.style = cloneModel(options.targetStyle)
  const source = originalSource && options.externalFormat ? externalCell(originalSource, target, options) : originalSource
  if (mode === 'formulas') {
    if (!source?.formula) return originalTarget ? cloneModel(originalTarget) : undefined
    target.formula = options.preserveFormulaReferences || options.externalFormat ? source.formula
      : translateFormulaReferenceAreas(source.formula, options.rowDelta, options.columnDelta)
    target.value = null; target.display = target.formula; target.dataType = 'formula'; target.cachedDataType = null
    delete target.richText; delete target.sparkline; delete target.image; invalidateContentMetadata(target)
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
    if (source?.image && !source.formula) {
      target.image = options.preserveFormulaReferences ? cloneModel(source.image) : copySheetCellImage(source.image)
      if (source.metadata?.publicBody) target.metadata = { ...target.metadata, publicBody: cloneModel(source.metadata.publicBody),
        sheetImagePublicDrawingId: source.metadata.sheetImagePublicDrawingId ?? source.image.id }
    }
    else delete target.image
    return originalTarget || source && !options.sourceIsVirtual ? { ...target, id: `${position.row}:${position.col}`, ...position } : undefined
  }
  if (options.preserveRangeThemeSource && mode === 'all' && !options.externalFormat && !options.contentOnly) {
    if (!source || options.sourceIsVirtual) return undefined
    const cell = cloneModel(source)
    if (cell.image && !options.preserveFormulaReferences) cell.image = copySheetCellImage(cell.image)
    return { ...cell, id: `${position.row}:${position.col}`, ...position }
  }
  if (options.sourceIsVirtual && source && !options.contentOnly && !options.targetHasInheritedStyle && isDefaultSheetCellStyle(source.style)
    && (mode === 'all' || !originalTarget)) return undefined
  if (mode === 'formats') {
    if (!source && !originalTarget) return undefined
    target.style = explicitSheetCellStyle(source?.style ?? defaultCellStyle())
    target.authoredStyle = cloneModel(target.style)
    delete target.themeStyle
    delete target.authoredStyleReference; delete target.clearedStyleProperties
    if (!target.formula) target.display = presentation(target.value, target, options)
    return { ...target, id: `${position.row}:${position.col}`, ...position }
  }
  if (options.contentOnly) {
    const content = createSheetClipboardPastedCell(source, target, position, source?.formula ? 'formulas' : 'values', options)
    if (content) content.hyperlink = source?.hyperlink ?? null
    return content
  }
  if (!source && (mode === 'all' || target.style.border === undefined)) return undefined
  const cell = source ? cloneModel(source) : createCell(position.row, position.col, null)
  if (cell.image && !options.preserveFormulaReferences) cell.image = copySheetCellImage(cell.image)
  if (options.externalFormat === 'text') cell.style = cloneModel(target.style)
  else cell.style = explicitSheetCellStyle(cell.style)
  if (mode === 'exceptBorders') {
    cell.style.border = cloneModel(target.style.border ?? {})
  }
  // Pasted formatting keeps the source presentation even under different destination axis defaults.
  cell.authoredStyle = options.externalFormat === 'text' ? cloneModel(target.authoredStyle ?? target.style) : cloneModel(cell.style)
  if (options.externalFormat !== 'text') delete cell.themeStyle
  else cell.themeStyle = target.themeStyle
  if (options.externalFormat !== 'text') { delete cell.authoredStyleReference; delete cell.clearedStyleProperties }
  else { cell.authoredStyleReference = target.authoredStyleReference; cell.clearedStyleProperties = target.clearedStyleProperties }
  if (cell.formula && !options.preserveFormulaReferences) {
    cell.formula = options.externalFormat ? cell.formula : translateFormulaReferenceAreas(cell.formula, options.rowDelta, options.columnDelta)
    cell.value = null; cell.display = cell.formula; cell.dataType = 'formula'; cell.cachedDataType = null
    delete cell.richText; invalidateContentMetadata(cell)
  } else if (!cell.formula) cell.display = presentation(cell.value, cell, options)
  if (cell.sparkline && !options.preserveFormulaReferences) cell.sparkline.sourceRange = options.externalFormat ? cell.sparkline.sourceRange
    : translateFormulaReferenceAreas(cell.sparkline.sourceRange, options.rowDelta, options.columnDelta)
  return { ...cell, id: `${position.row}:${position.col}`, ...position }
}
