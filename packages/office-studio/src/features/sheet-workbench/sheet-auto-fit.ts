import type { ExcelCellView, ExcelSheetView } from '../../services/excelViewAdapter.ts'
import { parseCellKey, sheetViewCellAt } from '../../services/excelViewAdapter.ts'
import { presentSheetCellNumber } from './sheet-number-presentation.ts'
import { sheetCellShrinkToFit, sheetCellTextRotation, sheetCellWrapStrategy } from './sheet-cell-layout.ts'
import { layoutSheetCellText, layoutSheetStackedText, rotatedSheetTextSize, sheetCellTextFont, sheetCellTextIndent, type SheetTextMeasure } from './sheet-cell-text-layout.ts'
import { hiddenSheetRows, hiddenSheetColumns, isSheetRowVisible, isSheetColumnVisible, sheetAxisViewHiddenSpans } from './sheet-row-visibility.ts'
import { positionSheetNumberFormatRuns } from './number-format-render.ts'
import { sheetAxisSpanAt, sheetColumnWidth, sheetRowHeight } from '../../domain/sheet-axis.ts'
import { layoutSheetCellRichText } from './sheet-cell-rich-text-layout.ts'
import { sheetCellTextAlignment } from './sheet-cell-alignment.ts'
import type { SheetStyleContext } from './sheet-style-inheritance.ts'
import { layoutSheetValidationTagCell, sheetValidationCellArrowSpace, sheetValidationListRuleAt,
  type SheetValidationTagOptionsAt } from '../sheet-validation/validation-tag-layout.ts'

export interface SheetAutoFitOptions {
  measure: SheetTextMeasure
  region?: string
  styleContext?: SheetStyleContext
  validationTagOptions?: SheetValidationTagOptionsAt
  columnWidth?: (column: number) => number
  mergeAt?: (row: number, column: number) => { startRow: number; endRow: number; startColumn: number; endColumn: number } | undefined
  /** Batch callers group sparse cells once instead of rescanning the complete sheet for each axis. */
  entries?: readonly (readonly [string, ExcelCellView])[]
  hiddenRows?: ReadonlySet<number>
  hiddenColumns?: ReadonlySet<number>
}

// Fixed-version engine-render/basics/const.ts uses 2000 CSS pixels for both.
export const SHEET_AUTO_FIT_MAX_ROW_HEIGHT = 2000
export const SHEET_AUTO_FIT_MAX_COLUMN_WIDTH = 2000

/** Auto-fit shares display values, font units, explicit breaks and grapheme wrapping with the Canvas layout. */
export function calculateSheetAxisAutoFit(sheet: ExcelSheetView, axis: 'row' | 'column', index: number, options: SheetAutoFitOptions): number {
  const count = axis === 'row' ? sheet.rowCount : sheet.columnCount
  if (!Number.isSafeInteger(index) || index < 0 || index >= count) throw new RangeError('自动适应的行列目标无效。')
  const hiddenRows = options.hiddenRows ?? hiddenSheetRows(sheet)
  const hiddenColumns = options.hiddenColumns ?? hiddenSheetColumns(sheet)
  const hiddenColumnSpans = axis === 'column' ? sheetAxisViewHiddenSpans(sheet, 'column') : undefined
  const widthOfColumn = options.columnWidth ?? (column => isSheetColumnVisible(sheet, column, hiddenColumns) ? sheetColumnWidth(sheet, column) : 0)
  let maximum = axis === 'row' ? sheet.defaultRowHeight ?? 26 : 2
  let hasContent = false
  for (const [key, source] of options.entries ?? Object.entries(sheet.cells)) {
    const position = parseCellKey(key)
    if (!position || position.row < 0 || position.col < 0 || position.row >= sheet.rowCount || position.col >= sheet.columnCount
      || (axis === 'row' ? position.row : position.col) !== index
      || (axis === 'column' ? hiddenColumns.has(position.col) || !!sheetAxisSpanAt(hiddenColumnSpans, position.col)?.value
        : !isSheetColumnVisible(sheet, position.col, hiddenColumns) || widthOfColumn(position.col) <= 0)
      || axis === 'column' && !isSheetRowVisible(sheet, position.row, hiddenRows)) continue
    const merge = options.mergeAt?.(position.row, position.col)
    if (merge && (axis === 'column' || merge.startRow !== position.row || merge.startColumn !== position.col || merge.endRow > merge.startRow)) continue
    const cell = sheet.cells[key] === source ? sheetViewCellAt(sheet, position.row, position.col, options.styleContext) : source
    const presentation = presentSheetCellNumber(cell, sheet.date1904 ?? false, options.region)
    if (!presentation.text && !presentation.runs?.some(run => run.kind === 'fill' && run.character.trim())) continue
    hasContent = true
    const base = sheetCellTextFont(cell), direction = sheetCellTextRotation(cell)
    const indent = sheetCellTextAlignment(cell) === 'center' ? 0 : sheetCellTextIndent(cell, base.fontSize)
    let columnWidth = widthOfColumn(position.col)
    if (merge) {
      columnWidth = 0
      for (let column = merge.startColumn; column <= merge.endColumn; column++) columnWidth += widthOfColumn(column)
    }
    const validation = sheetValidationListRuleAt(sheet.dataValidations, position.row, position.col)
    const tags = layoutSheetValidationTagCell(validation, cell, { width: columnWidth, height: sheetRowHeight(sheet, position.row),
      fit: axis === 'row' ? 'row' : 'column', measure: options.measure, options: options.validationTagOptions?.(position.row, position.col) })
    if (tags) {
      maximum = Math.max(maximum, axis === 'column' ? tags.size.width + tags.leftInset + tags.rightInset : tags.size.height + tags.verticalInset * 2)
      continue
    }
    const arrowSpace = sheetValidationCellArrowSpace(validation)
    const availableWidth = Math.max(0, columnWidth - 12 - indent - arrowSpace)
    const widthOf = (text: string) => options.measure(text, base.font)
    const rich = layoutSheetCellRichText(cell, presentation.text, { width: availableWidth, height: Infinity,
      wrap: axis === 'row' && sheetCellWrapStrategy(cell) === 'wrap', stacked: direction === 'v', measure: options.measure })
    const layout = rich ?? layoutSheetCellText(presentation.text, { width: availableWidth, lineHeight: base.lineHeight,
      wrap: axis === 'row' && sheetCellWrapStrategy(cell) === 'wrap', measure: widthOf })
    const stacked = !rich && direction === 'v' ? layoutSheetStackedText(presentation.text, { height: Infinity, lineHeight: base.fontSize, wrap: false, measure: widthOf }) : undefined
    // Fill characters consume the current column's free space; they do not force a wider column.
    const runs = !rich && !stacked && !/[\r\n]/u.test(presentation.text) && presentation.runs
      ? positionSheetNumberFormatRuns(presentation.runs, axis === 'column' ? 0 : availableWidth, widthOf) : undefined
    const width = rich?.width ?? stacked?.width ?? runs?.width ?? layout.width
    const height = rich?.height ?? stacked?.height ?? (runs ? base.lineHeight : layout.height)
    const size = rotatedSheetTextSize(width, height, direction === 'v' ? 0 : direction)
    const shrink = axis === 'row' && sheetCellShrinkToFit(cell) && size.width > availableWidth && size.width > 0
      ? availableWidth / size.width : 1
    const measured = axis === 'column' ? size.width + 12 + indent + arrowSpace : size.height * shrink + 6
    maximum = Math.max(maximum, measured)
  }
  if (!hasContent && axis === 'column') return widthOfColumn(index)
  return Math.min(axis === 'row' ? SHEET_AUTO_FIT_MAX_ROW_HEIGHT : SHEET_AUTO_FIT_MAX_COLUMN_WIDTH, Math.ceil(maximum))
}
