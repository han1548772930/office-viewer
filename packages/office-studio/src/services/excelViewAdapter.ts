import type { CellScalar, CellStyle, ChartModel, PivotCacheModel, PivotTableModel, SheetTableModel, WorkbookModel, SpreadsheetCell, SheetDataValidationRule, SheetRangeProtectionRule, SheetWorksheetProtectionRule, SheetOutlineState, SheetAutoFilter, SparklineSpec, SheetConditionalFormatRule, SlicerCacheModel, SlicerModel } from '../domain/model.ts'
import { classifySheetNumberFormat, formatSheetValue, isSheetTextNumberFormat, SHEET_NUMBER_FORMATS, type SheetNumberFormat } from '../features/sheet-workbench/sheet-number-format.ts'
import { refreshAutoFilter } from '../features/sheet-sort-filter/sort-filter-engine.ts'
import { getSheetOutlineHiddenSpans } from '../features/sheet-outline/sheet-outline.ts'
import { inferSheetCellValueType, normalizeSheetCellValueType, parseSheetCellText } from '../domain/sheet-cell-value.ts'
import type { SheetAxisSpan } from '../domain/sheet-axis.ts'
import { cloneModel } from '../domain/serialization.ts'
import type { SheetCellStylePatch, SheetCellStyleProperty, SheetStyleSource } from '../domain/model.ts'
import type { SheetRangeThemeRule, SheetRangeThemeStyle } from '../domain/model.ts'
import type { SheetBackgroundImage, SheetCellImage, SheetFloatingImage } from '../domain/sheet-images.ts'
import type { SheetDrawingShapeDTO, SheetDrawingState } from '../domain/sheet-drawings.ts'
import type { ISlicerStylesDefinition, ITimelineStylesDefinition } from '../domain/slicer-styles.ts'
import { composeSheetCellStyles, mergeSheetCellStylePatch, refreshWorkbookCellStyles, resolveSheetCellStyle, sheetStyleResetsAfterPatch,
  type SheetStyleContext } from '../features/sheet-workbench/sheet-style-inheritance.ts'

export interface ExcelCellView {
  value: string
  authoredStyle?: SheetCellStylePatch | null
  authoredStyleReference?: string | null
  clearedStyleProperties?: SheetCellStyleProperty[] | null
  themeStyle?: SheetStyleSource
  formula?: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  align?: CellStyle['horizontal']
  dataType?: string
  cachedDataType?: string | null
  vertical?: 'top' | 'middle' | 'bottom'
  background?: string | null
  color?: string | null
  fontSize?: number
  numberFormat?: SheetNumberFormat
  cachedValue?: string
  fontFamily?: string
  rawValue?: CellScalar
  displayValue?: string
  rawNumberFormat?: string
  wrap?: boolean
  wrapStrategy?: 'overflow' | 'wrap' | 'clip'
  border?: Record<string, unknown>
  rotation?: number
  rotationVertical?: boolean
  indent?: number
  shrinkToFit?: boolean
  hyperlink?: string | null
  sparkline?: SparklineSpec
  image?: SheetCellImage | null
  richText?: ReadonlyArray<Record<string, unknown>> | null
  /** Null explicitly clears preserved style metadata when formatting is reset. */
  styleMetadata?: Record<string, unknown> | null
  /** Preserved cell extensions also follow content when a merge moves its first non-empty cell. */
  metadata?: Record<string, unknown> | null
}
export interface ExcelSheetView {
  id: string
  name: string
  tabColor?: string
  cells: Record<string, ExcelCellView>
  images?: SheetFloatingImage[]
  drawingState?: SheetDrawingState
  shapes?: SheetDrawingShapeDTO[]
  backgroundImage?: SheetBackgroundImage | null
  defaultStyle?: SheetStyleSource
  rowStyles?: Record<number, SheetStyleSource>
  columnStyles?: Record<number, SheetStyleSource>
  rowStyleSpans?: SheetAxisSpan<SheetStyleSource>[]
  columnStyleSpans?: SheetAxisSpan<SheetStyleSource>[]
  rowStylePriority?: boolean
  rangeThemeRules?: SheetRangeThemeRule[]
  rangeThemes?: Record<string, SheetRangeThemeStyle>
  /** Workbook sources are carried with the detached view, including sparse blank-cell presentation. */
  workbookDefaultStyle?: SheetStyleSource
  styles?: Record<string, SheetCellStylePatch>
  rowCount: number
  columnCount: number
  merges?: string[]
  freeze?: { rows: number; columns: number }
  columnWidths?: Record<number, number>
  rowHeights?: Record<number, number>
  axisUnit?: 'cssPixel'
  rowHeightSpans?: SheetAxisSpan<number>[]
  columnWidthSpans?: SheetAxisSpan<number>[]
  rowAutoHeightSpans?: SheetAxisSpan<boolean>[]
  rowAutoHeights?: Record<number, number>
  hiddenRows?: number[]
  rowHiddenSpans?: SheetAxisSpan<boolean>[]
  columnHiddenSpans?: SheetAxisSpan<boolean>[]
  outline?: SheetOutlineState
  /** Derived collapsed rows. The domain outline is the persisted source of truth. */
  outlineHiddenRows?: number[]
  /** Derived collapsed columns. The domain outline is the persisted source of truth. */
  outlineHiddenColumns?: number[]
  /** Compact derived visibility; production views do not enumerate each collapsed index. */
  outlineHiddenRowSpans?: readonly SheetAxisSpan<boolean>[]
  outlineHiddenColumnSpans?: readonly SheetAxisSpan<boolean>[]
  /** Derived view visibility. This must never be written back as manually hidden rows. */
  filteredRows?: number[]
  autoFilter?: SheetAutoFilter
  hiddenColumns?: number[]
  showGridlines?: boolean
  visible?: boolean
  defaultColumnWidth?: number
  defaultRowHeight?: number
  defaultRowAutoHeight?: boolean
  date1904?: boolean
  dataValidations?: SheetDataValidationRule[]
  rangeProtectionRules?: SheetRangeProtectionRule[]
  worksheetProtection?: SheetWorksheetProtectionRule
  conditionalFormats?: SheetConditionalFormatRule[]
}
export interface ExcelWorkbookView { id?: string; name: string; sheets: ExcelSheetView[]; activeSheetId?: string; charts?: ChartModel[]; pivotTables?: PivotTableModel[]; pivotCaches?: PivotCacheModel[]; slicers?: SlicerModel[]; slicerCaches?: SlicerCacheModel[]; slicerStyles?: ISlicerStylesDefinition; timelineStyles?: ITimelineStylesDefinition; tables?: SheetTableModel[]; date1904?: boolean; calculationMode?: 'automatic' | 'manual'; metadata?: Record<string, unknown>; defaultStyle?: SheetStyleSource; styles?: Record<string, SheetCellStylePatch>; rangeThemes?: Record<string, SheetRangeThemeStyle> }
export const cellKey = (row: number, col: number): string => {
  let result = ''
  for (let index = col; index >= 0; index = Math.floor(index/26)-1) result = String.fromCharCode(index%26+65)+result
  return result+(row+1)
}
export const parseCellKey = (key: string): { row: number; col: number } | undefined => {
  const match = /^([A-Z]+)(\d+)$/.exec(key.toUpperCase())
  if (!match) return
  let col = 0
  for (const char of match[1]!) col = col*26+char.charCodeAt(0)-64
  return { row: Number(match[2])-1, col: col-1 }
}
export const toExcelCellView = (cell: SpreadsheetCell): ExcelCellView => ({
  authoredStyle: cell.authoredStyle === undefined ? undefined : cloneModel(cell.authoredStyle),
  authoredStyleReference: cell.authoredStyleReference,
  clearedStyleProperties: cell.clearedStyleProperties ? [...cell.clearedStyleProperties] : undefined,
  themeStyle: cell.themeStyle === undefined ? undefined : cloneModel(cell.themeStyle),
  value: cell.formula ? '' : cell.value === null ? '' : String(cell.value), formula: cell.formula ?? undefined,
  bold: cell.style.font.bold, italic: cell.style.font.italic, underline: cell.style.font.underline, strike: cell.style.font.strike,
  align: cell.style.horizontal,
  dataType: cell.dataType, cachedDataType: cell.cachedDataType,
  vertical: cell.style.vertical, background: cell.style.fill, color: cell.style.font.color,
  fontSize: cell.style.font.size, fontFamily: cell.style.font.family,
  numberFormat: classifySheetNumberFormat(cell.style.numberFormat) === 'custom' ? undefined : classifySheetNumberFormat(cell.style.numberFormat) as SheetNumberFormat,
  cachedValue: cell.formula ? cell.display : undefined, rawValue: cell.value, displayValue: cell.display,
  rawNumberFormat: cell.style.numberFormat, wrap: cell.style.wrap, wrapStrategy: cell.style.wrapStrategy, border: cell.style.border,
  rotation: cell.style.rotation === 255 ? 0 : cell.style.rotation,
  rotationVertical: cell.style.rotationVertical === true || cell.style.rotation === 255,
  indent: cell.style.indent, shrinkToFit: cell.style.shrinkToFit,
  hyperlink: cell.hyperlink ?? undefined,
  sparkline: cell.sparkline ? { ...cell.sparkline } : undefined,
  image: cell.image ? cloneModel(cell.image) : undefined,
  richText: cell.richText ? structuredClone(cell.richText) : undefined,
  styleMetadata: cell.style.metadata ? structuredClone(cell.style.metadata) : undefined,
  metadata: cell.metadata ? structuredClone(cell.metadata) : undefined,
})
export function toExcelWorkbookView(workbook: WorkbookModel): ExcelWorkbookView {
  return {
    id: workbook.id, name: workbook.name, activeSheetId: workbook.activeSheetId, date1904: workbook.date1904, calculationMode: workbook.calculationMode ?? 'automatic',
    defaultStyle: workbook.defaultStyle === undefined ? undefined : cloneModel(workbook.defaultStyle),
    styles: workbook.styles ? cloneModel(workbook.styles) : undefined,
    rangeThemes: workbook.rangeThemes ? cloneModel(workbook.rangeThemes) : undefined,
    metadata: workbook.metadata ? structuredClone(workbook.metadata) : undefined,
    charts: workbook.charts?.map(cloneChartModel),
    pivotTables: workbook.pivotTables?.map(clonePivotTableModel),
    pivotCaches: workbook.pivotCaches?.map(cache => cloneModel(cache)),
    slicers: workbook.slicers?.map(slicer => cloneModel(slicer)),
    slicerCaches: workbook.slicerCaches?.map(cache => cloneModel(cache)),
    slicerStyles: workbook.slicerStyles === undefined ? undefined : cloneModel(workbook.slicerStyles),
    timelineStyles: workbook.timelineStyles === undefined ? undefined : cloneModel(workbook.timelineStyles),
    tables: workbook.tables?.map(cloneSheetTable),
    sheets: workbook.sheets.map((sheet) => {
      const cells: Record<string, ExcelCellView> = {}
      let rowCount = Math.max(1, sheet.rowCount ?? 100)
      let columnCount = Math.max(1, sheet.columnCount ?? 26)
      for (const cell of Object.values(sheet.cells)) {
        cells[cellKey(cell.row, cell.col)] = toExcelCellView({ ...cell, style: resolveSheetCellStyle(sheet, cell.row, cell.col, cell, workbook) })
        rowCount = Math.max(rowCount, cell.row+1)
        columnCount = Math.max(columnCount, cell.col+1)
      }
      return {
        id: sheet.id, name: sheet.name, tabColor: sheet.tabColor, cells, rowCount, columnCount,
        images: sheet.images ? cloneModel(sheet.images) : undefined,
        drawingState: sheet.drawingState ? cloneModel(sheet.drawingState) : undefined,
        shapes: sheet.shapes ? cloneModel(sheet.shapes) : undefined,
        backgroundImage: sheet.backgroundImage ? cloneModel(sheet.backgroundImage) : undefined,
        defaultStyle: sheet.defaultStyle === undefined ? undefined : cloneModel(sheet.defaultStyle),
        rowStyles: sheet.rowStyles ? cloneModel(sheet.rowStyles) : undefined,
        columnStyles: sheet.columnStyles ? cloneModel(sheet.columnStyles) : undefined,
        rowStyleSpans: sheet.rowStyleSpans ? cloneModel(sheet.rowStyleSpans) : undefined,
        columnStyleSpans: sheet.columnStyleSpans ? cloneModel(sheet.columnStyleSpans) : undefined,
        rowStylePriority: sheet.rowStylePriority,
        rangeThemeRules: sheet.rangeThemeRules ? cloneModel(sheet.rangeThemeRules) : undefined,
        rangeThemes: workbook.rangeThemes ? cloneModel(workbook.rangeThemes) : undefined,
        workbookDefaultStyle: workbook.defaultStyle === undefined ? undefined : cloneModel(workbook.defaultStyle),
        styles: workbook.styles ? cloneModel(workbook.styles) : undefined,
        merges: sheet.mergedRanges.map((r) => cellKey(r.startRow,r.startCol)+':'+cellKey(r.endRow,r.endCol)),
        freeze: { rows: sheet.frozenRows, columns: sheet.frozenCols },
        columnWidths: { ...sheet.columnWidths }, rowHeights: { ...sheet.rowHeights },
        axisUnit: sheet.axisUnit ?? 'cssPixel',
        rowHeightSpans: sheet.rowHeightSpans?.map(span => ({ ...span })), columnWidthSpans: sheet.columnWidthSpans?.map(span => ({ ...span })),
        rowAutoHeightSpans: sheet.rowAutoHeightSpans?.map(span => ({ ...span })), rowAutoHeights: sheet.rowAutoHeights ? { ...sheet.rowAutoHeights } : undefined,
        hiddenRows: [...(sheet.hiddenRows ?? [])], hiddenColumns: [...(sheet.hiddenColumns ?? [])],
        rowHiddenSpans: sheet.rowHiddenSpans?.map(span => ({ ...span })), columnHiddenSpans: sheet.columnHiddenSpans?.map(span => ({ ...span })),
        outline: sheet.outline ? { ...sheet.outline, groups: sheet.outline.groups.map(group => ({ ...group })) } : undefined,
        outlineHiddenRowSpans: sheet.outline ? getSheetOutlineHiddenSpans(sheet.outline, 'row', rowCount) : [],
        outlineHiddenColumnSpans: sheet.outline ? getSheetOutlineHiddenSpans(sheet.outline, 'column', columnCount) : [],
        filteredRows: [...new Set([...(sheet.autoFilter?.filteredRows ?? []), ...(workbook.tables ?? []).filter(table => table.sheetId === sheet.id).flatMap(table => {
          if (!table.autoFilter) return []
          return refreshAutoFilter({ ...sheet, autoFilter: structuredClone(table.autoFilter) })
        })])],
        autoFilter: sheet.autoFilter ? structuredClone(sheet.autoFilter) : undefined,
        showGridlines: sheet.showGridlines, visible: sheet.visible,
        defaultColumnWidth: sheet.defaultColumnWidth, defaultRowHeight: sheet.defaultRowHeight,
        defaultRowAutoHeight: sheet.defaultRowAutoHeight,
        date1904: workbook.date1904,
        dataValidations: sheet.dataValidations?.map(rule => ({ ...rule, ranges: rule.ranges.map(range => ({ ...range })) })),
        rangeProtectionRules: sheet.rangeProtectionRules?.map(rule => ({ ...rule, ranges: rule.ranges.map(range => ({ ...range })), allowedUserIds: [...rule.allowedUserIds] })),
        worksheetProtection: sheet.worksheetProtection ? { ...sheet.worksheetProtection, allowedUserIds: [...sheet.worksheetProtection.allowedUserIds] } : undefined,
        conditionalFormats: sheet.conditionalFormats?.map(cloneConditionalFormatRule),
      }
    }),
  }
}
function rawValue(cell: ExcelCellView, original?: SpreadsheetCell): CellScalar {
  if (cell.formula) {
    if (cell.rawValue !== undefined) return cell.rawValue
    if (original?.formula === cell.formula) return original.value
    if (cell.cachedValue === undefined || cell.cachedValue === '') return null
    const type = cell.cachedDataType ?? (normalizeSheetCellValueType(cell.dataType) === 'formula' ? undefined : cell.dataType)
    return parseSheetCellText(cell.cachedValue, type)
  }
  if (original && cell.rawValue !== undefined && Object.is(cell.rawValue, original.value)
    && cell.value !== String(original.value ?? '') && cell.value !== original.display) {
    return parseSheetCellText(cell.value, cell.dataType)
  }
  if (cell.rawValue !== undefined) return cell.rawValue
  if (original && String(original.value ?? '') === cell.value) return original.value
  return parseSheetCellText(cell.value, cell.dataType ?? (isSheetTextNumberFormat(cell.rawNumberFormat) || cell.numberFormat === 'text' ? 'text' : undefined))
}
export interface ExcelCellViewMappingOptions {
  date1904?: boolean
  previousDate1904?: boolean
  numberFormatLocale?: string
}

/** The same cell mapping serves a single input command and a complete workbook view update. */
export function applyExcelCellView(position: { row: number; col: number }, cell: ExcelCellView, previous?: SpreadsheetCell,
  options: ExcelCellViewMappingOptions = {}): SpreadsheetCell {
  const { date1904, previousDate1904, numberFormatLocale } = options
  const id = position.row + ':' + position.col
  const numberFormat = cell.rawNumberFormat ?? (cell.numberFormat ? SHEET_NUMBER_FORMATS[cell.numberFormat] : previous?.style.numberFormat) ?? 'General'
  const value = rawValue(cell, previous)
  const formula = cell.formula ?? null
  const sameContent = previous !== undefined && previous.formula === formula && Object.is(previous.value, value)
  const dataType = sameContent ? cell.dataType ?? previous?.dataType : formula ? 'formula'
    : normalizeSheetCellValueType(cell.dataType) === 'formula' ? inferSheetCellValueType(value) : cell.dataType ?? inferSheetCellValueType(value)
  const cachedDataType = sameContent ? cell.cachedDataType === undefined ? previous?.cachedDataType : cell.cachedDataType
    : formula ? cell.cachedDataType === undefined ? value === null ? null : inferSheetCellValueType(value) : cell.cachedDataType : null
  const samePresentation = sameContent && previous?.style.numberFormat === numberFormat && date1904 === previousDate1904
  const preservedDisplay = samePresentation ? cell.displayValue ?? (formula ? cell.cachedValue : undefined) ?? previous?.display : undefined
  const display = preservedDisplay ?? (formula && value === null ? cell.displayValue ?? cell.cachedValue ?? (sameContent ? previous?.display : undefined) ?? formula
    : formatSheetValue(value, numberFormat, date1904, numberFormatLocale)
      ?? cell.displayValue ?? (formula ? cell.cachedValue ?? formula : sameContent ? previous?.display ?? cell.value : cell.value))
  const metadata = cell.metadata === null ? undefined : cell.metadata ? structuredClone(cell.metadata) : previous?.metadata ? { ...previous.metadata } : undefined
  if (metadata && (!sameContent || cell.richText === null)) delete metadata.formulaError
  if (metadata && (!sameContent || cell.image === null)) { delete metadata.publicBody; delete metadata.sheetImagePublicDrawingId }
  const rotationVertical = cell.rotationVertical ?? (cell.rotation === 255 ? true : previous?.style.rotationVertical)
  const rotation = cell.rotation === 255 ? 0 : cell.rotation ?? previous?.style.rotation
  const result: SpreadsheetCell = {
    ...previous, id, ...position, value, formula, dataType, cachedDataType, metadata,
    authoredStyle: cell.authoredStyle === null ? {} : cell.authoredStyle !== undefined ? cloneModel(cell.authoredStyle)
      : previous?.authoredStyle !== undefined ? cloneModel(previous.authoredStyle) : previous ? undefined
        : cell.authoredStyleReference ? {} : sheetViewCellStylePatch(cell),
    authoredStyleReference: cell.authoredStyleReference === null ? undefined : cell.authoredStyleReference ?? previous?.authoredStyleReference,
    clearedStyleProperties: cell.clearedStyleProperties === null ? undefined : cell.clearedStyleProperties === undefined
      ? previous?.clearedStyleProperties : [...cell.clearedStyleProperties],
    themeStyle: cell.themeStyle === undefined ? previous?.themeStyle : cell.themeStyle,
    hyperlink: cell.hyperlink === undefined ? previous?.hyperlink ?? null : cell.hyperlink,
    sparkline: cell.sparkline ? { ...cell.sparkline } : undefined,
    image: cell.image === null ? undefined : cell.image ? cloneModel(cell.image) : sameContent ? previous?.image : undefined,
    richText: cell.richText === null ? undefined : cell.richText ? structuredClone(cell.richText) : sameContent ? previous?.richText : undefined,
    display,
    style: {
      ...previous?.style,
      font: {
        family: cell.fontFamily ?? previous?.style.font.family ?? 'Aptos', size: cell.fontSize ?? previous?.style.font.size ?? 11,
        bold: cell.bold ?? previous?.style.font.bold ?? false, italic: cell.italic ?? previous?.style.font.italic ?? false,
        underline: cell.underline ?? previous?.style.font.underline ?? false, strike: cell.strike ?? previous?.style.font.strike ?? false,
        color: cell.color === undefined ? previous?.style.font.color ?? null : cell.color,
      },
      fill: cell.background === undefined ? previous?.style.fill ?? null : cell.background,
      horizontal: cell.align ?? previous?.style.horizontal ?? 'general', vertical: cell.vertical ?? previous?.style.vertical ?? 'middle',
      wrap: cell.wrap ?? previous?.style.wrap ?? false, wrapStrategy: cell.wrapStrategy ?? previous?.style.wrapStrategy, border: cell.border ?? previous?.style.border,
      rotation, rotationVertical, indent: cell.indent ?? previous?.style.indent,
      shrinkToFit: cell.shrinkToFit ?? previous?.style.shrinkToFit, numberFormat,
      metadata: cell.styleMetadata === null ? undefined : cell.styleMetadata ?? previous?.style.metadata,
    },
  }
  // A new view cell with explicit provenance can contain inherited presentation. Its direct cache starts sparse.
  if (!previous && cell.authoredStyle !== undefined) result.style = composeSheetCellStyles(result.authoredStyle)
  return result
}
export function applyExcelWorkbookView(workbook: WorkbookModel, view: ExcelWorkbookView, numberFormatLocale?: string): WorkbookModel {
  const date1904 = view.date1904 ?? workbook.date1904
  const result: WorkbookModel = {
    ...workbook, id: view.id ?? workbook.id, name: view.name, activeSheetId: view.activeSheetId ?? workbook.activeSheetId, date1904,
    metadata: view.metadata ? structuredClone(view.metadata) : workbook.metadata,
    defaultStyle: view.defaultStyle === undefined ? workbook.defaultStyle : cloneModel(view.defaultStyle),
    styles: view.styles ? cloneModel(view.styles) : workbook.styles,
    rangeThemes: view.rangeThemes === undefined ? workbook.rangeThemes : cloneModel(view.rangeThemes),
    charts: view.charts?.map(cloneChartModel) ?? workbook.charts,
    pivotTables: view.pivotTables?.map(clonePivotTableModel) ?? workbook.pivotTables,
    pivotCaches: view.pivotCaches?.map(cache => cloneModel(cache)) ?? workbook.pivotCaches,
    slicers: view.slicers?.map(slicer => cloneModel(slicer)) ?? workbook.slicers,
    slicerCaches: view.slicerCaches?.map(cache => cloneModel(cache)) ?? workbook.slicerCaches,
    slicerStyles: Object.hasOwn(view, 'slicerStyles') ? view.slicerStyles === undefined ? undefined : cloneModel(view.slicerStyles) : workbook.slicerStyles,
    timelineStyles: Object.hasOwn(view, 'timelineStyles') ? view.timelineStyles === undefined ? undefined : cloneModel(view.timelineStyles) : workbook.timelineStyles,
    tables: view.tables?.map(cloneSheetTable) ?? workbook.tables,
    calculationMode: view.calculationMode === 'manual' ? 'manual' : 'automatic',
    sheets: view.sheets.map((source) => {
      const original = workbook.sheets.find((sheet) => sheet.id === source.id)
      const cells: Record<string, SpreadsheetCell> = {}
      for (const [key, cell] of Object.entries(source.cells)) {
        const position = parseCellKey(key)
        if (!position) continue
        const id = position.row+':'+position.col
        const previous = original?.cells[id]
        cells[id] = applyExcelCellView(position, cell, previous, { date1904, previousDate1904: workbook.date1904, numberFormatLocale })
      }
      const next: WorkbookModel['sheets'][number] = {
        ...original, id: source.id, name: source.name, tabColor: source.tabColor ?? original?.tabColor, cells, visible: source.visible ?? original?.visible ?? true,
        images: source.images === undefined ? original?.images : cloneModel(source.images),
        drawingState: source.drawingState === undefined ? original?.drawingState : cloneModel(source.drawingState),
        shapes: source.shapes === undefined ? original?.shapes : cloneModel(source.shapes),
        backgroundImage: source.backgroundImage === null ? undefined : source.backgroundImage === undefined ? original?.backgroundImage : cloneModel(source.backgroundImage),
        defaultStyle: source.defaultStyle === undefined ? original?.defaultStyle : cloneModel(source.defaultStyle),
        rowStyles: source.rowStyles === undefined ? original?.rowStyles : cloneModel(source.rowStyles),
        columnStyles: source.columnStyles === undefined ? original?.columnStyles : cloneModel(source.columnStyles),
        rowStyleSpans: source.rowStyleSpans === undefined ? original?.rowStyleSpans : cloneModel(source.rowStyleSpans),
        columnStyleSpans: source.columnStyleSpans === undefined ? original?.columnStyleSpans : cloneModel(source.columnStyleSpans),
        rowStylePriority: source.rowStylePriority ?? original?.rowStylePriority,
        rangeThemeRules: source.rangeThemeRules === undefined ? original?.rangeThemeRules : cloneModel(source.rangeThemeRules),
        dataValidations: source.dataValidations?.map(rule => ({ ...rule, ranges: rule.ranges.map(range => ({ ...range })) })) ?? original?.dataValidations,
        conditionalFormats: (source.conditionalFormats ?? original?.conditionalFormats)?.map(cloneConditionalFormatRule),
        rangeProtectionRules: source.rangeProtectionRules?.map(rule => ({ ...rule, ranges: rule.ranges.map(range => ({ ...range })), allowedUserIds: [...rule.allowedUserIds] })) ?? original?.rangeProtectionRules,
        worksheetProtection: source.worksheetProtection ? { ...source.worksheetProtection, allowedUserIds: [...source.worksheetProtection.allowedUserIds] } : original?.worksheetProtection,
        rowCount: source.rowCount, columnCount: source.columnCount,
        columnWidths: { ...(source.columnWidths ?? original?.columnWidths) }, rowHeights: { ...(source.rowHeights ?? original?.rowHeights) },
        axisUnit: 'cssPixel',
        rowHeightSpans: (source.rowHeightSpans ?? original?.rowHeightSpans)?.map(span => ({ ...span })),
        columnWidthSpans: (source.columnWidthSpans ?? original?.columnWidthSpans)?.map(span => ({ ...span })),
        rowAutoHeightSpans: (source.rowAutoHeightSpans ?? original?.rowAutoHeightSpans)?.map(span => ({ ...span })),
        rowAutoHeights: source.rowAutoHeights ? { ...source.rowAutoHeights } : original?.rowAutoHeights ? { ...original.rowAutoHeights } : undefined,
        hiddenRows: [...(source.hiddenRows ?? original?.hiddenRows ?? [])], hiddenColumns: [...(source.hiddenColumns ?? original?.hiddenColumns ?? [])],
        rowHiddenSpans: (source.rowHiddenSpans ?? original?.rowHiddenSpans)?.map(span => ({ ...span })),
        columnHiddenSpans: (source.columnHiddenSpans ?? original?.columnHiddenSpans)?.map(span => ({ ...span })),
        outline: source.outline ? { ...source.outline, groups: source.outline.groups.map(group => ({ ...group })) } : original?.outline ? { ...original.outline, groups: original.outline.groups.map(group => ({ ...group })) } : undefined,
        showGridlines: source.showGridlines ?? original?.showGridlines ?? true,
        defaultColumnWidth: source.defaultColumnWidth ?? original?.defaultColumnWidth,
        defaultRowHeight: source.defaultRowHeight ?? original?.defaultRowHeight,
        defaultRowAutoHeight: source.defaultRowAutoHeight ?? original?.defaultRowAutoHeight,
        frozenRows: source.freeze?.rows ?? original?.frozenRows ?? 0, frozenCols: source.freeze?.columns ?? original?.frozenCols ?? 0,
        mergedRanges: source.merges?.map((merge) => {
          const [start,end] = merge.split(':')
          const first = parseCellKey(start ?? '') ?? {row:0,col:0}
          const last = parseCellKey(end ?? start ?? '') ?? first
          return { startRow:first.row,startCol:first.col,endRow:last.row,endCol:last.col }
        }) ?? [],
        autoFilter: source.autoFilter ? structuredClone(source.autoFilter)
          : Object.hasOwn(source, 'autoFilter') ? undefined
            : original?.autoFilter ? structuredClone(original.autoFilter) : undefined,
      }
      if (next.autoFilter) {
        next.autoFilter = { ...next.autoFilter, columns: { ...next.autoFilter.columns }, filteredRows: [...next.autoFilter.filteredRows] }
        refreshAutoFilter(next)
      }
      return next
    }),
  }
  refreshWorkbookCellStyles(result, workbook, true, numberFormatLocale)
  return result
}

export function sheetViewCellStylePatch(cell: ExcelCellView | undefined): SheetCellStylePatch {
  if (!cell) return {}
  const patch: SheetCellStylePatch = {}, font: SheetCellStylePatch['font'] = {}
  const fontFields = { fontFamily: 'family', fontSize: 'size', bold: 'bold', italic: 'italic', underline: 'underline', strike: 'strike', color: 'color' } as const
  for (const [field, key] of Object.entries(fontFields)) {
    const value = cell[field as keyof typeof fontFields]
    if (value !== undefined) Object.assign(font, { [key]: value })
  }
  if (Object.keys(font).length) patch.font = font
  const fields = { background: 'fill', align: 'horizontal', vertical: 'vertical', wrap: 'wrap', wrapStrategy: 'wrapStrategy',
    border: 'border', rotation: 'rotation', rotationVertical: 'rotationVertical', indent: 'indent', shrinkToFit: 'shrinkToFit', styleMetadata: 'metadata' } as const
  for (const [field, key] of Object.entries(fields)) {
    const value = cell[field as keyof typeof fields]
    if (value !== undefined && !(key === 'metadata' && value === null)) Object.assign(patch, { [key]: value })
  }
  if (cell.rawNumberFormat !== undefined) patch.numberFormat = cell.rawNumberFormat
  else if (cell.numberFormat !== undefined) patch.numberFormat = SHEET_NUMBER_FORMATS[cell.numberFormat]
  return patch
}

/** Read an effective cell without populating the document's sparse cell map. */
export function sheetViewCellAt(sheet: ExcelSheetView, row: number, col: number, context?: SheetStyleContext): ExcelCellView {
  const cell = sheet.cells[cellKey(row, col)]
  const inherited = sheetViewCellStylePatch(cell)
  const style = resolveSheetCellStyle(sheet, row, col, cell ? { style: composeSheetCellStyles(inherited),
    authoredStyle: cell.authoredStyle === null ? {} : cell.authoredStyle ?? (cell.authoredStyleReference ? {} : inherited),
    authoredStyleReference: cell.authoredStyleReference ?? undefined, clearedStyleProperties: cell.clearedStyleProperties ?? undefined,
    themeStyle: cell.themeStyle } : undefined,
    context ?? { defaultStyle: sheet.workbookDefaultStyle, styles: sheet.styles, rangeThemes: sheet.rangeThemes })
  const result = toExcelCellView({ id: `${row}:${col}`, row, col, value: cell?.rawValue ?? null, formula: cell?.formula ?? null,
    display: cell?.displayValue ?? cell?.value ?? '', style, authoredStyle: cell ? cell.authoredStyle ?? (cell.authoredStyleReference ? {} : inherited) : {}, themeStyle: cell?.themeStyle,
    authoredStyleReference: cell?.authoredStyleReference ?? undefined, clearedStyleProperties: cell?.clearedStyleProperties ?? undefined })
  return { ...result, ...cell, fontFamily: style.font.family, fontSize: style.font.size, bold: style.font.bold, italic: style.font.italic,
    underline: style.font.underline, strike: style.font.strike, color: style.font.color, background: style.fill, align: style.horizontal,
    vertical: style.vertical, rawNumberFormat: style.numberFormat, numberFormat: classifySheetNumberFormat(style.numberFormat) === 'custom' ? undefined
      : classifySheetNumberFormat(style.numberFormat) as SheetNumberFormat, wrap: style.wrap, wrapStrategy: style.wrapStrategy, border: style.border,
    rotation: style.rotation, rotationVertical: style.rotationVertical, indent: style.indent, shrinkToFit: style.shrinkToFit,
    styleMetadata: style.metadata, authoredStyle: cell ? cell.authoredStyle ?? (cell.authoredStyleReference ? {} : inherited) : {}, themeStyle: cell?.themeStyle }
}

export function sheetModelCellViewAt(workbook: WorkbookModel, sheet: SpreadsheetSheet, row: number, col: number): ExcelCellView {
  const cell = sheet.cells[`${row}:${col}`]
  return toExcelCellView({ ...(cell ?? { id: `${row}:${col}`, row, col, value: null, formula: null, display: '', authoredStyle: {} }),
    style: resolveSheetCellStyle(sheet, row, col, cell, workbook) })
}

/** Record explicit view formatting before exposing the composed presentation. */
export function setSheetViewCellStyleProperties(sheet: ExcelSheetView, row: number, col: number, patch: SheetCellStylePatch): void {
  const key = cellKey(row, col), current = sheet.cells[key]
  const authored = mergeSheetCellStylePatch(current?.authoredStyle ?? (current?.authoredStyleReference ? {} : sheetViewCellStylePatch(current)), patch)
  sheet.cells[key] = { ...(current ?? { value: '', rawValue: null }), authoredStyle: authored,
    clearedStyleProperties: sheetStyleResetsAfterPatch(current?.clearedStyleProperties ?? undefined, patch) ?? null }
  sheet.cells[key] = sheetViewCellAt(sheet, row, col)
}

function cloneSheetTable(table: SheetTableModel): SheetTableModel {
  return { ...table, range: { ...table.range }, columns: table.columns.map(column => ({ ...column })), autoFilter: table.autoFilter === null ? null : table.autoFilter ? structuredClone(table.autoFilter) : undefined }
}

function cloneChartModel(chart: ChartModel): ChartModel {
  return cloneModel(chart)
}

function clonePivotTableModel(pivot: PivotTableModel): PivotTableModel {
  return {
    ...pivot,
    source: { ...pivot.source, range: { ...pivot.source.range } },
    anchor: { ...pivot.anchor },
    fields: pivot.fields.map((field) => ({
      ...field,
      selectedItems: field.selectedItems ? [...field.selectedItems] : undefined,
      valueFilter: field.valueFilter ? { ...field.valueFilter } : undefined,
      numberGroup: field.numberGroup ? { ...field.numberGroup } : undefined,
      manualGroups: field.manualGroups?.map((group) => ({ ...group, members: [...group.members] })),
    })),
    options: { ...pivot.options },
    outputRange: pivot.outputRange ? { ...pivot.outputRange } : undefined,
    generatedCellKeys: pivot.generatedCellKeys ? [...pivot.generatedCellKeys] : undefined,
    drillDownMap: pivot.drillDownMap ? Object.fromEntries(Object.entries(pivot.drillDownMap).map(([key, entry]) => [key, {
      ...entry,
      fieldFilters: { ...entry.fieldFilters },
      fieldMembers: entry.fieldMembers ? Object.fromEntries(Object.entries(entry.fieldMembers).map(([fieldId, members]) => [fieldId, [...members]])) : undefined,
    }])) : undefined,
  }
}

function cloneConditionalFormatRule(rule: SheetConditionalFormatRule): SheetConditionalFormatRule {
  return cloneModel(rule)
}
