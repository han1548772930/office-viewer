import type { CellScalar, CellStyle, ChartModel, PivotTableModel, SheetTableModel, WorkbookModel, SpreadsheetCell, SheetDataValidationRule, SheetRangeProtectionRule, SheetWorksheetProtectionRule, SheetOutlineState, SheetAutoFilter, SparklineSpec, SheetConditionalFormatRule, SlicerModel } from '../domain/model.ts'
import { classifySheetNumberFormat, formatSheetValue, isSheetTextNumberFormat, SHEET_NUMBER_FORMATS, type SheetNumberFormat } from '../features/sheet-workbench/sheet-number-format.ts'
import { refreshAutoFilter } from '../features/sheet-sort-filter/sort-filter-engine.ts'
import { getCollapsedOutlineIndexes } from '../features/sheet-outline/sheet-outline.ts'
import { inferSheetCellValueType, normalizeSheetCellValueType, parseSheetCellText } from '../domain/sheet-cell-value.ts'
import type { SheetAxisSpan } from '../domain/sheet-axis.ts'
import { cloneModel } from '../domain/serialization.ts'

export interface ExcelCellView {
  value: string
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
export interface ExcelWorkbookView { id?: string; name: string; sheets: ExcelSheetView[]; activeSheetId?: string; charts?: ChartModel[]; pivotTables?: PivotTableModel[]; slicers?: SlicerModel[]; tables?: SheetTableModel[]; date1904?: boolean; calculationMode?: 'automatic' | 'manual'; metadata?: Record<string, unknown> }
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
  richText: cell.richText ? structuredClone(cell.richText) : undefined,
  styleMetadata: cell.style.metadata ? structuredClone(cell.style.metadata) : undefined,
  metadata: cell.metadata ? structuredClone(cell.metadata) : undefined,
})
export function toExcelWorkbookView(workbook: WorkbookModel): ExcelWorkbookView {
  return {
    id: workbook.id, name: workbook.name, activeSheetId: workbook.activeSheetId, date1904: workbook.date1904, calculationMode: workbook.calculationMode ?? 'automatic',
    metadata: workbook.metadata ? structuredClone(workbook.metadata) : undefined,
    charts: workbook.charts?.map(cloneChartModel),
    pivotTables: workbook.pivotTables?.map(clonePivotTableModel),
    slicers: workbook.slicers?.map((slicer) => ({ ...slicer, position: { ...slicer.position }, size: { ...slicer.size }, layout: slicer.layout ? { ...slicer.layout } : undefined, selectedItems: slicer.selectedItems ? [...slicer.selectedItems] : undefined, timelineRange: slicer.timelineRange ? { ...slicer.timelineRange } : undefined })),
    tables: workbook.tables?.map(cloneSheetTable),
    sheets: workbook.sheets.map((sheet) => {
      const cells: Record<string, ExcelCellView> = {}
      let rowCount = Math.max(1, sheet.rowCount ?? 100)
      let columnCount = Math.max(1, sheet.columnCount ?? 26)
      for (const cell of Object.values(sheet.cells)) {
        cells[cellKey(cell.row, cell.col)] = toExcelCellView(cell)
        rowCount = Math.max(rowCount, cell.row+1)
        columnCount = Math.max(columnCount, cell.col+1)
      }
      return {
        id: sheet.id, name: sheet.name, tabColor: sheet.tabColor, cells, rowCount, columnCount,
        merges: sheet.mergedRanges.map((r) => cellKey(r.startRow,r.startCol)+':'+cellKey(r.endRow,r.endCol)),
        freeze: { rows: sheet.frozenRows, columns: sheet.frozenCols },
        columnWidths: { ...sheet.columnWidths }, rowHeights: { ...sheet.rowHeights },
        axisUnit: sheet.axisUnit ?? 'cssPixel',
        rowHeightSpans: sheet.rowHeightSpans?.map(span => ({ ...span })), columnWidthSpans: sheet.columnWidthSpans?.map(span => ({ ...span })),
        rowAutoHeightSpans: sheet.rowAutoHeightSpans?.map(span => ({ ...span })), rowAutoHeights: sheet.rowAutoHeights ? { ...sheet.rowAutoHeights } : undefined,
        hiddenRows: [...(sheet.hiddenRows ?? [])], hiddenColumns: [...(sheet.hiddenColumns ?? [])],
        rowHiddenSpans: sheet.rowHiddenSpans?.map(span => ({ ...span })), columnHiddenSpans: sheet.columnHiddenSpans?.map(span => ({ ...span })),
        outline: sheet.outline ? { ...sheet.outline, groups: sheet.outline.groups.map(group => ({ ...group })) } : undefined,
        outlineHiddenRows: sheet.outline ? [...getCollapsedOutlineIndexes({ count: rowCount, groups: sheet.outline.groups.filter(group => group.axis === 'row') })] : [],
        outlineHiddenColumns: sheet.outline ? [...getCollapsedOutlineIndexes({ count: columnCount, groups: sheet.outline.groups.filter(group => group.axis === 'column') })] : [],
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
export function applyExcelWorkbookView(workbook: WorkbookModel, view: ExcelWorkbookView, numberFormatLocale?: string): WorkbookModel {
  const date1904 = view.date1904 ?? workbook.date1904
  return {
    ...workbook, id: view.id ?? workbook.id, name: view.name, activeSheetId: view.activeSheetId ?? workbook.activeSheetId, date1904,
    metadata: view.metadata ? structuredClone(view.metadata) : workbook.metadata,
    charts: view.charts?.map(cloneChartModel) ?? workbook.charts,
    pivotTables: view.pivotTables?.map(clonePivotTableModel) ?? workbook.pivotTables,
    slicers: view.slicers?.map((slicer) => ({ ...slicer, position: { ...slicer.position }, size: { ...slicer.size }, layout: slicer.layout ? { ...slicer.layout } : undefined, selectedItems: slicer.selectedItems ? [...slicer.selectedItems] : undefined, timelineRange: slicer.timelineRange ? { ...slicer.timelineRange } : undefined })) ?? workbook.slicers,
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
        const numberFormat = cell.rawNumberFormat ?? (cell.numberFormat ? SHEET_NUMBER_FORMATS[cell.numberFormat] : previous?.style.numberFormat) ?? 'General'
        const value = rawValue(cell, previous)
        const formula = cell.formula ?? null
        const sameContent = previous !== undefined && previous.formula === formula && Object.is(previous.value, value)
        const dataType = sameContent ? cell.dataType ?? previous?.dataType : formula ? 'formula'
          : normalizeSheetCellValueType(cell.dataType) === 'formula' ? inferSheetCellValueType(value) : cell.dataType ?? inferSheetCellValueType(value)
        const cachedDataType = sameContent ? cell.cachedDataType === undefined ? previous?.cachedDataType : cell.cachedDataType
          : formula ? cell.cachedDataType === undefined ? value === null ? null : inferSheetCellValueType(value) : cell.cachedDataType : null
        const samePresentation = sameContent && previous?.style.numberFormat === numberFormat && date1904 === workbook.date1904
        const preservedDisplay = samePresentation ? cell.displayValue ?? (formula ? cell.cachedValue : undefined) ?? previous?.display : undefined
        const display = preservedDisplay ?? (formula && value === null ? cell.displayValue ?? cell.cachedValue ?? (sameContent ? previous?.display : undefined) ?? formula
          : formatSheetValue(value, numberFormat, date1904, numberFormatLocale)
            ?? cell.displayValue ?? (formula ? cell.cachedValue ?? formula : sameContent ? previous?.display ?? cell.value : cell.value))
        const metadata = cell.metadata === null ? undefined : cell.metadata ? structuredClone(cell.metadata) : previous?.metadata ? { ...previous.metadata } : undefined
        if (metadata && (!sameContent || cell.richText === null)) delete metadata.formulaError
        const rotationVertical = cell.rotationVertical ?? (cell.rotation === 255 ? true : previous?.style.rotationVertical)
        const rotation = cell.rotation === 255 ? 0 : cell.rotation ?? previous?.style.rotation
        cells[id] = {
          ...previous, id, ...position, value, formula, dataType, cachedDataType, metadata,
          hyperlink: cell.hyperlink === undefined ? previous?.hyperlink ?? null : cell.hyperlink,
          sparkline: cell.sparkline ? { ...cell.sparkline } : undefined,
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
      }
      const next: WorkbookModel['sheets'][number] = {
        ...original, id: source.id, name: source.name, tabColor: source.tabColor ?? original?.tabColor, cells, visible: source.visible ?? original?.visible ?? true,
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
}

function cloneSheetTable(table: SheetTableModel): SheetTableModel {
  return { ...table, range: { ...table.range }, columns: table.columns.map(column => ({ ...column })), autoFilter: table.autoFilter ? structuredClone(table.autoFilter) : undefined }
}

function cloneChartModel(chart: ChartModel): ChartModel {
  return {
    ...chart,
    source: { range: { ...chart.source.range }, orientation: chart.source.orientation },
    position: { ...chart.position },
    size: { ...chart.size },
    legend: { ...chart.legend },
    xAxis: { ...chart.xAxis },
    yAxis: { ...chart.yAxis },
    valueFields: [...chart.valueFields],
    seriesToCategory: chart.seriesToCategory ? Object.fromEntries(Object.entries(chart.seriesToCategory).map(([field, categories]) => [field, [...categories]])) : undefined,
  }
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
