import type { CellScalar, ChartModel, PivotTableModel, WorkbookModel, SpreadsheetCell, SheetDataValidationRule, SheetRangeProtectionRule, SheetOutlineState, SparklineSpec } from '../domain/model.ts'
import { classifySheetNumberFormat, formatSheetValue, SHEET_NUMBER_FORMATS, type SheetNumberFormat } from '../features/sheet-workbench/sheet-number-format.ts'
import { refreshAutoFilter } from '../features/sheet-sort-filter/sort-filter-engine.ts'
import { getCollapsedOutlineIndexes } from '../features/sheet-outline/sheet-outline.ts'

export interface ExcelCellView {
  value: string
  formula?: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  align?: 'left' | 'center' | 'right'
  vertical?: 'top' | 'middle' | 'bottom'
  background?: string
  color?: string
  fontSize?: number
  numberFormat?: SheetNumberFormat
  cachedValue?: string
  fontFamily?: string
  rawValue?: CellScalar
  displayValue?: string
  rawNumberFormat?: string
  wrap?: boolean
  border?: Record<string, unknown>
  rotation?: number
  rotationVertical?: boolean
  indent?: number
  shrinkToFit?: boolean
  sparkline?: SparklineSpec
}
export interface ExcelSheetView {
  id: string
  name: string
  cells: Record<string, ExcelCellView>
  rowCount: number
  columnCount: number
  merges?: string[]
  freeze?: { rows: number; columns: number }
  columnWidths?: Record<number, number>
  rowHeights?: Record<number, number>
  hiddenRows?: number[]
  outline?: SheetOutlineState
  /** Derived collapsed rows. The domain outline is the persisted source of truth. */
  outlineHiddenRows?: number[]
  /** Derived collapsed columns. The domain outline is the persisted source of truth. */
  outlineHiddenColumns?: number[]
  /** Derived view visibility. This must never be written back as manually hidden rows. */
  filteredRows?: number[]
  hiddenColumns?: number[]
  showGridlines?: boolean
  visible?: boolean
  defaultColumnWidth?: number
  defaultRowHeight?: number
  date1904?: boolean
  dataValidations?: SheetDataValidationRule[]
  rangeProtectionRules?: SheetRangeProtectionRule[]
}
export interface ExcelWorkbookView { id?: string; name: string; sheets: ExcelSheetView[]; activeSheetId?: string; charts?: ChartModel[]; pivotTables?: PivotTableModel[]; date1904?: boolean; calculationMode?: 'automatic' | 'manual' }
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
const toViewCell = (cell: SpreadsheetCell): ExcelCellView => ({
  value: cell.formula ? '' : cell.value === null ? '' : String(cell.value), formula: cell.formula ?? undefined,
  bold: cell.style.font.bold, italic: cell.style.font.italic, underline: cell.style.font.underline, strike: cell.style.font.strike,
  align: cell.style.horizontal === 'justify' ? 'left' : cell.style.horizontal,
  vertical: cell.style.vertical, background: cell.style.fill, color: cell.style.font.color,
  fontSize: cell.style.font.size, fontFamily: cell.style.font.family,
  numberFormat: classifySheetNumberFormat(cell.style.numberFormat) === 'custom' ? undefined : classifySheetNumberFormat(cell.style.numberFormat) as SheetNumberFormat,
  cachedValue: cell.formula ? cell.display : undefined, rawValue: cell.value, displayValue: cell.display,
  rawNumberFormat: cell.style.numberFormat, wrap: cell.style.wrap, border: cell.style.border,
  rotation: cell.style.rotation === 255 ? 0 : cell.style.rotation,
  rotationVertical: cell.style.rotationVertical === true || cell.style.rotation === 255,
  indent: cell.style.indent, shrinkToFit: cell.style.shrinkToFit,
  sparkline: cell.sparkline ? { ...cell.sparkline } : undefined,
})
export function toExcelWorkbookView(workbook: WorkbookModel): ExcelWorkbookView {
  return {
    id: workbook.id, name: workbook.name, activeSheetId: workbook.activeSheetId, date1904: workbook.date1904, calculationMode: workbook.calculationMode ?? 'automatic',
    charts: workbook.charts?.map(cloneChartModel),
    pivotTables: workbook.pivotTables?.map((pivot) => ({ ...pivot, source: { ...pivot.source, range: { ...pivot.source.range } }, anchor: { ...pivot.anchor }, fields: pivot.fields.map((field) => ({ ...field })), options: { ...pivot.options } })),
    sheets: workbook.sheets.map((sheet) => {
      const cells: Record<string, ExcelCellView> = {}
      let rowCount = Math.max(100, sheet.rowCount ?? 0)
      let columnCount = Math.max(26, sheet.columnCount ?? 0)
      for (const cell of Object.values(sheet.cells)) {
        cells[cellKey(cell.row, cell.col)] = toViewCell(cell)
        rowCount = Math.max(rowCount, cell.row+1)
        columnCount = Math.max(columnCount, cell.col+1)
      }
      return {
        id: sheet.id, name: sheet.name, cells, rowCount, columnCount,
        merges: sheet.mergedRanges.map((r) => cellKey(r.startRow,r.startCol)+':'+cellKey(r.endRow,r.endCol)),
        freeze: { rows: sheet.frozenRows, columns: sheet.frozenCols },
        columnWidths: { ...sheet.columnWidths }, rowHeights: { ...sheet.rowHeights },
        hiddenRows: [...(sheet.hiddenRows ?? [])], hiddenColumns: [...(sheet.hiddenColumns ?? [])],
        outline: sheet.outline ? { ...sheet.outline, groups: sheet.outline.groups.map(group => ({ ...group })) } : undefined,
        outlineHiddenRows: sheet.outline ? [...getCollapsedOutlineIndexes({ count: rowCount, groups: sheet.outline.groups.filter(group => group.axis === 'row') })] : [],
        outlineHiddenColumns: sheet.outline ? [...getCollapsedOutlineIndexes({ count: columnCount, groups: sheet.outline.groups.filter(group => group.axis === 'column') })] : [],
        filteredRows: [...(sheet.autoFilter?.filteredRows ?? [])],
        showGridlines: sheet.showGridlines, visible: sheet.visible,
        defaultColumnWidth: sheet.defaultColumnWidth, defaultRowHeight: sheet.defaultRowHeight,
        date1904: workbook.date1904,
        dataValidations: sheet.dataValidations?.map(rule => ({ ...rule, ranges: rule.ranges.map(range => ({ ...range })) })),
        rangeProtectionRules: sheet.rangeProtectionRules?.map(rule => ({ ...rule, ranges: rule.ranges.map(range => ({ ...range })), allowedUserIds: [...rule.allowedUserIds] })),
      }
    }),
  }
}
function rawValue(cell: ExcelCellView, original?: SpreadsheetCell): CellScalar {
  if (cell.formula) return cell.rawValue !== undefined ? cell.rawValue : original?.formula === cell.formula ? original.value : cell.cachedValue ?? null
  if (original && String(original.value ?? '') === cell.value) return original.value
  if (cell.value.trim() && /^[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?$/i.test(cell.value)) return Number(cell.value)
  return cell.value
}
export function applyExcelWorkbookView(workbook: WorkbookModel, view: ExcelWorkbookView): WorkbookModel {
  return {
    ...workbook, id: view.id ?? workbook.id, name: view.name, activeSheetId: view.activeSheetId ?? workbook.activeSheetId, date1904: view.date1904 ?? workbook.date1904,
    charts: view.charts?.map(cloneChartModel) ?? workbook.charts,
    calculationMode: view.calculationMode === 'manual' ? 'manual' : 'automatic',
    sheets: view.sheets.map((source) => {
      const original = workbook.sheets.find((sheet) => sheet.id === source.id)
      const cells: Record<string, SpreadsheetCell> = {}
      for (const [key, cell] of Object.entries(source.cells)) {
        const position = parseCellKey(key)
        if (!position) continue
        const id = position.row+':'+position.col
        const previous = original?.cells[id]
        const preserveDisplay = previous && previous.formula === (cell.formula ?? null) && String(previous.value ?? '') === cell.value
        const numberFormat = cell.rawNumberFormat ?? (cell.numberFormat ? SHEET_NUMBER_FORMATS[cell.numberFormat] : previous?.style.numberFormat) ?? 'General'
        const value = rawValue(cell, previous)
        const display = formatSheetValue(value, numberFormat, view.date1904 ?? workbook.date1904)
        const rotationVertical = cell.rotationVertical ?? (cell.rotation === 255 ? true : previous?.style.rotationVertical)
        const rotation = cell.rotation === 255 ? 0 : cell.rotation ?? previous?.style.rotation
        cells[id] = {
          ...previous, id, ...position, value, formula: cell.formula ?? null,
          sparkline: cell.sparkline ? { ...cell.sparkline } : undefined,
          display: display ?? (cell.formula ? cell.cachedValue ?? cell.formula : preserveDisplay ? previous.display : cell.value),
          style: {
            ...previous?.style,
            font: {
              family: cell.fontFamily ?? previous?.style.font.family ?? 'Aptos', size: cell.fontSize ?? previous?.style.font.size ?? 11,
              bold: cell.bold ?? previous?.style.font.bold ?? false, italic: cell.italic ?? previous?.style.font.italic ?? false,
              underline: cell.underline ?? previous?.style.font.underline ?? false, strike: cell.strike ?? previous?.style.font.strike ?? false,
              color: cell.color ?? previous?.style.font.color ?? '#1f2937',
            },
            fill: cell.background ?? previous?.style.fill ?? '#fff',
            horizontal: cell.align ?? previous?.style.horizontal ?? 'left', vertical: cell.vertical ?? previous?.style.vertical ?? 'middle',
            wrap: cell.wrap ?? previous?.style.wrap ?? false, border: cell.border ?? previous?.style.border,
            rotation, rotationVertical, indent: cell.indent ?? previous?.style.indent,
            shrinkToFit: cell.shrinkToFit ?? previous?.style.shrinkToFit, numberFormat,
          },
        }
      }
      const next: WorkbookModel['sheets'][number] = {
        ...original, id: source.id, name: source.name, cells, visible: source.visible ?? original?.visible ?? true,
        dataValidations: source.dataValidations?.map(rule => ({ ...rule, ranges: rule.ranges.map(range => ({ ...range })) })) ?? original?.dataValidations,
        rangeProtectionRules: source.rangeProtectionRules?.map(rule => ({ ...rule, ranges: rule.ranges.map(range => ({ ...range })), allowedUserIds: [...rule.allowedUserIds] })) ?? original?.rangeProtectionRules,
        rowCount: source.rowCount, columnCount: source.columnCount,
        columnWidths: { ...(source.columnWidths ?? original?.columnWidths) }, rowHeights: { ...(source.rowHeights ?? original?.rowHeights) },
        hiddenRows: [...(source.hiddenRows ?? original?.hiddenRows ?? [])], hiddenColumns: [...(source.hiddenColumns ?? original?.hiddenColumns ?? [])],
        outline: source.outline ? { ...source.outline, groups: source.outline.groups.map(group => ({ ...group })) } : original?.outline ? { ...original.outline, groups: original.outline.groups.map(group => ({ ...group })) } : undefined,
        showGridlines: source.showGridlines ?? original?.showGridlines ?? true,
        defaultColumnWidth: source.defaultColumnWidth ?? original?.defaultColumnWidth,
        defaultRowHeight: source.defaultRowHeight ?? original?.defaultRowHeight,
        frozenRows: source.freeze?.rows ?? original?.frozenRows ?? 0, frozenCols: source.freeze?.columns ?? original?.frozenCols ?? 0,
        mergedRanges: source.merges?.map((merge) => {
          const [start,end] = merge.split(':')
          const first = parseCellKey(start ?? '') ?? {row:0,col:0}
          const last = parseCellKey(end ?? start ?? '') ?? first
          return { startRow:first.row,startCol:first.col,endRow:last.row,endCol:last.col }
        }) ?? [],
      }
      if (next.autoFilter) {
        next.autoFilter = { ...next.autoFilter, columns: { ...next.autoFilter.columns }, filteredRows: [...next.autoFilter.filteredRows] }
        refreshAutoFilter(next)
      }
      return next
    }),
  }
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
  }
}
