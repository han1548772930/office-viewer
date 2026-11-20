import type { CellScalar, ChartModel, WorkbookModel, SpreadsheetCell } from '../domain/model.ts'

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
  numberFormat?: 'general' | 'number' | 'percent'
  cachedValue?: string
  fontFamily?: string
  rawValue?: CellScalar
  displayValue?: string
  rawNumberFormat?: string
  wrap?: boolean
  border?: Record<string, unknown>
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
  hiddenColumns?: number[]
  showGridlines?: boolean
  visible?: boolean
  defaultColumnWidth?: number
  defaultRowHeight?: number
}
export interface ExcelWorkbookView { id?: string; name: string; sheets: ExcelSheetView[]; activeSheetId?: string; charts?: ChartModel[] }
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
  numberFormat: cell.style.numberFormat === 'General' ? 'general' : cell.style.numberFormat.includes('%') ? 'percent' : 'number',
  cachedValue: cell.formula ? cell.display : undefined, rawValue: cell.value, displayValue: cell.display,
  rawNumberFormat: cell.style.numberFormat, wrap: cell.style.wrap, border: cell.style.border,
})
export function toExcelWorkbookView(workbook: WorkbookModel): ExcelWorkbookView {
  return {
    id: workbook.id, name: workbook.name, activeSheetId: workbook.activeSheetId,
    charts: workbook.charts?.map(cloneChartModel),
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
        showGridlines: sheet.showGridlines, visible: sheet.visible,
        defaultColumnWidth: sheet.defaultColumnWidth, defaultRowHeight: sheet.defaultRowHeight,
      }
    }),
  }
}
function rawValue(cell: ExcelCellView, original?: SpreadsheetCell): CellScalar {
  if (cell.formula) return original?.formula === cell.formula ? original.value : cell.cachedValue ?? null
  if (original && String(original.value ?? '') === cell.value) return original.value
  if (cell.value.trim() && /^[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?$/i.test(cell.value)) return Number(cell.value)
  return cell.value
}
export function applyExcelWorkbookView(workbook: WorkbookModel, view: ExcelWorkbookView): WorkbookModel {
  return {
    ...workbook, id: view.id ?? workbook.id, name: view.name, activeSheetId: view.activeSheetId ?? workbook.activeSheetId,
    charts: view.charts?.map(cloneChartModel) ?? workbook.charts,
    sheets: view.sheets.map((source) => {
      const original = workbook.sheets.find((sheet) => sheet.id === source.id)
      const cells: Record<string, SpreadsheetCell> = {}
      for (const [key, cell] of Object.entries(source.cells)) {
        const position = parseCellKey(key)
        if (!position) continue
        const id = position.row+':'+position.col
        const previous = original?.cells[id]
        const preserveDisplay = previous && previous.formula === (cell.formula ?? null) && String(previous.value ?? '') === cell.value
        cells[id] = {
          ...previous, id, ...position, value: rawValue(cell,previous), formula: cell.formula ?? null,
          display: cell.formula ? cell.cachedValue ?? cell.formula : preserveDisplay ? previous.display : cell.value,
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
            numberFormat: cell.numberFormat === 'percent' ? '0%' : cell.numberFormat === 'number'
              ? (previous?.style.numberFormat !== 'General' ? previous?.style.numberFormat ?? '#,##0.00' : '#,##0.00')
              : cell.rawNumberFormat ?? previous?.style.numberFormat ?? 'General',
          },
        }
      }
      return {
        ...original, id: source.id, name: source.name, cells, visible: source.visible ?? original?.visible ?? true,
        rowCount: source.rowCount, columnCount: source.columnCount,
        columnWidths: { ...(source.columnWidths ?? original?.columnWidths) }, rowHeights: { ...(source.rowHeights ?? original?.rowHeights) },
        hiddenRows: [...(source.hiddenRows ?? original?.hiddenRows ?? [])], hiddenColumns: [...(source.hiddenColumns ?? original?.hiddenColumns ?? [])],
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
