import type { CellRange, ChartModel, CellScalar, SpreadsheetCell, WorkbookModel } from '../domain/model'
import { createCell, defaultCellStyle } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { cloneChartOptions } from '../features/charts/chart-options'
import type { ExcelCellView, ExcelWorkbookView } from '../services/excelViewAdapter'
import { GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH, SpreadsheetGeometry, type GridRange, type SpreadsheetRenderSheet } from '../rendering/spreadsheet/layout'

export type ExcelWorkbookChartView = ExcelWorkbookView & { charts?: ChartModel[] }

/**
 * The editor view deliberately stores display-friendly A1 keys. The chart
 * engine uses the sparse domain model and row:column keys, so the conversion
 * lives at this UI boundary instead of leaking view concerns into chart code.
 */
export function toChartWorkbook(view: ExcelWorkbookChartView): WorkbookModel {
  return {
    kind: 'workbook',
    name: view.name,
    activeSheetId: view.activeSheetId ?? view.sheets[0]?.id ?? '',
    date1904: view.date1904 ?? false,
    modified: false,
    id: view.id ?? 'chart-overlay-workbook',
    charts: view.charts?.map((chart) => cloneChart(chart)),
    tables: view.tables?.map(table => cloneModel(table)),
    pivotTables: view.pivotTables?.map((pivot) => ({ ...pivot, source: { ...pivot.source, range: { ...pivot.source.range } }, anchor: { ...pivot.anchor }, fields: pivot.fields.map((field) => ({ ...field })), options: { ...pivot.options } })),
    sheets: view.sheets.map((source) => ({
      id: source.id,
      name: source.name,
      visible: source.visible ?? true,
      cells: Object.fromEntries(Object.entries(source.cells).flatMap(([address, value]) => {
        const position = parseA1(address)
        if (!position) return []
        return [[`${position.row}:${position.col}`, toDomainCell(position.row, position.col, value)]]
      })),
      mergedRanges: (source.merges ?? []).flatMap((merge) => {
        const [first, last] = merge.split(':')
        const start = parseA1(first ?? '')
        const end = parseA1(last ?? first ?? '') ?? start
        return start && end ? [{ startRow: start.row, startCol: start.col, endRow: end.row, endCol: end.col }] : []
      }),
      columnWidths: { ...(source.columnWidths ?? {}) },
      rowHeights: { ...(source.rowHeights ?? {}) },
      frozenRows: source.freeze?.rows ?? 0,
      frozenCols: source.freeze?.columns ?? 0,
      showGridlines: source.showGridlines ?? true,
      hiddenRows: [...(source.hiddenRows ?? [])],
      hiddenColumns: [...(source.hiddenColumns ?? [])],
      autoFilter: source.autoFilter ? cloneModel(source.autoFilter) : undefined,
      outline: source.outline ? cloneModel(source.outline) : undefined,
      rowHiddenSpans: source.rowHiddenSpans?.map(span => ({ ...span })), columnHiddenSpans: source.columnHiddenSpans?.map(span => ({ ...span })),
      rowHeightSpans: source.rowHeightSpans?.map(span => ({ ...span })), columnWidthSpans: source.columnWidthSpans?.map(span => ({ ...span })),
      rowAutoHeightSpans: source.rowAutoHeightSpans?.map(span => ({ ...span })), rowAutoHeights: { ...source.rowAutoHeights },
      defaultRowAutoHeight: source.defaultRowAutoHeight,
      defaultColumnWidth: source.defaultColumnWidth,
      defaultRowHeight: source.defaultRowHeight,
      rowCount: source.rowCount,
      columnCount: source.columnCount,
    })),
  }
}

export function cloneChart(chart: ChartModel): ChartModel {
  return {
    ...chart,
    source: { ...chart.source, range: { ...chart.source.range } },
    position: { ...chart.position },
    size: { ...chart.size },
    legend: cloneModel(chart.legend),
    xAxis: cloneModel(chart.xAxis),
    yAxis: cloneModel(chart.yAxis),
    palette: chart.palette ? [...chart.palette] : undefined,
    options: cloneChartOptions(chart.options),
    valueFields: [...chart.valueFields],
    seriesToCategory: chart.seriesToCategory ? Object.fromEntries(Object.entries(chart.seriesToCategory).map(([field, categories]) => [field, [...categories]])) : undefined,
  }
}

export function chartSourceFromGridRange(range: GridRange): CellRange {
  return {
    startRow: range.startRow,
    startCol: range.startColumn,
    endRow: range.endRow,
    endCol: range.endColumn,
  }
}

/** Position charts below their source range in the same coordinate system as the grid. */
export function chartPositionForRange(range: GridRange, sheet: SpreadsheetRenderSheet): { x: number; y: number } {
  const geometry = new SpreadsheetGeometry(sheet)
  return {
    x: GRID_HEADER_WIDTH + geometry.columns.offset(range.startColumn),
    y: GRID_HEADER_HEIGHT + geometry.rows.offset(range.endRow + 2),
  }
}

function toDomainCell(row: number, col: number, cell: ExcelCellView): SpreadsheetCell {
  const value = cell.rawValue ?? parseScalar(cell.value)
  const result = createCell(row, col, value)
  result.formula = cell.formula ?? null
  result.display = cell.displayValue ?? cell.cachedValue ?? cell.value
  result.style = {
    ...defaultCellStyle(),
    font: {
      ...defaultCellStyle().font,
      family: cell.fontFamily ?? 'Aptos',
      size: cell.fontSize ?? 11,
      bold: cell.bold ?? false,
      italic: cell.italic ?? false,
      underline: cell.underline ?? false,
      strike: cell.strike ?? false,
      color: cell.color ?? '#1f2937',
    },
    fill: cell.background ?? '#ffffff',
    horizontal: cell.align ?? 'general',
    vertical: cell.vertical ?? 'middle',
    wrap: cell.wrap ?? false,
    numberFormat: cell.rawNumberFormat ?? cell.numberFormat ?? 'General',
    border: cell.border,
  }
  return result
}

function parseScalar(value: string): CellScalar {
  const trimmed = value.trim()
  if (trimmed === '') return ''
  if (/^(?:true|false)$/i.test(trimmed)) return trimmed.toLowerCase() === 'true'
  if (/^[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?$/i.test(trimmed)) return Number(trimmed)
  return value
}

function parseA1(address: string): { row: number; col: number } | undefined {
  const match = /^\$?([A-Z]+)\$?(\d+)$/i.exec(address.trim())
  if (!match) return undefined
  let col = 0
  for (const character of match[1]!.toUpperCase()) col = col * 26 + character.charCodeAt(0) - 64
  const row = Number(match[2]) - 1
  return Number.isSafeInteger(row) && row >= 0 ? { row, col: col - 1 } : undefined
}
