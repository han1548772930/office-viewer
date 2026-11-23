import type { WorkbookModel } from '../../domain/model'
import type { SheetAxisSpan } from '../../domain/sheet-axis'
import { SpreadsheetGeometry } from '../../rendering/spreadsheet/layout'
import { refreshAutoFilter } from '../sheet-sort-filter/sort-filter-engine'
import { chartColumnName, parseChartSourceRange } from './chart-settings-draft'

export type ChartPlacementGeometry = Pick<SpreadsheetGeometry, 'rows' | 'columns'>

/** Use Canvas axis geometry without constructing a cell-content index. Active table filters still read their source ranges. */
export function chartPlacementGeometry(workbook?: WorkbookModel, sheetId?: string): ChartPlacementGeometry {
  const sheet = workbook?.sheets.find(item => item.id === sheetId)
  if (workbook && !sheet) throw new Error('放置图表的工作表不存在。')
  const collapsed = sheet?.outline?.groups.filter(group => group.collapsed) ?? []
  const collapsedSpans = (axis: 'row' | 'column'): SheetAxisSpan<boolean>[] => collapsed.filter(group => group.axis === axis)
    .map(group => ({ start: group.start, end: group.end, value: true }))
  const filteredRows = new Set(sheet?.autoFilter?.filteredRows ?? [])
  for (const table of workbook?.tables ?? []) {
    if (!sheet || table.sheetId !== sheet.id || !table.autoFilter) continue
    const filter = { ...table.autoFilter, range: { ...table.autoFilter.range }, filteredRows: [] }
    for (const row of refreshAutoFilter({ ...sheet, autoFilter: filter })) filteredRows.add(row)
  }
  return new SpreadsheetGeometry({
    id: sheet?.id ?? 'chart-placement', name: sheet?.name ?? '', cells: {},
    rowCount: sheet?.rowCount ?? 1_048_576, columnCount: sheet?.columnCount ?? 16_384,
    defaultRowHeight: sheet?.defaultRowHeight, defaultColumnWidth: sheet?.defaultColumnWidth,
    defaultRowAutoHeight: sheet?.defaultRowAutoHeight, rowAutoHeights: sheet?.rowAutoHeights,
    rowHeights: sheet?.rowHeights, columnWidths: sheet?.columnWidths,
    rowHeightSpans: sheet?.rowHeightSpans, columnWidthSpans: sheet?.columnWidthSpans,
    rowAutoHeightSpans: sheet?.rowAutoHeightSpans,
    rowHiddenSpans: [...(sheet?.rowHiddenSpans ?? []), ...collapsedSpans('row')],
    columnHiddenSpans: [...(sheet?.columnHiddenSpans ?? []), ...collapsedSpans('column')],
    hiddenRows: sheet?.hiddenRows, hiddenColumns: sheet?.hiddenColumns, filteredRows: [...filteredRows],
  })
}

export function chartCellPosition(address: string, geometry: ChartPlacementGeometry = chartPlacementGeometry()): { x: number; y: number } {
  const range = parseChartSourceRange(address)
  if (!range || address.includes(':')) throw new TypeError('请输入有效的放置单元格，例如 F2。')
  if (range.startRow >= geometry.rows.count || range.startCol >= geometry.columns.count) throw new RangeError('放置单元格超出了工作表边界。')
  const position = { x: geometry.columns.offset(range.startCol), y: geometry.rows.offset(range.startRow) }
  if (!Number.isFinite(position.x) || position.x < 0 || !Number.isFinite(position.y) || position.y < 0) throw new RangeError('工作表尺寸累积后无法生成有限图表坐标。')
  return position
}

export function chartPositionCell(position: { x: number; y: number }, geometry: ChartPlacementGeometry = chartPlacementGeometry()): string {
  if (!Number.isFinite(position.x) || position.x < 0 || !Number.isFinite(position.y) || position.y < 0) throw new RangeError('图表位置须为非负有限坐标。')
  return `${chartColumnName(geometry.columns.locate(position.x))}${geometry.rows.locate(position.y) + 1}`
}
