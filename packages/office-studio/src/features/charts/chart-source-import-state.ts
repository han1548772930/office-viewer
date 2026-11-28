import type { CellRange, SpreadsheetSheet, WorkbookModel } from '../../domain/model'
import { sheetRangeContainsCell } from '../sheet-workbench/sheet-range-command'
import type { ChartSourceImportPlan } from './chart-source-import'
import { resolveChartImageResource } from './chart-image-resources'

export interface ChartSourceImportState {
  /** Exact resource text, not a hash with an undocumented collision boundary. */
  catalog: string
  cells: Record<string, Record<string, string>>
  images: Record<string, string>
}

function sheetDescriptor(sheet: SpreadsheetSheet) {
  const { cells: _cells, images: _images, shapes: _shapes, drawingState: _drawingState,
    backgroundImage: _backgroundImage, dataValidations: _validations, conditionalFormats: _formats,
    rangeProtectionRules: _rangePermissions, worksheetProtection: _sheetPermissions, ...descriptor } = sheet
  return descriptor
}

/** Cell data stays sparse and is captured when read. Resource identities and
 * name/format descriptors are captured before the first cooperative yield. */
export function chartSourceImportCatalog(workbook: WorkbookModel): string {
  return JSON.stringify({ id: workbook.id, date1904: workbook.date1904, activeSheetId: workbook.activeSheetId,
    defaultStyle: workbook.defaultStyle, styles: workbook.styles, rangeThemes: workbook.rangeThemes,
    definedNames: workbook.metadata?.definedNames, localDefinedNames: workbook.metadata?.localDefinedNames,
    tables: workbook.tables, pivotTables: workbook.pivotTables, pivotCaches: workbook.pivotCaches,
    sheets: workbook.sheets.map(sheetDescriptor) })
}

/** Revisions are the normal concurrency boundary. This exact comparison also
 * rejects direct mutations of the publicly accessible live model, including
 * new sparse cells inside a copied whole-axis range and altered name sources. */
export function assertChartSourceImportState(workbook: WorkbookModel, plan: ChartSourceImportPlan): void {
  if (chartSourceImportCatalog(workbook) !== plan.sourceState.catalog) throw new Error('图表来源的名称、样式、表格、透视缓存或工作表结构已更改，请重新粘贴。')
  const chart = workbook.charts?.find(chart => chart.id === plan.sourceChartId)
  for (const [id, snapshot] of Object.entries(plan.sourceState.images)) {
    const current = chart && resolveChartImageResource(workbook, chart, id)
    if (!current || JSON.stringify(current) !== snapshot) throw new Error('图表来源图片资源在准备期间已更改，请重新粘贴。')
  }
  const ranges = new Map<string, CellRange[]>()
  for (const entry of plan.sourceRanges) {
    const current = ranges.get(entry.sheetId) ?? []; current.push(entry.range); ranges.set(entry.sheetId, current)
  }
  for (const [sheetId, snapshots] of Object.entries(plan.sourceState.cells)) {
    const sheet = workbook.sheets.find(sheet => sheet.id === sheetId)
    if (!sheet) throw new Error('图表来源工作表已经关闭，请重新粘贴。')
    const imported = ranges.get(sheetId) ?? []
    let count = 0
    for (const [key, cell] of Object.entries(sheet.cells)) {
      if (!imported.some(range => sheetRangeContainsCell(range, cell.row, cell.col))) continue
      count++
      if (!Object.hasOwn(snapshots, key) || JSON.stringify(cell) !== snapshots[key]) throw new Error('图表来源单元格在准备期间已更改，请重新粘贴。')
    }
    if (count !== Object.keys(snapshots).length) throw new Error('图表来源单元格在准备期间已删除，请重新粘贴。')
  }
}
