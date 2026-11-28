import type { PivotTableModel, SheetStyleSource, SpreadsheetSheet, WorkbookModel } from '../../domain/model'
import { cloneModel } from '../../domain/serialization'
import { assertSheetCommandRanges, sheetRangeContainsCell } from '../sheet-workbench/sheet-range-command'
import { getPivotSourceAvailability } from '../pivot/services/pivot-source-lifecycle'
import { CHART_IMAGE_RESOURCE_KEY, readChartImageResources, resolveChartImageResource } from './chart-image-resources'
import type { ChartSourceImportPlan } from './chart-source-import'
import { readFormulaImportSource } from '../formula/formula-import-context'

const equal = (first: unknown, second: unknown): boolean => JSON.stringify(first) === JSON.stringify(second)
const object = (value: unknown): value is Record<string, unknown> => !!value && typeof value === 'object' && !Array.isArray(value)

function assertNewIds(existing: readonly { id: string }[], imported: readonly { id: string }[]): void {
  const ids = new Set(existing.map(item => item.id))
  for (const item of imported) {
    if (!item.id || ids.has(item.id)) throw new Error('图表导入资源 ID 已被占用，请重新粘贴。')
    ids.add(item.id)
  }
}
function assertCatalog(existing: Record<string, unknown> | undefined, imported: Record<string, unknown>): void {
  for (const [id, value] of Object.entries(imported)) {
    if (existing && Object.hasOwn(existing, id) && !equal(existing[id], value)) throw new Error('图表导入的样式、主题或图片资源已更改，请重新粘贴。')
  }
}
function assertRange(sheet: SpreadsheetSheet, range: Parameters<typeof assertSheetCommandRanges>[0][number]): void {
  assertSheetCommandRanges([range])
  if (range.endRow >= (sheet.rowCount ?? 1_048_576) || range.endCol >= (sheet.columnCount ?? 16_384)) throw new RangeError('图表导入范围超出其工作表尺寸。')
}

/** The local numeric OOXML index is derived from sheet ID at application time.
 * A prepared plan can be applied after unrelated sheets are reordered. */
function reindexLocalNames(workbook: WorkbookModel, imported: SpreadsheetSheet[]): void {
  const indices = new Map(workbook.sheets.map((sheet, index) => [sheet.id, index]))
  for (const sheet of imported) {
    const used = new Set<string>()
    for (const key of ['definedNames', 'localDefinedNames'] as const) {
      const values = sheet.metadata?.[key]
      if (!Array.isArray(values)) continue
      for (const value of values) {
        if (!object(value) || typeof value.name !== 'string' || !value.name || value.sheetId !== sheet.id) throw new TypeError('图表导入的本地定义名称作用域无效。')
        const name = value.name.toLowerCase()
        if (used.has(name)) throw new TypeError('图表导入的本地定义名称重复。')
        used.add(name)
        if (value.localSheetId !== undefined) value.localSheetId = typeof value.localSheetId === 'string' ? sheet.id : indices.get(sheet.id)
        if (value.subUnitId !== undefined) value.subUnitId = sheet.id
        if (value.sheet !== undefined) value.sheet = sheet.name
      }
    }
  }
}
function assertBindings(workbook: WorkbookModel, plan: ChartSourceImportPlan, imported: SpreadsheetSheet[]): void {
  const sheets = new Map(workbook.sheets.map(sheet => [sheet.id, sheet]))
  const style = (value: SheetStyleSource | undefined): void => {
    if (typeof value === 'string' && !Object.hasOwn(workbook.styles ?? {}, value)) throw new Error('图表导入的单元格样式引用已经丢失。')
  }
  for (const sheet of imported) {
    style(sheet.defaultStyle)
    for (const axis of [sheet.rowStyles, sheet.columnStyles]) for (const value of Object.values(axis ?? {})) style(value)
    for (const spans of [sheet.rowStyleSpans, sheet.columnStyleSpans]) for (const span of spans ?? []) style(span.value)
    for (const cell of Object.values(sheet.cells)) {
      style(cell.authoredStyleReference); style(cell.themeStyle)
      const source = readFormulaImportSource(cell)
      if (source?.aliasSheetId && !sheets.has(source.aliasSheetId)) throw new TypeError('图表导入公式的名称别名所属工作表已丢失。')
    }
  }
  const tables = new Set((workbook.tables ?? []).map(table => table.id))
  for (const table of plan.tables) {
    const sheet = sheets.get(table.sheetId)
    if (!sheet) throw new TypeError('图表导入的结构化表格所属工作表已经丢失。')
    assertRange(sheet, table.range)
    if (table.columns.length !== table.range.endCol - table.range.startCol + 1
      || new Set(table.columns.map(column => column.id)).size !== table.columns.length
      || new Set(table.columns.map(column => column.name.toLowerCase())).size !== table.columns.length) throw new TypeError('图表导入的结构化表格列身份无效。')
    if (!tables.has(table.id)) throw new TypeError('图表导入的结构化表格未完整绑定。')
  }
  const reports: PivotTableModel[] = [...plan.pivotTables, ...(plan.chart.ownedPivot ? [plan.chart.ownedPivot] : [])]
  for (const report of reports) {
    const source = sheets.get(report.source.subUnitId), anchor = sheets.get(report.anchor.subUnitId)
    if (!source || !anchor || report.anchor.unitId !== workbook.id || getPivotSourceAvailability(workbook, report) !== 'available') throw new TypeError('图表导入的透视报告、数据源或缓存未完整绑定。')
    assertRange(source, report.source.range)
    assertRange(anchor, report.outputRange ?? { startRow: report.anchor.row, endRow: report.anchor.row, startCol: report.anchor.col, endCol: report.anchor.col })
    const cache = workbook.pivotCaches?.find(cache => cache.id === (report.pivotCacheId ?? report.source.pivotCacheId))
    if (!cache || report.fields.some(field => !sheetRangeContainsCell(report.source.range, report.source.range.startRow, field.sourceColumn)
      || cache.fields && !cache.fields.some(canonical => canonical.id === field.sourceFieldId && canonical.sourceColumn === field.sourceColumn))) throw new TypeError('图表导入的透视字段与其缓存来源不一致。')
  }
  const chart = plan.chart
  if (chart.sheetId !== plan.targetSheetId || chart.pivotTableId && chart.ownedPivot) throw new TypeError('图表导入的图表宿主或透视绑定无效。')
  const report = chart.pivotTableId ? workbook.pivotTables?.find(report => report.id === chart.pivotTableId) : chart.ownedPivot
  const source = sheets.get(chart.source.sheetId ?? chart.sheetId)
  if (!source || chart.pivotTableId && !report || chart.ownedPivot && chart.ownedPivot.anchor.subUnitId !== chart.sheetId) throw new TypeError('图表导入的图表来源未完整绑定。')
  if (!report) assertRange(source, chart.source.range)
  const imageId = chart.options?.wordCloud?.maskImageId ?? chart.options?.wordCloud?.maskImage
  if (imageId && !resolveChartImageResource(workbook, chart, imageId)) throw new TypeError('图表导入的图片引用未完整绑定。')
}

/** Stage every allocation and validation before assigning any caller property.
 * Existing target cells and resource identities are retained on both success
 * and failure; the host owns history and the final content publication. */
export function applyChartSourceImport(workbook: WorkbookModel, plan: ChartSourceImportPlan): void {
  if (plan.version !== 1 || workbook.id !== plan.targetUnitId || !workbook.sheets.some(sheet => sheet.id === plan.targetSheetId)) throw new TypeError('图表依赖计划不属于目标工作簿。')
  assertNewIds(workbook.sheets, plan.sheets)
  assertNewIds(workbook.tables ?? [], plan.tables)
  assertNewIds(workbook.pivotTables ?? [], plan.pivotTables)
  assertNewIds(workbook.pivotCaches ?? [], plan.pivotCaches)
  const sheetNames = new Set(workbook.sheets.map(sheet => sheet.name.toLowerCase()))
  for (const sheet of plan.sheets) {
    if (!sheet.name || sheetNames.has(sheet.name.toLowerCase())) throw new Error('图表导入的工作表名称已被占用，请重新粘贴。')
    sheetNames.add(sheet.name.toLowerCase())
  }
  const names = new Set((workbook.tables ?? []).map(table => table.name.toLowerCase()))
  for (const metadata of [workbook.metadata, ...workbook.sheets.map(sheet => sheet.metadata)]) for (const key of ['definedNames', 'localDefinedNames']) {
    const values = metadata?.[key]
    if (Array.isArray(values)) for (const value of values) if (object(value) && typeof value.name === 'string') names.add(value.name.toLowerCase())
  }
  for (const value of [...plan.tables, ...Object.values(plan.names).flat()]) {
    if (!value || typeof value.name !== 'string' || !value.name || names.has(value.name.toLowerCase())) throw new Error('图表导入的表格或定义名称已被占用，请重新粘贴。')
    names.add(value.name.toLowerCase())
  }
  assertCatalog(workbook.styles, plan.styles)
  // Theme lookup is case insensitive; an unrelated destination theme must not
  // start applying to a source rule after asynchronous preparation.
  const themes = new Map(Object.entries(workbook.rangeThemes ?? {}).map(([name, value]) => [name.toLowerCase(), value]))
  for (const [name, value] of Object.entries(plan.rangeThemes)) if (themes.has(name.toLowerCase()) && !equal(themes.get(name.toLowerCase()), value)) throw new Error('图表导入的范围主题名称已被占用，请重新粘贴。')
  const images = readChartImageResources(workbook)
  assertCatalog(images, plan.images)
  const importedSheets = cloneModel(plan.sheets), metadata = { ...workbook.metadata }
  for (const key of ['definedNames', 'localDefinedNames'] as const) if (plan.names[key]?.length) metadata[key] = [
    ...(Array.isArray(metadata[key]) ? metadata[key] : []), ...cloneModel(plan.names[key]!),
  ]
  if (Object.keys(plan.images).length) metadata[CHART_IMAGE_RESOURCE_KEY] = { ...images, ...cloneModel(plan.images) }
  const staged: WorkbookModel = { ...workbook, sheets: [...workbook.sheets, ...importedSheets], metadata,
    ...(plan.tables.length ? { tables: [...(workbook.tables ?? []), ...cloneModel(plan.tables)] } : {}),
    ...(plan.pivotTables.length ? { pivotTables: [...(workbook.pivotTables ?? []), ...cloneModel(plan.pivotTables)] } : {}),
    ...(plan.pivotCaches.length ? { pivotCaches: [...(workbook.pivotCaches ?? []), ...cloneModel(plan.pivotCaches)] } : {}),
    ...(Object.keys(plan.styles).length ? { styles: { ...workbook.styles, ...cloneModel(plan.styles) } } : {}),
    ...(Object.keys(plan.rangeThemes).length ? { rangeThemes: { ...workbook.rangeThemes, ...cloneModel(plan.rangeThemes) } } : {}),
  }
  reindexLocalNames(staged, importedSheets)
  readChartImageResources(staged)
  assertBindings(staged, plan, importedSheets)
  workbook.sheets = staged.sheets
  if (plan.tables.length) workbook.tables = staged.tables
  if (plan.pivotTables.length) workbook.pivotTables = staged.pivotTables
  if (plan.pivotCaches.length) workbook.pivotCaches = staged.pivotCaches
  if (Object.keys(plan.styles).length) workbook.styles = staged.styles
  if (Object.keys(plan.rangeThemes).length) workbook.rangeThemes = staged.rangeThemes
  workbook.metadata = staged.metadata
}
