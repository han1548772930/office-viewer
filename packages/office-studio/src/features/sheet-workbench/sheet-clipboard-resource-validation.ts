import type { CellRange, CellScalar, ChartAxisOptions, PivotFieldModel, PivotTableModel, SheetAutoFilter,
  SheetFilterCriteria, SheetTableModel, SlicerModel } from '../../domain/model.ts'
import { validateSlicerDrawing } from '../slicer/slicer-drawing.ts'
import { assertSheetCommandRanges, sheetRangeContainsCell, sheetRangesEqual } from './sheet-range-command.ts'
import { sheetClipboardBoundsContain, type SheetClipboardObjectBounds } from './sheet-clipboard-object-geometry.ts'
import type { SheetClipboardLocation, SheetClipboardWorkbookResources } from './sheet-clipboard-types.ts'

export const SHEET_CLIPBOARD_RESOURCE_BUDGET = 100_000
const identifier = (value: unknown): value is string => typeof value === 'string' && value.trim().length > 0
const record = (value: unknown): value is Record<string, unknown> => !!value && typeof value === 'object' && !Array.isArray(value)
const scalar = (value: unknown): value is CellScalar => value === null || typeof value === 'string' || typeof value === 'boolean'
  || typeof value === 'number' && Number.isFinite(value)
const index = (value: unknown, limit: number): value is number => typeof value === 'number' && Number.isSafeInteger(value) && value >= 0 && value < limit
const count = (value: unknown, limit: number): value is number => typeof value === 'number' && Number.isSafeInteger(value) && value > 0 && value <= limit
const stringItems = (items: unknown): items is string[] => Array.isArray(items) && items.every(item => typeof item === 'string')
const contains = (outer: CellRange, inner: CellRange) => sheetRangeContainsCell(outer, inner.startRow, inner.startCol)
  && sheetRangeContainsCell(outer, inner.endRow, inner.endCol)
const pivotRange = (pivot: PivotTableModel): CellRange => pivot.outputRange ?? {
  startRow: pivot.anchor.row, endRow: pivot.anchor.row, startCol: pivot.anchor.col, endCol: pivot.anchor.col,
}

export interface SheetClipboardResourceValidationContext {
  source: SheetClipboardLocation
  rows: readonly number[]
  columns: readonly number[]
  sourceOrigin?: { left: number; top: number }
  sourceBounds?: SheetClipboardObjectBounds
}

function lowerBound(values: readonly number[], value: number): number {
  let low = 0, high = values.length
  while (low < high) { const middle = (low + high) >>> 1; if (values[middle]! < value) low = middle + 1; else high = middle }
  return low
}
function fullyRepresented(values: readonly number[], first: number, last: number): boolean {
  return lowerBound(values, last + 1) - lowerBound(values, first) === last - first + 1
}

/** A dependency outside a copied rectangle can remain in the source workbook until planning resolves it. */
export function assertSheetClipboardResourceBindings(
  resources: SheetClipboardWorkbookResources, available: SheetClipboardWorkbookResources = resources, requireResolved = true,
): void {
  const tables = new Map((available.tables ?? []).map(table => [table.id, table]))
  const pivots = new Map((available.pivotTables ?? []).map(pivot => [pivot.id, pivot]))
  for (const chart of resources.charts ?? []) {
    if (!chart.pivotTableId) continue
    const pivot = pivots.get(chart.pivotTableId)
    if (!pivot) {
      if (requireResolved) throw new TypeError('剪贴板透视图缺少其透视数据源。')
      continue
    }
    if ((chart.source.sheetId ?? chart.sheetId) !== pivot.anchor.subUnitId || !sheetRangesEqual(chart.source.range, pivotRange(pivot))) {
      throw new TypeError('剪贴板透视图的数据源与透视输出不一致。')
    }
  }
  for (const slicer of resources.slicers ?? []) {
    if (slicer.tableId) {
      const table = tables.get(slicer.tableId)
      if (!table) {
        if (requireResolved) throw new TypeError('剪贴板切片器缺少其表格数据源。')
        continue
      }
      if (!table.columns.some(column => column.id === slicer.sourceFieldId)) throw new TypeError('剪贴板切片器的表格列绑定无效。')
    } else if (slicer.pivotTableId) {
      const pivot = pivots.get(slicer.pivotTableId)
      if (!pivot) {
        if (requireResolved) throw new TypeError('剪贴板切片器缺少其透视数据源。')
        continue
      }
      if (pivot.source.unitId !== slicer.unitId || !pivot.fields.some(field => field.id === slicer.sourceFieldId && field.area !== 'value')) {
        throw new TypeError('剪贴板切片器的透视字段绑定无效。')
      }
    }
  }
}

function assertSourceResources(resources: SheetClipboardWorkbookResources, context: SheetClipboardResourceValidationContext): void {
  const { source, rows, columns, sourceOrigin, sourceBounds } = context
  assertSheetCommandRanges([source.range])
  for (const table of resources.tables ?? []) {
    if (table.sheetId !== source.sheetId || !contains(source.range, table.range)
      || !fullyRepresented(columns, table.range.startCol, table.range.endCol)
      || table.showHeaderRow && rows[lowerBound(rows, table.range.startRow)] !== table.range.startRow) throw new TypeError('剪贴板表格不属于完整的复制源范围。')
  }
  for (const pivot of resources.pivotTables ?? []) {
    const output = pivotRange(pivot)
    if (pivot.anchor.unitId !== source.unitId || pivot.source.unitId !== source.unitId || pivot.anchor.subUnitId !== source.sheetId
      || !contains(source.range, output) || !fullyRepresented(rows, output.startRow, output.endRow)
      || !fullyRepresented(columns, output.startCol, output.endCol)) throw new TypeError('剪贴板透视对象不属于完整的复制源范围。')
  }
  if ((resources.charts?.length ?? 0) + (resources.slicers?.length ?? 0) === 0) return
  if (!sourceBounds || ![sourceBounds.left, sourceBounds.top, sourceBounds.width, sourceBounds.height].every(Number.isFinite)
    || sourceBounds.left < 0 || sourceBounds.top < 0 || sourceBounds.width <= 0 || sourceBounds.height <= 0
    || !sourceOrigin || sourceOrigin.left !== sourceBounds.left || sourceOrigin.top !== sourceBounds.top) {
    throw new TypeError('剪贴板绘图缺少一致的源范围几何。')
  }
  for (const chart of resources.charts ?? []) {
    if (chart.sheetId !== source.sheetId || !sheetClipboardBoundsContain(sourceBounds, { left: chart.position.x, top: chart.position.y, ...chart.size })) {
      throw new TypeError('剪贴板图表不属于完整选中的复制源范围。')
    }
  }
  for (const slicer of resources.slicers ?? []) {
    if (slicer.unitId !== source.unitId || slicer.hostSheetId !== source.sheetId || !slicer.layout || !sheetClipboardBoundsContain(sourceBounds, slicer.layout)) {
      throw new TypeError('剪贴板切片器不属于完整选中的复制源范围。')
    }
  }
}

function assertFilterCriterion(criterion: SheetFilterCriteria): void {
  if (!record(criterion)) throw new TypeError('剪贴板筛选条件无效。')
  if (criterion.kind === 'values') {
    if (!Array.isArray(criterion.selectedItems) || !criterion.selectedItems.every(scalar)) throw new TypeError('剪贴板筛选值无效。')
  } else if (criterion.kind !== 'condition' || !['equals', 'doesNotEqual', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual',
    'between', 'notBetween', 'contains', 'doesNotContain', 'beginsWith', 'endsWith', 'isBlank', 'isNotBlank', 'isError'].includes(criterion.operator)
    || criterion.value !== undefined && !scalar(criterion.value) || criterion.secondValue !== undefined && !scalar(criterion.secondValue)
    || criterion.caseSensitive !== undefined && typeof criterion.caseSensitive !== 'boolean') throw new TypeError('剪贴板筛选条件无效。')
}

function assertTableFilter(filter: SheetAutoFilter, table: SheetTableModel): void {
  if (!record(filter)) throw new TypeError('剪贴板表格筛选无效。')
  assertSheetCommandRanges([filter.range])
  if (!contains(table.range, filter.range) || !record(filter.columns) || !Array.isArray(filter.filteredRows)
    || filter.filteredRows.some(row => !index(row, 1_048_576) || row <= filter.range.startRow || row > filter.range.endRow)) {
    throw new TypeError('剪贴板表格筛选范围无效。')
  }
  for (const [column, criterion] of Object.entries(filter.columns)) {
    if (!/^\d+$/.test(column) || !index(Number(column), 16_384) || Number(column) < filter.range.startCol || Number(column) > filter.range.endCol) {
      throw new RangeError('剪贴板表格筛选列无效。')
    }
    assertFilterCriterion(criterion)
  }
}

function assertPivotField(field: PivotFieldModel, pivot: PivotTableModel): void {
  if (!record(field) || !identifier(field.id) || typeof field.name !== 'string' || !index(field.sourceColumn, 16_384)
    || field.sourceColumn < pivot.source.range.startCol || field.sourceColumn > pivot.source.range.endCol
    || !['row', 'column', 'filter', 'value'].includes(field.area) || !index(field.position, 16_384)
    || field.aggregation !== undefined && !['average', 'count', 'countNums', 'max', 'min', 'product', 'stdDev', 'stdDevp', 'sum', 'var', 'varp'].includes(field.aggregation)
    || field.sort !== undefined && !['ascending', 'descending'].includes(field.sort)
    || field.displayName !== undefined && typeof field.displayName !== 'string' || field.format !== undefined && typeof field.format !== 'string'
    || field.selectedItems !== undefined && !stringItems(field.selectedItems)) throw new TypeError('剪贴板透视字段无效。')
  if (field.valueFilter !== undefined) {
    const filter = field.valueFilter
    if (!record(filter) || !['valueGreaterThan', 'valueGreaterThanOrEqual', 'valueLessThan', 'valueLessThanOrEqual', 'valueEqual',
      'valueNotEqual', 'valueBetween', 'valueNotBetween'].includes(filter.operator) || !Number.isFinite(filter.expected)
      || filter.expected2 !== undefined && !Number.isFinite(filter.expected2)
      || ['valueBetween', 'valueNotBetween'].includes(filter.operator) && !Number.isFinite(filter.expected2)
      || !pivot.fields.some(candidate => candidate.id === filter.valueFieldId && candidate.area === 'value')) throw new TypeError('剪贴板透视值筛选无效。')
  }
  if (field.numberGroup !== undefined && (!record(field.numberGroup) || ![field.numberGroup.start, field.numberGroup.end, field.numberGroup.interval].every(Number.isFinite)
    || field.numberGroup.end < field.numberGroup.start || field.numberGroup.interval <= 0)) throw new TypeError('剪贴板透视数值分组无效。')
  if (field.dateGroupType !== undefined && !['year', 'quarter', 'month', 'day'].includes(field.dateGroupType)) throw new TypeError('剪贴板透视日期分组无效。')
  if (field.manualGroups !== undefined && (!Array.isArray(field.manualGroups) || field.manualGroups.some(group => !record(group)
    || !identifier(group.name) || !Array.isArray(group.members) || !group.members.every(scalar)))) throw new TypeError('剪贴板透视手动分组无效。')
}

function assertPivot(pivot: PivotTableModel): void {
  if (!record(pivot.source) || !record(pivot.anchor) || !identifier(pivot.source.unitId) || !identifier(pivot.source.subUnitId)
    || typeof pivot.source.sheetName !== 'string' || !identifier(pivot.anchor.unitId) || !identifier(pivot.anchor.subUnitId)
    || pivot.source.unitId !== pivot.anchor.unitId || !index(pivot.anchor.row, 1_048_576) || !index(pivot.anchor.col, 16_384)
    || !Array.isArray(pivot.fields) || pivot.fields.length > 16_384 || new Set(pivot.fields.map(field => field?.id)).size !== pivot.fields.length
    || !record(pivot.options) || !index(pivot.revision, Number.MAX_SAFE_INTEGER)) throw new TypeError('剪贴板透视对象无效。')
  assertSheetCommandRanges([pivot.source.range])
  if (pivot.source.range.endRow <= pivot.source.range.startRow) throw new RangeError('剪贴板透视数据源缺少记录行。')
  if (pivot.outputRange) {
    assertSheetCommandRanges([pivot.outputRange])
    if (pivot.outputRange.startRow !== pivot.anchor.row || pivot.outputRange.startCol !== pivot.anchor.col) throw new RangeError('剪贴板透视输出与锚点不一致。')
  }
  for (const key of ['repeatRowLabels', 'showRowGrandTotal', 'showColGrandTotal', 'showRowSubtotals', 'showColSubtotals'] as const) {
    if (typeof pivot.options[key] !== 'boolean') throw new TypeError('剪贴板透视选项无效。')
  }
  if (!['row', 'column'].includes(pivot.options.valuePosition) || pivot.dateGroupingMode !== undefined && !['derived', 'excelLike'].includes(pivot.dateGroupingMode)) {
    throw new TypeError('剪贴板透视布局选项无效。')
  }
  for (const field of pivot.fields) assertPivotField(field, pivot)
  const output = pivot.outputRange ?? { startRow: pivot.anchor.row, endRow: pivot.anchor.row, startCol: pivot.anchor.col, endCol: pivot.anchor.col }
  const assertOutputKey = (key: string): void => {
    const match = /^(\d+):(\d+)$/.exec(key)
    if (!match || !index(Number(match[1]), 1_048_576) || !index(Number(match[2]), 16_384)
      || !sheetRangeContainsCell(output, Number(match[1]), Number(match[2]))) throw new RangeError('剪贴板透视生成坐标无效。')
  }
  if (pivot.generatedCellKeys !== undefined) {
    if (!stringItems(pivot.generatedCellKeys) || new Set(pivot.generatedCellKeys).size !== pivot.generatedCellKeys.length) throw new TypeError('剪贴板透视生成坐标无效。')
    pivot.generatedCellKeys.forEach(assertOutputKey)
  }
  if (pivot.drillDownMap !== undefined) {
    if (!record(pivot.drillDownMap)) throw new TypeError('剪贴板透视下钻数据无效。')
    for (const [key, entry] of Object.entries(pivot.drillDownMap)) {
      assertOutputKey(key)
      if (!record(entry) || typeof entry.valueFieldId !== 'string' || !record(entry.fieldFilters) || !Object.values(entry.fieldFilters).every(scalar)
        || entry.fieldMembers !== undefined && (!record(entry.fieldMembers) || Object.values(entry.fieldMembers).some(items => !Array.isArray(items) || !items.every(scalar)))) {
        throw new TypeError('剪贴板透视下钻数据无效。')
      }
    }
  }
}

function assertChartAxis(axis: ChartAxisOptions): void {
  if (!record(axis) || !['linear', 'log'].includes(axis.scaleType) || axis.title !== undefined && typeof axis.title !== 'string'
    || axis.logBase !== undefined && (!Number.isFinite(axis.logBase) || axis.logBase < 2 || axis.logBase > 1000)) throw new TypeError('剪贴板图表坐标轴无效。')
}

function assertSlicer(slicer: SlicerModel): void {
  if (!identifier(slicer.unitId) || !identifier(slicer.hostSheetId) || !identifier(slicer.sourceFieldId) || typeof slicer.caption !== 'string'
    || !['slicer', 'timeline'].includes(slicer.type) || Boolean(slicer.tableId) === Boolean(slicer.pivotTableId)
    || slicer.tableId !== undefined && !identifier(slicer.tableId) || slicer.pivotTableId !== undefined && !identifier(slicer.pivotTableId)
    || !record(slicer.position) || !record(slicer.size) || !index(slicer.position.row, 1_048_576) || !index(slicer.position.column, 16_384)
    || !count(slicer.size.rows, 1_048_576) || !count(slicer.size.columns, 16_384)
    || slicer.position.row + slicer.size.rows > 1_048_576 || slicer.position.column + slicer.size.columns > 16_384
    || !index(slicer.revision, Number.MAX_SAFE_INTEGER) || slicer.selectedItems !== undefined && !stringItems(slicer.selectedItems)) throw new TypeError('剪贴板切片器无效。')
  if (slicer.drawing) {
    validateSlicerDrawing(slicer.drawing)
    for (const point of [slicer.drawing.from, slicer.drawing.to]) if (!index(point.row, 1_048_576) || !index(point.column, 16_384)) {
      throw new RangeError('剪贴板切片器绘图锚点超出工作表。')
    }
  }
  if (slicer.layout && (!record(slicer.layout) || ![slicer.layout.left, slicer.layout.top, slicer.layout.width, slicer.layout.height].every(Number.isFinite)
    || slicer.layout.width <= 0 || slicer.layout.height <= 0)) throw new TypeError('剪贴板切片器布局无效。')
  if (slicer.timelineRange && (!record(slicer.timelineRange) || !['day', 'month', 'quarter', 'year'].includes(slicer.timelineRange.level)
    || slicer.timelineRange.start !== undefined && !scalar(slicer.timelineRange.start) || slicer.timelineRange.end !== undefined && !scalar(slicer.timelineRange.end))) {
    throw new TypeError('剪贴板时间线范围无效。')
  }
  for (const key of ['level', 'selectionLevel'] as const) if (slicer[key] !== undefined && !['day', 'month', 'quarter', 'year'].includes(slicer[key]!)) throw new TypeError('剪贴板时间线粒度无效。')
  for (const key of ['showCaption', 'lockedPosition', 'showHeader', 'showSelectionLabel', 'showTimeLevel', 'showHorizontalScrollbar'] as const) {
    if (slicer[key] !== undefined && typeof slicer[key] !== 'boolean') throw new TypeError('剪贴板切片器显示选项无效。')
  }
  for (const key of ['startItem', 'scrollPosition'] as const) if (slicer[key] !== undefined && (!Number.isFinite(slicer[key]) || slicer[key]! < 0)) throw new RangeError('剪贴板切片器滚动位置无效。')
  for (const key of ['columnCount', 'rowHeight'] as const) if (slicer[key] !== undefined && (!Number.isFinite(slicer[key]) || slicer[key]! <= 0)) throw new RangeError('剪贴板切片器项目尺寸无效。')
}

/** Validate imported clipboard object contracts before they can be installed in a candidate. */
export function assertSheetClipboardWorkbookResources(resources: SheetClipboardWorkbookResources, context?: SheetClipboardResourceValidationContext): void {
  if (!record(resources)) throw new TypeError('剪贴板工作簿资源无效。')
  for (const key of ['tables', 'pivotTables', 'charts', 'slicers'] as const) {
    const entries = resources[key]
    if (entries === undefined) continue
    if (!Array.isArray(entries) || entries.length > SHEET_CLIPBOARD_RESOURCE_BUDGET
      || entries.some(item => !record(item) || !identifier(item.id)) || new Set(entries.map(item => item.id)).size !== entries.length) {
      throw new TypeError('剪贴板对象数量或标识无效。')
    }
  }
  const names = new Set<string>()
  for (const table of resources.tables ?? []) {
    assertSheetCommandRanges([table.range])
    if (!identifier(table.sheetId) || !identifier(table.name) || names.has(table.name.toLowerCase())
      || !Array.isArray(table.columns) || table.columns.length !== table.range.endCol - table.range.startCol + 1
      || table.columns.some(column => !record(column) || !identifier(column.id) || !identifier(column.name))
      || new Set(table.columns.map(column => column.id)).size !== table.columns.length
      || new Set(table.columns.map(column => column.name.toLowerCase())).size !== table.columns.length
      || typeof table.showHeaderRow !== 'boolean' || typeof table.showTotalsRow !== 'boolean' || typeof table.styleName !== 'string') throw new TypeError('剪贴板结构化表格无效。')
    names.add(table.name.toLowerCase())
    if (table.autoFilter) assertTableFilter(table.autoFilter, table)
  }
  for (const pivot of resources.pivotTables ?? []) assertPivot(pivot)
  for (const chart of resources.charts ?? []) {
    if (!identifier(chart.sheetId) || !['bar', 'line', 'area', 'pie', 'doughnut', 'stacked-column', 'scatter', 'radar'].includes(chart.type)
      || !record(chart.source) || !['columns', 'rows'].includes(chart.source.orientation) || chart.source.sheetId !== undefined && !identifier(chart.source.sheetId)
      || !record(chart.position) || !record(chart.size) || ![chart.position.x, chart.position.y, chart.size.width, chart.size.height].every(Number.isFinite)
      || chart.position.x < 0 || chart.position.y < 0 || chart.size.width <= 0 || chart.size.height <= 0 || chart.title !== undefined && typeof chart.title !== 'string'
      || !record(chart.legend) || typeof chart.legend.visible !== 'boolean' || !['top', 'right', 'bottom', 'left'].includes(chart.legend.position)
      || !Array.isArray(chart.valueFields) || !index(chart.categoryField, 1_048_576) || chart.valueFields.some(field => !index(field, 1_048_576))
      || new Set(chart.valueFields).size !== chart.valueFields.length || !index(chart.revision, Number.MAX_SAFE_INTEGER)
      || chart.pivotTableId !== undefined && !identifier(chart.pivotTableId)) throw new TypeError('剪贴板图表无效。')
    assertSheetCommandRanges([chart.source.range])
    assertChartAxis(chart.xAxis); assertChartAxis(chart.yAxis)
    const fieldCount = chart.source.orientation === 'rows' ? chart.source.range.endRow - chart.source.range.startRow + 1
      : chart.source.range.endCol - chart.source.range.startCol + 1
    if (!chart.pivotTableId && (chart.categoryField >= fieldCount || chart.valueFields.some(field => field >= fieldCount))) throw new RangeError('剪贴板图表字段超出数据源。')
    if (chart.seriesToCategory !== undefined) {
      if (!record(chart.seriesToCategory)) throw new TypeError('剪贴板图表系列映射无效。')
      for (const [field, categories] of Object.entries(chart.seriesToCategory)) if (!/^\d+$/.test(field) || !index(Number(field), fieldCount)
        || !Array.isArray(categories) || categories.length !== 1 || !index(categories[0], fieldCount)) throw new RangeError('剪贴板图表系列映射无效。')
    }
  }
  for (const slicer of resources.slicers ?? []) assertSlicer(slicer)
  assertSheetClipboardResourceBindings(resources, resources, false)
  if (context) assertSourceResources(resources, context)
}
