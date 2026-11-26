import type { CellRange, CellScalar, ChartAxisOptions, PivotFieldModel, PivotTableModel, SheetAutoFilter,
  SheetFilterCriteria, SheetSlicerConnections, SheetTableModel, SlicerCacheModel, SlicerModel } from '../../domain/model.ts'
import { validateSlicerDrawing } from '../slicer/slicer-drawing.ts'
import { drawingAffineBounds, normalizeDrawingAffineGeometry } from '../../domain/drawing-affine.ts'
import { GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH } from '../../rendering/spreadsheet/layout.ts'
import { chartDrawingWorldBounds } from '../charts/chart-drawing-affine.ts'
import { assertSheetCommandRanges, sheetRangeContainsCell, sheetRangesEqual } from './sheet-range-command.ts'
import { sheetClipboardBoundsContain, type SheetClipboardObjectBounds } from './sheet-clipboard-object-geometry.ts'
import type { SheetClipboardLocation, SheetClipboardWorkbookResources } from './sheet-clipboard-types.ts'
import { samePivotSource, sourceFieldId } from '../slicer/slicer-connections.ts'
import { isChartType } from '../charts/chart-catalog.ts'

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
  const controls = new Map((available.slicerCaches ?? []).map(cache => [cache.id, cache]))
  const sourceCaches = new Map((available.pivotCaches ?? []).map(cache => [cache.id, cache]))
  for (const cache of resources.slicerCaches ?? []) {
    if (cache.source.kind === 'table') continue
    const definition = cache.source, sourceCache = sourceCaches.get(definition.pivotCacheId)
    if (!sourceCache && requireResolved) throw new TypeError('剪贴板切片器缺少共享透视来源缓存。')
    if (sourceCache && sourceCache.source.unitId !== cache.unitId) throw new TypeError('剪贴板透视缓存不属于切片器工作簿。')
    if (sourceCache?.fields && !sourceCache.fields.some(field => field.id === definition.sourceFieldId)) {
      throw new TypeError('剪贴板切片器锁定的来源字段不存在。')
    }
    const check = (pivot: PivotTableModel | undefined): void => {
      if (!pivot) { if (requireResolved) throw new TypeError('剪贴板切片器缺少完整连接报告。'); return }
      if (pivot.source.unitId !== cache.unitId || (pivot.pivotCacheId ?? pivot.source.pivotCacheId) !== definition.pivotCacheId
        || sourceCache && !samePivotSource(sourceCache.source, pivot.source)
        || !pivot.fields.some(field => field.area !== 'value' && sourceFieldId(pivot, field) === definition.sourceFieldId)) {
        throw new TypeError('剪贴板连接报告的共享来源或字段无效。')
      }
    }
    for (const ref of cache.connections?.pivotTables ?? []) {
      const pivot = pivots.get(ref.pivotTableId)
      if (pivot && pivot.anchor.subUnitId !== ref.subUnitId) throw new TypeError('剪贴板透视连接宿主无效。')
      check(pivot)
    }
    for (const ref of cache.connections?.ownedPivotCharts ?? []) {
      const chart = available.charts?.find(chart => chart.id === ref.pivotChartId)
      if (chart && (chart.sheetId !== ref.subUnitId || !chart.ownedPivot || chart.pivotTableId)) throw new TypeError('剪贴板独立透视图连接无效。')
      check(chart?.ownedPivot)
    }
  }
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
    if (slicer.cacheId) {
      const cache = controls.get(slicer.cacheId)
      if (!cache) { if (requireResolved) throw new TypeError('剪贴板切片器缺少共享控件缓存。'); continue }
      if (cache.unitId !== slicer.unitId || cache.type !== slicer.type || (cache.source.kind === 'table' ? cache.source.columnId : cache.source.sourceFieldId) !== slicer.sourceFieldId) {
        throw new TypeError('剪贴板切片器与共享缓存不一致。')
      }
      if (cache.source.kind === 'pivot') continue
      const definition = cache.source, table = tables.get(definition.tableId)
      if (!table) { if (requireResolved) throw new TypeError('剪贴板切片器缺少表格来源。'); continue }
      if (table.sheetId !== definition.subUnitId || !table.columns.some(column => column.id === definition.columnId)
        || slicer.tableId !== undefined && slicer.tableId !== definition.tableId) throw new TypeError('剪贴板切片器共享表格来源无效。')
      continue
    }
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
      if (pivot.source.unitId !== slicer.unitId || !pivot.fields.some(field => sourceFieldId(pivot, field) === slicer.sourceFieldId && field.area !== 'value')) {
        throw new TypeError('剪贴板切片器的透视字段绑定无效。')
      }
    } else if (slicer.ownedPivotChartId) {
      const chart = available.charts?.find(chart => chart.id === slicer.ownedPivotChartId)
      if (!chart?.ownedPivot) { if (requireResolved) throw new TypeError('剪贴板切片器缺少独立透视图来源。'); continue }
      if (chart.pivotTableId || chart.ownedPivot.source.unitId !== slicer.unitId
        || !chart.ownedPivot.fields.some(field => field.area !== 'value' && sourceFieldId(chart.ownedPivot!, field) === slicer.sourceFieldId)) {
        throw new TypeError('剪贴板切片器的独立透视图字段无效。')
      }
    }
  }
}

function assertSourceResources(resources: SheetClipboardWorkbookResources, context: SheetClipboardResourceValidationContext): void {
  const { source, rows, columns, sourceOrigin, sourceBounds } = context
  assertSheetCommandRanges([source.range])
  if ((resources.slicerCaches ?? []).some(cache => cache.unitId !== source.unitId)
    || (resources.pivotCaches ?? []).some(cache => cache.source.unitId !== source.unitId)) throw new TypeError('剪贴板共享缓存不属于复制源工作簿。')
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
    if (chart.sheetId !== source.sheetId || !sheetClipboardBoundsContain(sourceBounds, chartDrawingWorldBounds(chart))) {
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
    || field.sourceFieldId !== undefined && !identifier(field.sourceFieldId)
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
    || pivot.pivotCacheId !== undefined && !identifier(pivot.pivotCacheId) || pivot.source.pivotCacheId !== undefined && !identifier(pivot.source.pivotCacheId)
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
  const hints = [slicer.tableId, slicer.pivotTableId, slicer.ownedPivotChartId].filter(value => value !== undefined)
  if (!identifier(slicer.unitId) || !identifier(slicer.hostSheetId) || !identifier(slicer.sourceFieldId) || typeof slicer.caption !== 'string'
    || !['slicer', 'timeline'].includes(slicer.type) || hints.length > 1 || !slicer.cacheId && hints.length !== 1
    || slicer.cacheId !== undefined && !identifier(slicer.cacheId) || slicer.ownedPivotChartId !== undefined && !identifier(slicer.ownedPivotChartId)
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
  if (slicer.drawingGeometry !== undefined) {
    const bounds = drawingAffineBounds(normalizeDrawingAffineGeometry(slicer.drawingGeometry))
    const expected = { ...bounds, left: GRID_HEADER_WIDTH + bounds.left, top: GRID_HEADER_HEIGHT + bounds.top }
    if (slicer.layout && (['left', 'top', 'width', 'height'] as const).some(key => Math.abs(slicer.layout![key] - expected[key])
      > Math.max(1e-7, Math.abs(expected[key]) * Number.EPSILON * 8))) throw new TypeError('剪贴板切片器布局与仿射几何不一致。')
  }
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

function assertConnections(connections: SheetSlicerConnections): void {
  if (!record(connections) || !Array.isArray(connections.pivotTables)
    || connections.ownedPivotCharts !== undefined && !Array.isArray(connections.ownedPivotCharts)) throw new TypeError('剪贴板切片器连接无效。')
  const keys = new Set<string>()
  for (const [kind, entries, idKey] of [
    ['pivotTable', connections.pivotTables, 'pivotTableId'], ['ownedPivotChart', connections.ownedPivotCharts ?? [], 'pivotChartId'],
  ] as const) for (const ref of entries) {
    if (!record(ref) || !identifier(ref.subUnitId) || !identifier(ref[idKey])) throw new TypeError('剪贴板切片器连接引用无效。')
    const key = JSON.stringify([kind, ref.subUnitId, ref[idKey]])
    if (keys.has(key)) throw new TypeError('剪贴板切片器连接重复。')
    keys.add(key)
  }
}

function assertSlicerCache(cache: SlicerCacheModel): void {
  const source = cache.source
  if (!identifier(cache.unitId) || !['slicer', 'timeline'].includes(cache.type) || !record(source)
    || !index(cache.revision, Number.MAX_SAFE_INTEGER) || cache.selectedItems !== undefined && !stringItems(cache.selectedItems)) {
    throw new TypeError('剪贴板切片器共享缓存无效。')
  }
  if (source.kind === 'table') {
    if (!identifier(source.subUnitId) || !identifier(source.tableId) || !identifier(source.columnId) || cache.type === 'timeline'
      || cache.connections !== undefined) throw new TypeError('剪贴板切片器表格来源无效。')
  } else if (source.kind !== 'pivot' || !identifier(source.pivotCacheId) || !identifier(source.sourceFieldId)) {
    throw new TypeError('剪贴板切片器透视来源无效。')
  }
  if (cache.connections !== undefined) assertConnections(cache.connections)
  if (cache.timelineRange !== undefined && (!record(cache.timelineRange) || !['day', 'month', 'quarter', 'year'].includes(cache.timelineRange.level)
    || cache.timelineRange.start !== undefined && !scalar(cache.timelineRange.start) || cache.timelineRange.end !== undefined && !scalar(cache.timelineRange.end))) {
    throw new TypeError('剪贴板时间线共享范围无效。')
  }
}

/** Validate imported clipboard object contracts before they can be installed in a candidate. */
export function assertSheetClipboardWorkbookResources(resources: SheetClipboardWorkbookResources, context?: SheetClipboardResourceValidationContext): void {
  if (!record(resources)) throw new TypeError('剪贴板工作簿资源无效。')
  for (const key of ['tables', 'pivotTables', 'pivotCaches', 'charts', 'slicers', 'slicerCaches'] as const) {
    const entries = resources[key]
    if (entries === undefined) continue
    if (!Array.isArray(entries) || entries.length > SHEET_CLIPBOARD_RESOURCE_BUDGET
      || entries.some(item => !record(item) || !identifier(item.id)) || new Set(entries.map(item => item.id)).size !== entries.length) {
      throw new TypeError('剪贴板对象数量或标识无效。')
    }
  }
  for (const cache of resources.pivotCaches ?? []) {
    if (!record(cache.source) || !identifier(cache.source.unitId) || !identifier(cache.source.subUnitId) || typeof cache.source.sheetName !== 'string'
      || !index(cache.revision, Number.MAX_SAFE_INTEGER) || cache.nextFieldId !== undefined && !index(cache.nextFieldId, Number.MAX_SAFE_INTEGER)) {
      throw new TypeError('剪贴板透视缓存来源无效。')
    }
    assertSheetCommandRanges([cache.source.range])
    if (cache.source.range.endRow <= cache.source.range.startRow) throw new RangeError('剪贴板透视缓存缺少记录行。')
    if (cache.fields !== undefined && (!Array.isArray(cache.fields) || cache.fields.length > cache.source.range.endCol - cache.source.range.startCol + 1
      || cache.fields.some(field => !record(field) || !identifier(field.id) || !index(field.sourceColumn, 16_384)
        || field.sourceColumn < cache.source.range.startCol || field.sourceColumn > cache.source.range.endCol
        || field.name !== undefined && typeof field.name !== 'string') || new Set(cache.fields.map(field => field.id)).size !== cache.fields.length
      || new Set(cache.fields.map(field => field.sourceColumn)).size !== cache.fields.length)) throw new TypeError('剪贴板透视缓存来源字段无效。')
  }
  let connectionCount = 0
  for (const cache of resources.slicerCaches ?? []) {
    assertSlicerCache(cache)
    connectionCount += (cache.connections?.pivotTables.length ?? 0) + (cache.connections?.ownedPivotCharts?.length ?? 0)
    if (connectionCount > SHEET_CLIPBOARD_RESOURCE_BUDGET) throw new RangeError('剪贴板切片器连接数量超出预算。')
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
    if (chart.ownedPivot !== undefined) {
      if (chart.pivotTableId) throw new TypeError('剪贴板透视图报告来源冲突。')
      assertPivot(chart.ownedPivot)
      if (chart.ownedPivot.anchor.subUnitId !== chart.sheetId) throw new TypeError('剪贴板独立透视图报告宿主无效。')
    }
    if (chart.drawingGeometry !== undefined) normalizeDrawingAffineGeometry(chart.drawingGeometry)
    if (!identifier(chart.sheetId) || !isChartType(chart.type)
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
