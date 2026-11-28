import type { CellRange, PivotTableModel, SheetAutoFilter, SheetTableModel, SlicerModel, WorkbookModel } from '../../domain/model.ts'
import { sheetImageClipboardBounds } from '../sheet-images/sheet-image-clipboard.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { rewriteFormulaImportCell } from '../formula/formula-import-context.ts'
import { resolveSlicerBounds, slicerGeometryWithBounds } from '../slicer/slicer-drawing.ts'
import { chartDrawingWorldBounds } from '../charts/chart-drawing-affine.ts'
import { assertSheetCommandRanges, sheetRangeContainsCell, sheetRangesEqual, sheetRangesIntersect } from './sheet-range-command.ts'
import { sheetClipboardBoundsContain, sheetClipboardBoundsDrawing, sheetClipboardObjectAxes, sheetClipboardObjectRange,
  sheetClipboardRangeBounds, type SheetClipboardObjectBounds } from './sheet-clipboard-object-geometry.ts'
import type { SheetClipboardLocation, SheetClipboardSnapshot, SheetClipboardWorkbookResources } from './sheet-clipboard-types.ts'
import type { SheetClipboardMove } from './sheet-clipboard-references.ts'
import type { SheetClipboardTargetGeometry } from './sheet-clipboard-target.ts'
import type { SheetPasteMode } from './sheet-paste.ts'
import { assertSheetClipboardResourceBindings, assertSheetClipboardWorkbookResources, SHEET_CLIPBOARD_RESOURCE_BUDGET } from './sheet-clipboard-resource-validation.ts'
import { normalizeSlicerCaches } from '../slicer/slicer-connections.ts'
import { copyPivotReportCache, copySlicerViewCache, reconcileSlicerConnectionResources } from '../slicer/slicer-connections-lifecycle.ts'
import { invalidatePivotSourceContentRanges, prepareCopiedPivotSourceReport, refreshPivotSourceWorksheetOutputs,
  transformWorkbookPivotSourceReports, type PivotSourceStructureTransform } from '../pivot/services/pivot-source-lifecycle.ts'
import { synchronizeLinkedPivotChartReports } from '../pivot/services/pivot-chart-report-lifecycle.ts'
export { assertSheetClipboardWorkbookResources } from './sheet-clipboard-resource-validation.ts'

const RESOURCE_BUDGET = SHEET_CLIPBOARD_RESOURCE_BUDGET
const contains = (outer: CellRange, inner: CellRange): boolean => sheetRangeContainsCell(outer, inner.startRow, inner.startCol)
  && sheetRangeContainsCell(outer, inner.endRow, inner.endCol)
const pivotRange = (pivot: PivotTableModel): CellRange => pivot.outputRange ?? {
  startRow: pivot.anchor.row, endRow: pivot.anchor.row, startCol: pivot.anchor.col, endCol: pivot.anchor.col,
}
const chartBounds = chartDrawingWorldBounds
const newId = (kind: string): string => `clipboard-${kind}-${crypto.randomUUID()}`
const movedRange = (range: CellRange, move: SheetClipboardMove): CellRange => {
  const row = move.targetRange.startRow - move.sourceRange.startRow, col = move.targetRange.startCol - move.sourceRange.startCol
  const next = { startRow: range.startRow + row, endRow: range.endRow + row, startCol: range.startCol + col, endCol: range.endCol + col }
  assertSheetCommandRanges([next]); return next
}

export function cloneSheetClipboardWorkbookResources(workbook: SheetClipboardWorkbookResources): SheetClipboardWorkbookResources {
  const result: SheetClipboardWorkbookResources = {}
  if (workbook.tables !== undefined) result.tables = cloneModel(workbook.tables)
  if (workbook.pivotTables !== undefined) result.pivotTables = cloneModel(workbook.pivotTables)
  if (workbook.pivotCaches !== undefined) result.pivotCaches = cloneModel(workbook.pivotCaches)
  if (workbook.charts !== undefined) result.charts = cloneModel(workbook.charts)
  if (workbook.slicers !== undefined) result.slicers = cloneModel(workbook.slicers)
  if (workbook.slicerCaches !== undefined) result.slicerCaches = cloneModel(workbook.slicerCaches)
  return result
}

/** Table/pivot content is captured only as complete objects; drawings follow the fixed contained-object rule. */
export function captureSheetClipboardWorkbookResources(workbook: WorkbookModel, location: SheetClipboardLocation, rows: readonly number[]) {
  // Materialize legacy identities in detached resource models, never in the copy source or its cell data.
  workbook = { ...workbook, ...cloneSheetClipboardWorkbookResources(workbook) }
  normalizeSlicerCaches(workbook)
  const sheet = workbook.sheets.find(item => item.id === location.sheetId)
  if (!sheet) throw new Error('复制资源的工作表已移除。')
  const bounds = sheetClipboardRangeBounds(workbook, sheet, location.range), axes = sheetClipboardObjectAxes(workbook, sheet)
  const rowSet = new Set(rows)
  const resources: SheetClipboardWorkbookResources = {
    tables: (workbook.tables ?? []).filter(table => table.sheetId === sheet.id && contains(location.range, table.range)
      && (!table.showHeaderRow || rowSet.has(table.range.startRow))),
    pivotTables: (workbook.pivotTables ?? []).filter(pivot => pivot.anchor.subUnitId === sheet.id && contains(location.range, pivotRange(pivot))
      && rows.filter(row => row >= pivotRange(pivot).startRow && row <= pivotRange(pivot).endRow).length === pivotRange(pivot).endRow - pivotRange(pivot).startRow + 1),
    charts: (workbook.charts ?? []).filter(chart => chart.sheetId === sheet.id && sheetClipboardBoundsContain(bounds, chartBounds(chart))),
    slicers: (workbook.slicers ?? []).flatMap(slicer => {
      if (slicer.hostSheetId !== sheet.id) return []
      const layout = resolveSlicerBounds(slicer, axes)
      // Collapsed controls have no visible contained footprint; their native models remain in the source.
      return layout.width > 0 && layout.height > 0 && sheetClipboardBoundsContain(bounds, layout) ? [{ ...slicer, layout }] : []
    }),
  }
  const controlCacheIds = new Set((resources.slicers ?? []).flatMap(view => view.cacheId ? [view.cacheId] : []))
  resources.slicerCaches = (workbook.slicerCaches ?? []).filter(cache => controlCacheIds.has(cache.id))
  const sourceCacheIds = new Set([...(resources.pivotTables ?? []).flatMap(pivot => pivot.pivotCacheId ? [pivot.pivotCacheId] : []),
    ...(resources.charts ?? []).flatMap(chart => chart.ownedPivot?.pivotCacheId ? [chart.ownedPivot.pivotCacheId] : []),
    ...resources.slicerCaches.flatMap(cache => cache.source.kind === 'pivot' ? [cache.source.pivotCacheId] : [])])
  resources.pivotCaches = (workbook.pivotCaches ?? []).filter(cache => sourceCacheIds.has(cache.id))
  return { resources: cloneSheetClipboardWorkbookResources(resources), origin: { left: bounds.left, top: bounds.top }, bounds }
}

function lowerBound(values: readonly number[], value: number): number {
  let low = 0, high = values.length
  while (low < high) { const middle = (low + high) >>> 1; if (values[middle]! < value) low = middle + 1; else high = middle }
  return low
}
function projectedRange(range: CellRange, snapshot: SheetClipboardSnapshot, geometry: SheetClipboardTargetGeometry, rowTile: number, colTile: number): CellRange {
  const top = lowerBound(snapshot.rows, range.startRow), bottom = lowerBound(snapshot.rows, range.endRow + 1) - 1
  const left = lowerBound(snapshot.columns, range.startCol), right = lowerBound(snapshot.columns, range.endCol + 1) - 1
  const startRow = geometry.rows[rowTile + top], endRow = geometry.rows[rowTile + bottom]
  const startCol = geometry.columns[colTile + left], endCol = geometry.columns[colTile + right]
  if (bottom < top || right < left || startRow === undefined || endRow === undefined || startCol === undefined || endCol === undefined
    || endRow - startRow !== bottom - top || endCol - startCol !== right - left) throw new RangeError('表格或透视对象需要连续的粘贴范围。')
  const next = { startRow, endRow, startCol, endCol }; assertSheetCommandRanges([next]); return next
}
function projectedCellKey(key: string, snapshot: SheetClipboardSnapshot, geometry: SheetClipboardTargetGeometry, rowTile: number, colTile: number): string {
  const match = /^(\d+):(\d+)$/.exec(key)
  if (!match) throw new TypeError('透视生成单元格坐标无效。')
  const row = Number(match[1]), col = Number(match[2]), rowIndex = lowerBound(snapshot.rows, row), colIndex = lowerBound(snapshot.columns, col)
  if (snapshot.rows[rowIndex] !== row || snapshot.columns[colIndex] !== col) throw new RangeError('剪贴板未包含完整透视输出。')
  const targetRow = geometry.rows[rowTile + rowIndex], targetCol = geometry.columns[colTile + colIndex]
  if (targetRow === undefined || targetCol === undefined) throw new RangeError('透视输出不能被截断。')
  return `${targetRow}:${targetCol}`
}
function projectedFilter(filter: SheetAutoFilter | null | undefined, snapshot: SheetClipboardSnapshot, geometry: SheetClipboardTargetGeometry, rowTile: number, colTile: number): SheetAutoFilter | null | undefined {
  if (!filter) return filter
  const columns: SheetAutoFilter['columns'] = {}
  for (const [column, criterion] of Object.entries(filter.columns)) {
    const index = lowerBound(snapshot.columns, Number(column)), destination = geometry.columns[colTile + index]
    if (snapshot.columns[index] === Number(column) && destination !== undefined) columns[destination] = cloneModel(criterion)
  }
  return { range: projectedRange(filter.range, snapshot, geometry, rowTile, colTile), columns, filteredRows: filter.filteredRows.flatMap(row => {
    const index = lowerBound(snapshot.rows, row), destination = geometry.rows[rowTile + index]
    return snapshot.rows[index] === row && destination !== undefined ? [destination] : []
  }) }
}
function movedFilter(filter: SheetAutoFilter | null | undefined, move: SheetClipboardMove): SheetAutoFilter | null | undefined {
  if (!filter) return filter
  const row = move.targetRange.startRow - move.sourceRange.startRow, col = move.targetRange.startCol - move.sourceRange.startCol
  return { range: movedRange(filter.range, move), columns: Object.fromEntries(Object.entries(filter.columns).map(([column, criterion]) => [Number(column) + col, cloneModel(criterion)])),
    filteredRows: filter.filteredRows.map(index => index + row) }
}
function relocatedSource(workbook: WorkbookModel, sheetId: string, range: CellRange, target: SheetClipboardLocation,
  snapshot: SheetClipboardSnapshot, geometry: SheetClipboardTargetGeometry, rowTile: number, colTile: number): { sheetId: string; range: CellRange } {
  if (snapshot.source && sheetId === snapshot.source.sheetId && contains(snapshot.source.range, range)) {
    return { sheetId: target.sheetId, range: projectedRange(range, snapshot, geometry, rowTile, colTile) }
  }
  if (snapshot.source?.unitId === workbook.id && workbook.sheets.some(sheet => sheet.id === sheetId)) return { sheetId, range: { ...range } }
  throw new Error('跨工作簿复制对象的数据源未包含在复制范围中；需要补齐该对象的数据源。')
}
function uniqueTableName(name: string, used: Set<string>): string {
  if (!used.has(name.toLowerCase())) { used.add(name.toLowerCase()); return name }
  let suffix = 2
  while (used.has(`${name}_${suffix}`.toLowerCase())) suffix++
  const next = `${name}_${suffix}`; used.add(next.toLowerCase()); return next
}
function assertNoTableOverlap(tables: readonly SheetTableModel[], candidate: SheetTableModel, ignored: ReadonlySet<string>): void {
  if (tables.some(table => !ignored.has(table.id) && table.id !== candidate.id && table.sheetId === candidate.sheetId && sheetRangesIntersect(table.range, candidate.range))) {
    throw new RangeError(`粘贴的表格“${candidate.name}”与目标表格重叠。`)
  }
}
function assertNoPivotOverlap(pivots: readonly PivotTableModel[], candidate: PivotTableModel, tables: readonly SheetTableModel[]): void {
  const range = pivotRange(candidate)
  if (candidate.source.subUnitId === candidate.anchor.subUnitId && sheetRangesIntersect(candidate.source.range, range)) throw new RangeError('透视输出不能覆盖其数据源。')
  if (pivots.some(pivot => pivot.id !== candidate.id && pivot.anchor.subUnitId === candidate.anchor.subUnitId && sheetRangesIntersect(pivotRange(pivot), range))
    || tables.some(table => table.sheetId === candidate.anchor.subUnitId && sheetRangesIntersect(table.range, range))) throw new RangeError('粘贴的透视输出与现有表格或透视对象重叠。')
}
function placedSlicer(workbook: WorkbookModel, slicer: SlicerModel, sheetId: string, bounds: SheetClipboardObjectBounds): SlicerModel {
  const anchors = sheetClipboardBoundsDrawing(workbook, sheetId, bounds)
  return { ...slicer, hostSheetId: sheetId, position: { row: anchors.from.row, column: anchors.from.column }, layout: { ...bounds },
    ...(slicer.drawingGeometry ? { drawingGeometry: slicerGeometryWithBounds(slicer.drawingGeometry, bounds) } : {}),
    drawing: { ...slicer.drawing, anchorType: slicer.drawing?.anchorType ?? 'position', ...anchors } }
}

/** Structured table names are identifiers; quoted text and bracketed column captions remain data. */
export function remapSheetClipboardTableFormula(formula: string, names: ReadonlyMap<string, string>): string {
  let result = '', cursor = 0
  while (cursor < formula.length) {
    const start = cursor, character = formula[cursor]
    if (character === '"') {
      cursor++
      while (cursor < formula.length) { if (formula[cursor++] !== '"') continue; if (formula[cursor] === '"') { cursor++; continue } break }
    } else if (character === '[') {
      let depth = 1; cursor++
      while (cursor < formula.length && depth) { if (formula[cursor] === '[') depth++; else if (formula[cursor] === ']') depth--; cursor++ }
    } else {
      const identifier = /^[\p{L}_\\][\p{L}\p{N}_.\\]*/u.exec(formula.slice(cursor))
      if (identifier) {
        cursor += identifier[0].length
        const mapped = formula[cursor] === '[' ? names.get(identifier[0].toLowerCase()) : undefined
        if (mapped) { result += mapped; continue }
      } else cursor++
    }
    result += formula.slice(start, cursor)
  }
  return result
}

/** Modify only a detached candidate. Identity-preserving moves and identity-remapped copies share one resource transaction. */
export function pasteSheetClipboardWorkbookResources(
  original: WorkbookModel, candidate: WorkbookModel, target: SheetClipboardLocation, snapshot: SheetClipboardSnapshot,
  geometry: SheetClipboardTargetGeometry, mode: SheetPasteMode, move?: SheetClipboardMove,
): SheetClipboardWorkbookResources | undefined {
  const content = mode !== 'formats' && mode !== 'columnWidths'
  const copied = snapshot.workbookResources
  const carryObjects = (mode === 'all' || mode === 'exceptBorders') && !snapshot.externalFormat && !geometry.contentOnlyMerge
  const stationaryPivotIds = new Set(!move && carryObjects && snapshot.source?.unitId === original.id
    && snapshot.source.sheetId === target.sheetId && geometry.rows.length === snapshot.rows.length
    && geometry.columns.length === snapshot.columns.length && geometry.rows.every((row, index) => row === snapshot.rows[index])
    && geometry.columns.every((column, index) => column === snapshot.columns[index])
    ? (copied?.pivotTables ?? []).filter(pivot => original.pivotTables?.some(current => current.id === pivot.id
      && JSON.stringify(current) === JSON.stringify(pivot))).map(pivot => pivot.id) : [])
  const movingPivots = new Set((original.pivotTables ?? []).filter(pivot => move && pivot.anchor.subUnitId === move.sourceSheetId
    && contains(move.sourceRange, pivotRange(pivot))).map(pivot => pivot.id))
  if (content) for (const pivot of original.pivotTables ?? []) {
    if (move && pivot.anchor.subUnitId === move.sourceSheetId && sheetRangesIntersect(move.sourceRange, pivotRange(pivot)) && !movingPivots.has(pivot.id)) {
      throw new RangeError('剪切需要包含完整透视输出。')
    }
    if (pivot.anchor.subUnitId === target.sheetId && !movingPivots.has(pivot.id) && !stationaryPivotIds.has(pivot.id)
      && geometry.ranges.some(range => sheetRangesIntersect(range, pivotRange(pivot)))) {
      throw new RangeError('不能直接覆盖透视输出，请先移动透视对象或修改数据源。')
    }
  }
  const resources = cloneSheetClipboardWorkbookResources(original)
  const pivotMoveTransform: PivotSourceStructureTransform | undefined = move ? {
    sheetId: move.sourceSheetId,
    range: range => contains(move.sourceRange, range) ? movedRange(range, move) : { ...range },
    column: (column, range) => contains(move.sourceRange, range) ? column + move.targetRange.startCol - move.sourceRange.startCol : column,
    sourceWorksheet: range => contains(move.sourceRange, range) ? { subUnitId: move.targetSheetId, sheetName: move.targetSheetName } : undefined,
    sourceDirty: range => contains(move.sourceRange, range),
    point: (row, col) => sheetRangeContainsCell(move.sourceRange, row, col)
      ? { row: row + move.targetRange.startRow - move.sourceRange.startRow, col: col + move.targetRange.startCol - move.sourceRange.startCol }
      : { row, col },
    anchorWorksheet: (row, col) => sheetRangeContainsCell(move.sourceRange, row, col) ? move.targetSheetId : undefined,
  } : undefined
  if (move) {
    const movingTableIds = new Set((resources.tables ?? []).filter(table => table.sheetId === move.sourceSheetId && contains(move.sourceRange, table.range)).map(table => table.id))
    if (movingTableIds.size && resources.tables?.some(table => !movingTableIds.has(table.id) && table.sheetId === target.sheetId
      && geometry.ranges.some(range => sheetRangesIntersect(range, table.range)))) throw new RangeError('移动表格的目标范围与现有表格重叠。')
    resources.tables = resources.tables?.map(table => {
      if (!movingTableIds.has(table.id)) return table
      const next = { ...table, sheetId: target.sheetId, range: movedRange(table.range, move), autoFilter: movedFilter(table.autoFilter, move) }
      assertNoTableOverlap(resources.tables ?? [], next, movingTableIds); return next
    })
    const pivotModels: WorkbookModel = { ...candidate, ...resources }
    transformWorkbookPivotSourceReports(original, pivotModels, pivotMoveTransform!)
    resources.pivotTables = pivotModels.pivotTables; resources.charts = pivotModels.charts
    const movingCharts = new Set((copied?.charts ?? []).map(chart => chart.id)), movingSlicers = new Set((copied?.slicers ?? []).map(slicer => slicer.id))
    const sourceSheet = original.sheets.find(sheet => sheet.id === move.sourceSheetId)!, targetSheet = candidate.sheets.find(sheet => sheet.id === target.sheetId)!
    const beforeBounds = sheetClipboardRangeBounds(original, sourceSheet, move.sourceRange), afterBounds = sheetClipboardRangeBounds(candidate, targetSheet, move.targetRange)
    const leftDelta = afterBounds.left - beforeBounds.left, topDelta = afterBounds.top - beforeBounds.top
    resources.charts = resources.charts?.map(chart => {
      const next = cloneModel(chart), sourceSheetId = chart.source.sheetId ?? chart.sheetId
      if (!chart.pivotTableId && !chart.ownedPivot && sourceSheetId === move.sourceSheetId && contains(move.sourceRange, chart.source.range)) {
        next.source = { ...chart.source, sheetId: move.targetSheetId, range: movedRange(chart.source.range, move) }
      }
      if (movingCharts.has(chart.id)) {
        next.sheetId = move.targetSheetId; next.position = { x: chart.position.x + leftDelta, y: chart.position.y + topDelta }
        if (next.drawingGeometry) { next.drawingGeometry.matrix.e += leftDelta; next.drawingGeometry.matrix.f += topDelta }
      }
      if (chart.pivotTableId) {
        const pivot = resources.pivotTables?.find(item => item.id === chart.pivotTableId)
        if (pivot?.outputRange) next.source = { ...next.source, sheetId: pivot.anchor.subUnitId, range: { ...pivot.outputRange } }
      }
      if (chart.ownedPivot && next.ownedPivot) {
        if (movingCharts.has(chart.id)) next.ownedPivot.anchor.subUnitId = move.targetSheetId
        const previous = original.charts?.find(previous => previous.id === chart.id)?.ownedPivot
        if (previous && JSON.stringify(previous) !== JSON.stringify(next.ownedPivot)) next.ownedPivot.revision = previous.revision + 1
      }
      const previous = original.charts?.find(previous => previous.id === chart.id)
      if (previous && JSON.stringify(previous) !== JSON.stringify(next)) next.revision = previous.revision + 1
      return next
    })
    resources.slicers = resources.slicers?.map(slicer => {
      if (!movingSlicers.has(slicer.id)) return slicer
      const bounds = resolveSlicerBounds(slicer, sheetClipboardObjectAxes(original, sourceSheet))
      return { ...placedSlicer(candidate, slicer, move.targetSheetId, { ...bounds, left: bounds.left + leftDelta, top: bounds.top + topDelta }), revision: slicer.revision + 1 }
    })
  } else if (carryObjects && copied) {
    const names = new Set((resources.tables ?? []).map(table => table.name.toLowerCase()))
    const sourceResources: WorkbookModel = { ...original, ...copied, id: snapshot.source?.unitId ?? original.id }
    for (let rowTile = 0; rowTile < geometry.rows.length; rowTile += snapshot.rows.length) for (let colTile = 0; colTile < geometry.columns.length; colTile += snapshot.columns.length) {
      const tableIds = new Map<string, { id: string; columns: Map<string, string> }>(), pivotIds = new Map<string, string>(), formulaNames = new Map<string, string>()
      const pivotCacheIds = new Map<string, string>(), slicerCacheIds = new Map<string, string>(), chartIds = new Map<string, string>()
      for (const table of copied.tables ?? []) {
        const range = projectedRange(table.range, snapshot, geometry, rowTile, colTile)
        if (snapshot.source?.unitId === candidate.id && table.sheetId === target.sheetId && sheetRangesEqual(range, table.range)) {
          tableIds.set(table.id, { id: table.id, columns: new Map(table.columns.map(column => [column.id, column.id])) }); continue
        }
        const next = { ...cloneModel(table), id: newId('table'), name: uniqueTableName(table.name, names), sheetId: target.sheetId, range,
          columns: table.columns.map(column => ({ ...column, id: newId('column') })), autoFilter: projectedFilter(table.autoFilter, snapshot, geometry, rowTile, colTile) }
        assertNoTableOverlap(resources.tables ?? [], next, new Set())
        tableIds.set(table.id, { id: next.id, columns: new Map(table.columns.map((column, index) => [column.id, next.columns[index]!.id])) })
        formulaNames.set(table.name.toLowerCase(), next.name)
        resources.tables ??= []; resources.tables.push(next)
      }
      for (const pivot of copied.pivotTables ?? []) {
        const range = projectedRange(pivotRange(pivot), snapshot, geometry, rowTile, colTile)
        if (stationaryPivotIds.has(pivot.id) && sheetRangesEqual(range, pivotRange(pivot))) {
          pivotIds.set(pivot.id, pivot.id); continue
        }
        const source = relocatedSource(candidate, pivot.source.subUnitId, pivot.source.range, target, snapshot, geometry, rowTile, colTile)
        const next = { ...cloneModel(pivot), id: newId('pivot'), source: { ...cloneModel(pivot.source), unitId: candidate.id!, subUnitId: source.sheetId,
          sheetName: candidate.sheets.find(sheet => sheet.id === source.sheetId)!.name, range: source.range },
          anchor: { unitId: candidate.id!, subUnitId: target.sheetId, row: range.startRow, col: range.startCol },
          fields: pivot.fields.map(field => ({ ...cloneModel(field), sourceColumn: source.range.startCol + field.sourceColumn - pivot.source.range.startCol })),
          ...(pivot.outputRange ? { outputRange: range } : {}), generatedCellKeys: pivot.generatedCellKeys?.map(key => projectedCellKey(key, snapshot, geometry, rowTile, colTile)),
          drillDownMap: pivot.drillDownMap ? Object.fromEntries(Object.entries(pivot.drillDownMap).map(([key, entry]) => [projectedCellKey(key, snapshot, geometry, rowTile, colTile), cloneModel(entry)])) : undefined,
          revision: 0 }
        const reportModels: WorkbookModel = { ...candidate, ...resources }
        copyPivotReportCache(sourceResources, reportModels, pivot, next, pivotCacheIds)
        prepareCopiedPivotSourceReport(next)
        resources.pivotCaches = reportModels.pivotCaches
        assertNoPivotOverlap(resources.pivotTables ?? [], next, resources.tables ?? [])
        pivotIds.set(pivot.id, next.id); resources.pivotTables ??= []; resources.pivotTables.push(next)
      }
      const targetSheet = candidate.sheets.find(sheet => sheet.id === target.sheetId)!, corner = sheetClipboardRangeBounds(candidate, targetSheet, {
        startRow: geometry.rows[rowTile]!, endRow: geometry.rows[rowTile]!, startCol: geometry.columns[colTile]!, endCol: geometry.columns[colTile]!,
      })
      const origin = snapshot.sourceOrigin
      if (!origin && ((copied.charts?.length ?? 0) > 0 || (copied.slicers?.length ?? 0) > 0)) throw new TypeError('剪贴板绘图缺少源位置。')
      const leftDelta = corner.left - (origin?.left ?? corner.left), topDelta = corner.top - (origin?.top ?? corner.top)
      for (const chart of copied.charts ?? []) {
        const next = { ...cloneModel(chart), id: newId('chart'), sheetId: target.sheetId,
          position: { x: chart.position.x + leftDelta, y: chart.position.y + topDelta }, revision: 0 }
        chartIds.set(chart.id, next.id)
        if (next.drawingGeometry) { next.drawingGeometry.matrix.e += leftDelta; next.drawingGeometry.matrix.f += topDelta }
        if (chart.pivotTableId) {
          next.pivotTableId = pivotIds.get(chart.pivotTableId) ?? (snapshot.source?.unitId === candidate.id ? chart.pivotTableId : undefined)
          const pivot = resources.pivotTables?.find(item => item.id === next.pivotTableId)
          if (!pivot) throw new Error('剪贴板透视图缺少其透视数据源。')
          next.source = { ...chart.source, sheetId: pivot.anchor.subUnitId, range: { ...pivotRange(pivot) } }
        } else if (!chart.ownedPivot) {
          const source = relocatedSource(candidate, chart.source.sheetId ?? chart.sheetId, chart.source.range, target, snapshot, geometry, rowTile, colTile)
          next.source = { ...chart.source, sheetId: source.sheetId, range: source.range }
        }
        if (chart.ownedPivot && next.ownedPivot) {
          if (chart.pivotTableId) throw new TypeError('剪贴板透视图不能同时持有两种报告。')
          const source = relocatedSource(candidate, chart.ownedPivot.source.subUnitId, chart.ownedPivot.source.range, target, snapshot, geometry, rowTile, colTile)
          next.ownedPivot = { ...next.ownedPivot, id: newId('owned-pivot'), revision: 0,
            anchor: { ...next.ownedPivot.anchor, unitId: candidate.id!, subUnitId: target.sheetId },
            source: { ...next.ownedPivot.source, unitId: candidate.id!, subUnitId: source.sheetId,
              sheetName: candidate.sheets.find(sheet => sheet.id === source.sheetId)!.name, range: source.range },
            fields: next.ownedPivot.fields.map(field => ({ ...field,
              sourceColumn: source.range.startCol + field.sourceColumn - chart.ownedPivot!.source.range.startCol })),
          }
          const reportModels: WorkbookModel = { ...candidate, ...resources }
          copyPivotReportCache(sourceResources, reportModels, chart.ownedPivot, next.ownedPivot, pivotCacheIds)
          prepareCopiedPivotSourceReport(next.ownedPivot, true)
          next.source = { ...next.source, sheetId: next.ownedPivot.source.subUnitId, range: cloneModel(next.ownedPivot.source.range) }
          resources.pivotCaches = reportModels.pivotCaches
        }
        sheetClipboardObjectRange(candidate, next.sheetId, chartBounds(next))
        resources.charts ??= []; resources.charts.push(next)
      }
      for (const slicer of copied.slicers ?? []) {
        const existingSheet = original.sheets.find(sheet => sheet.id === slicer.hostSheetId)
        const bounds = slicer.layout ?? (existingSheet ? resolveSlicerBounds(slicer, sheetClipboardObjectAxes(original, existingSheet)) : undefined)
        if (!bounds) throw new TypeError('剪贴板切片器缺少绘图布局。')
        const next = placedSlicer(candidate, { ...cloneModel(slicer), id: newId('slicer'), unitId: candidate.id!, revision: 0 }, target.sheetId,
          { ...bounds, left: bounds.left + leftDelta, top: bounds.top + topDelta })
        const controlModels: WorkbookModel = { ...candidate, ...resources }
        copySlicerViewCache(sourceResources, controlModels, slicer, next, { tableIds, pivotTableIds: pivotIds,
          pivotChartIds: chartIds, pivotCacheIds, slicerCacheIds })
        resources.slicerCaches = controlModels.slicerCaches
        resources.pivotCaches = controlModels.pivotCaches
        resources.slicers ??= []; resources.slicers.push(next)
      }
      if (formulaNames.size) for (let row = rowTile; row < rowTile + snapshot.rows.length; row++) for (let col = colTile; col < colTile + snapshot.columns.length; col++) {
        const cell = targetSheet.cells[`${geometry.rows[row]}:${geometry.columns[col]}`]
        if (cell?.formula) rewriteFormulaImportCell(cell, remapSheetClipboardTableFormula(cell.formula, formulaNames))
      }
      for (const entries of [resources.tables, resources.pivotTables, resources.pivotCaches, resources.charts, resources.slicers, resources.slicerCaches]) if ((entries?.length ?? 0) > RESOURCE_BUDGET) {
        throw new RangeError('粘贴的工作簿对象数量超出当前同步预算。')
      }
    }
  }
  if (move) for (const pivot of resources.pivotTables ?? []) {
    const previous = original.pivotTables?.find(item => item.id === pivot.id)
    if (JSON.stringify(previous) !== JSON.stringify(pivot)) assertNoPivotOverlap(resources.pivotTables ?? [], pivot, resources.tables ?? [])
  }
  if (move || content) {
    const models: WorkbookModel = { ...candidate, ...resources }
    if (content) {
      const targetIndex = models.sheets.findIndex(sheet => sheet.id === target.sheetId)
      if (models.sheets[targetIndex] === original.sheets.find(sheet => sheet.id === target.sheetId)) {
        models.sheets[targetIndex] = cloneModel(models.sheets[targetIndex]!)
      }
      const targetSheet = models.sheets[targetIndex]!
      // Adopt copied generated cells before rebuilding: a smaller matrix must retire its stale tail too.
      for (const row of geometry.rows) for (const column of geometry.columns) {
        const cell = targetSheet.cells[`${row}:${column}`]
        if (!cell?.metadata?.pivotTableId) continue
        const owner = models.pivotTables?.find(pivot => pivot.anchor.subUnitId === target.sheetId && contains(pivotRange(pivot), {
          startRow: row, endRow: row, startCol: column, endCol: column,
        }))
        const metadata = { ...cell.metadata }
        if (owner) { metadata.pivotTableId = owner.id; metadata.pivotGenerated = true }
        else { delete metadata.pivotTableId; delete metadata.pivotGenerated }
        cell.metadata = Object.keys(metadata).length ? metadata : undefined
      }
    }
    if (move) reconcileSlicerConnectionResources(original, models, pivotMoveTransform)
    if (content) invalidatePivotSourceContentRanges(original, models, [
      ...geometry.ranges.map(range => ({ sheetId: target.sheetId, range })),
      ...(move ? [{ sheetId: move.sourceSheetId, range: move.sourceRange }] : []),
    ])
    refreshPivotSourceWorksheetOutputs(original, models, (resources.pivotTables ?? [])
      .filter(report => !original.pivotTables?.some(previous => previous.id === report.id)).map(report => report.id))
    synchronizeLinkedPivotChartReports(original, models)
    Object.assign(resources, cloneSheetClipboardWorkbookResources(models))
  }
  const changed = JSON.stringify(cloneSheetClipboardWorkbookResources(original)) !== JSON.stringify(resources)
  if (changed) {
    const updated: SheetClipboardWorkbookResources = {}
    for (const key of ['tables', 'pivotTables', 'pivotCaches', 'charts', 'slicers', 'slicerCaches'] as const) {
      const previous = new Map((original[key] ?? []).map(item => [item.id, JSON.stringify(item)]))
      Object.assign(updated, { [key]: resources[key]?.filter(item => previous.get(item.id) !== JSON.stringify(item)) })
    }
    assertSheetClipboardWorkbookResources(updated)
    assertSheetClipboardResourceBindings(updated, resources)
    Object.assign(candidate, resources)
  }
  return changed ? resources : undefined
}

export function sheetClipboardWorkbookResourceChangedRanges(before: WorkbookModel, after: WorkbookModel): Record<string, CellRange[]> {
  const result: Record<string, CellRange[]> = Object.create(null)
  const add = (sheetId: string, range: CellRange) => { const ranges = result[sheetId] ?? []; ranges.push({ ...range }); result[sheetId] = ranges }
  const track = <T extends { id: string }>(originalItems: readonly T[], nextItems: readonly T[],
    footprint: (workbook: WorkbookModel, item: T) => { sheetId: string; range: CellRange } | undefined): void => {
    const originals = new Map(originalItems.map(item => [item.id, item])), next = new Map(nextItems.map(item => [item.id, item]))
    const changedIds = new Set([...originals.keys(), ...next.keys()].filter(id => JSON.stringify(originals.get(id)) !== JSON.stringify(next.get(id))))
    for (const { workbook, items } of [{ workbook: before, items: originalItems }, { workbook: after, items: nextItems }]) for (const item of items) {
      if (!changedIds.has(item.id)) continue
      const location = footprint(workbook, item)
      if (location) add(location.sheetId, location.range)
    }
  }
  track(before.tables ?? [], after.tables ?? [], (_workbook, table) => ({ sheetId: table.sheetId, range: table.range }))
  track(before.pivotTables ?? [], after.pivotTables ?? [], (_workbook, pivot) => ({ sheetId: pivot.anchor.subUnitId, range: pivotRange(pivot) }))
  track(before.charts ?? [], after.charts ?? [], (workbook, chart) => ({ sheetId: chart.sheetId,
    range: sheetClipboardObjectRange(workbook, chart.sheetId, chartBounds(chart)) }))
  track(before.slicers ?? [], after.slicers ?? [], (workbook, slicer) => {
    const sheet = workbook.sheets.find(sheet => sheet.id === slicer.hostSheetId)
    return sheet ? { sheetId: sheet.id, range: sheetClipboardObjectRange(workbook, sheet.id, resolveSlicerBounds(slicer, sheetClipboardObjectAxes(workbook, sheet))) } : undefined
  })
  for (const sheet of before.sheets) {
    const next = after.sheets.find(item => item.id === sheet.id)
    if (!next) continue
    track(sheet.images ?? [], next.images ?? [], (workbook, image) => {
      const owner = workbook.sheets.find(item => item.id === sheet.id)
      return owner ? { sheetId: owner.id, range: sheetClipboardObjectRange(workbook, owner.id, sheetImageClipboardBounds(workbook, owner, image)) } : undefined
    })
  }
  return result
}
