import type {
  CellScalar, ChartModel, PivotCacheModel, PivotFieldModel, PivotRecordSelection, PivotTableModel, PivotTableSource, SheetFilterCriteria, SheetSlicerConnections,
  SlicerCacheModel, SlicerCacheSource, SlicerConnectionCandidate, SlicerModel, SpreadsheetCell, SpreadsheetSheet, WorkbookModel,
} from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { equivalentModelData } from '../../platform/commands/model-equivalence.ts'
import { getPivotSourceRecordSelection, getPivotSourceSelectedItems } from '../pivot/services/pivot-engine.ts'
import { normalizeSlicerRecordSelection, slicerRecordSelectionFromItems, slicerRecordSelectionFromNativeValues, slicerRecordSelectionItems } from './slicer-record-selection.ts'
import { sheetCellRecordValue } from '../sheet-workbench/sheet-record-value.ts'
import type { RecordValue } from '../../domain/record-value.ts'
import { normalizeSlicerCacheOptions } from './slicer-cache-options.ts'
import { normalizeRetainedTimelineCache, readTimelineReportFilterState } from './timeline-state.ts'
import { pivotSourceFieldIdentity } from '../pivot/services/pivot-source-field.ts'
import type { TimelineDateFilterState } from '../../domain/timeline-filters.ts'

export type SlicerConnectionErrorCode = 'invalid-resource' | 'invalid-target' | 'source-missing' | 'field-missing' | 'permission-denied'

/** Ordinary business refusals are distinguishable from command/listener failures. */
export class SlicerConnectionError extends Error {
  constructor(readonly code: SlicerConnectionErrorCode, message: string) { super(message); this.name = 'SlicerConnectionError' }
}

export interface SlicerConnectionReport {
  kind: 'pivotTable' | 'ownedPivotChart'
  subUnitId: string
  id: string
  pivot: PivotTableModel
  field: PivotFieldModel
  chart?: ChartModel
}

export interface SlicerConnectionPlan {
  cacheId: string
  connections: SheetSlicerConnections
  reports: SlicerConnectionReport[]
  removedReports: SlicerConnectionReport[]
  changed: boolean
}

export interface SlicerResourceDiagnostic {
  kind: 'pivotReport' | 'slicerView'
  id: string
  code: SlicerConnectionErrorCode
  message: string
}

/** Exact source identity is a legacy fallback. Authored/imported IDs never coalesce by range. */
export function resolvePivotCacheId(workbook: WorkbookModel, pivot: PivotTableModel): string {
  assertSource(workbook, pivot.source)
  const authored = pivot.pivotCacheId ?? pivot.source.pivotCacheId
  if (authored !== undefined) return nonemptyId(authored, '透视缓存标识无效。')
  return `pivot-cache:${JSON.stringify(sourceKey(pivot.source))}`
}

export function ensurePivotCache(workbook: WorkbookModel, pivot: PivotTableModel): PivotCacheModel {
  const id = resolvePivotCacheId(workbook, pivot)
  const current = workbook.pivotCaches?.find(cache => cache.id === id)
  if (current && !samePivotSource(current.source, pivot.source)) {
    throw new SlicerConnectionError('invalid-resource', '同一透视缓存的报告引用了不同数据源。')
  }
  const fields = new Map((current?.fields ?? []).map(field => [field.id, { ...field }]))
  for (const field of pivot.fields) {
    const fieldId = pivotSourceFieldIdentity(pivot, field, current), previous = fields.get(fieldId)
    if (previous && previous.sourceColumn !== field.sourceColumn) throw new SlicerConnectionError('invalid-resource', '同一透视缓存字段引用了不同来源列。')
    fields.set(fieldId, { id: fieldId, sourceColumn: field.sourceColumn, name: field.name })
    field.sourceFieldId = fieldId
  }
  const occupiedColumns = new Set([...fields.values()].map(field => field.sourceColumn))
  let nextFieldId = current?.nextFieldId ?? 0
  for (let column = pivot.source.range.startCol; column <= pivot.source.range.endCol; column++) {
    if (occupiedColumns.has(column)) continue
    let fieldId = current ? `source-field:${++nextFieldId}` : String(column - pivot.source.range.startCol)
    while (fields.has(fieldId)) fieldId = `source-field:${++nextFieldId}`
    const sheet = workbook.sheets.find(candidate => candidate.id === pivot.source.subUnitId)!, header = sheet.cells[`${pivot.source.range.startRow}:${column}`]?.value
    fields.set(fieldId, { id: fieldId, sourceColumn: column, name: header == null ? undefined : String(header) })
    occupiedColumns.add(column)
  }
  if (current) { current.fields = [...fields.values()]; current.nextFieldId = nextFieldId; pivot.pivotCacheId = id; return current }
  const cache: PivotCacheModel = { id, source: cloneModel(pivot.source), fields: [...fields.values()], nextFieldId, revision: 0 }
  workbook.pivotCaches ??= []
  workbook.pivotCaches.push(cache)
  pivot.pivotCacheId = id
  return cache
}

export function samePivotSource(first: PivotTableSource, second: PivotTableSource): boolean {
  return equivalentModelData(sourceKey(first), sourceKey(second))
}

/** A source field is not the report-local instance ID (copies may have their own IDs). */
export function sourceFieldId(pivot: PivotTableModel, field: PivotFieldModel): string {
  return field.sourceFieldId ?? String(field.sourceColumn - pivot.source.range.startCol)
}

export function findSlicerReportField(pivot: PivotTableModel, id: string): PivotFieldModel | undefined {
  return pivot.fields.find(field => field.area !== 'value' && sourceFieldId(pivot, field) === id)
}

export function getSlicerCache(workbook: WorkbookModel, slicer: SlicerModel): SlicerCacheModel {
  assertControl(workbook, slicer)
  if (slicer.cacheId) {
    const cache = workbook.slicerCaches?.find(candidate => candidate.id === slicer.cacheId)
    if (!cache || cache.unitId !== workbook.id || cache.type !== slicer.type) {
      throw new SlicerConnectionError('invalid-resource', '切片器共享缓存不存在或不属于当前工作簿。')
    }
    assertCacheSource(workbook, cache.source)
    if (cache.source.kind === 'pivot' && cache.source.sourceFieldId !== slicer.sourceFieldId
      || cache.source.kind === 'table' && cache.source.columnId !== slicer.sourceFieldId) {
      throw new SlicerConnectionError('invalid-resource', '切片器源字段与共享缓存不一致。')
    }
    return cache
  }
  const legacy = legacyCache(workbook, slicer)
  const shared = workbook.slicerCaches?.find(cache => cache.id === legacy.id)
  if (shared && (shared.unitId !== legacy.unitId || shared.type !== legacy.type || !equivalentModelData(shared.source, legacy.source))) {
    throw new SlicerConnectionError('invalid-resource', '切片器缓存标识冲突。')
  }
  return shared ?? legacy
}

/** Used by mutations/normalization only. Read-only Facades never materialize resources. */
export function ensureSlicerCache(workbook: WorkbookModel, slicer: SlicerModel): SlicerCacheModel {
  const seed = getSlicerCache(workbook, slicer)
  if (slicer.cacheId) {
    if (seed.type === 'timeline') normalizeRetainedTimelineCache(seed, workbook.date1904)
    mirrorSlicerCache(slicer, seed)
    return seed
  }
  const existing = workbook.slicerCaches?.find(cache => cache.id === seed.id)
  if (existing && (existing.unitId !== seed.unitId || existing.type !== seed.type || !equivalentModelData(existing.source, seed.source))) {
    throw new SlicerConnectionError('invalid-resource', '切片器缓存标识冲突。')
  }
  if (seed.source.kind === 'pivot') {
    const report = legacyPivotReport(workbook, slicer)
    if (report) ensurePivotCache(workbook, report.pivot)
  }
  workbook.slicerCaches ??= []
  if (!existing) workbook.slicerCaches.push(seed)
  slicer.cacheId = seed.id
  const cache = existing ?? seed
  if (cache.type === 'timeline') normalizeRetainedTimelineCache(cache, workbook.date1904)
  slicer.sourceFieldId = cache.source.kind === 'pivot' ? cache.source.sourceFieldId : cache.source.columnId
  mirrorSlicerCache(slicer, cache)
  return cache
}

function mirrorSlicerCache(slicer: SlicerModel, cache: SlicerCacheModel): void {
  slicer.selectedItems = cache.selectedItems ? [...cache.selectedItems] : undefined
  slicer.timelineRange = cache.timelineRange ? cloneModel(cache.timelineRange) : undefined
  slicer.timelineDateRange = cache.timelineDateRange ? cloneModel(cache.timelineDateRange) : undefined
  slicer.timelineDateCondition = cache.timelineDateCondition ? cloneModel(cache.timelineDateCondition) : undefined
}

/** Retained unavailable sources are preserved. Their live operations fail explicitly. */
export function normalizeSlicerCaches(workbook: WorkbookModel): SlicerResourceDiagnostic[] {
  const diagnostics: SlicerResourceDiagnostic[] = []
  for (const pivot of [...(workbook.pivotTables ?? []), ...(workbook.charts ?? []).flatMap(chart => chart.ownedPivot ? [chart.ownedPivot] : [])]) {
    try {
      ensurePivotCache(workbook, pivot)
      for (const field of pivot.fields) field.sourceFieldId ??= sourceFieldId(pivot, field)
    } catch (error) {
      if (!(error instanceof SlicerConnectionError)) throw error
      diagnostics.push({ kind: 'pivotReport', id: pivot.id, code: error.code, message: error.message })
    }
  }
  for (const slicer of workbook.slicers ?? []) {
    try {
      const retained = slicer.cacheId ? workbook.slicerCaches?.find(cache => cache.id === slicer.cacheId && cache.type === slicer.type
        && cache.unitId === workbook.id && cache.source.kind === 'pivot' && cache.source.sourceFieldId === slicer.sourceFieldId) : undefined
      if (retained?.type === 'timeline') {
        // Retained date authoring survives a missing source or removed connected report.
        normalizeRetainedTimelineCache(retained, workbook.date1904)
        mirrorSlicerCache(slicer, retained)
      }
      const cache = ensureSlicerCache(workbook, slicer)
      const options = normalizeSlicerCacheOptions(cache.options)
      if (Object.keys(options).length) cache.options = options
      else delete cache.options
      if (cache.source.kind === 'pivot') getSlicerConnectionReports(workbook, slicer)
      normalizeSlicerCacheSelection(workbook, slicer, cache)
      mirrorSlicerCache(slicer, cache)
    } catch (error) {
      if (!(error instanceof SlicerConnectionError || error instanceof TypeError || error instanceof RangeError)) throw error
      diagnostics.push({ kind: 'slicerView', id: slicer.id, code: error instanceof SlicerConnectionError ? error.code : 'invalid-resource', message: error.message })
    }
  }
  if (diagnostics.length) workbook.metadata = { ...workbook.metadata, slicerResourceDiagnostics: diagnostics }
  else if (workbook.metadata && Object.hasOwn(workbook.metadata, 'slicerResourceDiagnostics')) {
    const { slicerResourceDiagnostics: _previous, ...metadata } = workbook.metadata
    workbook.metadata = metadata
  }
  return diagnostics
}

export function getSlicerConnections(workbook: WorkbookModel, slicer: SlicerModel): SheetSlicerConnections | undefined {
  const cache = getSlicerCache(workbook, slicer)
  return cache.source.kind === 'table' ? undefined : cloneModel(cache.connections ?? { pivotTables: [] })
}

export function getSlicerPivotSource(workbook: WorkbookModel, slicer: SlicerModel): PivotTableSource {
  const cache = getSlicerCache(workbook, slicer)
  if (cache.source.kind !== 'pivot') throw new SlicerConnectionError('invalid-target', '表格切片器没有透视缓存。')
  return pivotCacheSource(workbook, cache.source.pivotCacheId)
}

export function getSlicerSourceColumn(workbook: WorkbookModel, slicer: SlicerModel): number {
  const cache = getSlicerCache(workbook, slicer)
  if (cache.source.kind !== 'pivot') throw new SlicerConnectionError('invalid-target', '该切片器不是透视来源。')
  const source = cache.source, retained = workbook.pivotCaches?.find(candidate => candidate.id === source.pivotCacheId)
  const field = retained?.fields?.find(candidate => candidate.id === source.sourceFieldId)
  if (field) return field.sourceColumn
  if (retained?.fields) throw new SlicerConnectionError('field-missing', '切片器源字段已从缓存中删除。')
  if (/^\d+$/.test(source.sourceFieldId)) return pivotCacheSource(workbook, source.pivotCacheId).range.startCol + Number(source.sourceFieldId)
  const report = getSlicerConnectionReports(workbook, slicer)[0]
  if (report) return report.field.sourceColumn
  throw new SlicerConnectionError('field-missing', '切片器源字段没有保留列元数据。')
}

export interface SlicerSourceCellRange { sheet: SpreadsheetSheet; column: number; startRow: number; endRow: number }
export interface SlicerSourceRecord { row: number; value: CellScalar; recordValue: RecordValue; cell: SpreadsheetCell | undefined }

export function getSlicerSourceCellRange(workbook: WorkbookModel, slicer: SlicerModel): SlicerSourceCellRange {
  const cache = getSlicerCache(workbook, slicer)
  let sheetId: string, column: number, startRow: number, endRow: number
  if (cache.source.kind === 'table') {
    const source = cache.source, table = workbook.tables?.find(candidate => candidate.id === source.tableId)
    if (!table) throw new SlicerConnectionError('source-missing', '切片器表格来源不存在。')
    const index = table.columns.findIndex(candidate => candidate.id === source.columnId)
    if (index < 0) throw new SlicerConnectionError('field-missing', '切片器表格源字段不存在。')
    sheetId = table.sheetId; column = table.range.startCol + index
    startRow = table.range.startRow + (table.showHeaderRow ? 1 : 0)
    endRow = table.range.endRow - (table.showTotalsRow ? 1 : 0)
  } else {
    const source = getSlicerPivotSource(workbook, slicer)
    sheetId = source.subUnitId; column = getSlicerSourceColumn(workbook, slicer)
    startRow = source.range.startRow + 1; endRow = source.range.endRow
  }
  const sheet = workbook.sheets.find(candidate => candidate.id === sheetId)
  if (!sheet) throw new SlicerConnectionError('source-missing', '切片器数据源工作表不存在。')
  return { sheet, column, startRow, endRow }
}

/** The same source rows provide raw legacy values and distinct typed Date/Error identities. */
export function getSlicerSourceRecords(workbook: WorkbookModel, slicer: SlicerModel): SlicerSourceRecord[] {
  const { sheet, column, startRow, endRow } = getSlicerSourceCellRange(workbook, slicer), records: SlicerSourceRecord[] = []
  for (let row = startRow; row <= endRow; row++) {
    const cell = sheet.cells[`${row}:${column}`]
    records.push({ row, cell, value: cell?.value ?? null, recordValue: sheetCellRecordValue(cell, workbook.date1904) })
  }
  return records
}

/** Source raw values preserve scalar identity; labels are derived only by presentation consumers. */
export function getSlicerSourceValues(workbook: WorkbookModel, slicer: SlicerModel): CellScalar[] {
  const { sheet, column, startRow, endRow } = getSlicerSourceCellRange(workbook, slicer), values: CellScalar[] = []
  for (let row = startRow; row <= endRow; row++) values.push(sheet.cells[`${row}:${column}`]?.value ?? null)
  return values
}

export function getSlicerSourceRecordValues(workbook: WorkbookModel, slicer: SlicerModel): RecordValue[] {
  const { sheet, column, startRow, endRow } = getSlicerSourceCellRange(workbook, slicer), values: RecordValue[] = []
  for (let row = startRow; row <= endRow; row++) values.push(sheetCellRecordValue(sheet.cells[`${row}:${column}`], workbook.date1904))
  return values
}

export function getSlicerTableNativeFilter(workbook: WorkbookModel, slicer: SlicerModel): SheetFilterCriteria | undefined {
  const cache = getSlicerCache(workbook, slicer)
  if (cache.source.kind !== 'table') throw new SlicerConnectionError('invalid-target', '该切片器不是表格来源。')
  const source = cache.source, table = workbook.tables?.find(candidate => candidate.id === source.tableId)
  if (!table) throw new SlicerConnectionError('source-missing', '切片器表格来源不存在。')
  const index = table.columns.findIndex(candidate => candidate.id === source.columnId)
  if (index < 0) throw new SlicerConnectionError('field-missing', '切片器表格源字段不存在。')
  const filter = table.autoFilter?.columns[table.range.startCol + index]
  return filter === undefined ? undefined : cloneModel(filter)
}

function normalizeSlicerCacheSelection(workbook: WorkbookModel, slicer: SlicerModel, cache: SlicerCacheModel): void {
  if (cache.type === 'timeline') {
    normalizeRetainedTimelineCache(cache, workbook.date1904)
    // Member-only snapshots cannot be represented by one date interval without changing future-source behavior.
    if (cache.recordSelection !== undefined) cache.recordSelection = normalizeSlicerRecordSelection(cache.recordSelection)
    return
  }
  const values = getSlicerSourceValues(workbook, slicer), typedValues = getSlicerSourceRecordValues(workbook, slicer)
  if (cache.source.kind === 'table') {
    const filter = getSlicerTableNativeFilter(workbook, slicer)
    if (filter !== undefined && filter.kind !== 'values') return
    try { cache.recordSelection = filter ? slicerRecordSelectionFromNativeValues(filter) : undefined }
    catch (error) { throw new SlicerConnectionError('invalid-resource', error instanceof Error ? error.message : '表格原始成员选择无效。') }
    cache.selectedItems = slicerRecordSelectionItems(cache.recordSelection, values, typedValues)
    return
  }
  if (cache.recordSelection !== undefined) {
    let normalized: PivotRecordSelection | undefined
    try { normalized = normalizeSlicerRecordSelection(cache.recordSelection) }
    catch (error) { throw new SlicerConnectionError('invalid-resource', error instanceof Error ? error.message : '切片器原始成员选择无效。') }
    cache.recordSelection = normalized
    cache.selectedItems = slicerRecordSelectionItems(normalized, values, typedValues)
    return
  }
  if (cache.selectedItems === undefined) return
  // Unresolved legacy members keep their string rule until their raw type is observable.
  if (cache.selectedItems.some(item => !values.some(value => String(value ?? '') === item))) return
  cache.recordSelection = slicerRecordSelectionFromItems(values, cache.selectedItems)
}

export function getSlicerConnectionReports(workbook: WorkbookModel, slicer: SlicerModel): SlicerConnectionReport[] {
  const cache = getSlicerCache(workbook, slicer)
  if (cache.source.kind === 'table') return []
  return resolveConnectionCollection(workbook, cache.source, cache.connections ?? { pivotTables: [] }).reports
}

export function getSlicerConnectionCandidates(workbook: WorkbookModel, slicer: SlicerModel): SlicerConnectionCandidate[] {
  const cache = getSlicerCache(workbook, slicer)
  if (cache.source.kind === 'table') return []
  const source = cache.source
  const candidates: SlicerConnectionCandidate[] = []
  const append = (kind: SlicerConnectionCandidate['kind'], pivot: PivotTableModel, id: string, subUnitId: string, name: string) => {
    const sheet = workbook.sheets.find(candidate => candidate.id === subUnitId)
    if (!sheet || pivot.source.unitId !== workbook.id || !findSlicerReportField(pivot, source.sourceFieldId)) return
    try {
      if (resolvePivotCacheId(workbook, pivot) !== source.pivotCacheId
        || !samePivotSource(pivot.source, pivotCacheSource(workbook, source.pivotCacheId))) return
    } catch (error) { if (error instanceof SlicerConnectionError) return; throw error }
    candidates.push({ kind, subUnitId, id, name, sheetName: sheet.name })
  }
  for (const pivot of workbook.pivotTables ?? []) append('pivotTable', pivot, pivot.id, pivot.anchor.subUnitId, pivot.id)
  for (const chart of workbook.charts ?? []) {
    // Linked charts are represented once by their owner in the PivotTable list.
    if (chart.ownedPivot && !chart.pivotTableId) append('ownedPivotChart', chart.ownedPivot, chart.id, chart.sheetId, chart.title || chart.id)
  }
  return candidates
}

export function planSlicerConnections(workbook: WorkbookModel, slicer: SlicerModel, input: SheetSlicerConnections): SlicerConnectionPlan {
  const cache = getSlicerCache(workbook, slicer)
  if (cache.source.kind !== 'pivot') throw new SlicerConnectionError('invalid-target', '表格切片器不支持报告连接。')
  const next = resolveConnectionCollection(workbook, cache.source, input)
  const previous = getSlicerConnectionReports(workbook, slicer)
  const selected = new Set(next.reports.map(reportKey))
  return { cacheId: cache.id, connections: next.connections, reports: next.reports,
    removedReports: previous.filter(report => !selected.has(reportKey(report))),
    changed: !sameConnections(cache.connections ?? { pivotTables: [] }, next.connections) }
}

export function sameConnections(first: SheetSlicerConnections, second: SheetSlicerConnections): boolean {
  const keys = (value: SheetSlicerConnections) => [...value.pivotTables.map(ref => JSON.stringify(['pivotTable', ref.subUnitId, ref.pivotTableId])),
    ...(value.ownedPivotCharts ?? []).map(ref => JSON.stringify(['ownedPivotChart', ref.subUnitId, ref.pivotChartId]))].sort()
  return equivalentModelData(keys(first), keys(second))
}

export function reportKey(report: Pick<SlicerConnectionReport, 'kind' | 'subUnitId' | 'id'>): string {
  return JSON.stringify([report.kind, report.subUnitId, report.id])
}

function resolveConnectionCollection(workbook: WorkbookModel, source: Extract<SlicerCacheSource, { kind: 'pivot' }>, input: SheetSlicerConnections): { connections: SheetSlicerConnections; reports: SlicerConnectionReport[] } {
  if (!input || typeof input !== 'object' || Array.isArray(input) || !Array.isArray(input.pivotTables)
    || input.ownedPivotCharts !== undefined && !Array.isArray(input.ownedPivotCharts)
    || Object.keys(input).some(key => key !== 'pivotTables' && key !== 'ownedPivotCharts')) {
    throw new SlicerConnectionError('invalid-target', '切片器报告连接参数无效。')
  }
  const reports: SlicerConnectionReport[] = [], seen = new Set<string>()
  const add = (report: SlicerConnectionReport) => {
    assertReport(workbook, source, report)
    const key = reportKey(report)
    if (!seen.has(key)) { seen.add(key); reports.push(report) }
  }
  for (const ref of input.pivotTables) {
    assertReference(ref, 'pivotTableId')
    const pivot = workbook.pivotTables?.find(candidate => candidate.id === ref.pivotTableId)
    if (!pivot || pivot.anchor.subUnitId !== ref.subUnitId) throw new SlicerConnectionError('invalid-target', '报告连接的透视表或工作表不存在。')
    add({ kind: 'pivotTable', subUnitId: ref.subUnitId, id: pivot.id, pivot, field: requireSourceField(pivot, source.sourceFieldId) })
  }
  for (const ref of input.ownedPivotCharts ?? []) {
    assertReference(ref, 'pivotChartId')
    const chart = workbook.charts?.find(candidate => candidate.id === ref.pivotChartId)
    if (!chart || chart.sheetId !== ref.subUnitId || chart.pivotTableId && chart.ownedPivot) {
      throw new SlicerConnectionError('invalid-target', '报告连接的透视图表或来源无效。')
    }
    if (chart.pivotTableId) {
      const pivot = workbook.pivotTables?.find(candidate => candidate.id === chart.pivotTableId)
      if (!pivot) throw new SlicerConnectionError('invalid-target', '链接图表的所属透视表不存在。')
      add({ kind: 'pivotTable', subUnitId: pivot.anchor.subUnitId, id: pivot.id, pivot, field: requireSourceField(pivot, source.sourceFieldId) })
    } else if (chart.ownedPivot) {
      add({ kind: 'ownedPivotChart', subUnitId: chart.sheetId, id: chart.id, pivot: chart.ownedPivot,
        field: requireSourceField(chart.ownedPivot, source.sourceFieldId), chart })
    } else throw new SlicerConnectionError('invalid-target', '普通图表不能连接透视切片器。')
  }
  return { reports, connections: { pivotTables: reports.filter(report => report.kind === 'pivotTable').map(report => ({ subUnitId: report.subUnitId, pivotTableId: report.id })),
    ...(reports.some(report => report.kind === 'ownedPivotChart') ? { ownedPivotCharts: reports.filter(report => report.kind === 'ownedPivotChart').map(report => ({ subUnitId: report.subUnitId, pivotChartId: report.id })) } : {}) } }
}

function assertReport(workbook: WorkbookModel, source: Extract<SlicerCacheSource, { kind: 'pivot' }>, report: SlicerConnectionReport): void {
  if (!workbook.sheets.some(sheet => sheet.id === report.subUnitId) || report.pivot.anchor.unitId !== workbook.id
    || resolvePivotCacheId(workbook, report.pivot) !== source.pivotCacheId
    || !samePivotSource(report.pivot.source, pivotCacheSource(workbook, source.pivotCacheId))) {
    throw new SlicerConnectionError('invalid-target', '连接目标必须属于同一工作簿、透视缓存和源字段。')
  }
}

function legacyCache(workbook: WorkbookModel, slicer: SlicerModel): SlicerCacheModel {
  let source: SlicerCacheSource, connections: SheetSlicerConnections | undefined, items: string[] | undefined, recordSelection: PivotRecordSelection | undefined
  let timelineState: TimelineDateFilterState = {}
  if (slicer.tableId) {
    if (slicer.pivotTableId || slicer.ownedPivotChartId || slicer.type === 'timeline') throw new SlicerConnectionError('invalid-resource', '切片器来源冲突。')
    const table = workbook.tables?.find(candidate => candidate.id === slicer.tableId)
    if (!table) throw new SlicerConnectionError('source-missing', '切片器表格来源不存在。')
    source = { kind: 'table', subUnitId: table.sheetId, tableId: table.id, columnId: slicer.sourceFieldId }
    assertCacheSource(workbook, source)
    const column = table.columns.findIndex(candidate => candidate.id === source.columnId)
    const filter = table.autoFilter?.columns[table.range.startCol + column]
    recordSelection = filter?.kind === 'values' ? slicerRecordSelectionFromNativeValues(filter) : undefined
    const sheet = workbook.sheets.find(candidate => candidate.id === table.sheetId)!, values: CellScalar[] = [], typedValues: RecordValue[] = []
    for (let row = table.range.startRow + (table.showHeaderRow ? 1 : 0); row <= table.range.endRow - (table.showTotalsRow ? 1 : 0); row++) {
      const cell = sheet.cells[`${row}:${table.range.startCol + column}`]
      values.push(cell?.value ?? null)
      typedValues.push(sheetCellRecordValue(cell, workbook.date1904))
    }
    items = slicerRecordSelectionItems(recordSelection, values, typedValues)
  } else {
    const report = legacyPivotReport(workbook, slicer)
    if (!report) throw new SlicerConnectionError('source-missing', '切片器透视报告来源不存在。')
    source = { kind: 'pivot', pivotCacheId: resolvePivotCacheId(workbook, report.pivot), sourceFieldId: sourceFieldId(report.pivot, report.field) }
    connections = report.kind === 'pivotTable' ? { pivotTables: [{ subUnitId: report.subUnitId, pivotTableId: report.id }] }
      : { pivotTables: [], ownedPivotCharts: [{ subUnitId: report.subUnitId, pivotChartId: report.id }] }
    const canonicalFieldId = source.sourceFieldId
    const hasTypedState = report.pivot.fields.some(field => field.area !== 'value' && sourceFieldId(report.pivot, field) === canonicalFieldId
      && (field.recordSelection !== undefined || field.groupedItems !== undefined))
    recordSelection = hasTypedState ? normalizeSlicerRecordSelection(getPivotSourceRecordSelection(workbook, report.pivot, report.field.id)) : undefined
    items = hasTypedState ? getPivotSourceSelectedItems(workbook, report.pivot, report.field.id)
      : report.field.selectedItems ? [...report.field.selectedItems] : undefined
    if (slicer.type === 'timeline') timelineState = slicer.timelineDateRange !== undefined ? { range: cloneModel(slicer.timelineDateRange) }
      : slicer.timelineDateCondition !== undefined ? { condition: cloneModel(slicer.timelineDateCondition) } : readTimelineReportFilterState(report.pivot, report.field.id)
  }
  // Equal legacy primary reports share a cache; distinct reports can retain independent connections.
  const id = `slicer-cache:${JSON.stringify([slicer.type, source, connections ?? null])}`
  return { id, unitId: workbook.id!, type: slicer.type, source, ...(connections ? { connections } : {}),
    ...(recordSelection ? { recordSelection } : {}),
    ...(items ? { selectedItems: items } : {}), ...(slicer.timelineRange ? { timelineRange: cloneModel(slicer.timelineRange) } : {}),
    ...(timelineState.range ? { timelineDateRange: timelineState.range } : {}), ...(timelineState.condition ? { timelineDateCondition: timelineState.condition } : {}), revision: 0 }
}

function legacyPivotReport(workbook: WorkbookModel, slicer: SlicerModel): SlicerConnectionReport | undefined {
  if (slicer.pivotTableId && slicer.ownedPivotChartId) throw new SlicerConnectionError('invalid-resource', '切片器报告来源冲突。')
  if (slicer.pivotTableId) {
    const pivot = workbook.pivotTables?.find(candidate => candidate.id === slicer.pivotTableId)
    return pivot && { kind: 'pivotTable', subUnitId: pivot.anchor.subUnitId, id: pivot.id, pivot, field: legacySourceField(pivot, slicer.sourceFieldId) }
  }
  const chart = workbook.charts?.find(candidate => candidate.id === slicer.ownedPivotChartId)
  if (!chart?.ownedPivot || chart.pivotTableId) return undefined
  return { kind: 'ownedPivotChart', subUnitId: chart.sheetId, id: chart.id, pivot: chart.ownedPivot, field: legacySourceField(chart.ownedPivot, slicer.sourceFieldId), chart }
}

function pivotCacheSource(workbook: WorkbookModel, id: string): PivotTableSource {
  const retained = workbook.pivotCaches?.find(cache => cache.id === id)?.source
  if (retained) { assertSource(workbook, retained); return retained }
  for (const pivot of [...(workbook.pivotTables ?? []), ...(workbook.charts ?? []).flatMap(chart => chart.ownedPivot ? [chart.ownedPivot] : [])]) {
    try { if (resolvePivotCacheId(workbook, pivot) === id) return pivot.source }
    catch (error) { if (!(error instanceof SlicerConnectionError)) throw error }
  }
  throw new SlicerConnectionError('source-missing', '透视缓存数据源不存在。')
}

function assertCacheSource(workbook: WorkbookModel, source: SlicerCacheSource): void {
  if (source.kind === 'pivot') {
    const range = pivotCacheSource(workbook, nonemptyId(source.pivotCacheId, '透视缓存标识无效。')).range
    const retained = workbook.pivotCaches?.find(cache => cache.id === source.pivotCacheId)
    if (retained?.fields) {
      const field = retained.fields.find(candidate => candidate.id === source.sourceFieldId)
      if (!field || field.sourceColumn < range.startCol || field.sourceColumn > range.endCol) throw new SlicerConnectionError('field-missing', '切片器源字段已从缓存中删除。')
      return
    }
    const column = Number(source.sourceFieldId)
    if (!/^\d+$/.test(source.sourceFieldId) || !Number.isSafeInteger(column) || column < 0 || column > range.endCol - range.startCol) {
      // Imported caches may have stable nonnumeric source IDs on report fields.
      const retainedField = workbook.pivotCaches?.find(cache => cache.id === source.pivotCacheId)?.fields?.find(field => field.id === source.sourceFieldId)
      const compatible = retainedField && retainedField.sourceColumn >= range.startCol && retainedField.sourceColumn <= range.endCol
        || [...(workbook.pivotTables ?? []), ...(workbook.charts ?? []).flatMap(chart => chart.ownedPivot ? [chart.ownedPivot] : [])]
          .some(pivot => {
            try { return resolvePivotCacheId(workbook, pivot) === source.pivotCacheId && Boolean(findSlicerReportField(pivot, source.sourceFieldId)) }
            catch (error) { if (error instanceof SlicerConnectionError) return false; throw error }
          })
      if (!compatible) throw new SlicerConnectionError('field-missing', '切片器透视源字段不存在。')
    }
  } else if (source.kind === 'table') {
    const table = workbook.tables?.find(candidate => candidate.id === source.tableId && candidate.sheetId === source.subUnitId)
    if (!table || !workbook.sheets.some(sheet => sheet.id === source.subUnitId)) throw new SlicerConnectionError('source-missing', '切片器表格数据源不存在。')
    if (!table.columns.some(column => column.id === source.columnId)) throw new SlicerConnectionError('field-missing', '切片器表格源字段不存在。')
  } else throw new SlicerConnectionError('invalid-resource', '切片器缓存来源类型无效。')
}

function assertControl(workbook: WorkbookModel, slicer: SlicerModel): void {
  if (!workbook.id || slicer.unitId !== workbook.id || !workbook.sheets.some(sheet => sheet.id === slicer.hostSheetId)) {
    throw new SlicerConnectionError('invalid-resource', '切片器宿主不属于当前工作簿。')
  }
}

function assertSource(workbook: WorkbookModel, source: PivotTableSource): void {
  const sheet = workbook.sheets.find(candidate => candidate.id === source.subUnitId)
  if (source.unitId !== workbook.id || !sheet) throw new SlicerConnectionError('source-missing', '透视源工作表不属于当前工作簿。')
  const { range } = source
  if (!range || !['startRow', 'startCol', 'endRow', 'endCol'].every(key => Number.isSafeInteger(range[key as keyof typeof range]))
    || range.startRow < 0 || range.startCol < 0 || range.endRow <= range.startRow || range.endCol < range.startCol
    || range.endRow >= (sheet.rowCount ?? 1_048_576) || range.endCol >= (sheet.columnCount ?? 16_384)) {
    throw new SlicerConnectionError('invalid-resource', '透视缓存数据源范围无效。')
  }
}

function sourceKey(source: PivotTableSource): readonly unknown[] {
  return [source.unitId, source.subUnitId, source.range.startRow, source.range.startCol, source.range.endRow, source.range.endCol]
}

function requireSourceField(pivot: PivotTableModel, id: string): PivotFieldModel {
  const field = findSlicerReportField(pivot, id)
  if (!field) throw new SlicerConnectionError('field-missing', '连接报告中没有配置该源字段的标签字段。')
  return field
}

function legacySourceField(pivot: PivotTableModel, id: string): PivotFieldModel {
  const field = findSlicerReportField(pivot, id) ?? pivot.fields.find(candidate => candidate.id === id && candidate.area !== 'value')
  if (!field) throw new SlicerConnectionError('field-missing', '连接报告中没有配置该源字段的标签字段。')
  return field
}

function nonemptyId(value: string, message: string): string {
  if (typeof value !== 'string' || !value.trim()) throw new SlicerConnectionError('invalid-resource', message)
  return value
}

function assertReference(value: unknown, idKey: 'pivotTableId' | 'pivotChartId'): void {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new SlicerConnectionError('invalid-target', '连接目标引用无效。')
  const ref = value as Record<string, unknown>
  if (typeof ref.subUnitId !== 'string' || !ref.subUnitId || typeof ref[idKey] !== 'string' || !ref[idKey]
    || Object.keys(ref).some(key => key !== 'subUnitId' && key !== idKey)) throw new SlicerConnectionError('invalid-target', '连接目标必须包含有效的工作表和报告标识。')
}
