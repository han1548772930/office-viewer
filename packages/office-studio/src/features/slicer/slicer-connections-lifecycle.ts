import type { PivotTableModel, SheetSlicerConnections, SlicerModel, WorkbookModel } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { ensureSlicerCache, normalizeSlicerCaches, resolvePivotCacheId, sameConnections, samePivotSource, sourceFieldId } from './slicer-connections.ts'
import { reconcilePivotSourceCaches, type PivotSourceCacheTransform } from '../pivot/services/pivot-source-lifecycle.ts'

const newId = (kind: string): string => `copy-${kind}-${crypto.randomUUID()}`
const sourceKey = (pivot: PivotTableModel): string => JSON.stringify([pivot.source.unitId, pivot.source.subUnitId,
  pivot.source.range.startRow, pivot.source.range.startCol, pivot.source.range.endRow, pivot.source.range.endCol])
const fieldKey = sourceFieldId
const resourceProbe = (workbook: WorkbookModel): WorkbookModel => ({ ...workbook,
  pivotTables: workbook.pivotTables ? cloneModel(workbook.pivotTables) : undefined,
  charts: workbook.charts ? cloneModel(workbook.charts) : undefined,
  slicers: workbook.slicers ? cloneModel(workbook.slicers) : undefined,
  pivotCaches: workbook.pivotCaches ? cloneModel(workbook.pivotCaches) : undefined,
  slicerCaches: workbook.slicerCaches ? cloneModel(workbook.slicerCaches) : undefined,
})

export interface SlicerCopyConnectionsMapping {
  pivotTableIds: ReadonlyMap<string, string>
  pivotChartIds: ReadonlyMap<string, string>
  tableIds?: ReadonlyMap<string, { id: string; columns: ReadonlyMap<string, string> }>
  /** A batch shares identities for source caches and control caches within one tile/copy. */
  pivotCacheIds: Map<string, string>
  slicerCacheIds: Map<string, string>
}

/** Retain source-field identity while instance IDs and report locations can change. No source records are duplicated here. */
export function copyPivotReportCache(before: WorkbookModel, candidate: WorkbookModel, original: PivotTableModel,
  copied: PivotTableModel, ids: Map<string, string>): void {
  const originalCacheId = original.pivotCacheId ?? original.source.pivotCacheId ?? `pivot-cache:${sourceKey(original)}`
  copied.fields = copied.fields.map((field, index) => ({ ...field,
    sourceFieldId: original.fields[index] ? fieldKey(original, original.fields[index]!) : field.sourceFieldId ?? String(field.sourceColumn - copied.source.range.startCol) }))
  const unchanged = before.id === candidate.id && sourceKey(original) === sourceKey(copied)
  const key = JSON.stringify([originalCacheId, sourceKey(copied)])
  let id = unchanged ? originalCacheId : ids.get(key)
  if (!id) { id = newId('pivot-cache'); ids.set(key, id) }
  copied.pivotCacheId = id
  if (copied.source.pivotCacheId !== undefined) copied.source.pivotCacheId = id
  candidate.pivotCaches ??= []
  const existing = candidate.pivotCaches.find(cache => cache.id === id)
  if (existing && !samePivotSource(existing.source, copied.source)) {
    throw new TypeError('复制的透视缓存身份与其数据源不一致。')
  }
  if (!existing) {
    const originalCache = before.pivotCaches?.find(cache => cache.id === originalCacheId)
      ?? candidate.pivotCaches.find(cache => cache.id === originalCacheId && samePivotSource(cache.source, original.source))
    const delta = copied.source.range.startCol - original.source.range.startCol
    candidate.pivotCaches.push({ ...(originalCache ? cloneModel(originalCache) : {}), id, source: cloneModel(copied.source), revision: 0,
      ...(originalCache?.fields ? { fields: originalCache.fields.map(field => ({ ...cloneModel(field), sourceColumn: field.sourceColumn + delta })) } : {}) })
  }
}

function connectionReport(workbook: WorkbookModel, ref: SheetSlicerConnections['pivotTables'][number]): PivotTableModel {
  const pivot = workbook.pivotTables?.find(pivot => pivot.id === ref.pivotTableId && pivot.anchor.subUnitId === ref.subUnitId)
  if (!pivot) throw new TypeError('复制的切片器连接缺少完整透视报告。')
  return pivot
}

/** Complete connections are replaced as one snapshot; an unresolved foreign report never gets silently dropped. */
export function copySlicerViewCache(before: WorkbookModel, candidate: WorkbookModel, original: SlicerModel, copied: SlicerModel,
  mapping: SlicerCopyConnectionsMapping): void {
  const source = resourceProbe(before), originalControl = source.slicers?.find(view => view.id === original.id) ?? cloneModel(original)
  const retained = originalControl.cacheId ? source.slicerCaches?.find(cache => cache.id === originalControl.cacheId) : undefined
  const cache = retained ?? ensureSlicerCache(source, originalControl)
  if (!cache) throw new TypeError('复制的切片器缺少有效共享来源缓存。')
  if (cache.unitId !== originalControl.unitId || cache.type !== originalControl.type || cache.source.kind === 'pivot' && cache.source.sourceFieldId !== originalControl.sourceFieldId
    || cache.source.kind === 'table' && cache.source.columnId !== originalControl.sourceFieldId) throw new TypeError('复制的切片器缓存与来源控件不一致。')
  const next = cloneModel(cache)
  next.unitId = candidate.id!; next.revision = 0
  if (next.source.kind === 'table') {
    const mapped = mapping.tableIds?.get(next.source.tableId)
    const tableId = mapped?.id ?? (before.id === candidate.id ? next.source.tableId : undefined)
    const table = candidate.tables?.find(table => table.id === tableId)
    const columnId = mapped?.columns.get(next.source.columnId) ?? next.source.columnId
    if (!table || !table.columns.some(column => column.id === columnId)) throw new TypeError('复制的切片器缺少完整表格列来源。')
    next.source = { kind: 'table', subUnitId: table.sheetId, tableId: table.id, columnId }
    copied.tableId = table.id; copied.sourceFieldId = columnId
    delete copied.pivotTableId; delete copied.ownedPivotChartId
  } else {
    const pivotSource = next.source
    const connections: SheetSlicerConnections = { pivotTables: [] }
    for (const ref of next.connections?.pivotTables ?? []) {
      const id = mapping.pivotTableIds.get(ref.pivotTableId) ?? (before.id === candidate.id ? ref.pivotTableId : undefined)
      const pivot = candidate.pivotTables?.find(pivot => pivot.id === id)
      if (!pivot) throw new TypeError('跨工作簿复制切片器需要包含所有连接的透视报告和来源。')
      connections.pivotTables.push({ subUnitId: pivot.anchor.subUnitId, pivotTableId: pivot.id })
    }
    if (next.connections?.ownedPivotCharts !== undefined) {
      connections.ownedPivotCharts = next.connections.ownedPivotCharts.map(ref => {
        const id = mapping.pivotChartIds.get(ref.pivotChartId) ?? (before.id === candidate.id ? ref.pivotChartId : undefined)
        const chart = candidate.charts?.find(chart => chart.id === id)
        if (!chart?.ownedPivot) throw new TypeError('复制切片器需要包含所有独立透视图报告和来源。')
        return { subUnitId: chart.sheetId, pivotChartId: chart.id }
      })
    }
    const reports = [
      ...connections.pivotTables.map(ref => connectionReport(candidate, ref)),
      ...(connections.ownedPivotCharts ?? []).map(ref => candidate.charts!.find(chart => chart.id === ref.pivotChartId)!.ownedPivot!),
    ]
    const cacheIds = new Set(reports.map(report => resolvePivotCacheId(candidate, report)))
    if (cacheIds.size > 1) throw new TypeError('复制后的切片器连接来自不同透视缓存，请连同所有连接报告一起复制。')
    const cacheId = [...cacheIds][0]
    if (!cacheId && before.id !== candidate.id) throw new TypeError('未连接报告的切片器缺少可迁移的透视来源。')
    if (cacheId) pivotSource.pivotCacheId = cacheId
    for (const report of reports) if (!report.fields.some(field => field.area !== 'value' && fieldKey(report, field) === pivotSource.sourceFieldId)) {
      throw new TypeError('复制后的切片器报告缺少锁定的来源字段。')
    }
    if (next.connections !== undefined || connections.pivotTables.length || connections.ownedPivotCharts?.length) next.connections = connections
    const primary = connections.pivotTables[0]
    const owned = connections.ownedPivotCharts?.[0]
    if (primary) copied.pivotTableId = primary.pivotTableId; else delete copied.pivotTableId
    if (owned && !primary) copied.ownedPivotChartId = owned.pivotChartId; else delete copied.ownedPivotChartId
    copied.sourceFieldId = pivotSource.sourceFieldId
    delete copied.tableId
  }
  const same = before.id === candidate.id && JSON.stringify(cache.source) === JSON.stringify(next.source)
    && sameConnections(cache.connections ?? { pivotTables: [] }, next.connections ?? { pivotTables: [] })
  const key = JSON.stringify([cache.id, next.source, next.connections])
  let id = same ? cache.id : mapping.slicerCacheIds.get(key)
  if (!id) { id = newId('slicer-cache'); mapping.slicerCacheIds.set(key, id) }
  next.id = id; copied.cacheId = id
  candidate.slicerCaches ??= []
  if (!candidate.slicerCaches.some(cache => cache.id === id)) candidate.slicerCaches.push(next)
  const committed = candidate.slicerCaches.find(cache => cache.id === id)!
  copied.selectedItems = committed.selectedItems ? [...committed.selectedItems] : undefined
  copied.timelineRange = committed.timelineRange ? cloneModel(committed.timelineRange) : undefined
}

export interface SlicerCacheStructureTransform extends PivotSourceCacheTransform {}

/** A structural transaction has already identified actual deletions; keep every surviving view and shared source. */
export function reconcileSlicerConnectionResources(before: WorkbookModel, candidate: WorkbookModel, transform?: SlicerCacheStructureTransform): void {
  const previous = resourceProbe(before)
  normalizeSlicerCaches(previous)
  candidate.slicerCaches ??= cloneModel(previous.slicerCaches ?? [])
  candidate.pivotCaches ??= cloneModel(previous.pivotCaches ?? [])
  for (const view of candidate.slicers ?? []) view.cacheId ??= previous.slicers?.find(item => item.id === view.id)?.cacheId
  for (const report of candidate.pivotTables ?? []) report.pivotCacheId ??= previous.pivotTables?.find(item => item.id === report.id)?.pivotCacheId
  for (const chart of candidate.charts ?? []) if (chart.ownedPivot) chart.ownedPivot.pivotCacheId ??=
    previous.charts?.find(item => item.id === chart.id)?.ownedPivot?.pivotCacheId
  const removedSourceFields = reconcilePivotSourceCaches(previous, candidate, transform)
  const removedViews = new Set<string>(), removedControlCaches = new Set<string>()
  for (const cache of candidate.slicerCaches) {
    const baseline = JSON.stringify(cache), definition = cache.source
    if (definition.kind === 'pivot' && removedSourceFields.get(definition.pivotCacheId)?.has(definition.sourceFieldId)) {
      for (const view of candidate.slicers ?? []) if (view.cacheId === cache.id) removedViews.add(view.id)
      removedControlCaches.add(cache.id)
      continue
    }
    if (definition.kind === 'table') {
      const table = candidate.tables?.find(table => table.id === definition.tableId)
      if (!table || !table.columns.some(column => column.id === definition.columnId)) {
        if (!previous.tables?.some(table => table.id === definition.tableId && table.columns.some(column => column.id === definition.columnId))) continue
        for (const view of candidate.slicers ?? []) if (view.cacheId === cache.id) removedViews.add(view.id)
        removedControlCaches.add(cache.id)
        continue
      }
      definition.subUnitId = table.sheetId
    } else if (cache.connections) {
      cache.connections.pivotTables = cache.connections.pivotTables.flatMap(ref => {
        const report = candidate.pivotTables?.find(pivot => pivot.id === ref.pivotTableId)
        const old = previous.pivotTables?.find(pivot => pivot.id === ref.pivotTableId && pivot.anchor.subUnitId === ref.subUnitId)
        if (!old) return [ref]
        if (!report) return []
        const survives = report.fields.some(field => field.area !== 'value' && fieldKey(report, field) === definition.sourceFieldId)
        const existed = old?.fields.some(field => field.area !== 'value' && fieldKey(old, field) === definition.sourceFieldId)
        return survives || !existed ? [{ subUnitId: report.anchor.subUnitId, pivotTableId: report.id }] : []
      })
      if (cache.connections.ownedPivotCharts) cache.connections.ownedPivotCharts = cache.connections.ownedPivotCharts.flatMap(ref => {
        const chart = candidate.charts?.find(chart => chart.id === ref.pivotChartId)
        const old = previous.charts?.find(chart => chart.id === ref.pivotChartId && chart.sheetId === ref.subUnitId)?.ownedPivot
        if (!old) return [ref]
        if (!chart?.ownedPivot) return []
        const survives = chart.ownedPivot.fields.some(field => field.area !== 'value' && fieldKey(chart.ownedPivot!, field) === definition.sourceFieldId)
        const existed = old?.fields.some(field => field.area !== 'value' && fieldKey(old, field) === definition.sourceFieldId)
        return survives || !existed ? [{ subUnitId: chart.sheetId, pivotChartId: chart.id }] : []
      })
    }
    const changed = JSON.stringify(cache) !== baseline
    if (changed) cache.revision++
    for (const view of candidate.slicers ?? []) if (view.cacheId === cache.id) {
      const previousView = JSON.stringify(view)
      if (definition.kind === 'pivot') {
        const primary = cache.connections?.pivotTables[0], owned = cache.connections?.ownedPivotCharts?.[0]
        if (primary) view.pivotTableId = primary.pivotTableId; else delete view.pivotTableId
        if (!primary && owned) view.ownedPivotChartId = owned.pivotChartId; else delete view.ownedPivotChartId
        view.sourceFieldId = definition.sourceFieldId
      }
      view.selectedItems = cache.selectedItems ? [...cache.selectedItems] : undefined
      view.timelineRange = cache.timelineRange ? cloneModel(cache.timelineRange) : undefined
      if (changed || JSON.stringify(view) !== previousView) view.revision++
    }
  }
  if (removedViews.size) candidate.slicers = candidate.slicers?.filter(view => !removedViews.has(view.id))
  if (removedControlCaches.size) candidate.slicerCaches = candidate.slicerCaches.filter(cache => !removedControlCaches.has(cache.id))
  // Empty caches may still hold an explicitly disconnected source/selection, so retain them for later reconnection.
  normalizeSlicerCaches(candidate)
}
