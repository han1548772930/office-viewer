import type { PivotTableModel, SlicerModel, SlicerType, WorkbookModel } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { calculatePivotReport, renderPivotTable } from '../pivot/services/pivot-engine.ts'
import { createSlicer, updateSlicerPresentation } from './slicer-engine.ts'
import { ensurePivotCache, findSlicerReportField, getSlicerCache, getSlicerConnectionReports, SlicerConnectionError } from './slicer-connections.ts'
import { normalizeSheetSlicerInfo, normalizeSheetSlicerSource, normalizeSheetTimelineInfo, readSheetSlicerSource,
  resolveSheetSlicerSource, sameSheetSlicerSource, sheetSlicerPlacementPatch,
  type ISheetSlicerInfo, type ISheetTimelineInfo, type SheetSlicerSourceSpec } from './slicer-public-info.ts'
import { normalizeSlicerSettingsPatch, readSlicerSettings } from './slicer-settings.ts'
import { normalizeTimelineSettingsPatch, readTimelineSettings } from './timeline-settings.ts'
import { applyTimelineReportFilter, readTimelineFilterState } from './timeline-filter.ts'
import { applySlicerReportFilter } from './slicer-report-filter.ts'

const invalid = (message: string): never => { throw new SlicerConnectionError('invalid-target', message) }
/** Linked chart references normalize to their owner; every identity remains in this Workbook's original cache. */
export function canonicalSheetSlicerSource(workbook: WorkbookModel, input: SheetSlicerSourceSpec): SheetSlicerSourceSpec {
  const source = normalizeSheetSlicerSource(input)
  if (source.kind === 'table') {
    resolveSheetSlicerSource(workbook, source.columnId, { subUnitId: source.subUnitId, tableId: source.tableId })
    return source
  }
  const pivotTables = [...source.pivotTables], ownedPivotCharts: NonNullable<typeof source.ownedPivotCharts> = []
  for (const ref of source.ownedPivotCharts ?? []) {
    const resolved = resolveSheetSlicerSource(workbook, source.sourceFieldId, { unitId: workbook.id!, subUnitId: ref.subUnitId, pivotChartId: ref.pivotChartId })
    if (resolved.kind !== 'pivot' || resolved.pivotCacheId !== source.pivotCacheId) return invalid('切片器连接必须来自同一原始透视缓存。')
    pivotTables.push(...resolved.pivotTables); ownedPivotCharts.push(...resolved.ownedPivotCharts ?? [])
  }
  for (const ref of pivotTables) {
    const resolved = resolveSheetSlicerSource(workbook, source.sourceFieldId, { unitId: workbook.id!, ...ref })
    if (resolved.kind !== 'pivot' || resolved.pivotCacheId !== source.pivotCacheId) return invalid('切片器连接必须来自同一原始透视缓存。')
  }
  return normalizeSheetSlicerSource({ ...source, pivotTables, ...(source.ownedPivotCharts === undefined && !ownedPivotCharts.length ? {} : { ownedPivotCharts }) })
}

/** Original fields can filter a report even when no visible dimension uses them; value instances are preserved. */
function retainSourceFilterField(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string): boolean {
  if (findSlicerReportField(pivot, fieldId)) return false
  const cache = ensurePivotCache(workbook, pivot), source = cache.fields?.find(field => field.id === fieldId)
  if (!source) return invalid('切片器原始透视字段不存在。')
  let localId = fieldId
  for (let index = 2; pivot.fields.some(field => field.id === localId); index++) localId = `${fieldId}:slicer-filter:${index}`
  pivot.fields.push({ id: localId, sourceFieldId: fieldId, sourceColumn: source.sourceColumn,
    name: source.name ?? fieldId, area: 'filter', position: pivot.fields.filter(field => field.area === 'filter').length })
  pivot.revision += 1
  return true
}

/** Called only on the command host's detached Workbook candidate; no partial state is published. */
export function createSheetSlicerFromInfo(workbook: WorkbookModel, hostSheetId: string, viewId: string, kind: SlicerType,
  input: ISheetSlicerInfo | ISheetTimelineInfo): SlicerModel {
  const info = kind === 'timeline' ? normalizeSheetTimelineInfo(input as ISheetTimelineInfo) : normalizeSheetSlicerInfo(input as ISheetSlicerInfo)
  const dataSource = canonicalSheetSlicerSource(workbook, info.dataSource), source = dataSource.kind === 'table' ? dataSource
    : { kind: 'pivot' as const, pivotCacheId: dataSource.pivotCacheId, sourceFieldId: dataSource.sourceFieldId }
  if (!workbook.id || !workbook.sheets.some(sheet => sheet.id === hostSheetId)) return invalid('切片器工作簿或宿主无效。')
  if (workbook.slicers?.some(view => view.id === viewId)) return invalid('切片器标识已存在。')
  const changedReports: Array<{ pivot: PivotTableModel; chartId?: string }> = []
  if (dataSource.kind === 'pivot') {
    for (const ref of dataSource.pivotTables) {
      const pivot = workbook.pivotTables!.find(pivot => pivot.id === ref.pivotTableId && pivot.anchor.subUnitId === ref.subUnitId)!
      ensurePivotCache(workbook, pivot)
      if (retainSourceFilterField(workbook, pivot, dataSource.sourceFieldId)) changedReports.push({ pivot })
    }
    for (const ref of dataSource.ownedPivotCharts ?? []) {
      const chart = workbook.charts!.find(chart => chart.id === ref.pivotChartId && chart.sheetId === ref.subUnitId)!
      if (!chart.ownedPivot || chart.pivotTableId) return invalid('切片器连接的独立透视图无效。')
      ensurePivotCache(workbook, chart.ownedPivot)
      if (retainSourceFilterField(workbook, chart.ownedPivot, dataSource.sourceFieldId)) changedReports.push({ pivot: chart.ownedPivot, chartId: chart.id })
    }
    const cache = workbook.pivotCaches?.find(cache => cache.id === dataSource.pivotCacheId)
    if (!cache || !cache.fields?.some(field => field.id === dataSource.sourceFieldId)
      || cache.source.unitId !== workbook.id || !workbook.sheets.some(sheet => sheet.id === cache.source.subUnitId)) return invalid('切片器原始透视缓存或字段不存在。')
  }
  const matches = (cache: NonNullable<WorkbookModel['slicerCaches']>[number]) => cache.unitId === workbook.id && cache.type === kind
    && sameSheetSlicerSource(cache.source.kind === 'table' ? cache.source : { ...cache.source, ...(cache.connections ?? { pivotTables: [] }) }, dataSource)
  let cache = workbook.slicerCaches?.find(matches)
  if (!cache) {
    let cacheId = `${viewId}:cache`
    for (let index = 2; workbook.slicerCaches?.some(cache => cache.id === cacheId); index++) cacheId = `${viewId}:cache:${index}`
    cache = { id: cacheId, unitId: workbook.id, type: kind, source: cloneModel(source), revision: 0,
      ...(dataSource.kind === 'pivot' ? { connections: { pivotTables: cloneModel(dataSource.pivotTables),
        ...(dataSource.ownedPivotCharts === undefined ? {} : { ownedPivotCharts: cloneModel(dataSource.ownedPivotCharts) }) } } : {}) }
    workbook.slicerCaches ??= []; workbook.slicerCaches.push(cache)
  }
  const settings = kind === 'timeline' ? normalizeTimelineSettingsPatch(info.settings) : normalizeSlicerSettingsPatch(info.settings)
  const view = createSlicer(workbook, { id: viewId, unitId: workbook.id, hostSheetId, type: kind, cacheId: cache.id,
    sourceFieldId: source.kind === 'table' ? source.columnId : source.sourceFieldId, ...settings })
  if (info.anchor !== undefined || info.position !== undefined || info.size !== undefined) updateSlicerPresentation(view, sheetSlicerPlacementPatch(workbook, view, info), workbook)
  view.revision = 0
  const reports = getSlicerConnectionReports(workbook, view)
  if (kind === 'timeline') applyTimelineReportFilter(workbook, reports, readTimelineFilterState(cache))
  else if (cache.recordSelection !== undefined) applySlicerReportFilter(workbook, reports, cache.recordSelection)
  // Adding an original-only source label must also refresh its real report output.
  for (const entry of changedReports) {
    if (entry.chartId) {
      calculatePivotReport(workbook, entry.pivot)
      workbook.charts!.find(chart => chart.id === entry.chartId)!.revision += 1
    } else renderPivotTable(workbook, entry.pivot)
  }
  getSlicerCache(workbook, view)
  return view
}

/** Source and connections are locked; per-view settings and placement form one content transaction. */
export function updateSheetSlicerFromInfo(workbook: WorkbookModel, view: SlicerModel, input: ISheetSlicerInfo | ISheetTimelineInfo): void {
  const info = view.type === 'timeline' ? normalizeSheetTimelineInfo(input as ISheetTimelineInfo) : normalizeSheetSlicerInfo(input as ISheetSlicerInfo)
  const currentSource = readSheetSlicerSource(workbook, view)
  if (!currentSource || !sameSheetSlicerSource(currentSource, info.dataSource)) return invalid('活动切片器的来源与报告连接不能通过显示更新替换。')
  const currentSettings = view.type === 'timeline' ? readTimelineSettings(view) : readSlicerSettings(view)
  const optional = view.type === 'timeline'
    ? ['caption', 'style', 'scrollPosition', 'showHeader', 'showSelectionLabel', 'showTimeLevel', 'showHorizontalScrollbar']
    : ['caption', 'style', 'startItem', 'columnCount', 'showCaption', 'lockedPosition']
  const settings = { ...Object.fromEntries(optional.map(key => [key, undefined])), ...info.settings, name: info.settings.name ?? currentSettings.name }
  const patch = view.type === 'timeline' ? normalizeTimelineSettingsPatch(settings) : normalizeSlicerSettingsPatch(settings)
  updateSlicerPresentation(view, { ...patch, ...(info.anchor !== undefined || info.position !== undefined || info.size !== undefined
    ? sheetSlicerPlacementPatch(workbook, view, info) : {}) }, workbook)
}
