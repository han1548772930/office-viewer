import type { PivotTableModel, SlicerCacheModel, SlicerModel, WorkbookModel } from '../../domain/model.ts'
import type { ITimelineDateCondition, ITimelineRange, TimelineDateFilterState } from '../../domain/timeline-filters.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { equivalentModelData } from '../../platform/commands/model-equivalence.ts'
import { calculatePivotReport, PivotOutputConflictError, renderPivotTable, setPivotDateFilter,
  setPivotRecordSelection, setPivotTimelineDateCondition, setPivotTimelineDateRange } from '../pivot/services/pivot-engine.ts'
import { pivotFilterSourceKey } from '../pivot/services/pivot-filter-predicate.ts'
import { getSlicerCache, getSlicerConnectionReports, getSlicerSourceCellRange, getSlicerSourceRecordValues,
  SlicerConnectionError, type SlicerConnectionReport } from './slicer-connections.ts'
import { normalizeTimelineDateCondition, normalizeTimelineDateRange } from './timeline-date-filter.ts'
import { readTimelineCacheFilterState, readTimelineReportFilterState } from './timeline-state.ts'

/** Retained getters validate the hosted identity, never the current source or connected reports. */
function retainedCache(workbook: WorkbookModel, view: SlicerModel): SlicerCacheModel | undefined {
  if (view.type !== 'timeline') return undefined
  const current = workbook.slicers?.find(candidate => candidate.id === view.id && candidate.unitId === view.unitId
    && candidate.hostSheetId === view.hostSheetId && candidate.type === 'timeline')
  if (!current || current.unitId !== workbook.id || !workbook.sheets.some(sheet => sheet.id === current.hostSheetId)) return undefined
  const cache = workbook.slicerCaches?.find(candidate => candidate.id === current.cacheId && candidate.type === 'timeline'
    && candidate.unitId === workbook.id && candidate.source.kind === 'pivot' && candidate.source.sourceFieldId === current.sourceFieldId)
  return cache
}

export function readTimelineFilterState(cache: SlicerCacheModel): TimelineDateFilterState {
  try { return readTimelineCacheFilterState(cache) }
  catch (error) {
    if (!(error instanceof TypeError || error instanceof RangeError)) throw error
    throw new SlicerConnectionError('invalid-resource', error.message)
  }
}

export function readTimelineDateRange(workbook: WorkbookModel, view: SlicerModel): ITimelineRange | undefined {
  const cache = retainedCache(workbook, view)
  return cache ? readTimelineFilterState(cache).range : undefined
}

export function readTimelineDateCondition(workbook: WorkbookModel, view: SlicerModel): ITimelineDateCondition | undefined {
  const cache = retainedCache(workbook, view)
  return cache ? readTimelineFilterState(cache).condition : undefined
}

function normalizeState(input: TimelineDateFilterState): TimelineDateFilterState {
  if (!input || typeof input !== 'object' || Array.isArray(input)
    || Object.keys(input).some(key => key !== 'range' && key !== 'condition')
    || input.range !== undefined && input.condition !== undefined) throw new TypeError('Invalid Timeline filter state')
  if (input.range !== undefined) return { range: normalizeTimelineDateRange(input.range) }
  if (input.condition !== undefined) return { condition: normalizeTimelineDateCondition(input.condition) }
  return {}
}

/** All report fields, output footprints and owned-chart calculations are staged before publishing. */
export function applyTimelineReportFilter(workbook: WorkbookModel, reports: readonly SlicerConnectionReport[],
  input: TimelineDateFilterState, clearLegacyMembers = false): boolean {
  const state = normalizeState(input), nextReports = new Map<PivotTableModel, PivotTableModel>()
  const changes: SlicerConnectionReport[] = []
  for (const report of reports) {
    const current = readTimelineReportFilterState(report.pivot, report.field.id)
    const canonical = state.condition === undefined || report.pivot.fields.some(field => field.area !== 'value'
      && pivotFilterSourceKey(report.pivot, field) === pivotFilterSourceKey(report.pivot, report.field)
      && field.timelineDateCondition !== undefined)
    if (!clearLegacyMembers && canonical && equivalentModelData(current, state)) continue
    const next = cloneModel(report.pivot), field = next.fields.find(candidate => candidate.id === report.field.id && candidate.area !== 'value')
    if (!field) throw new SlicerConnectionError('field-missing', '时间线报告字段已更改。')
    if (clearLegacyMembers) setPivotRecordSelection(workbook, next, field.id, null, 'deferred')
    if (state.range !== undefined) setPivotTimelineDateRange(workbook, next, field.id, state.range, 'deferred')
    else if (state.condition !== undefined) setPivotTimelineDateCondition(workbook, next, field.id, state.condition, 'deferred')
    else {
      const source = pivotFilterSourceKey(next, field)
      for (const sibling of next.fields) if (sibling.area !== 'value' && pivotFilterSourceKey(next, sibling) === source
        && (sibling.dateFilter !== undefined || sibling.timelineDateRange !== undefined || sibling.timelineDateCondition !== undefined)) {
        setPivotDateFilter(workbook, next, sibling.id, null, 'deferred')
      }
    }
    if (equivalentModelData(next.fields, report.pivot.fields)) continue
    // Retiring an old member filter and authoring the date rule form one report edit.
    next.revision = report.pivot.revision + 1
    nextReports.set(report.pivot, next); changes.push(report)
  }
  if (!changes.length) return false
  const outputSheetIds = new Set(changes.filter(report => report.kind === 'pivotTable').map(report => report.subUnitId))
  const staged: WorkbookModel = { ...workbook,
    sheets: workbook.sheets.map(sheet => outputSheetIds.has(sheet.id) ? { ...sheet, cells: { ...sheet.cells } } : sheet),
    pivotTables: workbook.pivotTables?.map(pivot => nextReports.get(pivot) ?? pivot),
    charts: workbook.charts?.map(chart => chart.ownedPivot && nextReports.has(chart.ownedPivot)
      ? { ...chart, ownedPivot: nextReports.get(chart.ownedPivot) } : chart),
  }
  // Previous footprints are cleared together; newly overlapping outputs still refuse the entire edit.
  for (const report of changes) if (report.kind === 'pivotTable') {
    const next = nextReports.get(report.pivot)!, sheet = staged.sheets.find(candidate => candidate.id === report.subUnitId)
    if (!sheet) throw new SlicerConnectionError('invalid-target', '时间线报告工作表不存在。')
    for (const key of next.generatedCellKeys ?? []) if (sheet.cells[key]?.metadata?.pivotTableId === next.id) delete sheet.cells[key]
    next.generatedCellKeys = []; delete next.outputRange
  }
  try {
    for (const report of changes) {
      const next = nextReports.get(report.pivot)!
      if (report.kind === 'pivotTable') renderPivotTable(staged, next)
      else calculatePivotReport(staged, next)
    }
  } catch (error) {
    if (error instanceof PivotOutputConflictError) throw new SlicerConnectionError('invalid-target', error.message)
    throw error
  }
  for (const sheet of workbook.sheets) if (outputSheetIds.has(sheet.id)) sheet.cells = staged.sheets.find(candidate => candidate.id === sheet.id)!.cells
  for (const report of changes) {
    Object.assign(report.pivot, nextReports.get(report.pivot)!)
    if (report.chart) report.chart.revision += 1
  }
  return true
}

/** Canonical source filtering is shared; every hosted view keeps its own grain and navigation. */
export function updateTimelineCacheState(workbook: WorkbookModel, view: SlicerModel, input: TimelineDateFilterState, force = false): boolean {
  const state = normalizeState(input), cache = retainedCache(workbook, view)
  if (!cache) throw new SlicerConnectionError('invalid-resource', '时间线视图或共享缓存不存在。')
  const changed = force || !equivalentModelData(readTimelineFilterState(cache), state)
    || cache.recordSelection !== undefined || cache.selectedItems !== undefined || cache.timelineRange !== undefined || cache.timelineLegacyMemberFilterPending === true
  if (!changed) return false
  delete cache.recordSelection; delete cache.selectedItems; delete cache.timelineRange
  delete cache.timelineLegacyMemberFilterPending
  delete cache.timelineDateRange; delete cache.timelineDateCondition
  if (state.range !== undefined) cache.timelineDateRange = cloneModel(state.range)
  else if (state.condition !== undefined) cache.timelineDateCondition = cloneModel(state.condition)
  cache.revision += 1
  for (const control of workbook.slicers ?? []) if (control.cacheId === cache.id && control.type === 'timeline' && control.unitId === workbook.id) {
    delete control.selectedItems; delete control.timelineRange; delete control.timelineDateRange; delete control.timelineDateCondition
    if (state.range !== undefined) control.timelineDateRange = cloneModel(state.range)
    else if (state.condition !== undefined) control.timelineDateCondition = cloneModel(state.condition)
    control.revision += 1
  }
  return true
}

function setTimelineFilter(workbook: WorkbookModel, view: SlicerModel, state: TimelineDateFilterState): boolean {
  const cache = retainedCache(workbook, view)
  if (!cache) throw new SlicerConnectionError('invalid-resource', '时间线视图或共享缓存不存在。')
  // Live edits validate the original field even when no reports are connected.
  getSlicerCache(workbook, view); getSlicerSourceCellRange(workbook, view)
  readTimelineFilterState(cache)
  if ((state.range !== undefined || state.condition !== undefined)
    && !getSlicerSourceRecordValues(workbook, view).some(value => value.type === 'date')) {
    throw new SlicerConnectionError('invalid-target', '时间线需要真实日期源字段。')
  }
  const legacy = cache.recordSelection !== undefined || cache.selectedItems !== undefined || cache.timelineRange !== undefined || cache.timelineLegacyMemberFilterPending === true
  const filtered = applyTimelineReportFilter(workbook, getSlicerConnectionReports(workbook, view), state, legacy)
  return updateTimelineCacheState(workbook, view, state, filtered) || filtered
}

export function setTimelineDateRange(workbook: WorkbookModel, view: SlicerModel, input: ITimelineRange | undefined): boolean {
  const state: TimelineDateFilterState = input === undefined ? {} : { range: normalizeTimelineDateRange(input) }
  return setTimelineFilter(workbook, view, state)
}

/** Internal intent used by Pivot date-condition edits; the fixed public Timeline Facade exposes its getter. */
export function setTimelineDateCondition(workbook: WorkbookModel, view: SlicerModel, input: ITimelineDateCondition | undefined): boolean {
  const state: TimelineDateFilterState = input === undefined ? {} : { condition: normalizeTimelineDateCondition(input) }
  return setTimelineFilter(workbook, view, state)
}
