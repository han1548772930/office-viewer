import type { PivotFieldModel, PivotTableModel, SlicerCacheModel } from '../../domain/model.ts'
import type { TimelineDateFilterState } from '../../domain/timeline-filters.ts'
import { equivalentModelData } from '../../platform/commands/model-equivalence.ts'
import { normalizeTimelineDateCondition, normalizeTimelineDateRange, timelineDateConditionFromPivot } from './timeline-date-filter.ts'
import { timelineFilterStateFromLegacy } from './timeline-periods.ts'

/** Pure retained read; missing live sources do not change authored date rules. */
export function readTimelineCacheFilterState(cache: SlicerCacheModel): TimelineDateFilterState {
  if (cache.type !== 'timeline' || cache.source.kind !== 'pivot') throw new TypeError('时间线需要透视来源缓存。')
  if (cache.timelineDateRange !== undefined && cache.timelineDateCondition !== undefined) throw new TypeError('时间线范围与日期条件不能同时存在。')
  if (cache.timelineDateRange !== undefined) return { range: normalizeTimelineDateRange(cache.timelineDateRange) }
  if (cache.timelineDateCondition !== undefined) return { condition: normalizeTimelineDateCondition(cache.timelineDateCondition) }
  return {}
}

/** Read the source category across report-local dimension copies; numeric value instances never own date rules. */
export function readTimelineReportFilterState(pivot: PivotTableModel, fieldId: string): TimelineDateFilterState {
  const field = pivot.fields.find(candidate => candidate.id === fieldId && candidate.area !== 'value')
  if (!field) throw new TypeError('时间线报告日期字段不存在。')
  const sourceKey = (candidate: PivotFieldModel) => candidate.sourceFieldId ?? String(candidate.sourceColumn - pivot.source.range.startCol)
  const source = sourceKey(field)
  let result: TimelineDateFilterState = {}
  for (const candidate of pivot.fields) if (candidate.area !== 'value' && sourceKey(candidate) === source) {
    if (candidate.timelineDateRange !== undefined && (candidate.timelineDateCondition !== undefined || candidate.dateFilter !== undefined)) {
      throw new TypeError('报告时间线范围与日期条件冲突。')
    }
    const condition = candidate.timelineDateCondition === undefined ? candidate.dateFilter === undefined ? undefined
      : timelineDateConditionFromPivot(candidate.dateFilter) : normalizeTimelineDateCondition(candidate.timelineDateCondition)
    if (candidate.timelineDateCondition !== undefined && candidate.dateFilter !== undefined
      && !equivalentModelData(condition, timelineDateConditionFromPivot(candidate.dateFilter))) throw new TypeError('报告日期条件与时间线来源标记不一致。')
    const state: TimelineDateFilterState = candidate.timelineDateRange !== undefined ? { range: normalizeTimelineDateRange(candidate.timelineDateRange) }
      : condition !== undefined ? { condition } : {}
    if (state.range === undefined && state.condition === undefined) continue
    if ((result.range !== undefined || result.condition !== undefined) && !equivalentModelData(result, state)) {
      throw new TypeError('同一报告原始字段包含不同时间线日期规则。')
    }
    result = state
  }
  return result
}

/** Snapshot migration runs on decoded/detached state; ambiguous member-only legacy data remains retained. */
export function normalizeRetainedTimelineCache(cache: SlicerCacheModel, date1904: boolean): void {
  const authored = readTimelineCacheFilterState(cache)
  const state = authored.range !== undefined || authored.condition !== undefined || cache.timelineRange === undefined
    ? authored : timelineFilterStateFromLegacy(cache.timelineRange, date1904)
  const retireLegacyMembers = cache.timelineRange !== undefined || state.range !== undefined || state.condition !== undefined
  if (retireLegacyMembers && (cache.timelineRange !== undefined || cache.recordSelection !== undefined || cache.selectedItems !== undefined)) {
    cache.timelineLegacyMemberFilterPending = true
  }
  delete cache.timelineDateRange; delete cache.timelineDateCondition
  if (state.range !== undefined) cache.timelineDateRange = state.range
  if (state.condition !== undefined) cache.timelineDateCondition = state.condition
  if (retireLegacyMembers) { delete cache.timelineRange; delete cache.recordSelection; delete cache.selectedItems }
}
