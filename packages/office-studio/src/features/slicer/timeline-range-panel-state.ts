import { FOffice, type FSheetTimeline } from '../../api/facade.ts'
import type { OfficeStudio } from '../../api/office-studio.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { TimelineDateFilterOperator, type ITimelineDateCondition, type TimelineDateFilterState } from '../../domain/timeline-filters.ts'
import { equivalentModelData } from '../../platform/commands/model-equivalence.ts'
import { pivotChartFilterDateCalendarInput, pivotChartFilterDateCompatibilityDay, pivotChartFilterDateFromCalendarInput } from '../pivot/services/pivot-chart-field-filter-condition.ts'
import { canConfigurePanelSlicer, SlicerPanelError, type SlicerPanelTarget } from './slicer-panel-state.ts'
import { captureSlicerPanelOwner, slicerPanelOwnerContext, slicerPanelOwnerCurrent, type SlicerPanelOwner } from './slicer-panel-owner.ts'
import { withSlicerPanelCommandOwner } from './slicer-panel-command-owner.ts'
import { normalizeTimelineDateCondition, normalizeTimelineDateRange } from './timeline-date-filter.ts'
import { SLICER_PUBLIC_COMMANDS } from './slicer-public-info.ts'

export type TimelinePanelDateBounds = Awaited<ReturnType<FSheetTimeline['getDateBounds']>>
export interface TimelineRangePanelSnapshot {
  owner: SlicerPanelOwner; timelineId: string; cacheId?: string; date1904: boolean; state: TimelineDateFilterState
}
export interface TimelineDateFilterDraft {
  mode: 'none' | 'range' | 'condition'
  range: { startDate?: number; endDate?: number; useWholeDay?: boolean }
  condition: { operator: TimelineDateFilterOperator; first?: number; second?: number; useWholeDay?: boolean }
}
export interface TimelinePanelPreparedCommand { id: string; payloadKey: 'range' | 'condition' | 'params'; payload: object }
export const TIMELINE_DATE_CONDITION_OPTIONS: Array<{ label: string; value: TimelineDateFilterOperator }> = [
  { label: '等于', value: 'equal' }, { label: '不等于', value: 'notEqual' }, { label: '晚于', value: 'newerThan' },
  { label: '晚于或等于', value: 'newerThanOrEqual' }, { label: '早于', value: 'olderThan' }, { label: '早于或等于', value: 'olderThanOrEqual' },
  { label: '介于（含两端）', value: 'between' }, { label: '不介于（含两端）', value: 'notBetween' },
]
export const timelinePanelDateCalendarInput = pivotChartFilterDateCalendarInput
export const timelinePanelDateFromCalendarInput = pivotChartFilterDateFromCalendarInput
export const timelinePanelDateCompatibilityDay = pivotChartFilterDateCompatibilityDay
export const timelinePanelConditionBetween = (operator: TimelineDateFilterOperator): boolean => operator === 'between' || operator === 'notBetween'

/** Retained state is readable independently of a live source query; fixed conditions stay fixed conditions. */
export function readPanelTimelineDateFilter(studio: OfficeStudio, target: SlicerPanelTarget, timelineId: string): TimelineRangePanelSnapshot {
  const owner = captureSlicerPanelOwner(studio, target)
  const control = FOffice.newAPI(studio).getWorkbook(target.workbookId)?.getSheetBySheetId(target.sheetId)?.getTimeline(timelineId)
  if (!control) throw new SlicerPanelError('TIMELINE_NOT_FOUND', '所选时间线已删除，请重新选择。')
  const range = control.getDateRange(), condition = control.getDateCondition()
  if (range && condition) throw new SlicerPanelError('INVALID_TIMELINE_FILTER', '时间线不能同时保留范围和日期条件。')
  const state: TimelineDateFilterState = range ? { range: normalizeTimelineDateRange(range) }
    : condition ? { condition: normalizeTimelineDateCondition(condition) } : {}
  if (!slicerPanelOwnerCurrent(owner)) throw new SlicerPanelError('STALE_TIMELINE_FILTER', '时间线内容已更改，请重新读取日期筛选。')
  return { owner, timelineId, cacheId: control.getCacheId(), date1904: studio.getWorkbookModel(target.workbookId)?.date1904 === true, state: cloneModel(state) }
}

export function timelineRangePanelCurrent(snapshot: TimelineRangePanelSnapshot, published = false): boolean {
  if (!slicerPanelOwnerCurrent(snapshot.owner, published)) return false
  const { studio, target } = snapshot.owner
  try {
    const control = FOffice.newAPI(studio).getWorkbook(target.workbookId)?.getSheetBySheetId(target.sheetId)?.getTimeline(snapshot.timelineId)
    return Boolean(control && control.getCacheId() === snapshot.cacheId)
  } catch { return false }
}

/** Inclusive typed original-source bounds are never reconstructed from rendered members or a selection endpoint. */
export async function readPanelTimelineDateBounds(snapshot: TimelineRangePanelSnapshot, isCurrent: () => boolean = () => true): Promise<TimelinePanelDateBounds> {
  const current = () => isCurrent() && timelineRangePanelCurrent(snapshot)
  if (!current()) throw new SlicerPanelError('STALE_TIMELINE_BOUNDS', '日期来源读取已过期，请重新读取。')
  const { studio, target } = snapshot.owner
  const control = FOffice.newAPI(studio).getWorkbook(target.workbookId)?.getSheetBySheetId(target.sheetId)?.getTimeline(snapshot.timelineId)
  if (!control) throw new SlicerPanelError('TIMELINE_NOT_FOUND', '所选时间线已删除，请重新选择。')
  const bounds = await control.getDateBounds()
  if (!current()) throw new SlicerPanelError('STALE_TIMELINE_BOUNDS', '日期来源已更改，旧日期边界不能用于当前筛选。')
  if ((bounds.minDate !== undefined && !Number.isFinite(bounds.minDate)) || (bounds.maxDate !== undefined && !Number.isFinite(bounds.maxDate))
    || (bounds.minDate !== undefined && bounds.maxDate !== undefined && bounds.minDate > bounds.maxDate)
    || bounds.dateSystem !== (snapshot.date1904 ? 'date1904' : 'date1900')) {
    throw new SlicerPanelError('INVALID_TIMELINE_BOUNDS', '日期来源未提供可靠的日期边界。')
  }
  return cloneModel(bounds)
}

export function createTimelineDateFilterDraft(snapshot: TimelineRangePanelSnapshot): TimelineDateFilterDraft {
  const range = snapshot.state.range, condition = snapshot.state.condition
  const values = condition ? Array.isArray(condition.expected) ? condition.expected : [condition.expected] : []
  return cloneModel({ mode: range ? 'range' : condition ? 'condition' : 'none', range: range ?? {},
    condition: { operator: condition?.operator ?? 'equal', ...(values[0] !== undefined ? { first: values[0] } : {}),
      ...(values[1] !== undefined ? { second: values[1] } : {}), ...(condition?.useWholeDay !== undefined ? { useWholeDay: condition.useWholeDay } : {}) } })
}

export function buildTimelinePanelDateFilter(draft: TimelineDateFilterDraft): TimelineDateFilterState {
  if (draft.mode === 'none') return {}
  if (draft.mode === 'range') return { range: normalizeTimelineDateRange(draft.range) }
  if (draft.mode !== 'condition') throw new SlicerPanelError('INVALID_TIMELINE_FILTER', '请选择日期范围、日期条件或全部日期。')
  const { operator, first, second, useWholeDay } = draft.condition
  return { condition: normalizeTimelineDateCondition({ operator, expected: timelinePanelConditionBetween(operator) ? [first, second] : first,
    ...(useWholeDay !== undefined ? { useWholeDay } : {}) }) }
}

export function timelinePanelDateFilterDirty(draft: TimelineDateFilterDraft, original: TimelineDateFilterState): boolean {
  try { return !equivalentModelData(buildTimelinePanelDateFilter(draft), original) } catch { return true }
}
export function timelinePanelDateFilterError(draft: TimelineDateFilterDraft): string | undefined {
  try { buildTimelinePanelDateFilter(draft); return undefined }
  catch (failure) { return failure instanceof Error ? failure.message : String(failure) }
}

/** A deliberate calendar-day action creates next-day midnight; it does not relabel inclusive maxDate as exclusive. */
export function fillTimelinePanelSourceDays(draft: TimelineDateFilterDraft, bounds: TimelinePanelDateBounds): TimelineDateFilterDraft {
  if (bounds.minDate === undefined || bounds.maxDate === undefined) throw new SlicerPanelError('TIMELINE_DATES_UNAVAILABLE', '来源没有可靠的日期，请重新读取。')
  return { ...cloneModel(draft), mode: 'range', range: { startDate: Math.floor(bounds.minDate), endDate: Math.floor(bounds.maxDate) + 1 } }
}

export function timelinePanelDateFilterSummary(state: TimelineDateFilterState, date1904: boolean): string {
  const format = (value: number): string => timelinePanelDateCalendarInput(value, date1904) || String(value)
  if (state.range) return `从 ${format(state.range.startDate)} 起（含），到 ${format(state.range.endDate)} 前（不含）${state.range.useWholeDay ? '，按整日比较' : ''}`
  if (state.condition) {
    const { operator, expected, useWholeDay } = state.condition, label = TIMELINE_DATE_CONDITION_OPTIONS.find(option => option.value === operator)?.label ?? operator
    return `${label} ${Array.isArray(expected) ? expected.map(format).join(' 至 ') : format(expected)}${useWholeDay ? '，按整日比较' : ''}`
  }
  return '全部日期'
}

/** The condition intent shares Pivot's date engine; the fixed public Timeline Facade intentionally has no condition setter. */
export async function applyPanelTimelineDateFilter(snapshot: TimelineRangePanelSnapshot, draft: TimelineDateFilterDraft,
  isCurrent: () => boolean = () => true, prepared?: (command: TimelinePanelPreparedCommand) => void): Promise<boolean> {
  const current = () => isCurrent() && timelineRangePanelCurrent(snapshot)
  if (!current()) throw new SlicerPanelError('STALE_TIMELINE_FILTER', '日期筛选已过期，请重新读取后再保存。')
  const { studio, target } = snapshot.owner
  if (!canConfigurePanelSlicer(studio, target, snapshot.timelineId)) return false
  const control = FOffice.newAPI(studio).getWorkbook(target.workbookId)?.getSheetBySheetId(target.sheetId)?.getTimeline(snapshot.timelineId)
  if (!control) throw new SlicerPanelError('TIMELINE_NOT_FOUND', '所选时间线已删除，请重新选择。')
  const state = buildTimelinePanelDateFilter(draft), context = slicerPanelOwnerContext(snapshot.owner)
  if (state.range) {
    const payload = state.range, id = SLICER_PUBLIC_COMMANDS.dateRange
    prepared?.({ id, payloadKey: 'range', payload })
    return withSlicerPanelCommandOwner(snapshot.owner, snapshot.timelineId, id, payload, current, () => control.setDateRange(payload, context), 'range')
  }
  if (state.condition) {
    const payload: ITimelineDateCondition = state.condition, id = SLICER_PUBLIC_COMMANDS.dateCondition
    const params = { unitId: target.workbookId, slicerId: snapshot.timelineId, condition: payload, ...context }
    prepared?.({ id, payloadKey: 'condition', payload })
    return withSlicerPanelCommandOwner(snapshot.owner, snapshot.timelineId, id, payload, current, () => studio.executeCommand(id, params), 'condition')
  }
  // Undefined has no object identity. Match this exact shared-command params object so a late clear cannot cancel a newer clear.
  const id = SLICER_PUBLIC_COMMANDS.dateRange, params = { unitId: target.workbookId, slicerId: snapshot.timelineId, range: undefined, ...context }
  prepared?.({ id, payloadKey: 'params', payload: params })
  return withSlicerPanelCommandOwner(snapshot.owner, snapshot.timelineId, id, params, current, () => studio.executeCommand(id, params), 'params')
}
