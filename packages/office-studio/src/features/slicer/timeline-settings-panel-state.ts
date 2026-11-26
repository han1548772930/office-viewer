import { FOffice } from '../../api/facade.ts'
import type { OfficeStudio } from '../../api/office-studio.ts'
import type { SlicerModel } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { equivalentModelData } from '../../platform/commands/model-equivalence.ts'
import type { SlicerPresentationPatch } from './slicer-engine.ts'
import { buildPanelSlicerGeometryPatch, createSlicerPanelEditDraft, type SlicerPanelEditDraft } from './slicer-settings-panel-state.ts'
import { canConfigurePanelSlicer, SlicerPanelError, type SlicerPanelTarget } from './slicer-panel-state.ts'
import { captureSlicerPanelOwner, slicerPanelOwnerContext, slicerPanelOwnerCurrent, type SlicerPanelOwner } from './slicer-panel-owner.ts'
import { withSlicerPanelCommandOwner } from './slicer-panel-command-owner.ts'
import { normalizeTimelineLevel, toTimelinePublicLevel } from './timeline-date-filter.ts'
import { assertTimelineSettingsAvailable, normalizeTimelineSettingsPatch, type SheetTimelineSettings } from './timeline-settings.ts'
import { SLICER_PUBLIC_COMMANDS } from './slicer-public-info.ts'

export interface TimelineSettingsPanelSnapshot {
  owner: SlicerPanelOwner; timelineId: string; cacheId?: string; settings: SheetTimelineSettings
}
export type TimelinePanelGeometryPatch = Pick<SlicerPresentationPatch, 'layout' | 'anchorType' | 'keepRatio' | 'lockedPosition'>

/** Public numeric grains are adapted once into the local view draft; optional author keys stay optional. */
export function createTimelinePanelEditDraft(model: SlicerModel, settings: SheetTimelineSettings): SlicerPanelEditDraft {
  if (model.type !== 'timeline') throw new SlicerPanelError('INVALID_TIMELINE_TYPE', '请选择时间线设置。')
  return cloneModel({ ...createSlicerPanelEditDraft(model), name: settings.name,
    level: normalizeTimelineLevel(settings.level), selectionLevel: normalizeTimelineLevel(settings.selectionLevel),
    caption: settings.caption ?? settings.name, captionFromName: settings.caption === undefined })
}

export function readPanelTimelineSettings(studio: OfficeStudio, target: SlicerPanelTarget, timelineId: string): TimelineSettingsPanelSnapshot {
  const owner = captureSlicerPanelOwner(studio, target)
  const control = FOffice.newAPI(studio).getWorkbook(target.workbookId)?.getSheetBySheetId(target.sheetId)?.getTimeline(timelineId)
  const settings = control?.getSettings()
  if (!settings) throw new SlicerPanelError('TIMELINE_SETTINGS_UNAVAILABLE', '时间线显示设置不可用，请重新选择。')
  if (!slicerPanelOwnerCurrent(owner)) throw new SlicerPanelError('STALE_TIMELINE_SETTINGS', '时间线内容已更改，请重新读取设置。')
  return { owner, timelineId, cacheId: control?.getCacheId(), settings: cloneModel(settings) }
}

export function buildPanelTimelineSettingsPatch(current: SlicerPanelEditDraft, original: SlicerPanelEditDraft): Partial<SheetTimelineSettings> {
  const patch: Partial<SheetTimelineSettings> = {}
  for (const key of ['name', 'style', 'scrollPosition', 'showHeader', 'showSelectionLabel', 'showTimeLevel', 'showHorizontalScrollbar'] as const) {
    if (!equivalentModelData(current[key], original[key])) Object.assign(patch, { [key]: current[key] })
  }
  for (const key of ['level', 'selectionLevel'] as const) {
    if (!equivalentModelData(current[key], original[key])) Object.assign(patch, { [key]: toTimelinePublicLevel(normalizeTimelineLevel(current[key])) })
  }
  const caption = current.captionFromName ? undefined : current.caption, prior = original.captionFromName ? undefined : original.caption
  if (!equivalentModelData(caption, prior)) Object.assign(patch, { caption })
  return patch
}

export function buildPanelTimelineGeometryPatch(current: SlicerPanelEditDraft, original: SlicerPanelEditDraft): TimelinePanelGeometryPatch {
  const patch: TimelinePanelGeometryPatch = buildPanelSlicerGeometryPatch(current, original)
  if (!equivalentModelData(current.lockedPosition, original.lockedPosition)) patch.lockedPosition = current.lockedPosition
  return patch
}

export function timelineSettingsPanelCurrent(snapshot: TimelineSettingsPanelSnapshot, published = false): boolean {
  if (!slicerPanelOwnerCurrent(snapshot.owner, published)) return false
  const { studio, target } = snapshot.owner
  try {
    const control = FOffice.newAPI(studio).getWorkbook(target.workbookId)?.getSheetBySheetId(target.sheetId)?.getTimeline(snapshot.timelineId)
    return Boolean(control && control.getCacheId() === snapshot.cacheId)
  } catch { return false }
}

export function timelineSettingsPanelDraftError(snapshot: TimelineSettingsPanelSnapshot, current: SlicerPanelEditDraft, original: SlicerPanelEditDraft): string | undefined {
  try {
    normalizeTimelineSettingsPatch({ name: current.name, level: current.level, selectionLevel: current.selectionLevel,
      caption: current.captionFromName ? undefined : current.caption, scrollPosition: current.scrollPosition,
      showHeader: current.showHeader, showSelectionLabel: current.showSelectionLabel, showTimeLevel: current.showTimeLevel,
      showHorizontalScrollbar: current.showHorizontalScrollbar })
    const workbook = snapshot.owner.studio.getWorkbookModel(snapshot.owner.target.workbookId), model = workbook?.slicers?.find(view => view.id === snapshot.timelineId)
    if (!workbook || !model) return '所选时间线已删除，请重新选择。'
    assertTimelineSettingsAvailable(workbook, model, normalizeTimelineSettingsPatch(buildPanelTimelineSettingsPatch(current, original)))
    return undefined
  } catch (failure) { return failure instanceof Error ? failure.message : String(failure) }
}

/** Shared filtering is not a view setting. Mixed drawing edits remain one content transaction. */
export async function applyPanelTimelineSettings(snapshot: TimelineSettingsPanelSnapshot, input: Partial<SheetTimelineSettings>,
  geometry: TimelinePanelGeometryPatch = {}, isCurrent: () => boolean = () => true): Promise<boolean> {
  const current = () => isCurrent() && timelineSettingsPanelCurrent(snapshot)
  if (!current()) throw new SlicerPanelError('STALE_TIMELINE_SETTINGS', '时间线设置已过期，请重新读取后再保存。')
  const { studio, target } = snapshot.owner
  if (!canConfigurePanelSlicer(studio, target, snapshot.timelineId)) return false
  const workbook = studio.getWorkbookModel(target.workbookId), model = workbook?.slicers?.find(view => view.id === snapshot.timelineId)
  const control = FOffice.newAPI(studio).getWorkbook(target.workbookId)?.getSheetBySheetId(target.sheetId)?.getTimeline(snapshot.timelineId)
  if (!workbook || !model || !control) throw new SlicerPanelError('TIMELINE_NOT_FOUND', '所选时间线已删除，请重新选择。')
  const normalized = normalizeTimelineSettingsPatch(input)
  assertTimelineSettingsAvailable(workbook, model, normalized)
  const context = slicerPanelOwnerContext(snapshot.owner)
  if (Object.keys(geometry).length) {
    const combined = { ...normalized, ...cloneModel(geometry) }
    await withSlicerPanelCommandOwner(snapshot.owner, snapshot.timelineId, 'sheet.command.update-slicer-presentation', combined, current,
      () => control.updatePresentation(combined, context))
    return true
  }
  const patch: Partial<SheetTimelineSettings> = { ...input }
  for (const key of ['level', 'selectionLevel'] as const) if (Object.hasOwn(normalized, key)) {
    Object.assign(patch, { [key]: toTimelinePublicLevel(normalizeTimelineLevel(normalized[key])) })
  }
  return withSlicerPanelCommandOwner(snapshot.owner, snapshot.timelineId, SLICER_PUBLIC_COMMANDS.timelineSettings, patch, current,
    () => control.updateSettings(patch, context))
}
