import { FOffice } from '../../api/facade.ts'
import type { OfficeStudio } from '../../api/office-studio.ts'
import type { SlicerModel } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { equivalentModelData } from '../../platform/commands/model-equivalence.ts'
import type { SlicerPresentationPatch } from './slicer-engine.ts'
import { buildPanelSlicerPresentationPatch, canConfigurePanelSlicer, SlicerPanelError, type SlicerPanelTarget } from './slicer-panel-state.ts'
import { captureSlicerPanelOwner, slicerPanelOwnerContext, slicerPanelOwnerCurrent, type SlicerPanelOwner } from './slicer-panel-owner.ts'
import { assertSlicerSettingsAvailable, normalizeSlicerSettingsPatch, type SheetSlicerSettings } from './slicer-settings.ts'
import { withSlicerPanelCommandOwner } from './slicer-panel-command-owner.ts'

/** Only text input is normalized; absent optional author settings remain absent. */
export type SlicerPanelEditDraft = SlicerModel & { name: string; caption: string; captionFromName: boolean }
export type SlicerPanelGeometryPatch = Pick<SlicerPresentationPatch, 'layout' | 'anchorType' | 'keepRatio'>
export interface SlicerSettingsPanelSnapshot { owner: SlicerPanelOwner; slicerId: string; cacheId: string; settings: SheetSlicerSettings }

export function createSlicerPanelEditDraft(info: SlicerModel, settings?: SheetSlicerSettings): SlicerPanelEditDraft {
  const name = settings?.name ?? info.name ?? info.caption ?? info.id
  return cloneModel({ ...info, ...(settings ?? {}), name, caption: info.caption ?? info.name ?? '',
    captionFromName: info.caption === undefined,
    ...(info.type === 'slicer' ? { rowHeight: settings?.rowHeight ?? info.rowHeight ?? 24 } : {}) })
}

export function readPanelSlicerSettings(studio: OfficeStudio, target: SlicerPanelTarget, slicerId: string): SlicerSettingsPanelSnapshot {
  const owner = captureSlicerPanelOwner(studio, target)
  const control = FOffice.newAPI(studio).getWorkbook(target.workbookId)?.getSheetBySheetId(target.sheetId)?.getSlicer(slicerId)
  const settings = control?.getSettings(), cacheId = control?.getCacheId()
  if (!settings || !cacheId) throw new SlicerPanelError('SLICER_SETTINGS_UNAVAILABLE', '切片器显示设置或来源不可用，请重新选择。')
  if (!slicerPanelOwnerCurrent(owner)) throw new SlicerPanelError('STALE_SLICER_SETTINGS', '切片器内容已更改，请重新读取设置。')
  return { owner, slicerId, cacheId, settings: cloneModel(settings) }
}

/** Comparing author presence keeps explicit defaults and own undefined distinct. Validation happens before dispatch. */
export function buildPanelSlicerSettingsPatch(current: SlicerPanelEditDraft, original: SlicerPanelEditDraft): Partial<SheetSlicerSettings> {
  const patch: Partial<SheetSlicerSettings> = {}
  for (const key of ['name', 'rowHeight', 'startItem', 'columnCount', 'showCaption', 'lockedPosition', 'style'] as const) {
    if (!equivalentModelData(current[key], original[key])) Object.assign(patch, { [key]: current[key] })
  }
  const caption = current.captionFromName ? undefined : current.caption
  const priorCaption = original.captionFromName ? undefined : original.caption
  if (!equivalentModelData(caption, priorCaption)) Object.assign(patch, { caption })
  return patch
}

export function buildPanelSlicerGeometryPatch(current: SlicerPanelEditDraft, original: SlicerPanelEditDraft): SlicerPanelGeometryPatch {
  const presentation = buildPanelSlicerPresentationPatch(current, original), patch: SlicerPanelGeometryPatch = {}
  for (const key of ['layout', 'anchorType', 'keepRatio'] as const) if (Object.hasOwn(presentation, key)) Object.assign(patch, { [key]: presentation[key] })
  return patch
}

export function slicerSettingsPanelCurrent(snapshot: SlicerSettingsPanelSnapshot, published = false): boolean {
  if (!slicerPanelOwnerCurrent(snapshot.owner, published)) return false
  const { studio, target } = snapshot.owner
  try { return FOffice.newAPI(studio).getWorkbook(target.workbookId)?.getSheetBySheetId(target.sheetId)?.getSlicer(snapshot.slicerId)?.getCacheId() === snapshot.cacheId }
  catch { return false }
}

export function slicerSettingsPanelDraftError(snapshot: SlicerSettingsPanelSnapshot, current: SlicerPanelEditDraft, original: SlicerPanelEditDraft): string | undefined {
  try {
    normalizeSlicerSettingsPatch({ name: current.name, rowHeight: current.rowHeight, caption: current.captionFromName ? undefined : current.caption,
      startItem: current.startItem, columnCount: current.columnCount, showCaption: current.showCaption, lockedPosition: current.lockedPosition })
    const workbook = snapshot.owner.studio.getWorkbookModel(snapshot.owner.target.workbookId), control = workbook?.slicers?.find(slicer => slicer.id === snapshot.slicerId)
    if (!workbook || !control) return '所选切片器已删除，请重新选择。'
    assertSlicerSettingsAvailable(workbook, control, normalizeSlicerSettingsPatch(buildPanelSlicerSettingsPatch(current, original)))
    return undefined
  } catch (failure) { return failure instanceof Error ? failure.message : String(failure) }
}

/** Pure settings use the public setter. A mixed geometry edit remains one existing content transaction. */
export async function applyPanelSlicerSettings(snapshot: SlicerSettingsPanelSnapshot, input: Partial<SheetSlicerSettings>, geometry: SlicerPanelGeometryPatch = {}, isCurrent: () => boolean = () => true): Promise<boolean> {
  const current = () => isCurrent() && slicerSettingsPanelCurrent(snapshot)
  if (!current()) throw new SlicerPanelError('STALE_SLICER_SETTINGS', '切片器设置已过期，请重新读取后再保存。')
  const { studio, target } = snapshot.owner
  if (!canConfigurePanelSlicer(studio, target, snapshot.slicerId)) return false
  const workbook = studio.getWorkbookModel(target.workbookId), model = workbook?.slicers?.find(control => control.id === snapshot.slicerId)
  const control = FOffice.newAPI(studio).getWorkbook(target.workbookId)?.getSheetBySheetId(target.sheetId)?.getSlicer(snapshot.slicerId)
  if (!workbook || !model || !control) throw new SlicerPanelError('SLICER_NOT_FOUND', '所选切片器已删除，请重新选择。')
  const patch = normalizeSlicerSettingsPatch(input)
  assertSlicerSettingsAvailable(workbook, model, patch)
  const context = slicerPanelOwnerContext(snapshot.owner)
  if (Object.keys(geometry).length) {
    const combined = { ...patch, ...cloneModel(geometry) }
    await withSlicerPanelCommandOwner(snapshot.owner, snapshot.slicerId, 'sheet.command.update-slicer-presentation', combined, current,
      () => control.updatePresentation(combined, context))
    return true
  }
  return withSlicerPanelCommandOwner(snapshot.owner, snapshot.slicerId, 'sheet.command.update-slicer-settings', patch, current,
    () => control.updateSettings(patch, context))
}
