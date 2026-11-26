import { FOffice } from '../../api/facade.ts'
import type { OfficeStudio } from '../../api/office-studio.ts'
import type { SheetSlicerCacheOptions } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { equivalentModelData } from '../../platform/commands/model-equivalence.ts'
import { normalizeSlicerCacheOptionsPatch } from './slicer-cache-options.ts'
import { canConfigurePanelSlicer, SlicerPanelError, type SlicerPanelTarget } from './slicer-panel-state.ts'
import { captureSlicerPanelOwner, slicerPanelOwnerContext, slicerPanelOwnerCurrent, type SlicerPanelOwner } from './slicer-panel-owner.ts'
import { withSlicerPanelCommandOwner } from './slicer-panel-command-owner.ts'

export interface SlicerCacheOptionsPanelSnapshot { owner: SlicerPanelOwner; slicerId: string; cacheId: string; options: SheetSlicerCacheOptions }
export const SLICER_CACHE_OPTIONS_DEFAULT = '__document_default__'

export function readPanelSlicerCacheOptions(studio: OfficeStudio, target: SlicerPanelTarget, slicerId: string): SlicerCacheOptionsPanelSnapshot {
  const owner = captureSlicerPanelOwner(studio, target)
  const control = FOffice.newAPI(studio).getWorkbook(target.workbookId)?.getSheetBySheetId(target.sheetId)?.getSlicer(slicerId)
  const options = control?.getCacheOptions(), cacheId = control?.getCacheId()
  if (!options || !cacheId) throw new SlicerPanelError('SLICER_CACHE_OPTIONS_UNAVAILABLE', '切片器排序设置或来源不可用；时间线使用独立的时间设置。')
  if (!slicerPanelOwnerCurrent(owner)) throw new SlicerPanelError('STALE_SLICER_CACHE_OPTIONS', '切片器内容已更改，请重新读取排序设置。')
  return { owner, slicerId, cacheId, options: cloneModel(options) }
}

export function buildPanelSlicerCacheOptionsPatch(current: SheetSlicerCacheOptions, original: SheetSlicerCacheOptions): Partial<SheetSlicerCacheOptions> {
  const patch: Partial<SheetSlicerCacheOptions> = {}
  for (const key of ['sortOrder', 'customListSort', 'crossFilter', 'hideItemsWithNoData'] as const) {
    if (!equivalentModelData(current[key], original[key])) Object.assign(patch, { [key]: current[key] })
  }
  return patch
}

/** Select defaults are UI sentinels; only a real optional-property clear enters the command DTO. */
export function setPanelSlicerCacheOption(options: SheetSlicerCacheOptions, key: keyof SheetSlicerCacheOptions, value: unknown): SheetSlicerCacheOptions {
  const normalized = value === SLICER_CACHE_OPTIONS_DEFAULT ? undefined
    : (key === 'customListSort' || key === 'hideItemsWithNoData') && (value === 'true' || value === 'false') ? value === 'true' : value
  const patch = normalizeSlicerCacheOptionsPatch({ [key]: normalized }), next = cloneModel(options)
  if (patch[key] === undefined) delete next[key]
  else Object.assign(next, patch)
  return next
}

export function slicerCacheOptionsPanelCurrent(snapshot: SlicerCacheOptionsPanelSnapshot, published = false): boolean {
  if (!slicerPanelOwnerCurrent(snapshot.owner, published)) return false
  const { studio, target } = snapshot.owner
  try { return FOffice.newAPI(studio).getWorkbook(target.workbookId)?.getSheetBySheetId(target.sheetId)?.getSlicer(snapshot.slicerId)?.getCacheId() === snapshot.cacheId }
  catch { return false }
}

export async function applyPanelSlicerCacheOptions(snapshot: SlicerCacheOptionsPanelSnapshot, input: Partial<SheetSlicerCacheOptions>,
  isCurrent: () => boolean = () => true, prepared?: (patch: Partial<SheetSlicerCacheOptions>) => void): Promise<boolean> {
  const current = () => isCurrent() && slicerCacheOptionsPanelCurrent(snapshot)
  if (!current()) throw new SlicerPanelError('STALE_SLICER_CACHE_OPTIONS', '排序设置已过期，请重新读取后再保存。')
  const { studio, target } = snapshot.owner
  if (!canConfigurePanelSlicer(studio, target, snapshot.slicerId)) return false
  const control = FOffice.newAPI(studio).getWorkbook(target.workbookId)?.getSheetBySheetId(target.sheetId)?.getSlicer(snapshot.slicerId)
  if (!control) throw new SlicerPanelError('SLICER_NOT_FOUND', '所选切片器已删除，请重新选择。')
  const patch = normalizeSlicerCacheOptionsPatch(input), context = slicerPanelOwnerContext(snapshot.owner)
  // Expose the exact command DTO to this dialog's lifecycle observer before a Facade wait.
  prepared?.(patch)
  return withSlicerPanelCommandOwner(snapshot.owner, snapshot.slicerId, 'sheet.command.patch-slicer-cache-options', patch, current,
    () => control.setCacheOptions(patch, context))
}
