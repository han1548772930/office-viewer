import type { SlicerModel, TimelineLevel, WorkbookModel } from '../../domain/model.ts'
import type { TimelinePublicLevel } from '../../domain/timeline-filters.ts'
import { equivalentModelData } from '../../platform/commands/model-equivalence.ts'
import { SlicerConnectionError } from './slicer-connections.ts'
import { hasSlicerStyle } from './slicer-style-catalog.ts'
import { normalizeTimelineLevel, toTimelinePublicLevel } from './timeline-date-filter.ts'

/** Public grains use the fixed owner's numeric values; per-view serial navigation is not an item/px offset. */
export interface SheetTimelineSettings {
  name: string
  level: TimelinePublicLevel
  selectionLevel: TimelinePublicLevel
  caption?: string
  style?: string
  scrollPosition?: number
  showHeader?: boolean
  showSelectionLabel?: boolean
  showTimeLevel?: boolean
  showHorizontalScrollbar?: boolean
}
export type TimelineSettingsModelPatch = Partial<Omit<SheetTimelineSettings, 'level' | 'selectionLevel'> & { level: TimelineLevel; selectionLevel: TimelineLevel }>
export type TimelineSettingsInputPatch = Partial<Omit<SheetTimelineSettings, 'level' | 'selectionLevel'>
  & { level: TimelinePublicLevel | TimelineLevel; selectionLevel: TimelinePublicLevel | TimelineLevel }>
const keys = ['name', 'level', 'selectionLevel', 'caption', 'style', 'scrollPosition', 'showHeader',
  'showSelectionLabel', 'showTimeLevel', 'showHorizontalScrollbar'] as const
const required = new Set<string>(['name', 'level', 'selectionLevel'])
const invalid = (message: string): never => { throw new SlicerConnectionError('invalid-target', message) }

export function normalizeTimelineSettingsPatch(input: unknown): TimelineSettingsModelPatch {
  if (!input || typeof input !== 'object' || Array.isArray(input)) return invalid('时间线显示设置无效。')
  const value = input as Record<string, unknown>
  if (Object.keys(value).some(key => !(keys as readonly string[]).includes(key))) return invalid('时间线显示设置包含未知属性。')
  const result: TimelineSettingsModelPatch = {}
  for (const key of keys) if (Object.hasOwn(value, key)) {
    let current = value[key]
    if (current === undefined) {
      if (required.has(key)) return invalid(`时间线 ${key} 不能清除。`)
    } else if (key === 'name' || key === 'caption' || key === 'style') {
      if (typeof current !== 'string' || !current.trim()) return invalid('时间线名称、标题或样式不能为空。')
      if (key === 'caption' && current.length > 32767 || key === 'style' && current.length > 255) return invalid('时间线标题或样式名称过长。')
    } else if (key === 'level' || key === 'selectionLevel') {
      try { current = normalizeTimelineLevel(current) } catch (error) {
        if (!(error instanceof TypeError)) throw error
        return invalid('时间线显示或选择粒度无效。')
      }
    } else if (key === 'scrollPosition') {
      if (typeof current !== 'number' || !Number.isFinite(current)) return invalid('时间线导航位置必须为有限 Excel 序列值。')
    } else if (typeof current !== 'boolean') return invalid('时间线显示开关必须是布尔值。')
    Object.assign(result, { [key]: current })
  }
  return result
}
export function readTimelineSettings(slicer: SlicerModel): SheetTimelineSettings {
  if (slicer.type !== 'timeline') return invalid('普通切片器使用独立的显示设置。')
  const result: SheetTimelineSettings = { name: slicer.name ?? slicer.caption ?? slicer.id,
    level: toTimelinePublicLevel(slicer.level ?? 'month'), selectionLevel: toTimelinePublicLevel(slicer.selectionLevel ?? 'month') }
  for (const key of keys) if (!required.has(key) && slicer[key] !== undefined) Object.assign(result, { [key]: slicer[key] })
  return result
}
export function assertTimelineSettingsAvailable(workbook: WorkbookModel, slicer: SlicerModel, patch: TimelineSettingsModelPatch): void {
  if (slicer.type !== 'timeline') return invalid('普通切片器使用独立的显示设置。')
  if (patch.name !== undefined && (workbook.slicers ?? []).some(control => control.id !== slicer.id
    && (control.name ?? control.caption ?? control.id).trim().toLocaleLowerCase() === patch.name!.trim().toLocaleLowerCase())) return invalid('工作簿中已存在同名的切片器或时间线。')
  if (patch.style !== undefined && !hasSlicerStyle(workbook, 'timeline', patch.style)) return invalid('指定的时间线样式不存在。')
}
/** Only authored view properties change; the command host handles history/permission/cancellation at execution. */
export function patchTimelineSettings(workbook: WorkbookModel, slicer: SlicerModel, input: TimelineSettingsInputPatch): boolean {
  const patch = normalizeTimelineSettingsPatch(input); assertTimelineSettingsAvailable(workbook, slicer, patch)
  const next = { ...slicer }
  for (const key of keys) if (Object.hasOwn(patch, key)) {
    if (patch[key] === undefined) delete next[key]
    else Object.assign(next, { [key]: patch[key] })
  }
  if (equivalentModelData(slicer, next)) return false
  for (const key of keys) if (Object.hasOwn(patch, key)) {
    if (patch[key] === undefined) delete slicer[key]
    else Object.assign(slicer, { [key]: patch[key] })
  }
  slicer.revision += 1
  return true
}
