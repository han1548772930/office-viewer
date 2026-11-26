import type { SlicerModel, WorkbookModel } from '../../domain/model.ts'
import { SlicerConnectionError } from './slicer-connections.ts'
import { hasSlicerStyle } from './slicer-style-catalog.ts'

/** Public view dimensions are CSS pixels. Resource EMU precision is retained by the adapter. */
export interface SheetSlicerSettings {
  name: string
  rowHeight: number
  caption?: string
  startItem?: number
  columnCount?: number
  showCaption?: boolean
  lockedPosition?: boolean
  style?: string
}

const keys = ['name', 'rowHeight', 'caption', 'startItem', 'columnCount', 'showCaption', 'lockedPosition', 'style'] as const
const optional = new Set<string>(keys.filter(key => key !== 'name' && key !== 'rowHeight'))
const invalid = (message: string): never => { throw new SlicerConnectionError('invalid-target', message) }

/** An own undefined clears only an optional property; unrelated queued edits are never copied into the patch. */
export function normalizeSlicerSettingsPatch(input: unknown): Partial<SheetSlicerSettings> {
  if (!input || typeof input !== 'object' || Array.isArray(input)) return invalid('切片器显示设置无效。')
  const value = input as Record<string, unknown>
  if (Object.keys(value).some(key => !(keys as readonly string[]).includes(key))) return invalid('切片器显示设置包含未知属性。')
  const result: Partial<SheetSlicerSettings> = {}
  for (const key of keys) if (Object.hasOwn(value, key)) {
    let current = value[key]
    if (current === undefined) {
      if (!optional.has(key)) return invalid(`切片器 ${key} 不能清除。`)
    } else if (key === 'name' || key === 'caption' || key === 'style') {
      if (typeof current !== 'string' || !current.trim()) return invalid('切片器名称、标题或样式不能为空。')
      if (key === 'caption' && current.length > 32767 || key === 'style' && current.length > 255) return invalid('切片器标题或样式名称过长。')
    } else if (key === 'rowHeight') {
      if (typeof current !== 'number' || !Number.isFinite(current) || current <= 0
        || Math.round(current * 9525) > 4294967295 || Math.round(current * 9525) < 1) return invalid('切片器按钮高度超出范围。')
      current = Math.round(current * 9525) / 9525
    } else if (key === 'startItem' || key === 'columnCount') {
      if (typeof current !== 'number' || !Number.isSafeInteger(current)
        || current < (key === 'startItem' ? 0 : 1) || current > (key === 'startItem' ? 4294967295 : 20000)) return invalid('切片器起始项或列数超出范围。')
    } else if (typeof current !== 'boolean') return invalid('切片器显示开关必须是布尔值。')
    Object.assign(result, { [key]: current })
  }
  return result
}

export function readSlicerSettings(slicer: SlicerModel): SheetSlicerSettings {
  if (slicer.type !== 'slicer') return invalid('时间线使用独立的显示设置。')
  const result: SheetSlicerSettings = { name: slicer.name ?? slicer.caption ?? slicer.id, rowHeight: slicer.rowHeight ?? 24 }
  for (const key of keys) if (optional.has(key) && slicer[key] !== undefined) Object.assign(result, { [key]: slicer[key] })
  return result
}

export function assertSlicerSettingsAvailable(workbook: WorkbookModel, slicer: SlicerModel, patch: Partial<SheetSlicerSettings>): void {
  if (slicer.type !== 'slicer') return invalid('时间线使用独立的显示设置。')
  if (patch.name !== undefined && (workbook.slicers ?? []).some(control => control.id !== slicer.id
    && (control.name ?? control.caption ?? control.id).trim().toLowerCase() === patch.name!.trim().toLowerCase())) return invalid('工作簿中已存在同名的切片器或时间线。')
  if (patch.style !== undefined && !hasSlicerStyle(workbook, 'slicer', patch.style)) return invalid('指定的切片器样式不存在。')
}

/** Default names remain unique when several views share a cache. */
export function resolveSlicerViewName(workbook: WorkbookModel, name: string, explicit = false): string {
  const names = new Set((workbook.slicers ?? []).map(control => (control.name ?? control.caption ?? control.id).trim().toLowerCase()))
  if (!name.trim()) return invalid('切片器或时间线名称不能为空。')
  if (!names.has(name.trim().toLowerCase())) return name
  if (explicit) return invalid('工作簿中已存在同名的切片器或时间线。')
  for (let index = 2; ; index++) {
    const candidate = `${name} ${index}`
    if (!names.has(candidate.toLowerCase())) return candidate
  }
}
