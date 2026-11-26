import type { WorkbookModel } from '../domain/model.ts'
import { deserializeWorkbookOutlineResource } from '../features/sheet-outline/sheet-outline-resources.ts'
import { SHEET_OUTLINE_PLUGIN } from '../features/sheet-outline/sheet-outline-types.ts'

/** The fixed plugin is a sheet-id map; parse the complete candidate before applying it. */
export function importWorkbookOutlineResources(workbook: WorkbookModel, raw: unknown): void {
  if (raw === undefined) return
  if (!Array.isArray(raw)) throw new TypeError('工作簿插件资源必须是数组。')
  const matches = raw.filter((resource): resource is { name: string; data: unknown } => Boolean(resource && typeof resource === 'object'
    && 'name' in resource && resource.name === SHEET_OUTLINE_PLUGIN))
  if (matches.length > 1) throw new TypeError('工作簿包含重复的大纲插件资源。')
  if (!matches[0]) return
  const outlines = deserializeWorkbookOutlineResource(workbook, matches[0].data)
  for (const sheet of workbook.sheets) sheet.outline = outlines[sheet.id]
}
