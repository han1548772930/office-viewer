import type { OutlineAxis, SheetOutlineGroup, SheetOutlineState, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { assertOutlineAxis, createSheetOutlineState, validateSheetOutlineState } from './sheet-outline.ts'
import { DimensionOutlineErrorReason, SheetOutlineError, type IDimensionOutline, type SheetDimensionOutlineResource,
  type SheetOutlineScope, type SheetsOutlineResource } from './sheet-outline-types.ts'

const object = (value: unknown): value is Record<string, unknown> => Boolean(value && typeof value === 'object' && !Array.isArray(value))
export function sheetOutlineScope(workbook: WorkbookModel, sheet: SpreadsheetSheet): SheetOutlineScope {
  if (!workbook.id || !sheet.id) throw new SheetOutlineError(DimensionOutlineErrorReason.INVALID_RANGE, '大纲快照需要工作簿和工作表 ID。')
  return { unitId: workbook.id, subUnitId: sheet.id, rowCount: sheet.rowCount ?? 1_048_576, columnCount: sheet.columnCount ?? 16_384 }
}
export function getSheetDimensionOutlines(state: SheetOutlineState | undefined, scope: SheetOutlineScope, axis?: OutlineAxis): IDimensionOutline[] {
  const current = state ?? createSheetOutlineState(); validateSheetOutlineState(current, scope)
  if (axis !== undefined) assertOutlineAxis(axis)
  return current.groups.filter(group => axis === undefined || group.axis === axis).map(group => ({ ...cloneModel(group), unitId: scope.unitId, subUnitId: scope.subUnitId }))
}
/** Facade add uses a quantity; core commands and clear use inclusive ranges. */
export function sheetOutlinePublicAddRange(axis: OutlineAxis, start: number, count: number, axisCount: number) {
  assertOutlineAxis(axis)
  if (!Number.isSafeInteger(start) || start < 0 || !Number.isSafeInteger(count) || count <= 0 || !Number.isSafeInteger(start + count - 1)) {
    throw new SheetOutlineError(DimensionOutlineErrorReason.INVALID_RANGE, '大纲起点和数量必须是安全整数。')
  }
  if (!Number.isSafeInteger(axisCount) || axisCount < 1 || start + count > axisCount) {
    throw new SheetOutlineError(DimensionOutlineErrorReason.OUT_OF_BOUNDS, '大纲范围超出工作表。')
  }
  return { axis, start, end: start + count - 1 }
}
export function serializeSheetOutlineResource(state: SheetOutlineState | undefined, scope: SheetOutlineScope): SheetDimensionOutlineResource {
  const current = state ?? createSheetOutlineState()
  return { outlineList: getSheetDimensionOutlines(current, scope),
    officeStudio: { schemaVersion: 1, summaryAbove: current.summaryAbove, summaryLeft: current.summaryLeft } }
}
export function deserializeSheetOutlineResource(raw: unknown, scope: SheetOutlineScope): SheetOutlineState {
  if (!object(raw) || !Array.isArray(raw.outlineList)) throw new SheetOutlineError(DimensionOutlineErrorReason.INVALID_RANGE, '大纲资源缺少 outlineList。')
  const groups: SheetOutlineGroup[] = raw.outlineList.map(value => {
    if (!object(value) || typeof value.id !== 'string' || !value.id.trim() || typeof value.collapsed !== 'boolean'
      || typeof value.start !== 'number' || typeof value.end !== 'number'
      || value.unitId !== undefined && typeof value.unitId !== 'string' || value.subUnitId !== undefined && typeof value.subUnitId !== 'string') {
      throw new SheetOutlineError(DimensionOutlineErrorReason.INVALID_RANGE, '大纲资源成员无效。')
    }
    assertOutlineAxis(value.axis)
    // Rebind ownership from the destination model. Preserve unknown group
    // fields in the canonical object instead of silently discarding them.
    const preserved = cloneModel(value); delete preserved.unitId; delete preserved.subUnitId
    return { ...preserved, id: value.id, axis: value.axis, start: value.start, end: value.end, collapsed: value.collapsed }
  })
  const extension = raw.officeStudio
  if (extension !== undefined && (!object(extension) || extension.schemaVersion !== 1
    || typeof extension.summaryAbove !== 'boolean' || typeof extension.summaryLeft !== 'boolean')) {
    throw new SheetOutlineError(DimensionOutlineErrorReason.INVALID_RANGE, '本地大纲汇总位置资源无效。')
  }
  const state: SheetOutlineState = { groups, summaryAbove: object(extension) && extension.summaryAbove === true,
    summaryLeft: object(extension) && extension.summaryLeft === true }
  validateSheetOutlineState(state, scope)
  return state
}
export function serializeWorkbookOutlineResource(workbook: WorkbookModel): SheetsOutlineResource {
  return Object.fromEntries(workbook.sheets.filter(sheet => sheet.outline !== undefined).map(sheet =>
    [sheet.id, serializeSheetOutlineResource(sheet.outline, sheetOutlineScope(workbook, sheet))]))
}
/** Parse the entire plugin candidate before publishing any sheet state. An
 * explicit empty plugin resource clears outlines on every current sheet.
 */
export function deserializeWorkbookOutlineResource(workbook: WorkbookModel, raw: unknown): Record<string, SheetOutlineState> {
  let value: unknown = raw
  if (typeof raw === 'string') {
    try { value = JSON.parse(raw) } catch { throw new SheetOutlineError(DimensionOutlineErrorReason.INVALID_RANGE, '大纲资源不是有效 JSON。') }
  }
  if (!object(value)) throw new SheetOutlineError(DimensionOutlineErrorReason.INVALID_RANGE, '大纲插件资源必须是工作表映射。')
  const sheets = new Set(workbook.sheets.map(sheet => sheet.id))
  if (Object.keys(value).some(id => !sheets.has(id))) throw new SheetOutlineError(DimensionOutlineErrorReason.OUT_OF_BOUNDS, '大纲资源引用了不存在的工作表。')
  return Object.fromEntries(workbook.sheets.map(sheet => [sheet.id,
    Object.hasOwn(value, sheet.id) ? deserializeSheetOutlineResource(value[sheet.id], sheetOutlineScope(workbook, sheet)) : createSheetOutlineState()]))
}
