import type { SheetOutlineGroup, SheetOutlineState } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'

export enum DimensionOutlineAxis { ROW = 'row', COLUMN = 'column' }
export enum DimensionOutlineErrorReason {
  INVALID_RANGE = 'invalid-range', OUT_OF_BOUNDS = 'out-of-bounds', CROSSING = 'crossing', MAX_DEPTH = 'max-depth',
  MOVE_SPLITS_OUTLINE = 'move-splits-outline', CLEAR_RANGE_NOT_CONTAIN_OUTLINE = 'clear-range-not-contain-outline', UNKNOWN = 'unknown',
}
export class SheetOutlineError extends Error {
  readonly code: string
  constructor(readonly reason: DimensionOutlineErrorReason, message: string, readonly outlineId?: string) {
    super(message); this.name = 'SheetOutlineError'; this.code = `SHEET_OUTLINE_${reason.replaceAll('-', '_').toUpperCase()}`
  }
}
export interface IDimensionOutline extends SheetOutlineGroup { unitId: string; subUnitId: string }
export interface SheetOutlineExtent { rowCount: number; columnCount: number }
export interface SheetOutlineScope extends SheetOutlineExtent { unitId: string; subUnitId: string }
export interface SheetOutlineHiddenRange { start: number; end: number }
export interface SheetOutlineTreeNode extends SheetOutlineGroup { depth: number; parentId?: string; children: SheetOutlineTreeNode[] }
export interface SheetDimensionOutlineResource {
  outlineList: IDimensionOutline[]
  /** Local extension: the fixed plugin resource does not expose summary direction. */
  officeStudio?: { schemaVersion: 1; summaryAbove: boolean; summaryLeft: boolean }
}
export type SheetsOutlineResource = Record<string, SheetDimensionOutlineResource>
export const MAX_SHEET_OUTLINE_DEPTH = 8
export const SHEET_OUTLINE_PLUGIN = 'SHEET_OUTLINE_PLUGIN'

export function sheetOutlineErrorMessage(error: unknown): string {
  if (!(error instanceof SheetOutlineError)) return error instanceof Error ? error.message : '大纲分组操作失败。'
  const messages: Record<DimensionOutlineErrorReason, string> = {
    'invalid-range': '分组范围须使用非负整数，且起点不能超过终点。', 'out-of-bounds': '分组范围超出工作表行列边界。',
    crossing: '分组范围必须嵌套或彼此分离，不能交叉。', 'max-depth': '行列大纲最多支持八个嵌套分组。',
    'move-splits-outline': '当前移动会拆散大纲组，请移动完整组或在组内调整。',
    'clear-range-not-contain-outline': '选区内没有可以取消的大纲组。', unknown: '大纲分组操作失败。',
  }
  return messages[error.reason]
}
export function isIgnoredOutlineAddError(error: unknown): error is SheetOutlineError {
  return error instanceof SheetOutlineError && [DimensionOutlineErrorReason.INVALID_RANGE, DimensionOutlineErrorReason.OUT_OF_BOUNDS,
    DimensionOutlineErrorReason.CROSSING, DimensionOutlineErrorReason.MAX_DEPTH].includes(error.reason)
}
export function cloneSheetOutlineState(state: SheetOutlineState, allocateId?: () => string): SheetOutlineState {
  const ids = new Set<string>()
  const snapshot = cloneModel(state)
  return { ...snapshot, groups: snapshot.groups.map(group => {
    const id = allocateId ? allocateId() : group.id
    if (typeof id !== 'string' || !id.trim() || ids.has(id)) throw new SheetOutlineError(DimensionOutlineErrorReason.INVALID_RANGE, '分组 ID 无效或重复。', id)
    ids.add(id); return { ...group, id }
  }) }
}
