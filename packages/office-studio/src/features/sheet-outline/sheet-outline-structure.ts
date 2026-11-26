import type { SheetOutlineGroup, SheetOutlineState } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { assertSheetAxisEdit, sheetAxisMaximum, sheetAxisInsertionEdit, transformSheetAxisInterval, type SheetAxisEdit } from '../../domain/sheet-axis-structure.ts'
import { validateSheetOutlineState } from './sheet-outline.ts'
import { DimensionOutlineErrorReason, SheetOutlineError, type SheetOutlineExtent } from './sheet-outline-types.ts'

/** Structural edits share the project axis permutation. An outline remains one
 * interval; a move that would split it rejects the entire detached candidate.
 * Copy is a local extension: only fully contained groups are cloned.
 */
export function transformSheetOutlineForAxisEdit(state: SheetOutlineState | undefined, edit: SheetAxisEdit,
  beforeExtent: SheetOutlineExtent, afterExtent?: SheetOutlineExtent, allocateId: () => string = () => crypto.randomUUID()): SheetOutlineState | undefined {
  assertSheetAxisEdit(edit)
  if (!state) return undefined
  validateSheetOutlineState(state, beforeExtent)
  const count = edit.axis === 'row' ? beforeExtent.rowCount : beforeExtent.columnCount
  if (edit.kind === 'insert' ? edit.start > count : edit.start + edit.count > count) {
    throw new SheetOutlineError(DimensionOutlineErrorReason.OUT_OF_BOUNDS, '大纲结构编辑超出源行列。')
  }
  if ((edit.kind === 'move' || edit.kind === 'copy') && edit.destination > count) {
    throw new SheetOutlineError(DimensionOutlineErrorReason.OUT_OF_BOUNDS, '大纲结构编辑超出目标行列。')
  }
  const nextCount = edit.kind === 'insert' || edit.kind === 'copy' ? Math.min(sheetAxisMaximum(edit.axis), count + edit.count)
    : edit.kind === 'delete' ? count - edit.count : count
  const extent = afterExtent ?? { ...beforeExtent, [edit.axis === 'row' ? 'rowCount' : 'columnCount']: nextCount }
  const insertion = sheetAxisInsertionEdit(edit), groups: SheetOutlineGroup[] = []
  for (const group of state.groups) {
    if (group.axis !== edit.axis) { groups.push(group); continue }
    // Do not silently drop a group at the Office axis limit when the shared
    // range mapper clips overflowing coordinates.
    if (insertion?.kind === 'insert' && insertion.start <= group.end && group.end + insertion.count >= sheetAxisMaximum(edit.axis)) {
      throw new SheetOutlineError(DimensionOutlineErrorReason.OUT_OF_BOUNDS, '插入将大纲组推至 Office 行列上限之外。', group.id)
    }
    const intervals = transformSheetAxisInterval({ start: group.start, end: group.end }, edit)
    if (intervals.length > 1) throw new SheetOutlineError(DimensionOutlineErrorReason.MOVE_SPLITS_OUTLINE,
      `Move would split dimension outline: ${group.id}`, group.id)
    if (intervals[0]) groups.push(intervals[0].start === group.start && intervals[0].end === group.end ? group : { ...group, ...intervals[0] })
  }
  if (edit.kind === 'copy') {
    const ids = new Set(state.groups.map(group => group.id))
    for (const group of state.groups) if (group.axis === edit.axis && group.start >= edit.start && group.end < edit.start + edit.count) {
      const id = allocateId()
      if (typeof id !== 'string' || !id.trim() || ids.has(id)) throw new SheetOutlineError(DimensionOutlineErrorReason.INVALID_RANGE, '复制大纲组 ID 无效或重复。', id)
      ids.add(id)
      groups.push({ ...cloneModel(group), id, start: edit.destination + group.start - edit.start, end: edit.destination + group.end - edit.start })
    }
  }
  // Delete/move can make formerly distinct nested ranges equal. Keep the
  // first surviving identity and every collapsed cause, as in the fixed source.
  const normalized: SheetOutlineGroup[] = [], positions = new Map<string, number>()
  for (const group of groups) {
    const key = `${group.axis}:${group.start}:${group.end}`, position = positions.get(key)
    if (position === undefined) { positions.set(key, normalized.length); normalized.push(group) }
    else {
      const previous = normalized[position]!
      if (group.collapsed && !previous.collapsed) normalized[position] = { ...previous, collapsed: true }
    }
  }
  const candidate = { ...state, groups: normalized }
  validateSheetOutlineState(candidate, extent)
  if (normalized.length === state.groups.length && normalized.every((group, index) => group === state.groups[index])) return state
  return candidate
}
