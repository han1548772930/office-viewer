import type { PivotFieldModel, PivotTableOptions, PivotValuePosition } from '../../../domain/model.ts'
import type { PivotAxisGroup } from './pivot-axis-state.ts'

export interface PivotValuesPositionState { position: PivotValuePosition; index: number }
type FieldOrder = Pick<PivotFieldModel, 'id' | 'area' | 'position'>
export const PIVOT_VALUE_POSITIONS = Object.freeze({ None: -1, Column: 1, Row: 2 } as const)

export function normalizePivotValuePositionInput(value: unknown): PivotValuePosition {
  if (value === 'none' || value === -1) return 'none'
  if (value === 'column' || value === 1) return 'column'
  if (value === 'row' || value === 2) return 'row'
  throw new TypeError('Invalid pivot value position')
}

export function normalizePivotValuesPosition(position: unknown, index: unknown = -1): PivotValuesPositionState {
  if (position !== 'none' && position !== 'row' && position !== 'column') throw new TypeError('Invalid pivot value position')
  if (typeof index !== 'number' || !Number.isSafeInteger(index) || index < -1) throw new RangeError('Invalid pivot value index')
  return { position, index }
}

/** The configured snapshot state is separate from the effective clamped layout. */
export function getPivotValuesPosition(pivot: { options: Pick<PivotTableOptions, 'valuePosition' | 'valueIndex'> }): PivotValuesPositionState {
  return normalizePivotValuesPosition(pivot.options.valuePosition, pivot.options.valueIndex === undefined ? -1 : pivot.options.valueIndex)
}

/** Fixed engine layout keeps measures visible in columns when the special field is absent. */
export function getPivotValuesLayout(pivot: {
  options: Pick<PivotTableOptions, 'valuePosition' | 'valueIndex'>
  fields: readonly Pick<PivotFieldModel, 'area'>[]
}): { area: 'row' | 'column'; index: number } {
  const state = getPivotValuesPosition(pivot), area = state.position === 'row' ? 'row' : 'column'
  const length = pivot.fields.filter(field => field.area === area).length
  return { area, index: state.index === -1 ? length : Math.min(state.index, length) }
}

/** Dimensions keep their local positions; ΣValues owns a gap and never becomes a source field. */
export function adjustPivotValuesIndexForFieldEdit(pivot: {
  options: Pick<PivotTableOptions, 'valuePosition' | 'valueIndex'>
  fields: readonly FieldOrder[]
}, previous: readonly FieldOrder[], fieldId: string): void {
  const state = getPivotValuesPosition(pivot)
  if (state.position === 'none' || state.index === -1) return
  const before = previous.find(field => field.id === fieldId), after = pivot.fields.find(field => field.id === fieldId)
  if (before?.area === after?.area && before?.position === after?.position) return
  if (before?.area !== state.position && after?.area !== state.position) return
  let gap = Math.min(state.index, previous.filter(field => field.area === state.position).length)
  if (before?.area === state.position && before.position < gap) gap--
  // An explicit dimension insertion at the gap places that field before ΣValues.
  if (after?.area === state.position && after.position <= gap) gap++
  pivot.options.valueIndex = gap
}

/** Source-field retirement removes only the dimensions formerly preceding the gap. */
export function retainPivotValuesIndexAfterRemoval(pivot: {
  options: Pick<PivotTableOptions, 'valuePosition' | 'valueIndex'>
  fields: readonly FieldOrder[]
}, previous: readonly FieldOrder[]): void {
  const state = getPivotValuesPosition(pivot)
  if (state.position === 'none' || state.index === -1) return
  const before = previous.filter(field => field.area === state.position).sort((a, b) => a.position - b.position)
  const removed = new Set(before.filter(field => !pivot.fields.some(next => next.id === field.id && next.area === state.position)).map(field => field.id))
  if (!removed.size) return
  const gap = Math.min(state.index, before.length)
  pivot.options.valueIndex = gap - before.slice(0, gap).filter(field => removed.has(field.id)).length
}

/** Insert measures between dimension levels, preserving projected aggregate/collapse scopes. */
export function expandPivotValuesAxis<T extends { group: Pick<PivotAxisGroup, 'values'> | null }>(
  axis: readonly T[], fields: readonly Pick<PivotFieldModel, 'id'>[], index: number,
): Array<T & { valueFieldId: string }> {
  if (!Number.isSafeInteger(index) || index < 0) throw new RangeError('Invalid pivot value layout index')
  const result: Array<T & { valueFieldId: string }> = []
  const scope = (entry: T): string => index === 0 ? 'root'
    : entry.group ? JSON.stringify(entry.group.values.slice(0, index)) : 'grand-total'
  for (let start = 0; start < axis.length;) {
    let end = start + 1
    const key = scope(axis[start]!)
    while (end < axis.length && scope(axis[end]!) === key) end++
    for (const field of fields) for (let item = start; item < end; item++) {
      result.push({ ...axis[item]!, valueFieldId: field.id })
    }
    start = end
  }
  return result
}
