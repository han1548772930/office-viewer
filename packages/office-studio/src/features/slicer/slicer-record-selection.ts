import type { CellScalar, PivotRecordSelection, SheetValueFilter } from '../../domain/model.ts'
import { compilePivotRecordSelection, normalizePivotRecordSelection } from '../pivot/services/pivot-filter-predicate.ts'
import { isRecordValue, normalizeRecordSelection, pivotRecordValueKey, recordValueFromScalar,
  type PivotRecordValue, type RecordSelection, type RecordValue } from '../../domain/record-value.ts'

const itemKey = (value: CellScalar): string => String(value ?? '')

/** Include-empty is restrictive. All and exclude-empty both clear the member filter. */
export function normalizeSlicerRecordSelection(value: PivotRecordSelection | null | undefined): PivotRecordSelection | undefined {
  if (value == null) return undefined
  const selection = normalizePivotRecordSelection(value)
  return selection.mode === 'all' || selection.mode === 'exclude' && selection.values.length === 0 ? undefined : selection
}

export function slicerRecordSelectionFromNativeValues(filter: SheetValueFilter): PivotRecordSelection | undefined {
  if (filter.recordSelection) return normalizeSlicerRecordSelection(filter.recordSelection)
  return normalizeSlicerRecordSelection({ mode: filter.selectionMode ?? 'include', values: filter.selectedItems })
}

/** Converts legacy raw conditions using observed identities; absent conditions retain their original scalar type. */
export function slicerPublicRecordSelection(selection: PivotRecordSelection | undefined, values: readonly CellScalar[] = [],
  typedValues: readonly RecordValue[] = []): RecordSelection {
  const current = normalizeSlicerRecordSelection(selection)
  if (!current || current.mode === 'all') return { mode: 'all' }
  const typed: RecordValue[] = []
  for (const condition of current.values) {
    if (isRecordValue(condition)) { typed.push(condition); continue }
    let observed = false
    values.forEach((value, index) => {
      if (typeof value === typeof condition && value === condition) {
        observed = true; typed.push(typedValues[index] ?? recordValueFromScalar(value))
      }
    })
    if (!observed) typed.push(recordValueFromScalar(condition))
  }
  return normalizeRecordSelection({ mode: current.mode, values: typed })
}

export function sameSlicerRecordSelection(first: PivotRecordSelection | null | undefined, second: PivotRecordSelection | null | undefined): boolean {
  const left = normalizeSlicerRecordSelection(first), right = normalizeSlicerRecordSelection(second)
  if (!left || !right) return left === right
  if (left.mode === 'all' || right.mode === 'all') return left.mode === right.mode
  return left.mode === right.mode && left.values.length === right.values.length
    && left.values.every(value => right.values.some(member => pivotRecordValueKey(value) === pivotRecordValueKey(member)))
}

/** A fully checked raw display key includes every currently observed scalar type behind it. */
export function slicerRecordSelectionFromItems(values: readonly CellScalar[], items: readonly string[] | null): PivotRecordSelection | undefined {
  if (items === null) return undefined
  const selected = new Set(items)
  return normalizePivotRecordSelection({ mode: 'include', values: values.filter(value => selected.has(itemKey(value))) })
}

export function slicerRecordSelectionItems(selection: PivotRecordSelection | undefined, values: readonly CellScalar[], typedValues?: readonly RecordValue[]): string[] | undefined {
  const normalized = normalizeSlicerRecordSelection(selection)
  if (!normalized) return undefined
  const matches = compilePivotRecordSelection(normalized)
  return [...new Set(values.filter((value, index) => matches(value, typedValues?.[index])).map(itemKey))]
}

export function slicerRecordSelectionItemState(selection: PivotRecordSelection | undefined, values: readonly CellScalar[], typedValues?: readonly RecordValue[]): { checkedItems: string[]; partialItems: string[] } {
  const matches = selection ? compilePivotRecordSelection(selection) : () => true
  return slicerItemStateFromPredicate(values, (value, index) => matches(value, typedValues?.[index]))
}

export function slicerItemStateFromPredicate(values: readonly CellScalar[], matches: (value: CellScalar, index: number) => boolean): { checkedItems: string[]; partialItems: string[] } {
  const states = new Map<string, { selected: boolean; unselected: boolean }>()
  for (const [index, value] of values.entries()) {
    const key = itemKey(value), state = states.get(key) ?? { selected: false, unselected: false }
    if (matches(value, index)) state.selected = true; else state.unselected = true
    states.set(key, state)
  }
  return { checkedItems: [...states].filter(([, state]) => state.selected && !state.unselected).map(([key]) => key),
    partialItems: [...states].filter(([, state]) => state.selected && state.unselected).map(([key]) => key) }
}

/** Partial display keys preserve the cache's existing typed membership while checked keys select all observed raw types. */
export function planSlicerRecordItemStates(current: PivotRecordSelection | undefined, values: readonly CellScalar[],
  checkedItems: readonly string[], partialItems: readonly string[], typedValues?: readonly RecordValue[]): PivotRecordSelection {
  if (!Array.isArray(checkedItems) || !Array.isArray(partialItems)
    || [...checkedItems, ...partialItems].some(item => typeof item !== 'string')) throw new TypeError('Slicer item states require string keys')
  const checked = new Set(checkedItems), partial = new Set(partialItems), keys = new Set(values.map(itemKey))
  if ([...checked, ...partial].some(key => !keys.has(key))) throw new TypeError('Slicer item state contains an unknown item')
  if ([...partial].some(key => checked.has(key))) throw new TypeError('A slicer item cannot be both checked and partial')
  const normalized = normalizeSlicerRecordSelection(current), matches = normalized ? compilePivotRecordSelection(normalized) : () => true
  const mode = partial.size && normalized?.mode === 'exclude' ? 'exclude' : 'include'
  const typed = normalized && normalized.mode !== 'all' && normalized.values.some(isRecordValue)
  const observed: readonly PivotRecordValue[] = typed && typedValues ? typedValues : values
  const next: PivotRecordValue[] = values.flatMap((value, index) => (checked.has(itemKey(value))
    || partial.has(itemKey(value)) && matches(value, typedValues?.[index])) === (mode === 'include') ? [observed[index]!] : [])
  if (partial.size && normalized && normalized.mode !== 'all' && normalized.mode === mode) {
    for (const value of normalized.values) if (!observed.some(member => pivotRecordValueKey(member) === pivotRecordValueKey(value))) next.push(value)
  }
  return normalizePivotRecordSelection({ mode, values: next })
}
