export interface SlicerItemSelectionInput {
  value: string
  index: number
  anchorIndex?: number
  ctrlKey: boolean
  metaKey: boolean
  shiftKey: boolean
  multiSelect: boolean
}

export interface SlicerControlItemState {
  checkedItems: string[]
  partialItems: string[]
}

function assertCurrentItem(items: readonly string[], input: SlicerItemSelectionInput): void {
  if (!Number.isSafeInteger(input.index) || input.index < 0 || input.index >= items.length || items[input.index] !== input.value
    || input.anchorIndex !== undefined && (!Number.isSafeInteger(input.anchorIndex) || input.anchorIndex < 0 || input.anchorIndex >= items.length)) {
    throw new Error('切片器项目或选择锚点已更改。')
  }
}

/** A toggle only replaces the clicked caption. Partial raw memberships under other captions survive. */
export function slicerControlItemStates(items: readonly string[], state: SlicerControlItemState, input: SlicerItemSelectionInput): SlicerControlItemState {
  assertCurrentItem(items, input)
  const checked = new Set(state.checkedItems.filter(value => items.includes(value)))
  const partial = new Set(state.partialItems.filter(value => items.includes(value)))
  if ([...partial].some(value => checked.has(value))) throw new Error('切片器项目不能同时全部选中和部分选中。')
  if (input.shiftKey && input.anchorIndex !== undefined) {
    const first = Math.min(input.anchorIndex, input.index), last = Math.max(input.anchorIndex, input.index)
    return { checkedItems: items.slice(first, last + 1), partialItems: [] }
  }
  if (!(input.ctrlKey || input.metaKey || input.multiSelect)) return { checkedItems: [input.value], partialItems: [] }
  if (checked.has(input.value)) checked.delete(input.value)
  else checked.add(input.value)
  partial.delete(input.value)
  return { checkedItems: items.filter(value => checked.has(value)), partialItems: items.filter(value => partial.has(value)) }
}

/** Pointer, keyboard and Ant item buttons derive one native filter selection from the current source order. */
export function slicerControlItemSelection(items: readonly string[], selectedItems: readonly string[] | undefined, input: SlicerItemSelectionInput): string[] | null {
  assertCurrentItem(items, input)
  let selected = selectedItems == null ? [...items] : [...new Set(selectedItems.filter(value => items.includes(value)))]
  if (input.shiftKey && input.anchorIndex !== undefined) {
    const first = Math.min(input.anchorIndex, input.index), last = Math.min(items.length - 1, Math.max(input.anchorIndex, input.index))
    selected = items.slice(first, last + 1)
  } else if (input.ctrlKey || input.metaKey || input.multiSelect) {
    selected = selected.includes(input.value) ? selected.filter(value => value !== input.value) : [...selected, input.value]
  } else selected = [input.value]
  return selected.length === items.length ? null : selected
}
