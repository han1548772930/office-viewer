import type { SheetValidationListOption } from '../../features/sheet-validation/validation-list-input.ts'

export type SheetValidationListNavigationKey = 'ArrowUp' | 'ArrowDown' | 'Home' | 'End'
export function sheetValidationListNavigationIndex(key: string, index: number, length: number): number | undefined {
  if (!Number.isSafeInteger(length) || length <= 0) return undefined
  if (key === 'Home') return 0
  if (key === 'End') return length - 1
  if (!Number.isSafeInteger(index) || index < -1 || index >= length) return undefined
  if (key === 'ArrowDown') return index < 0 ? 0 : Math.min(length - 1, index + 1)
  if (key === 'ArrowUp') return index < 0 ? length - 1 : Math.max(0, index - 1)
  return undefined
}

interface NavigationState {
  owner?: object
  items: readonly SheetValidationListOption[]
  query: string
  unitId: string
  baseRevision: number
  ready: boolean
  current(): boolean
}
interface NavigationHost {
  state(): NavigationState
  activate(index: number): void
  scrollTo(item: SheetValidationListOption, align: 'top' | 'bottom' | 'auto'): void
  focusAt(index: number, current: () => boolean): boolean
  nextFrame(): Promise<void>
}

/** Logical result indices are independent of mounted Listy rows. The caller
 * invalidates this controller synchronously on owner/query/revision/IME changes;
 * every scroll, frame and focus boundary checks the complete captured context.
 */
export function createSheetValidationListNavigation(host: NavigationHost) {
  let generation = 0, stopped = false
  function current(initial: NavigationState, token: number): boolean {
    if (stopped || token !== generation) return false
    const latest = host.state()
    const valid = latest.ready && latest.owner === initial.owner && latest.items === initial.items && latest.query === initial.query
      && latest.unitId === initial.unitId && latest.baseRevision === initial.baseRevision && latest.current()
    return valid && !stopped && token === generation
  }
  async function focus(index: number, align: 'top' | 'bottom' | 'auto' = 'auto'): Promise<boolean> {
    const token = ++generation, initial = host.state()
    if (!initial.owner || !Number.isSafeInteger(index) || index < 0 || !initial.items[index] || !current(initial, token)) return false
    host.activate(index)
    if (!current(initial, token)) return false
    host.scrollTo(initial.items[index]!, align)
    if (!current(initial, token)) return false
    // Dynamic row measurements may settle over more than one frame. Each wait
    // is bounded and an obsolete request never focuses a newly mounted row.
    for (let attempt = 0; attempt < 4; attempt++) {
      await host.nextFrame()
      if (!current(initial, token)) return false
      if (host.focusAt(index, () => current(initial, token))) return current(initial, token)
      if (!current(initial, token)) return false
      host.scrollTo(initial.items[index]!, align)
      if (!current(initial, token)) return false
    }
    return false
  }
  return { focus, invalidate: () => { generation++ }, dispose: () => { stopped = true; generation++ },
    get generation() { return generation } }
}
