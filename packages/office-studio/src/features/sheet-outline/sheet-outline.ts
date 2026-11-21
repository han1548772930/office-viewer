import type { OutlineAxis, SheetOutlineGroup, SheetOutlineState } from '../../domain/model.ts'
export type { OutlineAxis, SheetOutlineGroup, SheetOutlineState } from '../../domain/model.ts'

export interface OutlineAxisState {
  count: number
  groups: readonly SheetOutlineGroup[]
}

export type SheetOutlineOperation = 'group' | 'toggle' | 'remove'

export function createSheetOutlineState(): SheetOutlineState {
  return { summaryAbove: false, summaryLeft: false, groups: [] }
}

export function addSheetOutlineGroup(
  state: SheetOutlineState,
  axis: OutlineAxis,
  start: number,
  end: number,
  count: number,
  id: string,
): SheetOutlineState {
  if (!id.trim()) throw new TypeError('Outline group id is required')
  validateRange(start, end, count)
  if (state.groups.some(group => group.id === id)) throw new Error(`Outline group already exists: ${id}`)
  const groups = state.groups.filter(group => group.axis === axis)
  if (groups.some(group => overlaps(start, end, group.start, group.end) && !contains(start, end, group.start, group.end) && !contains(group.start, group.end, start, end))) {
    throw new RangeError('Outline groups must be nested or disjoint')
  }
  return { ...state, groups: [...state.groups, { id, axis, start, end, collapsed: false }] }
}

export function removeSheetOutlineGroup(state: SheetOutlineState, id: string): SheetOutlineState {
  const groups = state.groups.filter(group => group.id !== id)
  if (groups.length === state.groups.length) return state
  return { ...state, groups }
}

export function setSheetOutlineCollapsed(state: SheetOutlineState, id: string, collapsed: boolean): SheetOutlineState {
  let found = false
  const groups = state.groups.map(group => {
    if (group.id !== id) return group
    found = true
    return { ...group, collapsed }
  })
  if (!found) throw new Error(`Outline group not found: ${id}`)
  return { ...state, groups }
}

export function setSheetOutlineSummaryPosition(state: SheetOutlineState, axis: OutlineAxis, atStart: boolean): SheetOutlineState {
  return axis === 'row' ? { ...state, summaryAbove: atStart } : { ...state, summaryLeft: atStart }
}

export function setSheetOutlineLevel(state: SheetOutlineState, axis: OutlineAxis, level: number): SheetOutlineState {
  if (!Number.isSafeInteger(level) || level < 1) throw new RangeError('Outline level must be a positive integer')
  const groups = state.groups.filter(group => group.axis === axis).sort((left, right) => left.start - right.start || right.end - left.end)
  const updated = groups.map(group => {
    const depth = groups.filter(parent => parent.start <= group.start && parent.end >= group.end && (parent.start < group.start || parent.end > group.end)).length
    return { ...group, collapsed: depth >= level - 1 }
  })
  const otherAxis = state.groups.filter(group => group.axis !== axis)
  return { ...state, groups: [...otherAxis, ...updated] }
}

export function applySheetOutlineOperation(
  state: SheetOutlineState,
  axis: OutlineAxis,
  operation: SheetOutlineOperation,
  start: number,
  end: number,
  count: number,
  id: string,
): SheetOutlineState {
  if (operation === 'group') return addSheetOutlineGroup(state, axis, start, end, count, id)
  const matching = state.groups
    .filter(group => group.axis === axis && group.start <= start && group.end >= end)
    .sort((left, right) => (left.end - left.start) - (right.end - right.start))
  const target = matching[0]
  if (!target) throw new Error(`No outline group contains the selected ${axis} range`)
  return operation === 'remove'
    ? removeSheetOutlineGroup(state, target.id)
    : setSheetOutlineCollapsed(state, target.id, !target.collapsed)
}

export function getOutlineDepth(state: OutlineAxisState, index: number): number {
  assertIndex(index, state.count)
  return state.groups.reduce((depth, group) => group.start <= index && index <= group.end ? depth + 1 : depth, 0)
}

export function getCollapsedOutlineIndexes(state: OutlineAxisState): Set<number> {
  const hidden = new Set<number>()
  for (const group of state.groups) {
    validateRange(group.start, group.end, state.count)
    if (group.axis !== 'row' && group.axis !== 'column') throw new TypeError('Invalid outline axis')
    if (!group.collapsed) continue
    for (let index = group.start; index <= group.end; index += 1) hidden.add(index)
  }
  return hidden
}

export function getOutlineLevels(state: OutlineAxisState): Array<Array<[number, number]>> {
  const groups = [...state.groups].sort((left, right) => left.start - right.start || right.end - left.end)
  const levels: Array<Array<[number, number]>> = []
  for (const group of groups) {
    validateRange(group.start, group.end, state.count)
    const level = groups.filter(parent => parent.start <= group.start && parent.end >= group.end && (parent.start < group.start || parent.end > group.end)).length
    levels[level] ??= []
    levels[level]!.push([group.start, group.end])
  }
  return levels
}

function validateRange(start: number, end: number, count: number): void {
  if (!Number.isSafeInteger(count) || count < 1) throw new RangeError('Outline axis count must be a positive integer')
  if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 0 || end < start || end >= count) {
    throw new RangeError('Outline range is outside its axis')
  }
  if (start === end) throw new RangeError('An outline group must contain at least two rows or columns')
}

function assertIndex(index: number, count: number): void {
  if (!Number.isSafeInteger(index) || index < 0 || index >= count) throw new RangeError('Outline index is outside its axis')
}

function overlaps(aStart: number, aEnd: number, bStart: number, bEnd: number): boolean {
  return aStart <= bEnd && bStart <= aEnd
}

function contains(outerStart: number, outerEnd: number, innerStart: number, innerEnd: number): boolean {
  return outerStart <= innerStart && outerEnd >= innerEnd
}
