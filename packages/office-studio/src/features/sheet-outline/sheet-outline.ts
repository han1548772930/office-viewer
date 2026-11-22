import type { CellRange, OutlineAxis, SheetOutlineGroup, SheetOutlineState, SpreadsheetSheet } from '../../domain/model.ts'
import { canEditProtectedRanges, canEditWorksheetProtection } from '../sheet-protection/range-protection.ts'
export type { OutlineAxis, SheetOutlineGroup, SheetOutlineState } from '../../domain/model.ts'

export interface OutlineAxisState {
  count: number
  groups: readonly SheetOutlineGroup[]
}

export type SheetOutlineOperation = 'group' | 'toggle' | 'remove'

type OutlineRange = { axis: OutlineAxis; start: number; end: number }

export type SheetOutlineAction =
  | (OutlineRange & { kind: SheetOutlineOperation })
  | { kind: 'toggle-group'; axis: OutlineAxis; groupId: string }
  | { kind: 'set-level'; axis: OutlineAxis; level: number }
  | { kind: 'summary-position'; axis: OutlineAxis; atStart: boolean }

export const CHANGE_SHEET_OUTLINE_COMMAND = 'sheet.command.change-outline'
export type ChangeSheetOutlineInput = { unitId: string; sheetId: string; action: SheetOutlineAction }

type OutlineWorksheet = Pick<SpreadsheetSheet, 'outline' | 'worksheetProtection' | 'rangeProtectionRules'> & { rowCount: number; columnCount: number }
const MAX_OUTLINE_DEPTH = 8

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
  const adjacent = new Set<SheetOutlineGroup>()
  let mergedStart = start, mergedEnd = end
  const axisGroups = state.groups.filter(group => group.axis === axis)
  for (let changed = true; changed;) {
    changed = false
    for (const group of axisGroups) {
      if (adjacent.has(group) || (group.end + 1 !== mergedStart && mergedEnd + 1 !== group.start)) continue
      adjacent.add(group)
      mergedStart = Math.min(mergedStart, group.start)
      mergedEnd = Math.max(mergedEnd, group.end)
      changed = true
    }
  }
  const existing = [...adjacent].sort(compareOutlineRanges)[0]
  const group: SheetOutlineGroup = existing
    ? { ...existing, id, start: mergedStart, end: mergedEnd }
    : { id, axis, start: mergedStart, end: mergedEnd, collapsed: false }
  const groups = [...state.groups.filter(candidate => !adjacent.has(candidate)), group]
  for (const depth of outlineGroupDepths(groups.filter(candidate => candidate.axis === axis)).values()) {
    if (depth >= MAX_OUTLINE_DEPTH) throw new RangeError(`Outline depth cannot exceed ${MAX_OUTLINE_DEPTH}`)
  }
  return { ...state, groups }
}

export function removeSheetOutlineGroup(state: SheetOutlineState, id: string): SheetOutlineState {
  const groups = state.groups.filter(group => group.id !== id)
  if (groups.length === state.groups.length) return state
  return { ...state, groups }
}

export function setSheetOutlineCollapsed(state: SheetOutlineState, id: string, collapsed: boolean): SheetOutlineState {
  const group = state.groups.find(candidate => candidate.id === id)
  if (!group) throw new Error(`Outline group not found: ${id}`)
  if (group.collapsed === collapsed) return state
  return { ...state, groups: state.groups.map(candidate => candidate === group ? { ...group, collapsed } : candidate) }
}

export function setSheetOutlineSummaryPosition(state: SheetOutlineState, axis: OutlineAxis, atStart: boolean): SheetOutlineState {
  const key = axis === 'row' ? 'summaryAbove' : 'summaryLeft'
  return state[key] === atStart ? state : { ...state, [key]: atStart }
}

export function setSheetOutlineLevel(state: SheetOutlineState, axis: OutlineAxis, level: number): SheetOutlineState {
  if (!Number.isSafeInteger(level) || level < 1) throw new RangeError('Outline level must be a positive integer')
  const depths = outlineGroupDepths(state.groups.filter(group => group.axis === axis))
  const groups = state.groups.map(group => {
    const depth = depths.get(group.id)
    if (depth === undefined || group.collapsed === (depth >= level - 1)) return group
    return group.collapsed === (depth >= level - 1) ? group : { ...group, collapsed: depth >= level - 1 }
  })
  return groups.some((group, index) => group !== state.groups[index]) ? { ...state, groups } : state
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
  validateRange(start, end, count)
  if (operation === 'remove') {
    const groups = state.groups.filter(group => group.axis !== axis || !contains(start, end, group.start, group.end))
    return groups.length === state.groups.length ? state : { ...state, groups }
  }
  const matching = state.groups
    .filter(group => group.axis === axis && group.start <= start && group.end >= end)
    .reverse().sort((left, right) => (left.end - left.start) - (right.end - right.start))
  const target = matching[0]
  if (!target) throw new Error(`No outline group contains the selected ${axis} range`)
  return setSheetOutlineCollapsed(state, target.id, !target.collapsed)
}

export function getSheetOutlineUpdate(sheet: OutlineWorksheet, action: SheetOutlineAction, userId: string): SheetOutlineState | undefined {
  const { axis } = action
  if (axis !== 'row' && axis !== 'column') throw new TypeError('Invalid outline axis')
  if (!canEditWorksheetProtection(sheet, userId)) throw new Error('此工作表受保护，当前用户没有编辑权限。')
  const before = sheet.outline ?? createSheetOutlineState()
  const count = axis === 'row' ? sheet.rowCount : sheet.columnCount
  let after: SheetOutlineState
  let selection: OutlineRange[] = []
  switch (action.kind) {
    case 'summary-position':
      after = setSheetOutlineSummaryPosition(before, axis, action.atStart)
      break
    case 'set-level':
      after = setSheetOutlineLevel(before, axis, action.level)
      break
    case 'toggle-group': {
      const group = before.groups.find(candidate => candidate.axis === axis && candidate.id === action.groupId)
      if (!group) return undefined
      after = setSheetOutlineCollapsed(before, group.id, !group.collapsed)
      break
    }
    default:
      selection = [action]
      after = applySheetOutlineOperation(before, axis, action.kind, action.start, action.end, count, action.kind === 'group' ? crypto.randomUUID() : '')
  }
  if (before === after) return undefined
  const previous = new Map(before.groups.map(group => [group.id, group]))
  const next = new Map(after.groups.map(group => [group.id, group]))
  const affected = [...selection, ...before.groups.filter(group => next.get(group.id) !== group), ...after.groups.filter(group => previous.get(group.id) !== group)]
  if (!canEditProtectedRanges(sheet, affected.map(range => outlineCellRange(sheet, range)), userId)) {
    throw new Error('该区域受保护，当前用户没有编辑权限。')
  }
  return after
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
  const groups = [...state.groups].sort(compareOutlineRanges)
  const depths = outlineGroupDepths(groups)
  const levels: Array<Array<[number, number]>> = []
  for (const group of groups) {
    validateRange(group.start, group.end, state.count)
    const level = depths.get(group.id)!
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
}

function assertIndex(index: number, count: number): void {
  if (!Number.isSafeInteger(index) || index < 0 || index >= count) throw new RangeError('Outline index is outside its axis')
}

function compareOutlineRanges(left: SheetOutlineGroup, right: SheetOutlineGroup): number {
  return left.start - right.start || right.end - left.end
}

export function outlineGroupDepths(groups: readonly SheetOutlineGroup[]): Map<string, number> {
  const depths = new Map<string, number>()
  const parents: SheetOutlineGroup[] = []
  for (const group of [...groups].sort(compareOutlineRanges)) {
    while (parents.length && group.start > parents[parents.length - 1]!.end) parents.pop()
    const parent = parents.at(-1)
    if (parent && group.end > parent.end) throw new RangeError('Outline groups must be nested or disjoint')
    depths.set(group.id, parents.length)
    parents.push(group)
  }
  return depths
}

function outlineCellRange(sheet: OutlineWorksheet, range: OutlineRange): CellRange {
  return range.axis === 'row'
    ? { startRow: range.start, endRow: range.end, startCol: 0, endCol: sheet.columnCount - 1 }
    : { startRow: 0, endRow: sheet.rowCount - 1, startCol: range.start, endCol: range.end }
}

function contains(outerStart: number, outerEnd: number, innerStart: number, innerEnd: number): boolean {
  return outerStart <= innerStart && outerEnd >= innerEnd
}
