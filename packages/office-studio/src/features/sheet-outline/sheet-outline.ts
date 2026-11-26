import type { CellRange, OutlineAxis, SheetOutlineGroup, SheetOutlineState, SpreadsheetSheet } from '../../domain/model.ts'
import { canEditProtectedRanges, canEditWorksheetProtection } from '../sheet-protection/range-protection.ts'
import { DimensionOutlineErrorReason, MAX_SHEET_OUTLINE_DEPTH, SheetOutlineError, type SheetOutlineExtent,
  type SheetOutlineHiddenRange, type SheetOutlineTreeNode } from './sheet-outline-types.ts'
export type { OutlineAxis, SheetOutlineGroup, SheetOutlineState } from '../../domain/model.ts'

export interface OutlineAxisState {
  count: number
  groups: readonly SheetOutlineGroup[]
}

export type SheetOutlineOperation = 'group' | 'toggle' | 'remove' | 'collapse' | 'expand'

export type OutlineRange = { axis: OutlineAxis; start: number; end: number }

export type SheetOutlineAction =
  | (OutlineRange & { kind: SheetOutlineOperation; groupId?: string })
  | { kind: 'toggle-group'; axis: OutlineAxis; groupId: string }
  | { kind: 'set-collapsed'; axis: OutlineAxis; groupId: string; collapsed: boolean }
  | { kind: 'remove-group'; axis: OutlineAxis; groupId: string }
  | (OutlineRange & { kind: 'clear' })
  | { kind: 'set-level'; axis: OutlineAxis; level: number }
  | { kind: 'summary-position'; axis: OutlineAxis; atStart: boolean }
  | { kind: 'summary-settings'; axis?: never; summaryAbove: boolean; summaryLeft: boolean }

export const CHANGE_SHEET_OUTLINE_COMMAND = 'sheet.command.change-outline'
export type ChangeSheetOutlineInput = { unitId: string; sheetId: string; action: SheetOutlineAction;
  expectedRevision?: number; expectedHistoryGeneration?: number; expectedUserId?: string; host?: 'api' | 'workbench' }

export type OutlineWorksheet = Pick<SpreadsheetSheet, 'outline' | 'worksheetProtection' | 'rangeProtectionRules'> & SheetOutlineExtent

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
  assertOutlineAxis(axis)
  validateSheetOutlineState(state)
  if (typeof id !== 'string' || !id.trim()) throw new SheetOutlineError(DimensionOutlineErrorReason.INVALID_RANGE, 'Outline group id is required')
  validateRange(start, end, count)
  if (state.groups.some(group => group.id === id)) throw new SheetOutlineError(DimensionOutlineErrorReason.INVALID_RANGE, `Outline group already exists: ${id}`, id)
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
    ? { ...existing, start: mergedStart, end: mergedEnd }
    : { id, axis, start: mergedStart, end: mergedEnd, collapsed: false }
  const groups = [...state.groups.filter(candidate => !adjacent.has(candidate)), group]
  for (const depth of outlineGroupDepths(groups.filter(candidate => candidate.axis === axis)).values()) {
    if (depth >= MAX_SHEET_OUTLINE_DEPTH) throw new SheetOutlineError(DimensionOutlineErrorReason.MAX_DEPTH, `Outline depth cannot exceed ${MAX_SHEET_OUTLINE_DEPTH}`, group.id)
  }
  return { ...state, groups }
}

export function removeSheetOutlineGroup(state: SheetOutlineState, id: string): SheetOutlineState {
  validateSheetOutlineState(state)
  const groups = state.groups.filter(group => group.id !== id)
  if (groups.length === state.groups.length) return state
  return { ...state, groups }
}

export function setSheetOutlineCollapsed(state: SheetOutlineState, id: string, collapsed: boolean): SheetOutlineState {
  validateSheetOutlineState(state)
  if (typeof collapsed !== 'boolean') throw new SheetOutlineError(DimensionOutlineErrorReason.INVALID_RANGE, 'Outline collapsed state must be boolean', id)
  const group = state.groups.find(candidate => candidate.id === id)
  if (!group) throw new SheetOutlineError(DimensionOutlineErrorReason.UNKNOWN, `Outline group not found: ${id}`, id)
  const matches = (candidate: SheetOutlineGroup) => candidate.axis === group.axis && candidate.start === group.start && candidate.end === group.end
  if (state.groups.every(candidate => !matches(candidate) || candidate.collapsed === collapsed)) return state
  return { ...state, groups: state.groups.map(candidate => matches(candidate) && candidate.collapsed !== collapsed ? { ...candidate, collapsed } : candidate) }
}

export function setSheetOutlineSummaryPosition(state: SheetOutlineState, axis: OutlineAxis, atStart: boolean): SheetOutlineState {
  assertOutlineAxis(axis)
  validateSheetOutlineState(state)
  if (typeof atStart !== 'boolean') throw new SheetOutlineError(DimensionOutlineErrorReason.INVALID_RANGE, 'Outline summary direction must be boolean')
  const key = axis === 'row' ? 'summaryAbove' : 'summaryLeft'
  return state[key] === atStart ? state : { ...state, [key]: atStart }
}

export function setSheetOutlineLevel(state: SheetOutlineState, axis: OutlineAxis, level: number): SheetOutlineState {
  assertOutlineAxis(axis)
  validateSheetOutlineState(state)
  if (!Number.isSafeInteger(level) || level < 1 || level > MAX_SHEET_OUTLINE_DEPTH + 1) throw new SheetOutlineError(DimensionOutlineErrorReason.INVALID_RANGE, 'Outline level must be between 1 and 9')
  const depths = outlineGroupDepths(state.groups.filter(group => group.axis === axis))
  const sameRangeDepth = new Map<string, number>()
  for (const group of state.groups) if (group.axis === axis) {
    const key = `${group.start}:${group.end}`, depth = depths.get(group.id)!
    sameRangeDepth.set(key, Math.max(sameRangeDepth.get(key) ?? 0, depth))
  }
  const groups = state.groups.map(group => {
    const depth = group.axis === axis ? sameRangeDepth.get(`${group.start}:${group.end}`) : undefined
    if (depth === undefined || group.collapsed === (depth >= level - 1)) return group
    return { ...group, collapsed: depth >= level - 1 }
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
  assertOutlineAxis(axis)
  validateSheetOutlineState(state)
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
  if (!target) throw new SheetOutlineError(DimensionOutlineErrorReason.UNKNOWN, '选区没有所属的大纲组。')
  if (operation !== 'toggle' && operation !== 'collapse' && operation !== 'expand') throw new SheetOutlineError(DimensionOutlineErrorReason.INVALID_RANGE, 'Invalid outline operation')
  return setSheetOutlineCollapsed(state, target.id, operation === 'toggle' ? !target.collapsed : operation === 'collapse')
}

export function getSheetOutlineUpdate(sheet: OutlineWorksheet, action: SheetOutlineAction, userId: string): SheetOutlineState | undefined {
  if (!action || typeof action !== 'object') throw new SheetOutlineError(DimensionOutlineErrorReason.INVALID_RANGE, 'Outline action is required')
  if (!canEditWorksheetProtection(sheet, userId)) throw new Error('此工作表受保护，当前用户没有编辑权限。')
  const before = sheet.outline ?? createSheetOutlineState()
  validateSheetOutlineState(before, sheet)
  if (action.kind === 'summary-settings') {
    const after = setSheetOutlineSummaryPosition(setSheetOutlineSummaryPosition(before, 'row', action.summaryAbove), 'column', action.summaryLeft)
    if (after === before) return undefined
    const affected = before.groups.filter(group => group.axis === 'row' ? action.summaryAbove !== before.summaryAbove : action.summaryLeft !== before.summaryLeft)
    assertOutlineRangesEditable(sheet, affected, userId)
    return after
  }
  const { axis } = action
  assertOutlineAxis(axis)
  const count = axis === 'row' ? sheet.rowCount : sheet.columnCount
  let after: SheetOutlineState
  let selection: OutlineRange[] = []
  switch (action.kind) {
    case 'summary-position':
      after = setSheetOutlineSummaryPosition(before, axis, action.atStart)
      selection = before.groups.filter(group => group.axis === axis)
      break
    case 'set-level':
      after = setSheetOutlineLevel(before, axis, action.level)
      break
    case 'toggle-group':
    case 'set-collapsed':
    case 'remove-group': {
      const group = before.groups.find(candidate => candidate.axis === axis && candidate.id === action.groupId)
      if (!group) return undefined
      after = action.kind === 'remove-group' ? removeSheetOutlineGroup(before, group.id)
        : setSheetOutlineCollapsed(before, group.id, action.kind === 'set-collapsed' ? action.collapsed : !group.collapsed)
      break
    }
    case 'clear':
      selection = [action]
      after = applySheetOutlineOperation(before, axis, 'remove', action.start, action.end, count, '')
      break
    default:
      selection = [action]
      after = applySheetOutlineOperation(before, axis, action.kind, action.start, action.end, count,
        action.kind === 'group' ? action.groupId ?? crypto.randomUUID() : '')
  }
  if (before === after) return undefined
  const previous = new Map(before.groups.map(group => [group.id, group]))
  const next = new Map(after.groups.map(group => [group.id, group]))
  const affected = [...selection, ...before.groups.filter(group => next.get(group.id) !== group), ...after.groups.filter(group => previous.get(group.id) !== group)]
  assertOutlineRangesEditable(sheet, affected, userId)
  return after
}

export function getOutlineDepth(state: OutlineAxisState, index: number): number {
  assertIndex(index, state.count)
  return state.groups.reduce((depth, group) => group.start <= index && index <= group.end ? depth + 1 : depth, 0)
}

export function getCollapsedOutlineIndexes(state: OutlineAxisState): Set<number> {
  const hidden = new Set<number>()
  for (const range of getCollapsedOutlineRanges(state)) {
    for (let index = range.start; index <= range.end; index += 1) hidden.add(index)
  }
  return hidden
}

/** Compact derived visibility. Manual hides, filters and zero sizes remain separate sources. */
export function getCollapsedOutlineRanges(state: OutlineAxisState): SheetOutlineHiddenRange[] {
  validateOutlineAxisState(state)
  const ranges: SheetOutlineHiddenRange[] = []
  for (const group of [...state.groups].sort(compareOutlineRanges)) {
    validateRange(group.start, group.end, state.count); assertOutlineAxis(group.axis)
    if (!group.collapsed) continue
    const previous = ranges.at(-1)
    if (!previous || group.start > previous.end + 1) ranges.push({ start: group.start, end: group.end })
    else previous.end = Math.max(previous.end, group.end)
  }
  return ranges
}

export function getSheetOutlineHiddenSpans(state: SheetOutlineState, axis: OutlineAxis, count: number) {
  assertOutlineAxis(axis)
  validateSheetOutlineState(state)
  return getCollapsedOutlineRanges({ count, groups: state.groups.filter(group => group.axis === axis) }).map(range => ({ ...range, value: true }))
}

export function getOutlineLevels(state: OutlineAxisState): Array<Array<[number, number]>> {
  validateOutlineAxisState(state)
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
  if (!Number.isSafeInteger(count) || count < 1) throw new SheetOutlineError(DimensionOutlineErrorReason.INVALID_RANGE, 'Outline axis count must be a positive integer')
  if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 0 || end < start) throw new SheetOutlineError(DimensionOutlineErrorReason.INVALID_RANGE, 'Outline range is outside its axis')
  if (end >= count) throw new SheetOutlineError(DimensionOutlineErrorReason.OUT_OF_BOUNDS, 'Outline range is outside its axis')
}

function assertIndex(index: number, count: number): void {
  if (!Number.isSafeInteger(count) || count < 1 || !Number.isSafeInteger(index) || index < 0 || index >= count) throw new RangeError('Outline index is outside its axis')
}

function compareOutlineRanges(left: SheetOutlineGroup, right: SheetOutlineGroup): number {
  return left.start - right.start || right.end - left.end
}

export function outlineGroupDepths(groups: readonly SheetOutlineGroup[]): Map<string, number> {
  if (!Array.isArray(groups)) throw new SheetOutlineError(DimensionOutlineErrorReason.INVALID_RANGE, 'Outline groups must be an array')
  for (const group of groups) {
    if (!group || typeof group.id !== 'string' || !group.id.trim()) throw new SheetOutlineError(DimensionOutlineErrorReason.INVALID_RANGE, 'Outline identity is invalid')
    assertOutlineAxis(group.axis)
    validateRange(group.start, group.end, group.axis === 'row' ? 1_048_576 : 16_384)
  }
  const depths = new Map<string, number>()
  const parentsByAxis = new Map<OutlineAxis, SheetOutlineGroup[]>()
  for (const group of [...groups].sort((left, right) => left.axis.localeCompare(right.axis) || compareOutlineRanges(left, right))) {
    assertOutlineAxis(group.axis)
    const parents = parentsByAxis.get(group.axis) ?? []
    parentsByAxis.set(group.axis, parents)
    while (parents.length && group.start > parents[parents.length - 1]!.end) parents.pop()
    const parent = parents.at(-1)
    if (parent && group.end > parent.end) throw new SheetOutlineError(DimensionOutlineErrorReason.CROSSING, 'Outline groups must be nested or disjoint', group.id)
    depths.set(group.id, parents.length)
    parents.push(group)
  }
  return depths
}

export function validateSheetOutlineState(state: SheetOutlineState, extent?: SheetOutlineExtent): void {
  if (!state || !Array.isArray(state.groups) || typeof state.summaryAbove !== 'boolean' || typeof state.summaryLeft !== 'boolean') {
    throw new SheetOutlineError(DimensionOutlineErrorReason.INVALID_RANGE, 'Outline state is invalid')
  }
  if (extent && (!Number.isSafeInteger(extent.rowCount) || extent.rowCount < 1 || extent.rowCount > 1_048_576
    || !Number.isSafeInteger(extent.columnCount) || extent.columnCount < 1 || extent.columnCount > 16_384)) {
    throw new SheetOutlineError(DimensionOutlineErrorReason.OUT_OF_BOUNDS, 'Outline worksheet extent is invalid')
  }
  const ids = new Set<string>()
  for (const group of state.groups) {
    if (!group || typeof group.id !== 'string' || !group.id.trim() || ids.has(group.id) || typeof group.collapsed !== 'boolean') {
      throw new SheetOutlineError(DimensionOutlineErrorReason.INVALID_RANGE, 'Outline identity or collapsed state is invalid', group?.id)
    }
    assertOutlineAxis(group.axis)
    validateRange(group.start, group.end, extent ? group.axis === 'row' ? extent.rowCount : extent.columnCount
      : group.axis === 'row' ? 1_048_576 : 16_384)
    ids.add(group.id)
  }
  for (const [id, depth] of outlineGroupDepths(state.groups)) if (depth >= MAX_SHEET_OUTLINE_DEPTH) {
    throw new SheetOutlineError(DimensionOutlineErrorReason.MAX_DEPTH, `Outline depth cannot exceed ${MAX_SHEET_OUTLINE_DEPTH}`, id)
  }
}

export function buildSheetOutlineTree(state: SheetOutlineState, axis?: OutlineAxis): SheetOutlineTreeNode[] {
  validateSheetOutlineState(state)
  if (axis !== undefined) assertOutlineAxis(axis)
  const roots: SheetOutlineTreeNode[] = [], parents = new Map<OutlineAxis, SheetOutlineTreeNode[]>()
  const groups = state.groups.filter(group => axis === undefined || group.axis === axis)
    .sort((left, right) => left.axis.localeCompare(right.axis) || compareOutlineRanges(left, right))
  for (const group of groups) {
    const stack = parents.get(group.axis) ?? []; parents.set(group.axis, stack)
    while (stack.length && group.start > stack.at(-1)!.end) stack.pop()
    const parent = stack.at(-1), node: SheetOutlineTreeNode = { ...group, depth: stack.length + 1,
      ...(parent ? { parentId: parent.id } : {}), children: [] }
    if (parent) parent.children.push(node); else roots.push(node)
    stack.push(node)
  }
  return roots
}

export function assertOutlineAxis(axis: unknown): asserts axis is OutlineAxis {
  if (axis !== 'row' && axis !== 'column') throw new SheetOutlineError(DimensionOutlineErrorReason.INVALID_RANGE, 'Invalid outline axis')
}

function validateOutlineAxisState(state: OutlineAxisState): void {
  if (!state || !Array.isArray(state.groups) || !Number.isSafeInteger(state.count) || state.count < 1) {
    throw new SheetOutlineError(DimensionOutlineErrorReason.INVALID_RANGE, 'Outline axis state is invalid')
  }
  const ids = new Set<string>()
  for (const group of state.groups) {
    if (!group || typeof group.id !== 'string' || !group.id.trim() || ids.has(group.id) || typeof group.collapsed !== 'boolean') {
      throw new SheetOutlineError(DimensionOutlineErrorReason.INVALID_RANGE, 'Outline identity or collapsed state is invalid')
    }
    assertOutlineAxis(group.axis); validateRange(group.start, group.end, state.count); ids.add(group.id)
  }
}

function assertOutlineRangesEditable(sheet: OutlineWorksheet, ranges: readonly OutlineRange[], userId: string): void {
  if (!canEditProtectedRanges(sheet, ranges.map(range => outlineCellRange(sheet, range)), userId)) throw new Error('该区域受保护，当前用户没有编辑权限。')
}

function outlineCellRange(sheet: OutlineWorksheet, range: OutlineRange): CellRange {
  return range.axis === 'row'
    ? { startRow: range.start, endRow: range.end, startCol: 0, endCol: sheet.columnCount - 1 }
    : { startRow: 0, endRow: sheet.rowCount - 1, startCol: range.start, endCol: range.end }
}

function contains(outerStart: number, outerEnd: number, innerStart: number, innerEnd: number): boolean {
  return outerStart <= innerStart && outerEnd >= innerEnd
}
