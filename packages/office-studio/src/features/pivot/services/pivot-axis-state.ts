import type { CellScalar, PivotCollapseInfo, PivotCollapseState, PivotFieldModel, PivotTableModel } from '../../../domain/model.ts'

export type PivotAxisArea = 'row' | 'column'
export interface PivotAxisFieldState {
  id: string
  name: string
  position: number
  collapsed: boolean
  partiallyCollapsed: boolean
  canExpand: boolean
  canCollapse: boolean
}
export interface PivotAxisState {
  area: PivotAxisArea
  fields: PivotAxisFieldState[]
  /** The shallowest collapsed boundary, or the last field when fully expanded. */
  categoryLevelId?: string
  categoryLevelIndex: number
  mixed: boolean
  canExpand: boolean
  canCollapse: boolean
}
export interface PivotAxisGroup {
  values: CellScalar[]
  key: string
  memberValues: Record<string, CellScalar[]>
}
export interface PivotAxisProjectionEntry {
  group: PivotAxisGroup
  subtotalLevel?: number
  collapsedLevel?: number
}

const record = (value: unknown): value is Record<string, unknown> => !!value && typeof value === 'object' && !Array.isArray(value)
const key = (value: CellScalar): string => String(value ?? '')
const hasCollapsed = (value: PivotCollapseState | undefined): boolean => value === true || record(value) && Object.values(value).some(item => item === true)

export function pivotAxisFields(pivot: PivotTableModel, area: PivotAxisArea): PivotFieldModel[] {
  if (area !== 'row' && area !== 'column') throw new TypeError('Pivot hierarchy requires a row or column axis')
  return pivot.fields.filter(field => field.area === area).sort((a, b) => a.position - b.position)
}

/** Commands resolve canonical source targets before passing report-local IDs to this service. */
export function pivotAxisField(pivot: PivotTableModel, fieldId: string, requireParent = false): PivotFieldModel {
  if (typeof fieldId !== 'string' || !fieldId.trim()) throw new TypeError('Pivot hierarchy field ID is invalid')
  const field = pivot.fields.find(field => field.id === fieldId)
  if (!field || field.area !== 'row' && field.area !== 'column') throw new Error('Pivot hierarchy requires a configured row or column field')
  if (requireParent && pivotAxisFields(pivot, field.area).at(-1)?.id === field.id) throw new Error('A leaf field has no hierarchy to collapse')
  return field
}

export function pivotMemberCollapsed(info: PivotCollapseInfo | undefined, fieldId: string, value: CellScalar): boolean {
  const state = info?.[fieldId]
  return state === true || record(state) && Object.hasOwn(state, key(value)) && state[key(value)] === true
}

export function getPivotCollapse(pivot: PivotTableModel, fieldId: string, item?: string): boolean {
  pivotAxisField(pivot, fieldId)
  if (item !== undefined && typeof item !== 'string') throw new TypeError('Pivot collapse member key is invalid')
  const state = normalizePivotCollapseInfo(pivot, pivot.collapseInfo ?? {})[fieldId]
  return item === undefined ? state === true : state === true || record(state) && Object.hasOwn(state, item) && state[item] === true
}

/** Validate restored/local state separately from mutable source-member availability. */
export function normalizePivotCollapseInfo(pivot: PivotTableModel, input: unknown,
  members?: (field: PivotFieldModel) => readonly CellScalar[]): PivotCollapseInfo {
  if (!record(input)) throw new TypeError('Pivot collapse state must be a field map')
  const result: PivotCollapseInfo = Object.create(null)
  for (const [id, state] of Object.entries(input)) {
    const field = pivotAxisField(pivot, id, true)
    if (typeof state === 'boolean') { if (state) result[id] = true; continue }
    if (!record(state) || Object.values(state).some(value => typeof value !== 'boolean')) throw new TypeError('Pivot collapse member state must contain booleans')
    const available = members?.(field), mapped: Record<string, boolean> = Object.create(null)
    for (const [item, collapsed] of Object.entries(state)) {
      if (available && available.filter(value => key(value) === item).length !== 1) throw new Error('Pivot collapse member is missing or ambiguous in the grouped source')
      if (collapsed) mapped[item] = true
    }
    if (Object.keys(mapped).length) result[id] = mapped
  }
  return result
}

/** Sanctioned field removal/movement retires states that no longer have a hierarchy. */
export function retainPivotCollapseInfo(pivot: PivotTableModel, input: PivotCollapseInfo | undefined,
  previousFields?: readonly PivotFieldModel[]): PivotCollapseInfo {
  const parents = new Map([...pivotAxisFields(pivot, 'row').slice(0, -1), ...pivotAxisFields(pivot, 'column').slice(0, -1)].map(field => [field.id, field]))
  return Object.fromEntries(Object.entries(input ?? {}).filter(([id]) => {
    const field = parents.get(id), old = previousFields?.find(field => field.id === id)
    return field && (!previousFields || old && old.sourceColumn === field.sourceColumn && old.sourceFieldId === field.sourceFieldId)
  }))
}

export function preparePivotCollapse(pivot: PivotTableModel, fieldId: string, collapse: boolean,
  item: string | undefined, members: readonly CellScalar[]): PivotCollapseInfo {
  pivotAxisField(pivot, fieldId, true)
  if (typeof collapse !== 'boolean' || item !== undefined && typeof item !== 'string') throw new TypeError('Pivot collapse requires a boolean and an optional member key')
  const next = normalizePivotCollapseInfo(pivot, pivot.collapseInfo ?? {})
  if (item === undefined) {
    if (collapse) next[fieldId] = true
    else delete next[fieldId]
    return next
  }
  if (members.filter(value => key(value) === item).length !== 1) throw new Error('Pivot collapse member is missing or ambiguous in the grouped source')
  const state = next[fieldId]
  if ((state === true || record(state) && state[item] === true) === collapse) return next
  // Expanding one member of a fully collapsed field materializes the existing grouped members.
  const mapped: Record<string, boolean> = Object.assign(Object.create(null), state === true
    ? Object.fromEntries(members.map(value => [key(value), true])) : record(state) ? state : {})
  if (collapse) mapped[item] = true
  else delete mapped[item]
  if (Object.keys(mapped).length) next[fieldId] = mapped
  else delete next[fieldId]
  return next
}

export function equalPivotCollapseInfo(left: PivotCollapseInfo | undefined, right: PivotCollapseInfo | undefined): boolean {
  const stable = (info: PivotCollapseInfo | undefined) => Object.entries(info ?? {}).sort(([a], [b]) => a.localeCompare(b))
    .map(([id, state]) => [id, typeof state === 'boolean' ? state : Object.entries(state).sort(([a], [b]) => a.localeCompare(b))])
  return JSON.stringify(stable(left)) === JSON.stringify(stable(right))
}

export function preparePivotCategoryLevel(pivot: PivotTableModel, fieldId: string): PivotCollapseInfo {
  const rows = pivotAxisFields(pivot, 'row'), index = rows.findIndex(field => field.id === fieldId)
  if (index < 0) throw new Error('Pivot category level requires a configured row field')
  const next = normalizePivotCollapseInfo(pivot, pivot.collapseInfo ?? {})
  rows.slice(0, -1).forEach((field, level) => {
    if (level >= index) next[field.id] = true
    else delete next[field.id]
  })
  return next
}

export function pivotCategoryStepField(pivot: PivotTableModel, direction: 'expand' | 'collapse'): string | undefined {
  if (direction !== 'expand' && direction !== 'collapse') throw new TypeError('Pivot category step direction is invalid')
  const parents = pivotAxisFields(pivot, 'row').slice(0, -1)
  return (direction === 'expand' ? parents.find(field => hasCollapsed(pivot.collapseInfo?.[field.id]))
    : parents.reverse().find(field => pivot.collapseInfo?.[field.id] !== true))?.id
}

export function getPivotAxisState(pivot: PivotTableModel, area: PivotAxisArea = 'row'): PivotAxisState {
  const axis = pivotAxisFields(pivot, area), info = normalizePivotCollapseInfo(pivot, pivot.collapseInfo ?? {})
  const fields = axis.map((field, index): PivotAxisFieldState => {
    const state = info[field.id], parent = index < axis.length - 1
    return { id: field.id, name: field.displayName ?? field.name, position: index, collapsed: state === true,
      partiallyCollapsed: record(state) && hasCollapsed(state), canExpand: parent && hasCollapsed(state), canCollapse: parent && state !== true }
  })
  const boundary = fields.findIndex(field => field.canExpand), index = boundary < 0 ? fields.length - 1 : boundary
  return { area, fields, categoryLevelId: fields[index]?.id, categoryLevelIndex: index,
    mixed: fields.some(field => field.partiallyCollapsed), canExpand: fields.some(field => field.canExpand), canCollapse: fields.some(field => field.canCollapse) }
}

/** Grouping must not silently detach a currently addressable collapsed member from its identity. */
export function assertPivotCollapseGroupingChange(state: PivotCollapseState | undefined,
  previous: readonly CellScalar[], next: readonly CellScalar[]): void {
  if (!record(state)) return
  for (const [item, collapsed] of Object.entries(state)) if (collapsed && previous.some(value => key(value) === item)
    && next.filter(value => key(value) === item).length !== 1) throw new Error('Reset or remap collapsed members before changing their grouping')
}

/** Project existing ordered groups; aggregation stays entirely in the original Pivot engine. */
export function projectPivotAxis(groups: readonly PivotAxisGroup[], fields: readonly PivotFieldModel[],
  info: PivotCollapseInfo | undefined, showSubtotals: boolean): PivotAxisProjectionEntry[] {
  const prefixes = fields.slice(0, -1).map(() => new Map<string, { group: PivotAxisGroup; members: Map<string, Set<CellScalar>> }>())
  for (const group of groups) for (let level = 0; level < fields.length - 1; level++) {
    const values = group.values.slice(0, level + 1), groupKey = JSON.stringify(values)
    let prefix = prefixes[level]!.get(groupKey)
    if (!prefix) {
      prefix = { group: { values, key: groupKey, memberValues: {} }, members: new Map() }
      prefixes[level]!.set(groupKey, prefix)
    }
    for (const field of fields.slice(0, level + 1)) {
      const members = prefix.members.get(field.id) ?? new Set<CellScalar>()
      prefix.members.set(field.id, members)
      for (const value of group.memberValues[field.id] ?? []) members.add(value)
    }
  }
  for (const level of prefixes) for (const prefix of level.values()) {
    prefix.group.memberValues = Object.fromEntries([...prefix.members].map(([id, members]) => [id, [...members]]))
  }
  const result: PivotAxisProjectionEntry[] = [], emitted = new Set<string>()
  groups.forEach((group, index) => {
    const collapsed = fields.slice(0, -1).findIndex((field, level) => pivotMemberCollapsed(info, field.id, group.values[level] ?? null))
    if (collapsed < 0) result.push({ group })
    else {
      const prefix = prefixes[collapsed]!.get(JSON.stringify(group.values.slice(0, collapsed + 1)))!.group
      const identity = JSON.stringify([collapsed, prefix.key])
      if (!emitted.has(identity)) { emitted.add(identity); result.push({ group: prefix, collapsedLevel: collapsed }) }
    }
    if (!showSubtotals) return
    const next = groups[index + 1], lastParent = collapsed < 0 ? fields.length - 2 : collapsed - 1
    for (let level = lastParent; level >= 0; level--) {
      const values = group.values.slice(0, level + 1)
      if (next && values.every((value, depth) => Object.is(value, next.values[depth]))) continue
      const prefix = prefixes[level]!.get(JSON.stringify(values))
      if (prefix) result.push({ group: prefix.group, subtotalLevel: level })
    }
  })
  return result
}
