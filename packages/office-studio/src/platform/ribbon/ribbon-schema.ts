import type { Component } from 'vue'

/** State is supplied by the active editor selection, never stored by a toolbar item. */
export interface RibbonItem {
  id: string
  type: 'button' | 'select' | 'menu' | 'split'
  title: string
  tooltip?: string
  icon?: Component
  order?: number
  hidden?: boolean
  disabled?: boolean
  disabledReason?: string
  activated?: boolean
  commandId?: string
  params?: unknown
  width?: number
  showLabel?: boolean
  shortcut?: string
  swatch?: string
  value?: string | number | null
  options?: RibbonOption[]
  placeholder?: string
  valueParams?: (value: string | number) => unknown
  children?: RibbonItem[]
  popupComponent?: Component
  popupProps?: Record<string, unknown>
  popupCommandId?: string
  gridLayout?: RibbonGridLayout
}

export interface RibbonGridLayout {
  row: number
  column: number
  rowSpan?: number
  columnSpan?: number
  width?: number
  minWidth?: number
  iconSize?: number
  showLabel?: boolean
}

export interface RibbonOption {
  label: string
  value: string | number
  disabled?: boolean
  commandId?: string
  params?: unknown
}

export interface RibbonGroup {
  id: string
  title?: string
  order?: number
  hidden?: boolean
  items: RibbonItem[]
}

export interface RibbonTab {
  id: string
  title: string
  order?: number
  hidden?: boolean
  contextual?: boolean
  groups: RibbonGroup[]
}

/** Fixed reference Ribbon/ClassicMenu/ToolbarButton dimensions, in CSS pixels. */
export const CLASSIC_RIBBON = Object.freeze({
  tabHeight: 36, tabButtonHeight: 28, toolbarHeight: 40,
  buttonHeight: 24, selectorHeight: 28, iconSize: 16,
  itemGap: 8, groupPadding: 8, separatorWidth: 1,
  horizontalPadding: 12, moreWidth: 32,
})

export const GRID_RIBBON = Object.freeze({
  tabHeight: 32,
  toolbarHeight: 73,
  rowHeight: 28,
  rowGap: 8,
  verticalPadding: 4,
  groupPadding: 8,
  compactGroupPadding: 6,
  itemGap: 8,
  compactItemGap: 4,
  moreWidth: 64,
  buttonSize: 28,
  largeButtonSize: 56,
  largeIconSize: 24,
})

function ordered<T extends { order?: number }>(values: readonly T[]): T[] {
  return values.map((value, index) => ({ value, index }))
    .sort((a, b) => (a.value.order ?? 0) - (b.value.order ?? 0) || a.index - b.index)
    .map(({ value }) => value)
}

export function visibleRibbonItems(items: readonly RibbonItem[]): RibbonItem[] {
  return ordered(items).flatMap((item) => {
    if (item.hidden) return []
    if (item.type === 'menu' || item.type === 'split') {
      const children = visibleRibbonItems(item.children ?? [])
      if (!children.length && item.type === 'menu') return []
      return [{ ...item, children }]
    }
    return [item]
  })
}

/** Hidden controls and empty groups/tabs do not consume layout or keyboard positions. */
export function visibleRibbonTabs(tabs: readonly RibbonTab[]): RibbonTab[] {
  return ordered(tabs).flatMap((tab) => {
    if (tab.hidden) return []
    const groups = ordered(tab.groups).flatMap((group) => {
      if (group.hidden) return []
      const items = visibleRibbonItems(group.items)
      return items.length ? [{ ...group, items }] : []
    })
    return groups.length ? [{ ...tab, groups }] : []
  })
}

export function resolveRibbonTab(tabs: readonly RibbonTab[], requested?: string, previousRegular?: string): string | undefined {
  const visible = visibleRibbonTabs(tabs)
  if (visible.some((tab) => tab.id === requested)) return requested
  if (visible.some((tab) => tab.id === previousRegular)) return previousRegular
  return visible[0]?.id
}

export function nextRibbonTab(tabs: readonly RibbonTab[], current: string, key: string, direction: 'ltr' | 'rtl' = 'ltr'): string | undefined {
  const visible = visibleRibbonTabs(tabs)
  if (!visible.length) return undefined
  const index = Math.max(0, visible.findIndex((tab) => tab.id === current))
  if (key === 'Home') return visible[0]?.id
  if (key === 'End') return visible.at(-1)?.id
  if (key !== 'ArrowLeft' && key !== 'ArrowRight') return undefined
  const forward = (key === 'ArrowRight') !== (direction === 'rtl')
  return visible[(index + (forward ? 1 : -1) + visible.length) % visible.length]?.id
}

export interface RibbonMeasurement { id: string; groupId: string; width: number }

/** Collapse a suffix in source order, including each visible group's padding and borders. */
export function fitRibbonItems(items: readonly RibbonMeasurement[], availableWidth: number): { visibleIds: string[]; overflowIds: string[] } {
  const usableWidth = Math.max(0, availableWidth - 2 * CLASSIC_RIBBON.horizontalPadding)
  const widths = items.map((item, index) => Math.max(0, item.width) + (
    index > 0 && item.groupId === items[index - 1]?.groupId
      ? CLASSIC_RIBBON.itemGap
      : 2 * CLASSIC_RIBBON.groupPadding + (index > 0 ? CLASSIC_RIBBON.separatorWidth : 0)
  ))
  if (widths.reduce((sum, width) => sum + width, 0) <= usableWidth) return { visibleIds: items.map((item) => item.id), overflowIds: [] }
  const limit = Math.max(0, usableWidth - CLASSIC_RIBBON.moreWidth)
  let count = 0
  let occupied = 0
  for (const width of widths) {
    // More has an inline divider only when at least one group remains visible.
    if (occupied + width + CLASSIC_RIBBON.separatorWidth > limit) break
    occupied += width
    count++
  }
  return { visibleIds: items.slice(0, count).map((item) => item.id), overflowIds: items.slice(count).map((item) => item.id) }
}

export interface RibbonGridPlacement {
  item: RibbonItem
  row: number
  column: number
  rowSpan: number
  columnSpan: number
  showLabel: boolean
  width?: number
  minWidth?: number
  iconSize?: number
}

export interface RibbonGridGroupWidth {
  regular: number
  compact: number
}

export function placeRibbonGridItems(items: readonly RibbonItem[]): RibbonGridPlacement[] {
  const occupied = new Set<string>()
  const placements = new Map<string, RibbonGridPlacement>()
  const fallbackItems: Array<{ item: RibbonItem; layout?: RibbonGridPlacement }> = []

  for (const item of items) {
    const layout = item.gridLayout
    if (!layout) {
      fallbackItems.push({ item })
      continue
    }

    const rowSpan = layout.rowSpan ?? 1
    const columnSpan = layout.columnSpan ?? 1
    const cells = ribbonGridCells(layout.row, layout.column, rowSpan, columnSpan)
    if (!isValidRibbonGridLayout(layout, rowSpan, cells)) {
      fallbackItems.push({ item })
      continue
    }

    const placement: RibbonGridPlacement = {
      item,
      row: layout.row,
      column: layout.column,
      rowSpan,
      columnSpan,
      showLabel: layout.showLabel ?? false,
      width: layout.width,
      minWidth: layout.minWidth,
      iconSize: layout.iconSize,
    }
    if (cells.some((cell) => occupied.has(cell))) {
      fallbackItems.push({ item, layout: placement })
      continue
    }

    cells.forEach((cell) => occupied.add(cell))
    placements.set(item.id, placement)
  }

  for (const { item, layout } of fallbackItems) {
    const rowSpan = layout?.rowSpan ?? 1
    const columnSpan = layout?.columnSpan ?? 1
    const { row, column, cells } = firstAvailableRibbonGridCells(occupied, rowSpan, columnSpan)
    cells.forEach((cell) => occupied.add(cell))
    placements.set(item.id, {
      item,
      row,
      column,
      rowSpan,
      columnSpan,
      showLabel: layout?.showLabel ?? item.showLabel ?? false,
      width: layout?.width ?? item.width,
      minWidth: layout?.minWidth,
      iconSize: layout?.iconSize,
    })
  }

  return items.map((item) => placements.get(item.id)!).sort((a, b) => a.row - b.row || a.column - b.column)
}

export function groupRibbonGridOverflowItems(items: readonly RibbonItem[]): RibbonItem[][] {
  const placements = placeRibbonGridItems(items)
  const groups: RibbonItem[][] = []
  placements.forEach((placement, index) => {
    const previous = placements[index - 1]
    if (!previous || placement.row !== previous.row || Boolean(placement.item.icon) !== Boolean(previous.item.icon)) {
      groups.push([])
    }
    groups[groups.length - 1]?.push(placement.item)
  })
  return groups
}

export function fitRibbonGridGroups(groups: readonly RibbonGridGroupWidth[], availableWidth: number): { visibleCount: number; compact: boolean } {
  const regularTotal = groups.reduce((sum, group) => sum + Math.max(0, group.regular), 0)
  const compact = regularTotal > Math.max(0, availableWidth)
  const widths = groups.map((group) => Math.max(0, compact ? group.compact : group.regular))
  const totalWidth = widths.reduce((sum, width) => sum + width, 0)
  const limit = totalWidth <= availableWidth ? availableWidth : Math.max(0, availableWidth - GRID_RIBBON.moreWidth)
  let used = 0
  let visibleCount = 0
  for (const width of widths) {
    if (used + width > limit) break
    used += width
    visibleCount += 1
  }
  return { visibleCount, compact }
}

function isValidRibbonGridLayout(layout: RibbonGridLayout, rowSpan: number, cells: string[]): boolean {
  return cells.length > 0
    && layout.row + rowSpan - 1 <= 2
    && (layout.width === undefined || (Number.isFinite(layout.width) && layout.width > 0))
    && (layout.minWidth === undefined || (Number.isFinite(layout.minWidth) && layout.minWidth > 0 && layout.minWidth <= (layout.width ?? 0)))
    && (layout.iconSize === undefined || (Number.isFinite(layout.iconSize) && layout.iconSize > 0))
}

function firstAvailableRibbonGridCells(occupied: Set<string>, rowSpan: number, columnSpan: number): { row: number; column: number; cells: string[] } {
  for (let column = 1; ; column += 1) {
    for (let row = 1; row + rowSpan - 1 <= 2; row += 1) {
      const cells = ribbonGridCells(row, column, rowSpan, columnSpan)
      if (cells.every((cell) => !occupied.has(cell))) return { row, column, cells }
    }
  }
}

function ribbonGridCells(row: number, column: number, rowSpan: number, columnSpan: number): string[] {
  if (![row, column, rowSpan, columnSpan].every(Number.isInteger) || Math.min(row, column, rowSpan, columnSpan) < 1) return []
  const cells: string[] = []
  for (let currentRow = row; currentRow < row + rowSpan; currentRow += 1) {
    for (let currentColumn = column; currentColumn < column + columnSpan; currentColumn += 1) {
      cells.push(`${currentRow}:${currentColumn}`)
    }
  }
  return cells
}

export interface RibbonInvocation { commandId: string; params?: unknown }

/** A second guard at dispatch time prevents disabled or stale option invocations. */
export function ribbonInvocation(item: RibbonItem, value?: string | number): RibbonInvocation | undefined {
  if (item.hidden || item.disabled || item.type === 'menu') return undefined
  if (item.type === 'select') {
    if (value === undefined) return undefined
    const option = item.options?.find((candidate) => candidate.value === value)
    if (!option || option.disabled) return undefined
    const commandId = option.commandId ?? item.commandId
    if (!commandId) return undefined
    return { commandId, params: option.params !== undefined ? option.params : item.valueParams ? item.valueParams(value) : value }
  }
  return item.commandId ? { commandId: item.commandId, params: item.params } : undefined
}
