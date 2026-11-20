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
  value?: string | number | null
  options?: RibbonOption[]
  placeholder?: string
  valueParams?: (value: string | number) => unknown
  children?: RibbonItem[]
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

export const CLASSIC_RIBBON = Object.freeze({ tabHeight: 36, toolbarHeight: 40, itemGap: 4, groupGap: 12, horizontalPadding: 12, moreWidth: 36 })

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

/** Collapse a suffix in source order; account for gaps only between remaining items/groups. */
export function fitRibbonItems(items: readonly RibbonMeasurement[], availableWidth: number): { visibleIds: string[]; overflowIds: string[] } {
  const usableWidth = Math.max(0, availableWidth - 2 * CLASSIC_RIBBON.horizontalPadding)
  const widths = items.map((item, index) => Math.max(0, item.width) + (index === 0 ? 0 : item.groupId === items[index - 1]?.groupId ? CLASSIC_RIBBON.itemGap : CLASSIC_RIBBON.groupGap))
  if (widths.reduce((sum, width) => sum + width, 0) <= usableWidth) return { visibleIds: items.map((item) => item.id), overflowIds: [] }
  const limit = Math.max(0, usableWidth - CLASSIC_RIBBON.moreWidth - CLASSIC_RIBBON.itemGap)
  let count = 0
  let occupied = 0
  for (const width of widths) {
    if (occupied + width > limit) break
    occupied += width
    count++
  }
  return { visibleIds: items.slice(0, count).map((item) => item.id), overflowIds: items.slice(count).map((item) => item.id) }
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
