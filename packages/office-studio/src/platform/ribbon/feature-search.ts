import { Keyboard, Search } from 'lucide-vue-next'
import { equivalentModelData } from '../commands/model-equivalence.ts'
import { ribbonInvocation, visibleRibbonTabs, type RibbonGroup, type RibbonInvocation, type RibbonItem, type RibbonTab } from './ribbon-schema.ts'
import { OPEN_FEATURE_SEARCH, TOGGLE_SHORTCUT_PANEL } from './workbench-operations'

export function featureSearchRibbonItem(title: string): RibbonItem {
  return { id: OPEN_FEATURE_SEARCH, commandId: OPEN_FEATURE_SEARCH, type: 'button', icon: Search, title,
    order: 2, shortcut: 'Ctrl+Shift+P', gridLayout: { row: 2, column: 2 } }
}

export function shortcutPanelRibbonItem(title: string, activated?: boolean): RibbonItem {
  return { id: TOGGLE_SHORTCUT_PANEL, commandId: TOGGLE_SHORTCUT_PANEL, type: 'button', icon: Keyboard, title, activated,
    order: 1, shortcut: 'Ctrl+\\', gridLayout: { row: 1, column: 2 } }
}

export interface FeatureSearchResult extends RibbonInvocation {
  id: string
  title: string
  path: string
  keywords: string
  contextId?: string
}

export function collectFeatureSearchResults(tabs: readonly RibbonTab[], ribbonLabel: string, context?: {
  title: string
  groups: readonly RibbonGroup[]
}): FeatureSearchResult[] {
  const results: FeatureSearchResult[] = []
  const byCommand = new Map<string, RibbonInvocation[]>()
  function add(id: string, title: string, parents: readonly string[], invocation: RibbonInvocation, tooltip?: string, contextId?: string) {
    if (invocation.commandId === OPEN_FEATURE_SEARCH) return
    const path = parents.filter(Boolean).join(' / ')
    const seen = byCommand.get(invocation.commandId) ?? []
    if (seen.some(item => equivalentModelData(item.params, invocation.params))) return
    seen.push(invocation)
    byCommand.set(invocation.commandId, seen)
    results.push({ ...invocation, id, title, path, contextId, keywords: [title, path, tooltip, invocation.commandId].join(' ') })
  }
  function visitItems(items: readonly RibbonItem[], parents: string[], keys: string[], isContext = false) {
    for (const item of items) {
      if (item.hidden || item.disabled) continue
      const key = [...keys, item.id]
      const invocation = ribbonInvocation(item)
      if (invocation && !item.inlineComponent) add(JSON.stringify(key), item.title, parents, invocation, item.tooltip, isContext ? item.id : undefined)
      for (const option of item.options ?? []) {
        const action = ribbonInvocation(item, option.value)
        if (action) add(JSON.stringify([...key, option.value]), option.label, [...parents, item.title], action, undefined, isContext ? item.id : undefined)
      }
      if (item.children) visitItems(item.children, [...parents, item.title], key, isContext)
    }
  }
  for (const tab of visibleRibbonTabs(tabs)) {
    for (const group of tab.groups) {
      visitItems(group.items, [ribbonLabel, tab.title, group.title ?? ''], ['ribbon', tab.id, group.id])
    }
  }
  if (context) {
    for (const group of context.groups) {
      if (!group.hidden) visitItems(group.items, [context.title, group.title ?? ''], ['context', group.id], true)
    }
  }
  return results
}
