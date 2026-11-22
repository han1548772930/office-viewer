import { h, type Component, type VNodeChild } from 'vue'
import type { MenuItemType } from 'antdv-next'
import type { RibbonItem } from './ribbon-schema'
import { ribbonInvocation, visibleRibbonItems, type RibbonInvocation } from './ribbon-schema'

export interface RibbonMenuAction {
  item: RibbonItem
  value?: string | number
  disabled: boolean
  path: readonly RibbonItem[]
}

export interface RibbonMenuModel {
  items: MenuItemType[]
  selectedKeys: string[]
  actions: Map<string, RibbonMenuAction>
  source: readonly RibbonItem[]
}

function itemIcon(item: RibbonItem): Component | undefined {
  return item.icon
}

function makeLabel(item: RibbonItem, title = item.title, renderItem?: (item: RibbonItem) => VNodeChild) {
  if (title === item.title) {
    const custom = renderItem?.(item)
    if (custom) return custom
  }
  return h('span', { class: 'ribbon-menu-item-label' }, [
    h('span', { class: 'ribbon-menu-item-text' }, title),
    item.shortcut ? h('span', { class: 'ribbon-menu-item-shortcut' }, item.shortcut) : null,
  ])
}

export function createRibbonMenuModel(items: readonly RibbonItem[], renderItem?: (item: RibbonItem) => VNodeChild): RibbonMenuModel {
  const actions = new Map<string, RibbonMenuAction>()
  const selectedKeys: string[] = []

  function createItems(source: readonly RibbonItem[], ancestorDisabled = false, ancestors: readonly RibbonItem[] = []): MenuItemType[] {
    const originals = new Map(source.map(item => [item.id, item]))
    return visibleRibbonItems(source).map((item) => {
      const original = originals.get(item.id) ?? item
      const path = [...ancestors, original]
      const icon = itemIcon(item)
      const disabled = ancestorDisabled || Boolean(item.disabled)
      const base = {
        'data-ribbon-item': item.id,
        disabled,
        icon: item.swatch ? h('span', { 'aria-hidden': true, style: { display: 'inline-block', width: '14px', height: '14px', borderRadius: '2px', border: '1px solid var(--ribbon-border)', backgroundColor: item.swatch } }) : icon ? h(icon, { size: 16, style: { color: item.iconColor } }) : undefined,
        title: item.disabledReason && item.disabled ? `${item.title}: ${item.disabledReason}` : item.title,
      }

      if (item.type === 'select') {
        const options = (item.options ?? []).map((option, index) => {
          const key = `${item.id}::option:${index}`
          if (item.value === option.value) selectedKeys.push(key)
          actions.set(key, { item: original, value: option.value, disabled: disabled || Boolean(option.disabled), path })
          return {
            key,
            ribbonItemId: item.id,
            label: makeLabel(item, option.label),
            disabled: disabled || option.disabled,
            title: option.label,
          }
        })
        return { key: item.id, ribbonItemId: item.id, ...base, label: makeLabel(item, item.title, renderItem), disabled: disabled || !options.length, children: options }
      }

      const children = original.children ?? []
      if ((item.type === 'menu' || item.type === 'split') && children.length) {
        const submenuItems: MenuItemType[] = []
        if (item.type === 'split' && item.commandId) {
          const key = `${item.id}::main`
          if (item.activated) selectedKeys.push(key)
          actions.set(key, { item: original, disabled, path })
          submenuItems.push({ key, ribbonItemId: item.id, label: makeLabel(item), disabled, icon: base.icon, title: item.title })
          submenuItems.push({ type: 'divider', key: `${item.id}::divider` })
        }
        submenuItems.push(...createItems(children, disabled, path))
        return { key: item.id, ribbonItemId: item.id, ...base, label: makeLabel(item, item.title, renderItem), disabled: disabled || !submenuItems.length, children: submenuItems }
      }

      actions.set(item.id, { item: original, disabled, path })
      if (item.activated) selectedKeys.push(item.id)
      return { key: item.id, ribbonItemId: item.id, ...base, label: makeLabel(item, item.title, renderItem), disabled }
    })
  }

  const menuItems = createItems(items)
  return { items: menuItems, actions, selectedKeys, source: items }
}

/** Recheck the complete menu path when a submenu selection is dispatched. */
export function ribbonMenuInvocation(model: RibbonMenuModel, key: string): RibbonInvocation | undefined {
  const action = model.actions.get(key)
  const root = action?.path[0]
  if (!action || !root || !model.source.includes(root)) return undefined
  for (const [index, item] of action.path.entries()) {
    if (item.hidden || item.disabled || (index > 0 && !action.path[index - 1]?.children?.includes(item))) return undefined
  }
  return ribbonInvocation(action.item, action.value)
}
