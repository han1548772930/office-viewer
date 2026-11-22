import { h, type Component, type VNodeChild } from 'vue'
import type { MenuItemType } from 'antdv-next'
import type { RibbonItem } from './ribbon-schema'

export interface RibbonMenuAction {
  item: RibbonItem
  value?: string | number
}

export interface RibbonMenuModel {
  items: MenuItemType[]
  actions: Map<string, RibbonMenuAction>
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

  function createItems(source: readonly RibbonItem[]): MenuItemType[] {
    return source.map((item) => {
      const icon = itemIcon(item)
      const base = {
        'data-ribbon-item': item.id,
        disabled: item.disabled,
        icon: icon ? h(icon, { size: 16 }) : item.swatch ? h('span', { 'aria-hidden': true, style: { display: 'inline-block', width: '14px', height: '14px', borderRadius: '2px', border: '1px solid var(--ribbon-border)', backgroundColor: item.swatch } }) : undefined,
        title: item.disabledReason && item.disabled ? `${item.title}: ${item.disabledReason}` : item.title,
      }

      if (item.type === 'select') {
        const options = (item.options ?? []).map((option, index) => {
          const key = `${item.id}::option:${index}`
          actions.set(key, { item, value: option.value })
          return {
            key,
            ribbonItemId: item.id,
            label: makeLabel(item, option.label),
            disabled: item.disabled || option.disabled,
            title: option.label,
          }
        })
        return { key: item.id, ribbonItemId: item.id, ...base, label: makeLabel(item, item.title, renderItem), disabled: item.disabled || !options.length, children: options }
      }

      const children = item.children ?? []
      if ((item.type === 'menu' || item.type === 'split') && children.length) {
        const submenuItems: MenuItemType[] = []
        if (item.type === 'split' && item.commandId) {
          const key = `${item.id}::main`
          actions.set(key, { item })
          submenuItems.push({ key, ribbonItemId: item.id, label: makeLabel(item), disabled: item.disabled, icon: base.icon, title: item.title })
          submenuItems.push({ type: 'divider', key: `${item.id}::divider` })
        }
        submenuItems.push(...createItems(children))
        return { key: item.id, ribbonItemId: item.id, ...base, label: makeLabel(item, item.title, renderItem), disabled: item.disabled || !submenuItems.length, children: submenuItems }
      }

      actions.set(item.id, { item })
      return { key: item.id, ribbonItemId: item.id, ...base, label: makeLabel(item, item.title, renderItem), disabled: item.disabled }
    })
  }

  return { items: createItems(items), actions }
}
