<script setup lang="ts">
import { computed, h, inject, nextTick, onBeforeUnmount, onMounted, ref, useId, useSlots, watch } from 'vue'
import { Button, Menu, Popover, Tooltip } from 'antdv-next'
import type { MenuInfo, MenuItemType, MenuRef } from 'antdv-next'
import { ChevronDown, Database, EllipsisVertical, Eye, House, LayoutPanelLeft, ListPlus, MoreVertical, Sigma } from 'lucide-vue-next'
import RibbonItemControl from './RibbonItem.vue'
import RibbonGridToolbar from './RibbonGridToolbar.vue'
import { CLASSIC_RIBBON, GRID_RIBBON, fitRibbonItems, nextRibbonTab, resolveRibbonTab, ribbonInvocation, simpleRibbonTab, visibleRibbonTabs, type RibbonGroup, type RibbonItem, type RibbonTab, type RibbonVariant } from './ribbon-schema'
import { OfficeDirectionKey, OfficeTranslatorKey } from '../../i18n/office-i18n'
import { provideRibbonPopupScope } from './ribbon-popup-scope'
import { useRibbonTheme } from './ribbon-theme'
import { createRibbonMenuModel, renderOverflowControl, type RibbonMenuAction } from './ribbon-command-menu'

const props = withDefaults(defineProps<{ tabs: RibbonTab[]; modelValue?: string; direction?: 'ltr' | 'rtl'; label?: string; variant?: RibbonVariant }>(), { modelValue: undefined, label: '功能区', variant: 'classic' })
const emit = defineEmits<{ 'update:modelValue': [id: string]; execute: [commandId: string, params?: unknown] }>()
const id = useId()
const slots = useSlots()
const translate = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const workspaceDirection = inject(OfficeDirectionKey, ref<'ltr' | 'rtl'>('ltr'))
const direction = computed(() => props.direction ?? workspaceDirection.value)
const popupScope = provideRibbonPopupScope(direction)
const themeStyle = useRibbonTheme()
const ribbonStyle = computed(() => ({
  '--ribbon-tab-height': `${props.variant === 'grid' ? GRID_RIBBON.tabHeight : CLASSIC_RIBBON.tabHeight}px`,
  '--ribbon-tab-button-height': `${props.variant === 'grid' ? GRID_RIBBON.tabHeight : CLASSIC_RIBBON.tabButtonHeight}px`,
  '--ribbon-toolbar-height': `${props.variant === 'grid' ? GRID_RIBBON.toolbarHeight : CLASSIC_RIBBON.toolbarHeight}px`,
  '--ribbon-button-height': `${CLASSIC_RIBBON.buttonHeight}px`,
  '--ribbon-item-gap': `${CLASSIC_RIBBON.itemGap}px`,
  '--ribbon-group-padding': `${CLASSIC_RIBBON.groupPadding}px`,
  '--ribbon-separator-width': `${CLASSIC_RIBBON.separatorWidth}px`,
  '--ribbon-horizontal-padding': `${props.variant === 'grid' ? GRID_RIBBON.groupPadding : CLASSIC_RIBBON.horizontalPadding}px`,
}))
const selectedId = ref(props.modelValue)
const previousRegular = ref<string>()
const tabs = computed(() => visibleRibbonTabs(props.tabs))
const activeId = computed(() => resolveRibbonTab(props.tabs, props.modelValue ?? selectedId.value, previousRegular.value))
const activeTab = computed(() => props.variant === 'simple' ? simpleRibbonTab(tabs.value) : tabs.value.find((tab) => tab.id === activeId.value))
const singleRow = computed(() => props.variant === 'simple' || props.variant === 'collapsed')
const selector = ref<HTMLElement>()
const selectorOpen = ref(false)
const selectorMenu = ref<MenuRef | null>(null)
const tabIcons = { home: House, insert: ListPlus, formulas: Sigma, data: Database, view: Eye, layout: LayoutPanelLeft }
const categoryItems = computed<MenuItemType[]>(() => tabs.value.map(tab => ({ key: tab.id, label: tab.title, icon: h(tabIcons[tab.id as keyof typeof tabIcons] ?? EllipsisVertical, { size: 18 }) })))
function closeSelector() { selectorOpen.value = false; selector.value?.querySelector('button')?.focus({ preventScroll: true }) }
async function openSelector() { selectorOpen.value = true; await nextTick(); selectorMenu.value?.focus() }
const toolbar = ref<HTMLElement>()
const measurer = ref<HTMLElement>()
const tablist = ref<HTMLElement>()
const moreButton = ref<HTMLElement>()
const moreMenuRef = ref<MenuRef | null>(null)
const availableWidth = ref(Number.POSITIVE_INFINITY)
const measuredWidths = ref<Record<string, number>>({})
const moreOpen = ref(false)
let observer: ResizeObserver | undefined
let disposed = false
let measurePending = false

const fitted = computed(() => fitRibbonItems((activeTab.value?.groups ?? []).flatMap((group) => group.items.map((item) => ({ id: item.id, groupId: group.id, width: measuredWidths.value[item.id] ?? item.width ?? (item.type === 'select' ? 96 : 32) }))), availableWidth.value))
function groupsFor(ids: readonly string[]): RibbonGroup[] {
  const visible = new Set(ids)
  return (activeTab.value?.groups ?? []).flatMap((group) => {
    const items = group.items.filter((item) => visible.has(item.id))
    return items.length ? [{ ...group, items }] : []
  })
}
const visibleGroups = computed(() => groupsFor(fitted.value.visibleIds))
const overflowGroups = computed(() => groupsFor(fitted.value.overflowIds))
const overflowMenuModel = computed(() => {
  const items: MenuItemType[] = []
  const actions = new Map<string, RibbonMenuAction>()
  overflowGroups.value.forEach((group, index) => {
    const model = createRibbonMenuModel(group.items, (item) => slots[`item-${item.id}`]?.({
      item,
      overflow: true,
      execute: (commandId: string, params?: unknown) => {
        closeMore()
        executeSlot(item, commandId, params)
      },
    }) ?? renderOverflowControl(item, (id, params) => { closeMore(); executeSlot(item, id, params) }))
    if (index > 0) items.push({ type: 'divider', key: `group-divider:${group.id}` })
    if (group.title) items.push({ type: 'group', key: `group:${group.id}`, label: group.title, children: model.items })
    else items.push(...model.items)
    model.actions.forEach((action, key) => {
      const rootItemId = key.split('::', 1)[0]
      if (!slots[`item-${rootItemId}`] && !action.item.inlineComponent) actions.set(key, action)
    })
  })
  return { items, actions }
})

function selectTab(tabId: string) {
  const tab = tabs.value.find((candidate) => candidate.id === tabId)
  if (!tab) return
  selectedId.value = tabId
  if (!tab.contextual) previousRegular.value = tabId
  moreOpen.value = false
  emit('update:modelValue', tabId)
}

function tabStyle(tab: RibbonTab) {
  if (props.variant !== 'grid') return undefined
  const selected = tab.id === activeId.value
  return {
    height: `${GRID_RIBBON.tabHeight}px`,
    padding: '0 12px',
    borderRadius: '0',
    background: 'transparent',
    color: selected ? 'var(--ribbon-primary)' : 'var(--ribbon-text-secondary)',
    fontWeight: 600,
    boxShadow: 'none',
  }
}

function onTabKeydown(event: KeyboardEvent) {
  const target = event.target instanceof HTMLElement ? event.target.closest<HTMLElement>('[role="tab"]') : null
  if (!target) return
  const targetId = nextRibbonTab(props.tabs, target.dataset.tabId ?? '', event.key, direction.value)
  if (!targetId) return
  event.preventDefault()
  event.stopPropagation()
  selectTab(targetId)
  Array.from(tablist.value?.querySelectorAll<HTMLButtonElement>('[role="tab"]') ?? []).find((button) => button.dataset.tabId === targetId)?.focus()
}

function execute(commandId: string, params?: unknown) {
  emit('execute', commandId, params)
}

function executeSlot(item: RibbonItem, commandId: string, params?: unknown) {
  if (!item.disabled && !item.hidden) execute(commandId, params)
}

async function measure() {
  if (measurePending || disposed || props.variant === 'grid') return
  measurePending = true
  await nextTick()
  measurePending = false
  if (disposed) return
  if (toolbar.value) availableWidth.value = toolbar.value.getBoundingClientRect().width - (selector.value?.getBoundingClientRect().width ?? 0)
  const widths: Record<string, number> = {}
  for (const element of measurer.value?.querySelectorAll<HTMLElement>('[data-measure-item]') ?? []) {
    const itemId = element.dataset.measureItem
    if (itemId) widths[itemId] = element.getBoundingClientRect().width
  }
  measuredWidths.value = widths
}

function closeMore(event?: KeyboardEvent) {
  if (event) { event.preventDefault(); event.stopPropagation() }
  moreOpen.value = false
  moreButton.value?.querySelector<HTMLButtonElement>('button')?.focus()
}

async function openMore() {
  popupScope.close()
  moreOpen.value = true
  await nextTick()
  moreMenuRef.value?.focus()
}

function onMoreKeydown(event: KeyboardEvent) {
  if (event.defaultPrevented) return
  if (event.key === 'Escape') { closeMore(event); return }
}

function executeOverflowItem({ key }: MenuInfo) {
  const action = overflowMenuModel.value.actions.get(String(key))
  if (!action || action.item.disabled || action.item.hidden) return
  const invocation = ribbonInvocation(action.item, action.value)
  if (invocation) {
    closeMore()
    execute(invocation.commandId, invocation.params)
  }
}

watch(activeId, (value) => {
  popupScope.close()
  const tab = tabs.value.find((candidate) => candidate.id === value)
  if (tab && !tab.contextual) previousRegular.value = tab.id
  moreOpen.value = false
  if (value && props.modelValue !== undefined && props.modelValue !== value) emit('update:modelValue', value)
}, { immediate: true })
watch(() => props.variant, () => { moreOpen.value = selectorOpen.value = false; popupScope.close(); void measure() })
watch(selector, (element, previous) => { if (previous) observer?.unobserve(previous); if (element) observer?.observe(element) }, { flush: 'post' })
watch(selectorOpen, open => { if (open) { moreOpen.value = false; popupScope.close() } })
watch(moreOpen, (open) => { if (open) popupScope.close() })
watch(activeTab, () => { void measure() }, { flush: 'post' })
watch(toolbar, (element, previous) => {
  if (previous) observer?.unobserve(previous)
  if (element) observer?.observe(element)
  void measure()
}, { flush: 'post' })
watch(() => overflowGroups.value.length, (count) => { if (!count) moreOpen.value = false })
onMounted(() => {
  observer = new ResizeObserver(() => { void measure() })
  if (toolbar.value) observer.observe(toolbar.value)
  if (selector.value) observer.observe(selector.value)
  void measure()
})
onBeforeUnmount(() => { disposed = true; observer?.disconnect() })
</script>

<template>
  <section v-if="tabs.length" class="workbench-ribbon" data-testid="workbench-ribbon" :dir="direction" :aria-label="label" :style="[ribbonStyle, themeStyle]">
    <div v-if="!singleRow" ref="tablist" class="workbench-ribbon-tabs" :class="{ 'is-grid': variant === 'grid' }" role="tablist" :aria-label="label" aria-orientation="horizontal" @keydown="onTabKeydown">
      <Button v-for="tab in tabs" :id="`${id}-tab-${tab.id}`" :key="tab.id" type="text" role="tab" class="workbench-ribbon-tab" :class="{ 'is-selected': tab.id === activeId, 'is-contextual': tab.contextual, 'is-grid-tab': variant === 'grid' }" :style="tabStyle(tab)" :aria-selected="tab.id === activeId" :aria-controls="`${id}-panel`" :tabindex="tab.id === activeId ? 0 : -1" :data-tab-id="tab.id" @mousedown.prevent @click="selectTab(tab.id)">{{ tab.title }}</Button>
      <div v-if="$slots['tab-extra']" class="ribbon-tab-extra"><slot name="tab-extra" /></div>
    </div>
    <div v-else-if="$slots['tab-extra']" class="ribbon-tab-extra ribbon-single-row-extra"><slot name="tab-extra" /></div>
    <div :id="`${id}-panel`" ref="toolbar" class="workbench-ribbon-toolbar" :class="{ 'is-grid': variant === 'grid', 'is-single-row': singleRow }" :role="singleRow ? 'group' : 'tabpanel'" :aria-label="singleRow ? label : undefined" :aria-labelledby="singleRow ? undefined : `${id}-tab-${activeId}`">
      <RibbonGridToolbar v-if="variant === 'grid' && activeTab" :key="activeTab.id" :groups="activeTab.groups" :title="activeTab.title" :direction="direction" :align-start="activeTab.id === 'data'" @execute="execute">
        <template v-for="(_, name) in $slots" #[name]="slotProps"><slot :name="name" v-bind="slotProps" /></template>
      </RibbonGridToolbar>
      <template v-else>
        <span v-if="variant === 'collapsed'" ref="selector" class="ribbon-category-anchor">
          <Popover v-model:open="selectorOpen" :trigger="['hover', 'click']" :placement="direction === 'rtl' ? 'bottomRight' : 'bottomLeft'">
            <Button type="text" size="small" class="ribbon-category-toggle" aria-haspopup="menu" :aria-expanded="selectorOpen" @mousedown.prevent @keydown.down.prevent="openSelector" @keydown.esc.stop.prevent="closeSelector">{{ activeTab?.title }}<ChevronDown :size="14" /></Button>
            <template #content><Menu ref="selectorMenu" class="ribbon-category-menu" :items="categoryItems" :selected-keys="activeId ? [activeId] : []" @click="({ key }) => { selectTab(String(key)); closeSelector() }" @keydown.esc.stop.prevent="closeSelector" /></template>
          </Popover>
        </span>
        <div class="ribbon-visible-groups" role="toolbar" :aria-label="activeTab?.title">
          <div v-for="group in visibleGroups" :key="group.id" class="workbench-ribbon-group" role="group" :aria-label="group.title">
            <template v-for="item in group.items" :key="item.id">
              <div v-if="$slots[`item-${item.id}`]" class="ribbon-custom-item" :data-ribbon-item="item.id"><slot :name="`item-${item.id}`" :item="item" :execute="(commandId: string, params?: unknown) => executeSlot(item, commandId, params)" /></div>
              <RibbonItemControl v-else :item="item" @execute="execute" />
            </template>
          </div>
        </div>
        <span v-if="overflowGroups.length" ref="moreButton" class="ribbon-more-anchor" :class="{ 'has-visible-groups': visibleGroups.length > 0 }">
          <Popover v-model:open="moreOpen" trigger="click" :placement="direction === 'rtl' ? 'bottomLeft' : 'bottomRight'" :align="{ overflow: { shiftX: true, adjustX: true, adjustY: true } }" :destroy-on-hidden="true">
            <Tooltip :title="translate('common.more', '更多功能')"><Button type="text" size="small" class="ribbon-more-button" :aria-label="translate('common.more', '更多功能')" aria-haspopup="menu" :aria-expanded="moreOpen" @mousedown.prevent @keydown.down.prevent="openMore" @keydown.esc="closeMore"><MoreVertical :size="CLASSIC_RIBBON.iconSize" /></Button></Tooltip>
            <template #content>
              <Menu ref="moreMenuRef" class="ribbon-overflow-menu" :dir="direction" :style="themeStyle" :aria-label="translate('common.more', '更多功能')" :items="overflowMenuModel.items" :selectable="false" :trigger-sub-menu-action="'click'" @click="executeOverflowItem" @keydown.esc.capture="onMoreKeydown" />
            </template>
          </Popover>
        </span>
      </template>
    </div>
    <div v-if="variant !== 'grid'" ref="measurer" class="ribbon-measurer" aria-hidden="true" inert>
      <template v-for="group in activeTab?.groups ?? []" :key="group.id">
        <span v-for="item in group.items" :key="item.id" :data-measure-item="item.id" class="ribbon-measure-item">
          <slot v-if="$slots[`item-${item.id}`]" :name="`item-${item.id}`" :item="item" :measurement="true" :execute="() => undefined" />
          <RibbonItemControl v-else :item="item" measurement />
        </span>
      </template>
    </div>
  </section>
</template>

<style scoped>
.workbench-ribbon { position: relative; flex: 0 0 auto; min-width: 0; background: var(--ribbon-background, #fff); color: var(--ribbon-text, #334155); border-bottom: 1px solid var(--ribbon-border, #e2e8f0); }
.workbench-ribbon-tabs { display: flex; height: var(--ribbon-tab-height); min-width: 0; gap: 4px; align-items: center; justify-content: safe center; padding: 0 var(--ribbon-horizontal-padding); overflow-x: auto; scrollbar-width: none; border-radius: 6px; background: var(--ribbon-surface, #f8fafc); box-sizing: border-box; }
.workbench-ribbon-tabs::-webkit-scrollbar { display: none; }
.workbench-ribbon-tabs.is-grid { align-items: flex-end; justify-content: flex-start; gap: 0; padding-inline: 8px; border-radius: 0; }
.workbench-ribbon-tab { position: relative; flex: 0 0 auto; height: var(--ribbon-tab-button-height); min-width: 0; padding: 4px 8px; border-radius: 2px; color: var(--ribbon-text-secondary, #475569); font-size: 14px; line-height: 20px; }
.workbench-ribbon-tab.is-selected { color: var(--ribbon-primary, #315bc2); background: var(--ribbon-active-background, #e7edfa); font-weight: 600; box-shadow: 0 1px 2px #1d2c4914; }
.workbench-ribbon-tab.is-contextual { color: #7450af; }
.workbench-ribbon-tabs.is-grid .workbench-ribbon-tab.is-grid-tab { border: 0 !important; border-radius: 0; background: transparent; color: var(--ribbon-text-secondary, #475569); font-weight: 600; }
.workbench-ribbon-tabs.is-grid .workbench-ribbon-tab.is-grid-tab::after { position: absolute; inset-inline: 0; bottom: 0; height: 3px; background: transparent; content: ''; }
.workbench-ribbon-tabs.is-grid .workbench-ribbon-tab.is-grid-tab.is-selected { background: transparent; color: var(--ribbon-primary, #315bc2); box-shadow: none; }
.workbench-ribbon-tabs.is-grid .workbench-ribbon-tab.is-grid-tab.is-selected::after { background: var(--ribbon-primary, #315bc2); }
.workbench-ribbon-tab:focus-visible { outline: 2px solid var(--ribbon-primary, #5274d7); outline-offset: -3px; }
.ribbon-tab-extra { display: flex; flex: 0 0 auto; align-items: center; margin-inline-start: auto; }
.ribbon-tab-extra:empty { display: none; }
.workbench-ribbon-toolbar { display: flex; min-width: 0; height: var(--ribbon-toolbar-height); align-items: center; justify-content: center; padding: 0 var(--ribbon-horizontal-padding); box-sizing: border-box; }
.workbench-ribbon-toolbar.is-grid { display: block; padding: 0; }
.workbench-ribbon-toolbar.is-single-row { justify-content: flex-start; }
.ribbon-category-anchor { display: inline-flex; flex: none; padding-inline-end: 8px; }
.ribbon-category-anchor .ribbon-category-toggle { height: 28px; padding-inline: 12px 8px; gap: 6px; border-radius: 14px; background: var(--ribbon-text); color: var(--ribbon-background); white-space: nowrap; }
.ribbon-category-menu { max-width: 384px; min-width: 160px; }
.ribbon-single-row-extra { justify-content: flex-end; }
.ribbon-visible-groups { display: flex; min-width: 0; align-items: center; }
.workbench-ribbon-group { display: flex; flex: 0 0 auto; gap: var(--ribbon-item-gap); align-items: center; padding: 0 var(--ribbon-group-padding); box-sizing: border-box; }
.workbench-ribbon-group + .workbench-ribbon-group { border-inline-start: var(--ribbon-separator-width) solid var(--ribbon-border, #e2e8f0); }
.ribbon-more-anchor { display: flex; flex: 0 0 auto; padding-inline-start: var(--ribbon-group-padding); }
.ribbon-more-anchor.has-visible-groups { border-inline-start: var(--ribbon-separator-width) solid var(--ribbon-border, #e2e8f0); }
.ribbon-more-button { display: inline-flex; height: var(--ribbon-button-height); width: var(--ribbon-button-height); min-width: var(--ribbon-button-height); align-items: center; justify-content: center; padding: 0; }
.ribbon-custom-item { display: inline-flex; align-items: center; min-height: 28px; flex: 0 0 auto; }
.ribbon-measurer { position: absolute; inset-inline-start: 0; top: 0; display: flex; width: 0; height: 0; overflow: hidden; visibility: hidden; pointer-events: none; contain: layout; }
.ribbon-measure-item { display: inline-flex; flex: 0 0 auto; }
.ribbon-overflow-menu { box-sizing: border-box; width: max-content; min-width: min(210px, calc(100vw - 48px)); max-width: calc(100vw - 48px); max-height: min(520px, 70vh); overflow: auto; }
.ribbon-overflow-menu :deep(.ant-menu-item-group-title) { padding: 3px 8px; color: var(--ribbon-text-secondary, #64748b); font-size: 11px; }
.ribbon-overflow-menu :deep(.ant-menu-item), .ribbon-overflow-menu :deep(.ant-menu-submenu-title) { min-height: 30px; height: auto; line-height: 22px; padding-block: 4px; white-space: normal; }
.ribbon-overflow-menu :deep(.ribbon-menu-item-label) { display: flex; min-width: 0; align-items: center; gap: 12px; }
.ribbon-overflow-menu :deep(.ribbon-menu-item-text) { min-width: 0; overflow-wrap: anywhere; }
.ribbon-overflow-menu :deep(.ribbon-menu-item-shortcut) { flex: 0 0 auto; margin-inline-start: auto; color: var(--ribbon-text-secondary, #94a3b8); font-size: 11px; }
</style>
