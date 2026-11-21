<script setup lang="ts">
import { computed, inject, nextTick, onBeforeUnmount, onMounted, ref, useId, watch } from 'vue'
import { Button, Popover, Tooltip } from 'antdv-next'
import { EllipsisVertical } from 'lucide-vue-next'
import RibbonItemControl from './RibbonItem.vue'
import RibbonGridToolbar from './RibbonGridToolbar.vue'
import { CLASSIC_RIBBON, GRID_RIBBON, fitRibbonItems, nextRibbonTab, resolveRibbonTab, visibleRibbonTabs, type RibbonGroup, type RibbonItem, type RibbonTab } from './ribbon-schema'
import { OfficeTranslatorKey } from '../../i18n/office-i18n'

const props = withDefaults(defineProps<{ tabs: RibbonTab[]; modelValue?: string; direction?: 'ltr' | 'rtl'; label?: string; variant?: 'classic' | 'grid' }>(), { modelValue: undefined, direction: 'ltr', label: '功能区', variant: 'classic' })
const emit = defineEmits<{ 'update:modelValue': [id: string]; execute: [commandId: string, params?: unknown] }>()
const id = useId()
const translate = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
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
const activeTab = computed(() => tabs.value.find((tab) => tab.id === activeId.value))
const toolbar = ref<HTMLElement>()
const measurer = ref<HTMLElement>()
const tablist = ref<HTMLElement>()
const moreButton = ref<HTMLElement>()
const morePanel = ref<HTMLElement>()
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
    color: selected ? '#315bc2' : '#475569',
    fontWeight: 600,
    boxShadow: 'none',
  }
}

function onTabKeydown(event: KeyboardEvent) {
  const target = event.target instanceof HTMLElement ? event.target.closest<HTMLElement>('[role="tab"]') : null
  if (!target) return
  const targetId = nextRibbonTab(props.tabs, target.dataset.tabId ?? '', event.key, props.direction)
  if (!targetId) return
  event.preventDefault()
  event.stopPropagation()
  selectTab(targetId)
  Array.from(tablist.value?.querySelectorAll<HTMLButtonElement>('[role="tab"]') ?? []).find((button) => button.dataset.tabId === targetId)?.focus()
}

function execute(commandId: string, params?: unknown) {
  moreOpen.value = false
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
  if (toolbar.value) availableWidth.value = toolbar.value.getBoundingClientRect().width
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
  moreOpen.value = true
  await nextTick()
  morePanel.value?.querySelector<HTMLElement>('button:not(:disabled), [role="combobox"]')?.focus()
}

function onMoreKeydown(event: KeyboardEvent) {
  if (event.key === 'Escape') { closeMore(event); return }
  if (event.target instanceof HTMLElement && event.target.closest('[role="combobox"]')) return
  const controls = Array.from(morePanel.value?.querySelectorAll<HTMLElement>('button:not(:disabled), [role="combobox"]') ?? [])
  const current = controls.findIndex((control) => control === event.target || control.contains(event.target as Node))
  let next: number
  if (event.key === 'ArrowDown') next = (current + 1) % controls.length
  else if (event.key === 'ArrowUp') next = (current - 1 + controls.length) % controls.length
  else if (event.key === 'Home') next = 0
  else if (event.key === 'End') next = controls.length - 1
  else return
  event.preventDefault()
  controls[next]?.focus()
}

watch(activeId, (value) => {
  const tab = tabs.value.find((candidate) => candidate.id === value)
  if (tab && !tab.contextual) previousRegular.value = tab.id
  moreOpen.value = false
  if (value && props.modelValue !== undefined && props.modelValue !== value) emit('update:modelValue', value)
}, { immediate: true })
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
  void measure()
})
onBeforeUnmount(() => { disposed = true; observer?.disconnect() })
</script>

<template>
  <section v-if="tabs.length" class="workbench-ribbon" data-testid="workbench-ribbon" :dir="direction" :aria-label="label" :style="ribbonStyle">
    <div ref="tablist" class="workbench-ribbon-tabs" :class="{ 'is-grid': variant === 'grid' }" role="tablist" :aria-label="label" aria-orientation="horizontal" @keydown="onTabKeydown">
      <Button v-for="tab in tabs" :id="`${id}-tab-${tab.id}`" :key="tab.id" type="text" role="tab" class="workbench-ribbon-tab" :class="{ 'is-selected': tab.id === activeId, 'is-contextual': tab.contextual, 'is-grid-tab': variant === 'grid' }" :style="tabStyle(tab)" :aria-selected="tab.id === activeId" :aria-controls="`${id}-panel`" :tabindex="tab.id === activeId ? 0 : -1" :data-tab-id="tab.id" @mousedown.prevent @click="selectTab(tab.id)">{{ tab.title }}</Button>
      <div v-if="$slots['tab-extra']" class="ribbon-tab-extra"><slot name="tab-extra" /></div>
    </div>
    <div :id="`${id}-panel`" ref="toolbar" class="workbench-ribbon-toolbar" :class="{ 'is-grid': variant === 'grid' }" role="tabpanel" :aria-labelledby="`${id}-tab-${activeId}`">
      <RibbonGridToolbar v-if="variant === 'grid' && activeTab" :groups="activeTab.groups" :title="activeTab.title" :direction="direction" @execute="execute" />
      <template v-else>
        <div class="ribbon-visible-groups" role="toolbar" :aria-label="activeTab?.title">
          <div v-for="group in visibleGroups" :key="group.id" class="workbench-ribbon-group" role="group" :aria-label="group.title">
            <template v-for="item in group.items" :key="item.id">
              <div v-if="$slots[`item-${item.id}`]" class="ribbon-custom-item" :data-ribbon-item="item.id"><slot :name="`item-${item.id}`" :item="item" :execute="(commandId: string, params?: unknown) => executeSlot(item, commandId, params)" /></div>
              <RibbonItemControl v-else :item="item" @execute="execute" />
            </template>
          </div>
        </div>
        <span v-if="overflowGroups.length" ref="moreButton" class="ribbon-more-anchor" :class="{ 'has-visible-groups': visibleGroups.length > 0 }">
          <Popover v-model:open="moreOpen" trigger="click" placement="bottomRight" :align="{ overflow: { shiftX: true, adjustX: true, adjustY: true } }" :destroy-on-hidden="true">
            <Tooltip :title="translate('common.more', '更多功能')"><Button type="text" size="small" class="ribbon-more-button" :aria-label="translate('common.more', '更多功能')" aria-haspopup="menu" :aria-expanded="moreOpen" @mousedown.prevent @keydown.down.prevent="openMore" @keydown.esc="closeMore"><EllipsisVertical :size="CLASSIC_RIBBON.iconSize" /></Button></Tooltip>
            <template #content>
              <div ref="morePanel" class="ribbon-overflow-menu" role="menu" :aria-label="translate('common.more', '更多功能')" @keydown="onMoreKeydown">
                <div v-for="group in overflowGroups" :key="group.id" class="ribbon-overflow-group" role="group" :aria-label="group.title">
                  <div v-if="group.title" class="ribbon-overflow-title">{{ group.title }}</div>
                  <template v-for="item in group.items" :key="item.id">
                    <div v-if="$slots[`item-${item.id}`]" class="ribbon-custom-item"><slot :name="`item-${item.id}`" :item="item" :overflow="true" :execute="(commandId: string, params?: unknown) => executeSlot(item, commandId, params)" /></div>
                    <RibbonItemControl v-else :item="item" overflow @execute="execute" />
                  </template>
                </div>
              </div>
            </template>
          </Popover>
        </span>
      </template>
    </div>
    <div v-if="variant === 'classic'" ref="measurer" class="ribbon-measurer" aria-hidden="true" inert>
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
.workbench-ribbon { position: relative; flex: 0 0 auto; min-width: 0; background: #fff; color: #334155; border-bottom: 1px solid #e2e8f0; }
.workbench-ribbon-tabs { display: flex; height: var(--ribbon-tab-height); min-width: 0; gap: 4px; align-items: center; justify-content: safe center; padding: 0 var(--ribbon-horizontal-padding); overflow-x: auto; scrollbar-width: none; border-radius: 6px; background: #f8fafc; box-sizing: border-box; }
.workbench-ribbon-tabs::-webkit-scrollbar { display: none; }
.workbench-ribbon-tabs.is-grid { align-items: flex-end; justify-content: flex-start; gap: 0; padding-inline: 8px; border-radius: 0; }
.workbench-ribbon-tab { position: relative; flex: 0 0 auto; height: var(--ribbon-tab-button-height); min-width: 0; padding: 4px 8px; border-radius: 2px; color: #475569; font-size: 14px; line-height: 20px; }
.workbench-ribbon-tab.is-selected { color: #315bc2; background: #e7edfa; font-weight: 600; box-shadow: 0 1px 2px #1d2c4914; }
.workbench-ribbon-tab.is-contextual { color: #7450af; }
.workbench-ribbon-tabs.is-grid .workbench-ribbon-tab.is-grid-tab { border: 0 !important; border-radius: 0; background: transparent; color: #475569; font-weight: 600; }
.workbench-ribbon-tabs.is-grid .workbench-ribbon-tab.is-grid-tab::after { position: absolute; inset-inline: 0; bottom: 0; height: 3px; background: transparent; content: ''; }
.workbench-ribbon-tabs.is-grid .workbench-ribbon-tab.is-grid-tab.is-selected { background: transparent; color: #315bc2; box-shadow: none; }
.workbench-ribbon-tabs.is-grid .workbench-ribbon-tab.is-grid-tab.is-selected::after { background: #315bc2; }
.workbench-ribbon-tab:focus-visible { outline: 2px solid #5274d7; outline-offset: -3px; }
.ribbon-tab-extra { display: flex; flex: 0 0 auto; align-items: center; margin-inline-start: auto; }
.ribbon-tab-extra:empty { display: none; }
.workbench-ribbon-toolbar { display: flex; min-width: 0; height: var(--ribbon-toolbar-height); align-items: center; justify-content: center; padding: 0 var(--ribbon-horizontal-padding); box-sizing: border-box; }
.workbench-ribbon-toolbar.is-grid { display: block; padding: 0; }
.ribbon-visible-groups { display: flex; min-width: 0; align-items: center; }
.workbench-ribbon-group { display: flex; flex: 0 0 auto; gap: var(--ribbon-item-gap); align-items: center; padding: 0 var(--ribbon-group-padding); box-sizing: border-box; }
.workbench-ribbon-group + .workbench-ribbon-group { border-inline-start: var(--ribbon-separator-width) solid #e2e8f0; }
.ribbon-more-anchor { display: flex; flex: 0 0 auto; padding-inline-start: var(--ribbon-group-padding); }
.ribbon-more-anchor.has-visible-groups { border-inline-start: var(--ribbon-separator-width) solid #e2e8f0; }
.ribbon-more-button { display: inline-flex; height: var(--ribbon-button-height); width: var(--ribbon-button-height); min-width: var(--ribbon-button-height); align-items: center; justify-content: center; padding: 0; }
.ribbon-custom-item { display: inline-flex; align-items: center; min-height: 28px; flex: 0 0 auto; }
.ribbon-measurer { position: absolute; inset-inline-start: 0; top: 0; display: flex; width: 0; height: 0; overflow: hidden; visibility: hidden; pointer-events: none; contain: layout; }
.ribbon-measure-item { display: inline-flex; flex: 0 0 auto; }
.ribbon-overflow-menu { display: flex; flex-direction: column; width: max-content; min-width: min(210px, calc(100vw - 48px)); max-width: calc(100vw - 48px); max-height: min(520px, 70vh); overflow: auto; }
.ribbon-overflow-group { display: flex; flex-direction: column; gap: 3px; padding: 4px 0; }
.ribbon-overflow-group + .ribbon-overflow-group { border-top: 1px solid #e2e8f0; }
.ribbon-overflow-title { padding: 3px 6px; font-size: 11px; color: #64748b; }
</style>
