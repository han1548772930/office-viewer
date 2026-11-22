<script setup lang="ts">
import { computed, inject, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import { Button, Popover, Tooltip } from 'antdv-next'
import { ChevronDown, MoreHorizontal } from 'lucide-vue-next'
import RibbonItemControl from './RibbonItem.vue'
import { fitRibbonGridGroups, GRID_RIBBON, placeRibbonGridItems, type RibbonGroup, type RibbonItem } from './ribbon-schema'
import { OfficeTranslatorKey } from '../../i18n/office-i18n'
import { RibbonPopupScopeKey } from './ribbon-popup-scope'
import { useRibbonTheme } from './ribbon-theme'
import RibbonOverflowPanel from './RibbonOverflowPanel.vue'

const props = withDefaults(defineProps<{ groups: RibbonGroup[]; title: string; direction?: 'ltr' | 'rtl'; alignStart?: boolean }>(), { direction: 'ltr', alignStart: false })
const emit = defineEmits<{ execute: [commandId: string, params?: unknown] }>()
const translate = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const popupScope = inject(RibbonPopupScopeKey, undefined)
const themeStyle = useRibbonTheme()

const container = ref<HTMLElement>()
const moreAnchor = ref<HTMLElement>()
const overflowMenuRef = ref<{ focus: () => void } | null>(null)
const groupElements = new Map<string, HTMLElement>()
const visibleGroupCount = ref(props.groups.length)
const moreOpen = ref(false)
const placements = computed(() => props.groups.map((group) => ({ group, items: placeRibbonGridItems(group.items) })))
const overflowGroups = computed(() => placements.value.slice(visibleGroupCount.value))
const hiddenGroups = computed(() => overflowGroups.value.map(({ group }) => group))
const gridStyle = {
  '--ribbon-grid-toolbar-height': `${GRID_RIBBON.toolbarHeight}px`,
  '--ribbon-grid-row-height': `${GRID_RIBBON.rowHeight}px`,
  '--ribbon-grid-row-gap': `${GRID_RIBBON.rowGap}px`,
  '--ribbon-grid-vertical-padding': `${GRID_RIBBON.verticalPadding}px`,
  '--ribbon-grid-group-padding': `${GRID_RIBBON.groupPadding}px`,
  '--ribbon-grid-compact-group-padding': `${GRID_RIBBON.compactGroupPadding}px`,
  '--ribbon-grid-item-gap': `${GRID_RIBBON.itemGap}px`,
  '--ribbon-grid-compact-item-gap': `${GRID_RIBBON.compactItemGap}px`,
  '--ribbon-grid-more-width': `${GRID_RIBBON.moreWidth}px`,
  '--ribbon-grid-button-size': `${GRID_RIBBON.buttonSize}px`,
  '--ribbon-grid-large-button-size': `${GRID_RIBBON.largeButtonSize}px`,
}
let observer: ResizeObserver | undefined
let disposed = false
let measuring = false

function setGroupElement(id: string, element: Element | null) {
  if (element instanceof HTMLElement) groupElements.set(id, element)
  else groupElements.delete(id)
}

function groupWidths(): number[] {
  return props.groups.map((group) => {
    const element = groupElements.get(group.id)
    return element ? Number.parseFloat(getComputedStyle(element).width) || 0 : 0
  })
}

function setDensity(compact: boolean) {
  if (!container.value) return
  container.value.style.setProperty('--ribbon-grid-density', compact ? '1' : '0')
  container.value.style.setProperty('--ribbon-grid-group-padding', `${compact ? GRID_RIBBON.compactGroupPadding : GRID_RIBBON.groupPadding}px`)
  container.value.style.setProperty('--ribbon-grid-item-gap', `${compact ? GRID_RIBBON.compactItemGap : GRID_RIBBON.itemGap}px`)
}

function availableWidth(): number {
  if (!container.value) return 0
  const style = getComputedStyle(container.value)
  const horizontalInsets = [style.paddingInlineStart, style.paddingInlineEnd, style.borderInlineStartWidth, style.borderInlineEndWidth]
    .reduce((sum, value) => sum + (Number.parseFloat(value) || 0), 0)
  return Math.max(0, (Number.parseFloat(style.width) || 0) - horizontalInsets)
}

function measure() {
  if (!container.value || disposed || measuring) return
  measuring = true
  const width = availableWidth()
  setDensity(false)
  const regularWidths = groupWidths()
  setDensity(true)
  const compactWidths = groupWidths()
  const fit = fitRibbonGridGroups(props.groups.map((_, index) => ({
    regular: regularWidths[index] ?? 0,
    compact: compactWidths[index] ?? 0,
  })), width)
  setDensity(fit.compact)
  visibleGroupCount.value = fit.visibleCount
  measuring = false
}

function observeGroups() {
  if (!observer) return
  observer.disconnect()
  if (container.value) observer.observe(container.value)
  groupElements.forEach((element) => observer?.observe(element))
  if (moreAnchor.value) observer.observe(moreAnchor.value)
  measure()
}

async function openOverflow() {
  if (!overflowGroups.value.length) return
  popupScope?.close()
  moreOpen.value = true
  await nextTick()
  overflowMenuRef.value?.focus()
}

function closeOverflow(event?: KeyboardEvent) {
  if (event) {
    event.preventDefault()
    event.stopPropagation()
  }
  moreOpen.value = false
  moreAnchor.value?.querySelector<HTMLButtonElement>('button')?.focus()
}

function executeSlot(item: RibbonItem, commandId: string, params?: unknown) {
  if (!item.disabled && !item.hidden) emit('execute', commandId, params)
}

watch(() => props.groups, async () => {
  await nextTick()
  observeGroups()
}, { flush: 'post' })
watch(() => overflowGroups.value.length, (count) => {
  if (!count) moreOpen.value = false
})
watch(moreOpen, () => { popupScope?.close() })
watch(visibleGroupCount, () => { moreOpen.value = false; popupScope?.close() })

onMounted(() => {
  observer = new ResizeObserver(measure)
  observeGroups()
})

onBeforeUnmount(() => {
  disposed = true
  observer?.disconnect()
})
</script>

<template>
  <div ref="container" class="ribbon-grid-toolbar" role="toolbar" :aria-label="title" :dir="direction" :style="[gridStyle, themeStyle]" data-testid="ribbon-grid-toolbar">
    <div class="ribbon-grid-groups">
      <div
        v-for="({ group, items }, index) in placements"
        :key="group.id"
        :ref="element => setGroupElement(group.id, element as Element | null)"
        class="ribbon-grid-group"
        role="group"
        :aria-label="group.title"
        :class="{ 'is-hidden': index >= visibleGroupCount }"
        :data-group-id="group.id"
        :aria-hidden="index >= visibleGroupCount ? 'true' : undefined"
        :inert="index >= visibleGroupCount"
      >
        <div class="ribbon-grid-cells">
          <div
            v-for="placement in items"
            :key="placement.item.id"
            class="ribbon-grid-cell"
            :class="{ 'is-full-width': placement.width !== undefined, 'is-large': placement.rowSpan > 1 }"
            :style="{
              gridRow: `${placement.row} / span ${placement.rowSpan}`,
              gridColumn: `${placement.column} / span ${placement.columnSpan}`,
              width: placement.width === undefined
                ? undefined
                : placement.minWidth === undefined
                  ? `${placement.width}px`
                  : `calc(${placement.width}px - ${placement.width - placement.minWidth}px * var(--ribbon-grid-density, 0))`,
            }"
          >
            <div v-if="$slots[`item-${placement.item.id}`]" class="ribbon-grid-custom-item" :data-ribbon-item="placement.item.id">
              <slot :name="`item-${placement.item.id}`" :item="placement.item" :grid="true" :execute="(commandId: string, params?: unknown) => executeSlot(placement.item, commandId, params)" />
            </div>
            <RibbonItemControl
              v-else
              :item="placement.item"
              grid
              :large="placement.rowSpan > 1"
              :align-start="alignStart && placement.rowSpan === 1"
              :show-label="placement.showLabel"
              :icon-size="placement.iconSize"
              :full-width="placement.width !== undefined"
              @execute="(commandId, params) => emit('execute', commandId, params)"
            />
          </div>
        </div>
        <span class="ribbon-grid-divider" aria-hidden="true" />
      </div>
    </div>
    <span
      ref="moreAnchor"
      class="ribbon-grid-more-anchor"
      :class="{ 'is-hidden': overflowGroups.length === 0 }"
      :aria-hidden="overflowGroups.length === 0 ? 'true' : undefined"
      :inert="overflowGroups.length === 0"
    >
      <Popover v-model:open="moreOpen" trigger="click" :placement="direction === 'rtl' ? 'bottomLeft' : 'bottomRight'" :align="{ overflow: { shiftX: true, adjustX: true, adjustY: true } }" :destroy-on-hidden="true">
        <Tooltip :title="translate('common.more', '更多功能')">
          <Button
            type="text"
            class="ribbon-grid-more-button"
            :aria-label="translate('common.more', '更多功能')"
            aria-haspopup="dialog"
            :aria-expanded="moreOpen"
            @mousedown.prevent
            @keydown.down.prevent="openOverflow"
            @keydown.esc="closeOverflow"
          >
            <MoreHorizontal :size="20" />
            <span class="ribbon-grid-more-label">{{ translate('common.more', '更多') }}<ChevronDown :size="12" /></span>
          </Button>
        </Tooltip>
        <template #content>
          <RibbonOverflowPanel ref="overflowMenuRef" grid :groups="hiddenGroups" :direction="direction" :label="translate('common.more', '更多功能')" @execute="(id, params) => emit('execute', id, params)" @close="closeOverflow">
            <template v-for="(_, name) in $slots" #[name]="slotProps"><slot :name="name" v-bind="slotProps" /></template>
          </RibbonOverflowPanel>
        </template>
      </Popover>
    </span>
  </div>
</template>

<style scoped>
.ribbon-grid-toolbar { position: relative; display: flex; box-sizing: border-box; width: 100%; min-width: 0; height: var(--ribbon-grid-toolbar-height); align-items: stretch; overflow: hidden; border-bottom: 1px solid var(--ribbon-border, #e2e8f0); background: var(--ribbon-surface, #f9fafb); color: var(--ribbon-text, #334155); font-size: 12px; }
.ribbon-grid-groups { display: flex; min-width: 0; overflow: hidden; }
.ribbon-grid-group { position: relative; display: flex; box-sizing: border-box; width: max-content; flex: 0 0 auto; }
.ribbon-grid-group.is-hidden { position: absolute; visibility: hidden; pointer-events: none; }
.ribbon-grid-cells { display: grid; min-height: 0; flex: 1; grid-auto-flow: column; grid-template-rows: repeat(2, var(--ribbon-grid-row-height)); align-content: center; column-gap: var(--ribbon-grid-item-gap); row-gap: var(--ribbon-grid-row-gap); padding: var(--ribbon-grid-vertical-padding) var(--ribbon-grid-group-padding); }
.ribbon-grid-cell { display: flex; min-width: var(--ribbon-grid-button-size); align-items: stretch; }
.ribbon-grid-cell.is-large { min-width: var(--ribbon-grid-large-button-size); }
.ribbon-grid-cell > :deep(.workbench-ribbon-item) { width: 100%; }
.ribbon-grid-custom-item { display: flex; width: 100%; min-width: 0; height: 100%; align-items: center; }
.ribbon-grid-divider { flex: 0 0 auto; margin-block: 8px; border-inline-end: 1px solid var(--ribbon-border, #d1d5db); opacity: .5; pointer-events: none; }
.ribbon-grid-more-anchor { display: flex; box-sizing: border-box; width: var(--ribbon-grid-more-width); flex: 0 0 var(--ribbon-grid-more-width); align-items: stretch; padding: 8px; }
.ribbon-grid-more-anchor.is-hidden { position: absolute; visibility: hidden; pointer-events: none; }
.ribbon-grid-more-button { display: inline-flex; box-sizing: border-box; width: 56px; min-width: 56px; height: 100%; flex-direction: column; align-items: center; justify-content: center; gap: 4px; padding: 2px; font-size: 12px; line-height: 16px; }
.ribbon-grid-more-label { display: flex; align-items: center; gap: 4px; }
.ribbon-grid-more-button:focus-visible { outline: 2px solid var(--ribbon-primary, #5274d7); outline-offset: -2px; }
</style>
