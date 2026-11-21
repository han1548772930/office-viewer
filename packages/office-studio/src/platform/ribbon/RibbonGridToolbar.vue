<script setup lang="ts">
import { computed, inject, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import { Button, Popover, Tooltip } from 'antdv-next'
import { Ellipsis, ChevronDown } from 'lucide-vue-next'
import RibbonItemControl from './RibbonItem.vue'
import { fitRibbonGridGroups, GRID_RIBBON, groupRibbonGridOverflowItems, placeRibbonGridItems, type RibbonGroup } from './ribbon-schema'
import { OfficeTranslatorKey } from '../../i18n/office-i18n'

const props = withDefaults(defineProps<{ groups: RibbonGroup[]; title: string; direction?: 'ltr' | 'rtl' }>(), { direction: 'ltr' })
const emit = defineEmits<{ execute: [commandId: string, params?: unknown] }>()
const translate = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)

const container = ref<HTMLElement>()
const moreAnchor = ref<HTMLElement>()
const overflowPanel = ref<HTMLElement>()
const groupElements = new Map<string, HTMLElement>()
const visibleGroupCount = ref(props.groups.length)
const moreOpen = ref(false)
const placements = computed(() => props.groups.map((group) => ({ group, items: placeRibbonGridItems(group.items) })))
const overflowGroups = computed(() => placements.value.slice(visibleGroupCount.value))
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
  return props.groups.map((group) => groupElements.get(group.id)?.getBoundingClientRect().width ?? 0)
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
  return Math.max(0, container.value.getBoundingClientRect().width - horizontalInsets)
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
  moreOpen.value = true
  await nextTick()
  overflowPanel.value?.querySelector<HTMLElement>('button:not(:disabled), [role="combobox"]')?.focus()
}

function closeOverflow(event?: KeyboardEvent) {
  if (event) {
    event.preventDefault()
    event.stopPropagation()
  }
  moreOpen.value = false
  moreAnchor.value?.querySelector<HTMLButtonElement>('button')?.focus()
}

function onOverflowKeydown(event: KeyboardEvent) {
  if (event.key === 'Escape') {
    closeOverflow(event)
    return
  }
  if (event.target instanceof HTMLElement && event.target.closest('[role="combobox"]')) return
  const controls = Array.from(overflowPanel.value?.querySelectorAll<HTMLElement>('button:not(:disabled), [role="combobox"]') ?? [])
  if (!controls.length) return
  const index = controls.findIndex((control) => control === event.target || control.contains(event.target as Node))
  let next: number
  if (event.key === 'ArrowDown') next = (index + 1) % controls.length
  else if (event.key === 'ArrowUp') next = (index - 1 + controls.length) % controls.length
  else if (event.key === 'Home') next = 0
  else if (event.key === 'End') next = controls.length - 1
  else return
  event.preventDefault()
  controls[next]?.focus()
}

watch(() => props.groups, async () => {
  await nextTick()
  observeGroups()
}, { flush: 'post' })
watch(() => overflowGroups.value.length, (count) => {
  if (!count) moreOpen.value = false
})

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
  <div ref="container" class="ribbon-grid-toolbar" role="toolbar" :aria-label="title" :dir="direction" :style="gridStyle" data-testid="ribbon-grid-toolbar">
    <div class="ribbon-grid-groups">
      <div
        v-for="({ group, items }, index) in placements"
        :key="group.id"
        :ref="element => setGroupElement(group.id, element as Element | null)"
        class="ribbon-grid-group"
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
            <RibbonItemControl
              :item="placement.item"
              grid
              :large="placement.rowSpan > 1"
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
      <Popover v-model:open="moreOpen" trigger="click" placement="bottomRight" :align="{ overflow: { shiftX: true, adjustX: true, adjustY: true } }" :destroy-on-hidden="true">
        <Tooltip :title="translate('common.more', '更多功能')">
          <Button
            type="text"
            class="ribbon-grid-more-button"
            :aria-label="translate('common.more', '更多功能')"
            aria-haspopup="menu"
            :aria-expanded="moreOpen"
            @mousedown.prevent
            @keydown.down.prevent="openOverflow"
            @keydown.esc="closeOverflow"
          >
            <Ellipsis :size="20" />
            <span>{{ translate('common.more', '更多') }}</span>
            <ChevronDown :size="12" />
          </Button>
        </Tooltip>
        <template #content>
          <div ref="overflowPanel" class="ribbon-grid-overflow" role="menu" :aria-label="translate('common.more', '更多功能')" @keydown="onOverflowKeydown">
            <section v-for="{ group } in overflowGroups" :key="group.id" class="ribbon-grid-overflow-group" role="group" :aria-label="group.title">
              <h3 v-if="group.title">{{ group.title }}</h3>
              <div v-for="(row, rowIndex) in groupRibbonGridOverflowItems(group.items)" :key="`${group.id}-${rowIndex}`" class="ribbon-grid-overflow-row">
                <RibbonItemControl
                  v-for="item in row"
                  :key="item.id"
                  :item="item"
                  grid
                  overflow
                  :show-label="item.gridLayout?.showLabel"
                  :icon-size="item.gridLayout?.iconSize"
                  full-width
                  @execute="(commandId, params) => emit('execute', commandId, params)"
                />
              </div>
            </section>
          </div>
        </template>
      </Popover>
    </span>
  </div>
</template>

<style scoped>
.ribbon-grid-toolbar { position: relative; display: flex; box-sizing: border-box; width: 100%; min-width: 0; height: var(--ribbon-grid-toolbar-height); align-items: stretch; overflow: hidden; padding: 0 var(--ribbon-grid-group-padding); background: #f8fafc; color: #334155; }
.ribbon-grid-groups { display: flex; min-width: 0; overflow: hidden; }
.ribbon-grid-group { position: relative; display: flex; box-sizing: border-box; width: max-content; flex: 0 0 auto; padding-inline: var(--ribbon-grid-group-padding); }
.ribbon-grid-group.is-hidden { position: absolute; visibility: hidden; pointer-events: none; }
.ribbon-grid-cells { display: grid; min-height: 0; flex: 1; grid-auto-flow: column; grid-template-rows: repeat(2, var(--ribbon-grid-row-height)); align-content: center; column-gap: var(--ribbon-grid-item-gap); row-gap: var(--ribbon-grid-row-gap); padding: var(--ribbon-grid-vertical-padding) 0; }
.ribbon-grid-cell { display: flex; min-width: var(--ribbon-grid-button-size); align-items: stretch; }
.ribbon-grid-cell.is-large { min-width: var(--ribbon-grid-large-button-size); }
.ribbon-grid-cell.is-full-width > :deep(.workbench-ribbon-item) { width: 100%; }
.ribbon-grid-divider { position: absolute; inset-block: 14px 6px; inset-inline-end: 0; border-inline-end: 1px solid #cbd5e1; opacity: .6; pointer-events: none; }
.ribbon-grid-more-anchor { display: flex; box-sizing: border-box; width: var(--ribbon-grid-more-width); flex: 0 0 var(--ribbon-grid-more-width); align-items: center; padding-block: 8px; padding-inline-start: var(--ribbon-grid-group-padding); }
.ribbon-grid-more-anchor.is-hidden { position: absolute; visibility: hidden; pointer-events: none; }
.ribbon-grid-more-button { display: inline-flex; box-sizing: border-box; width: 56px; min-width: 56px; height: 56px; flex-direction: column; align-items: center; justify-content: center; gap: 2px; padding: 2px; font-size: 11px; line-height: 14px; }
.ribbon-grid-more-button:focus-visible { outline: 2px solid #5274d7; outline-offset: -2px; }
.ribbon-grid-overflow { display: flex; box-sizing: border-box; width: min(360px, calc(100vw - 24px)); min-width: min(240px, calc(100vw - 24px)); max-width: calc(100vw - 24px); max-height: min(520px, 70vh); flex-direction: column; gap: 4px; overflow: auto; padding: 4px 8px; }
.ribbon-grid-overflow-group { display: flex; flex-direction: column; gap: 4px; padding-block: 4px; }
.ribbon-grid-overflow-group + .ribbon-grid-overflow-group { border-top: 1px solid #e2e8f0; }
.ribbon-grid-overflow-group h3 { margin: 0; padding: 2px 4px; color: #64748b; font-size: 11px; font-weight: 500; }
.ribbon-grid-overflow-row { display: flex; flex-wrap: wrap; align-items: center; gap: 4px; }
</style>
