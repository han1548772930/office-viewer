<script setup lang="ts">
import { computed, inject, nextTick, onBeforeUnmount, ref, useId, watch } from 'vue'
import { Button, Menu, Popover, Select, Tooltip } from 'antdv-next'
import type { MenuInfo } from 'antdv-next'
import { ChevronDown } from 'lucide-vue-next'
import { CLASSIC_RIBBON, GRID_RIBBON, ribbonInvocation, visibleRibbonItems, type RibbonItem } from './ribbon-schema'
import { OfficeTranslatorKey } from '../../i18n/office-i18n'
import { provideRibbonPopupScope, RibbonPopupScopeKey } from './ribbon-popup-scope'
import { useRibbonTheme } from './ribbon-theme'
import { createRibbonMenuModel } from './ribbon-command-menu'

const props = withDefaults(defineProps<{ item: RibbonItem; overflow?: boolean; measurement?: boolean; grid?: boolean; large?: boolean; alignStart?: boolean; showLabel?: boolean; iconSize?: number; fullWidth?: boolean }>(), { overflow: false, measurement: false, grid: false, large: false, alignStart: false, fullWidth: false })
const emit = defineEmits<{ execute: [commandId: string, params?: unknown] }>()
const translate = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const localOpen = ref(false)
const root = ref<HTMLElement>()
const panel = ref<HTMLElement>()
const menuRef = ref<InstanceType<typeof Menu> | null>(null)
const popupId = useId()
const popupScope = inject(RibbonPopupScopeKey, undefined)
const childPopupScope = provideRibbonPopupScope(popupScope?.direction)
const direction = childPopupScope.direction
const themeStyle = useRibbonTheme()
const trigger = ref<HTMLElement>()
const tooltipOpen = ref(false)
const open = computed({
  get: () => popupScope ? popupScope.activeId.value === popupId : localOpen.value,
  set: (value: boolean) => {
    if (value && (props.item.disabled || props.item.hidden || props.measurement)) return
    if (popupScope) {
      if (value) popupScope.activeId.value = popupId
      else if (popupScope.activeId.value === popupId) popupScope.close()
    } else localOpen.value = value
    if (!value) childPopupScope.close()
    tooltipOpen.value = false
  },
})
const classicStyle = { '--ribbon-button-height': `${CLASSIC_RIBBON.buttonHeight}px`, '--ribbon-selector-height': `${CLASSIC_RIBBON.selectorHeight}px` }
const children = computed(() => visibleRibbonItems(props.item.children ?? []))
const menuModel = computed(() => createRibbonMenuModel(children.value))
const showLabel = computed(() => {
  if (props.overflow || (props.grid && props.large)) return true
  const requested = props.showLabel ?? props.item.showLabel
  return requested ?? !props.item.icon
})
const iconSize = computed(() => props.iconSize ?? (props.grid && props.large ? GRID_RIBBON.largeIconSize : CLASSIC_RIBBON.iconSize))
const description = computed(() => props.item.disabledReason && props.item.disabled ? `${props.item.title}：${props.item.disabledReason}` : props.item.tooltip ?? props.item.title)
const isDropdown = computed(() => (props.item.type === 'menu' || props.item.type === 'split') && (children.value.length > 0 || Boolean(props.item.popupComponent)))
const style = computed(() => props.item.width && !props.overflow && !props.grid ? { width: `${props.item.width}px` } : undefined)
const itemStyle = computed(() => [classicStyle, style.value, themeStyle.value, { '--ribbon-icon-size': `${iconSize.value}px`, '--ribbon-grid-button-size': `${GRID_RIBBON.buttonSize}px` }])

function execute(value?: unknown) {
  if (props.measurement) return
  const invocation = ribbonInvocation(props.item, typeof value === 'string' || typeof value === 'number' ? value : undefined)
  if (invocation) {
    open.value = false
    emit('execute', invocation.commandId, invocation.params)
  }
}

function forward(commandId: string, params?: unknown) {
  if (!commandId || props.item.disabled || props.item.hidden || props.measurement) return
  open.value = false
  emit('execute', commandId, params)
}

function executeMenuItem({ key }: MenuInfo) {
  const action = menuModel.value.actions.get(String(key))
  if (!action || action.item.disabled || action.item.hidden) return
  const invocation = ribbonInvocation(action.item, action.value)
  if (!invocation) return
  open.value = false
  emit('execute', invocation.commandId, invocation.params)
}

async function openMenu(event: KeyboardEvent, last = false) {
  if (props.item.disabled || props.measurement || !isDropdown.value) return
  if (event.currentTarget instanceof HTMLElement) trigger.value = event.currentTarget
  open.value = true
  await nextTick()
  const controls = menuRef.value?.$el?.querySelectorAll<HTMLElement>('[role="menuitem"]:not([aria-disabled="true"])')
  const target = last ? controls?.[controls.length - 1] : controls?.[0]
  if (target) target.focus()
  else if (menuRef.value) menuRef.value.focus()
  else panel.value?.querySelector<HTMLElement>('button:not(:disabled), input:not(:disabled), [tabindex="0"]')?.focus()
}

function closeMenu() {
  open.value = false
  const target = trigger.value ?? root.value?.querySelector<HTMLButtonElement>('[aria-haspopup="menu"]')
  if (target?.isConnected) target.focus({ preventScroll: true })
}

function onPanelKeydown(event: KeyboardEvent) {
  if (event.defaultPrevented) return
  if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); closeMenu(); return }
  if (event.key === 'Tab') { open.value = false; return }
  if (event.target instanceof HTMLElement && event.target.matches('input, textarea, [contenteditable="true"]')) return
  if (event.target instanceof HTMLElement && event.target.closest('[role="combobox"]')) return
}

watch(() => props.item.disabled || props.item.hidden || !isDropdown.value, (unavailable) => {
  if (unavailable && props.item.type !== 'select') open.value = false
})
watch(() => props.item.disabled || props.item.hidden, (unavailable) => { if (unavailable) open.value = false })
watch(open, (value) => { if (!value) { childPopupScope.close(); tooltipOpen.value = false } })
onBeforeUnmount(() => { open.value = false })
</script>

<template>
  <span ref="root" class="workbench-ribbon-item" :class="{ 'is-overflow': overflow, 'is-active': item.activated, 'has-label': showLabel, 'has-icon': Boolean(item.icon), 'is-selector': isDropdown || item.type === 'select', 'is-grid': grid, 'is-grid-large': grid && large, 'is-grid-full-width': grid && fullWidth, 'is-grid-align-start': grid && !large && alignStart && Boolean(item.icon) && !showLabel }" :style="itemStyle" :data-ribbon-item="item.id">
    <Tooltip :title="measurement ? undefined : description" :open="!open && tooltipOpen" :mouse-enter-delay="0.6" @open-change="value => { tooltipOpen = !open && value }">
      <span v-if="item.type === 'select'" class="ribbon-select-wrap">
        <label :for="`${popupId}-select`" :class="overflow ? 'ribbon-select-label' : grid && (large || showLabel) ? 'ribbon-grid-select-title' : 'ribbon-sr-only'">{{ item.title }}</label>
        <Select
          :id="`${popupId}-select`" :value="item.value == null ? undefined : item.value" :options="item.options ?? []" :disabled="item.disabled || measurement"
          :placeholder="item.placeholder ?? (item.value === null ? translate('common.mixed', '混合') : item.title)" :aria-label="item.title" size="small"
          class="ribbon-select" :style="{ width: overflow ? '150px' : grid && fullWidth ? '100%' : `${item.width ?? 96}px`, height: `${grid && !overflow ? GRID_RIBBON.rowHeight : CLASSIC_RIBBON.selectorHeight}px`, alignItems: 'center', flex: overflow ? '1 1 0' : undefined, minWidth: overflow || (grid && fullWidth) ? '0' : undefined, maxWidth: overflow ? '150px' : undefined }" :popup-match-select-width="false"
          :open="open" @open-change="value => { open = value }" @change="execute"
        />
      </span>
      <span v-else-if="isDropdown" class="ribbon-dropdown" :class="{ 'ribbon-split': item.type === 'split' }">
        <Button v-if="item.type === 'split'" type="text" size="small" class="ribbon-button" :class="{ 'is-active': item.activated }" :disabled="item.disabled || measurement" :aria-label="item.title" :aria-pressed="item.activated === undefined ? undefined : item.activated" @mousedown.prevent @click="execute()">
          <component :is="item.icon" v-if="item.icon" :size="iconSize" /><span v-if="showLabel">{{ item.title }}</span>
        </Button>
        <Popover v-model:open="open" trigger="click" :placement="direction === 'rtl' ? 'bottomRight' : 'bottomLeft'" :align="{ overflow: { shiftX: true, adjustX: true, adjustY: true } }" :destroy-on-hidden="!item.popupComponent">
          <Button type="text" size="small" class="ribbon-button" :class="{ 'ribbon-split-arrow': item.type === 'split', 'is-active': item.activated }" :disabled="item.disabled || measurement" :aria-label="item.type === 'split' ? `${item.title}选项` : item.title" :aria-pressed="item.type === 'menu' ? item.activated : undefined" :aria-haspopup="item.popupComponent ? 'dialog' : 'menu'" :aria-expanded="open" :aria-controls="popupId" @mousedown.prevent @keydown.down.prevent.stop="openMenu" @keydown.up.prevent.stop="openMenu($event, true)" @keydown.esc.prevent.stop="closeMenu">
            <component :is="item.icon" v-if="item.type === 'menu' && item.icon" :size="iconSize" /><span v-if="item.type === 'menu' && showLabel">{{ item.title }}</span><ChevronDown :size="12" />
          </Button>
          <template #content>
            <div v-if="item.popupComponent" :id="popupId" ref="panel" class="ribbon-component-popup" :dir="direction" :style="themeStyle" :aria-label="item.title" @keydown="onPanelKeydown">
              <component :is="item.popupComponent" v-bind="item.popupProps" @insert="(value: string) => forward(item.popupCommandId ?? item.commandId ?? '', value)" />
            </div>
            <Menu v-else :id="popupId" ref="menuRef" class="ribbon-command-menu" :dir="direction" :style="themeStyle" :aria-label="item.title" :items="menuModel.items" :selectable="false" :trigger-sub-menu-action="'click'" @click="executeMenuItem" @keydown.esc.capture.stop.prevent="closeMenu" />
          </template>
        </Popover>
      </span>
      <Button v-else type="text" size="small" class="ribbon-button" :class="{ 'is-active': item.activated }" :disabled="item.disabled || measurement" :aria-label="item.title" :aria-pressed="item.activated === undefined ? undefined : item.activated" @mousedown.prevent @click="execute()">
        <span v-if="item.swatch" class="ribbon-color-swatch" :style="{ backgroundColor: item.swatch }" aria-hidden="true" />
        <component :is="item.icon" v-else-if="item.icon" :size="iconSize" /><span v-if="showLabel">{{ item.title }}</span><span v-if="overflow && item.shortcut" class="ribbon-shortcut">{{ item.shortcut }}</span>
      </Button>
    </Tooltip>
  </span>
</template>

<style scoped>
.workbench-ribbon-item { display: inline-flex; flex: 0 0 auto; min-width: var(--ribbon-button-height); height: var(--ribbon-selector-height); align-items: center; vertical-align: middle; }
.workbench-ribbon-item.is-grid { box-sizing: border-box; width: 100%; min-width: var(--ribbon-grid-button-size); height: 100%; }
.ribbon-button { display: inline-flex; min-width: var(--ribbon-button-height); height: var(--ribbon-button-height); padding: 0; gap: 4px; align-items: center; justify-content: center; font-size: 14px; white-space: nowrap; color: var(--ribbon-text, #334155); border-radius: 4px; }
.workbench-ribbon-item.is-grid :deep(.ant-tooltip-trigger) { display: flex; width: 100%; height: 100%; }
.workbench-ribbon-item.is-grid .ribbon-button { box-sizing: border-box; width: 100%; min-width: var(--ribbon-grid-button-size); height: 100%; padding-inline: 4px; font-size: 12px; }
.workbench-ribbon-item.is-grid .ribbon-button :deep(svg) { width: var(--ribbon-icon-size); height: var(--ribbon-icon-size); }
.workbench-ribbon-item.is-grid.is-grid-large .ribbon-button { flex-direction: column; gap: 4px; padding: 4px 6px; white-space: normal; }
.workbench-ribbon-item.is-grid.is-grid-large .ribbon-button :deep(svg) { flex: 0 0 auto; }
.has-label .ribbon-button { padding-inline: 6px; }
.workbench-ribbon-item:not(.has-icon) .ribbon-button { padding-inline: 8px; }
.is-selector .ribbon-button { height: var(--ribbon-selector-height); }
.ribbon-button :deep(svg) { flex-shrink: 0; }
.ribbon-color-swatch { display: inline-block; width: 14px; height: 14px; flex: 0 0 14px; border: 1px solid #aeb7c2; border-radius: 2px; }
.ribbon-button.is-active { background: var(--ribbon-active-background, #e7edfa); color: var(--ribbon-primary, #315bc2); }
.ribbon-button:focus-visible { outline: 2px solid var(--ribbon-primary, #5274d7); outline-offset: -2px; }
.ribbon-select-wrap { display: flex; align-items: center; gap: 12px; }
.workbench-ribbon-item.is-grid .ribbon-select-wrap { box-sizing: border-box; width: 100%; height: 100%; flex-direction: column; align-items: stretch; justify-content: center; gap: 2px; }
.ribbon-grid-select-title { display: block; overflow: hidden; font-size: 11px; line-height: 14px; text-align: center; text-overflow: ellipsis; white-space: nowrap; }
.workbench-ribbon-item .ribbon-select { min-width: 52px; height: var(--ribbon-selector-height); font-size: 14px; }
.workbench-ribbon-item.is-grid .ribbon-select, .workbench-ribbon-item.is-grid .ribbon-select :deep(input) { font-size: 12px; }
.workbench-ribbon-item.is-grid-align-start .ribbon-button { justify-content: flex-start; }
.ribbon-sr-only { position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px; overflow: hidden; clip-path: inset(50%); white-space: nowrap; border: 0; }
.ribbon-dropdown { display: flex; align-items: center; }
.ribbon-split .ribbon-button:not(.ribbon-split-arrow) { min-width: var(--ribbon-selector-height); padding-inline: 4px; }
.ribbon-dropdown .ribbon-split-arrow { min-width: 24px; width: 24px; padding: 0; border-inline-start: 1px solid var(--ribbon-border, #e2e8f0); border-start-start-radius: 0; border-end-start-radius: 0; }
.workbench-ribbon-item.is-grid .ribbon-dropdown { width: 100%; height: 100%; }
.workbench-ribbon-item.is-grid .ribbon-split .ribbon-button:not(.ribbon-split-arrow) { min-width: 20px; flex: 1 1 auto; }
.workbench-ribbon-item.is-grid .ribbon-split-arrow { width: 16px; min-width: 16px; flex: 0 0 16px; justify-content: flex-start; border: 0; }
.workbench-ribbon-item.is-grid.is-grid-large .ribbon-split { position: relative; }
.workbench-ribbon-item.is-grid.is-grid-large .ribbon-split .ribbon-button:not(.ribbon-split-arrow) { padding-bottom: 16px; }
.workbench-ribbon-item.is-grid.is-grid-large .ribbon-split-arrow { position: absolute; inset-inline: 0; bottom: 0; width: 100%; height: 16px; justify-content: center; padding: 0; }
.ribbon-command-menu { box-sizing: border-box; min-width: min(180px, calc(100vw - 48px)); max-width: calc(100vw - 48px); max-height: min(460px, 65vh); overflow: auto; color: var(--ribbon-text, #334155); }
.ribbon-command-menu :deep(.ant-menu-item), .ribbon-command-menu :deep(.ant-menu-submenu-title) { min-height: 30px; height: auto; line-height: 22px; padding-block: 4px; white-space: normal; }
.ribbon-command-menu :deep(.ribbon-menu-item-label) { display: flex; min-width: 0; align-items: center; gap: 12px; }
.ribbon-command-menu :deep(.ribbon-menu-item-text) { min-width: 0; overflow-wrap: anywhere; }
.ribbon-command-menu :deep(.ribbon-menu-item-shortcut) { flex: 0 0 auto; margin-inline-start: auto; color: var(--ribbon-text-secondary, #94a3b8); font-size: 11px; }
.ribbon-component-popup { box-sizing: border-box; max-width: calc(100vw - 24px); max-height: min(520px, 70vh); overflow: auto; }
.ribbon-button:not(:disabled):hover { background: var(--ribbon-hover-background, #f1f5f9); }
.ribbon-button:disabled { color: var(--ribbon-disabled, #94a3b8); background: transparent; cursor: not-allowed; }
.is-overflow { width: 100%; min-width: 0; height: auto; min-height: 30px; }
.workbench-ribbon-item.is-grid.is-overflow { height: auto; min-height: 30px; }
.is-overflow > :deep(span), .is-overflow .ribbon-dropdown, .is-overflow .ribbon-select-wrap { width: 100%; }
.is-overflow .ribbon-button { width: 100%; height: auto; min-height: 28px; justify-content: flex-start; white-space: normal; text-align: start; }
.workbench-ribbon-item.is-grid.is-overflow .ribbon-button { height: auto; min-height: 28px; }
.is-overflow .ribbon-select-wrap { justify-content: space-between; }
.is-overflow .ribbon-select { flex: 1; min-width: 0; max-width: 150px; }
.is-overflow .ribbon-select-label { flex: 0 0 auto; font-size: 13px; white-space: nowrap; }
.is-overflow .ribbon-split-arrow { width: 24px; flex: 0 0 24px; }
.ribbon-shortcut { margin-inline-start: auto; padding-inline-start: 16px; color: var(--ribbon-text-secondary, #94a3b8); font-size: 11px; }
</style>
