<script setup lang="ts">
import { computed, nextTick, ref, useId } from 'vue'
import { Button, Popover, Select, Tooltip } from 'antdv-next'
import { ChevronDown } from 'lucide-vue-next'
import { ribbonInvocation, visibleRibbonItems, type RibbonItem } from './ribbon-schema'

const props = withDefaults(defineProps<{ item: RibbonItem; overflow?: boolean; measurement?: boolean }>(), { overflow: false, measurement: false })
const emit = defineEmits<{ execute: [commandId: string, params?: unknown] }>()
const open = ref(false)
const root = ref<HTMLElement>()
const panel = ref<HTMLElement>()
const popupId = useId()
const children = computed(() => visibleRibbonItems(props.item.children ?? []))
const showLabel = computed(() => props.overflow || props.item.showLabel || !props.item.icon)
const description = computed(() => props.item.disabledReason && props.item.disabled ? `${props.item.title}：${props.item.disabledReason}` : props.item.tooltip ?? props.item.title)
const isDropdown = computed(() => (props.item.type === 'menu' || props.item.type === 'split') && children.value.length > 0)
const style = computed(() => props.item.width && !props.overflow ? { width: `${props.item.width}px` } : undefined)

function execute(value?: unknown) {
  if (props.measurement) return
  const invocation = ribbonInvocation(props.item, typeof value === 'string' || typeof value === 'number' ? value : undefined)
  if (invocation) emit('execute', invocation.commandId, invocation.params)
}

function forward(commandId: string, params?: unknown) {
  if (props.item.disabled || props.item.hidden || props.measurement) return
  open.value = false
  emit('execute', commandId, params)
}

async function openMenu() {
  if (props.item.disabled || props.measurement || !isDropdown.value) return
  open.value = true
  await nextTick()
  panel.value?.querySelector<HTMLElement>('button:not(:disabled), input:not(:disabled), [tabindex="0"]')?.focus()
}

function closeMenu() {
  open.value = false
  root.value?.querySelector<HTMLButtonElement>('button:not(:disabled)')?.focus()
}

function onPanelKeydown(event: KeyboardEvent) {
  if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); closeMenu(); return }
  if (event.target instanceof HTMLElement && event.target.closest('[role="combobox"]')) return
  const controls = Array.from(panel.value?.querySelectorAll<HTMLElement>('button:not(:disabled), [role="combobox"]') ?? [])
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
</script>

<template>
  <span ref="root" class="workbench-ribbon-item" :class="{ 'is-overflow': overflow, 'is-active': item.activated }" :style="style" :data-ribbon-item="item.id">
    <Tooltip :title="measurement ? undefined : description" :mouse-enter-delay="0.6">
      <span v-if="item.type === 'select'" class="ribbon-select-wrap">
        <label :for="`${popupId}-select`" :class="overflow ? 'ribbon-select-label' : 'ribbon-sr-only'">{{ item.title }}</label>
        <Select
          :id="`${popupId}-select`" :value="item.value == null ? undefined : item.value" :options="item.options ?? []" :disabled="item.disabled || measurement"
          :placeholder="item.placeholder ?? (item.value === null ? '混合' : item.title)" :aria-label="item.title" size="small"
          class="ribbon-select" :style="{ width: overflow ? '150px' : `${item.width ?? 96}px` }" :popup-match-select-width="false"
          @change="execute"
        />
      </span>
      <span v-else-if="isDropdown" class="ribbon-dropdown" :class="{ 'ribbon-split': item.type === 'split' }">
        <Button v-if="item.type === 'split'" type="text" size="small" class="ribbon-button" :class="{ 'is-active': item.activated }" :disabled="item.disabled || measurement" :aria-label="item.title" :aria-pressed="item.activated === undefined ? undefined : item.activated" @mousedown.prevent @click="execute()">
          <component :is="item.icon" v-if="item.icon" :size="16" /><span v-if="showLabel">{{ item.title }}</span>
        </Button>
        <Popover v-model:open="open" trigger="click" placement="bottomLeft" :destroy-on-hidden="true">
          <Button type="text" size="small" class="ribbon-button" :class="{ 'ribbon-split-arrow': item.type === 'split', 'is-active': item.activated }" :disabled="item.disabled || measurement" :aria-label="item.type === 'split' ? `${item.title}选项` : item.title" aria-haspopup="menu" :aria-expanded="open" :aria-controls="popupId" @mousedown.prevent @keydown.down.prevent="openMenu" @keydown.esc.prevent="closeMenu">
            <component :is="item.icon" v-if="item.type === 'menu' && item.icon" :size="16" /><span v-if="item.type === 'menu' && showLabel">{{ item.title }}</span><ChevronDown :size="12" />
          </Button>
          <template #content>
            <div :id="popupId" ref="panel" class="ribbon-command-menu" role="menu" :aria-label="item.title" @keydown="onPanelKeydown">
              <RibbonItem v-for="child in children" :key="child.id" :item="child" overflow @execute="forward" />
            </div>
          </template>
        </Popover>
      </span>
      <Button v-else type="text" size="small" class="ribbon-button" :class="{ 'is-active': item.activated }" :disabled="item.disabled || measurement" :aria-label="item.title" :aria-pressed="item.activated === undefined ? undefined : item.activated" @mousedown.prevent @click="execute()">
        <component :is="item.icon" v-if="item.icon" :size="16" /><span v-if="showLabel">{{ item.title }}</span><span v-if="overflow && item.shortcut" class="ribbon-shortcut">{{ item.shortcut }}</span>
      </Button>
    </Tooltip>
  </span>
</template>

<style scoped>
.workbench-ribbon-item { display: inline-flex; flex: 0 0 auto; min-width: 28px; height: 28px; align-items: center; vertical-align: middle; }
.ribbon-button { display: inline-flex; min-width: 28px; height: 28px; padding: 0 6px; gap: 5px; align-items: center; justify-content: center; font-size: 13px; white-space: nowrap; color: #334155; border-radius: 4px; }
.ribbon-button.is-active { background: #e7edfa; color: #315bc2; }
.ribbon-button:focus-visible { outline: 2px solid #5274d7; outline-offset: -2px; }
.ribbon-select-wrap { display: flex; align-items: center; gap: 12px; }
.ribbon-select { min-width: 52px; font-size: 13px; }
.ribbon-sr-only { position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px; overflow: hidden; clip-path: inset(50%); white-space: nowrap; border: 0; }
.ribbon-dropdown { display: flex; align-items: center; }
.ribbon-split-arrow { min-width: 19px; padding: 0 3px; border-inline-start: 1px solid #e2e8f0; border-start-start-radius: 0; border-end-start-radius: 0; }
.ribbon-command-menu { display: flex; min-width: 180px; max-height: min(460px, 65vh); flex-direction: column; gap: 3px; overflow: auto; }
.is-overflow { width: 100%; height: auto; min-height: 30px; }
.is-overflow > :deep(span), .is-overflow .ribbon-dropdown, .is-overflow .ribbon-select-wrap { width: 100%; }
.is-overflow .ribbon-button { width: 100%; justify-content: flex-start; }
.is-overflow .ribbon-select-wrap { justify-content: space-between; }
.is-overflow .ribbon-select-label { font-size: 13px; white-space: nowrap; }
.is-overflow .ribbon-split-arrow { width: 24px; flex: 0 0 24px; }
.ribbon-shortcut { margin-inline-start: auto; padding-inline-start: 16px; color: #94a3b8; font-size: 11px; }
</style>
