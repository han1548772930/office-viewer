<script setup lang="ts">
import { computed, h, nextTick, onDeactivated, onMounted, onUnmounted, ref, shallowRef, watch } from 'vue'
import { Button, Dropdown, Tooltip } from 'antdv-next'
import { CHART_OBJECT_MENU, type ChartObjectAction, type ChartObjectCapabilities, type ChartObjectTarget } from '../chart-object-commands'

const props = defineProps<{
  selected: boolean
  chartId: string
  capture(): ChartObjectTarget | undefined
  current(target: ChartObjectTarget): boolean
  capabilities(target?: ChartObjectTarget): ChartObjectCapabilities
  execute(action: ChartObjectAction, target: ChartObjectTarget): Promise<boolean>
  restoreFocus(target?: ChartObjectTarget): void
  interceptContext?(event: MouseEvent): boolean
  interceptKeyboard?(event: KeyboardEvent): boolean
}>()
const emit = defineEmits<{ (event: 'select'): void }>()
const contextOpen = ref(false), moreOpen = ref(false)
const captured = shallowRef<ChartObjectTarget>()
const items = computed(() => {
  const capabilities = props.capabilities(captured.value)
  return CHART_OBJECT_MENU.filter(entry => !capabilities[entry.action].hidden).map(entry => ({
    key: entry.action, label: h(Tooltip, { title: capabilities[entry.action].reason }, () => h('span', entry.label)),
    disabled: capabilities[entry.action].disabled, danger: entry.action === 'delete', extra: entry.shortcut,
  }))
})
function open(kind: 'context' | 'more', value: boolean): void {
  if (value) {
    const target = props.capture()
    if (!target || !props.current(target)) return
    captured.value = target; emit('select')
    contextOpen.value = kind === 'context'; moreOpen.value = kind === 'more'
  } else {
    if (kind === 'context') contextOpen.value = false
    else moreOpen.value = false
  }
}
async function choose({ key }: { key: string | number }): Promise<void> {
  const target = captured.value, action = CHART_OBJECT_MENU.find(entry => entry.action === key)?.action
  contextOpen.value = false; moreOpen.value = false
  if (!target || !action || !props.current(target)) return
  const accepted = await props.execute(action, target)
  if (action !== 'edit' || !accepted) { await nextTick(); props.restoreFocus(target) }
}
function keyboard(event: KeyboardEvent): void {
  if (event.defaultPrevented || event.isComposing || event.keyCode === 229) return
  if (props.interceptKeyboard?.(event)) return
  if (event.key === 'ContextMenu' || event.key === 'F10' && event.shiftKey && !event.ctrlKey && !event.metaKey && !event.altKey) {
    event.preventDefault(); event.stopPropagation(); open('context', true)
  } else if (event.key === 'Escape' && (contextOpen.value || moreOpen.value)) {
    event.preventDefault(); event.stopPropagation(); close(true)
  }
}
function context(event: MouseEvent): void {
  if (props.interceptContext?.(event)) { contextOpen.value = false; moreOpen.value = false; event.preventDefault(); event.stopPropagation() }
}
function close(restore = false): void {
  const target = captured.value
  contextOpen.value = false; moreOpen.value = false; captured.value = undefined
  if (restore) props.restoreFocus(target)
}
function popupKeyDown(event: KeyboardEvent): void {
  if ((contextOpen.value || moreOpen.value) && event.key === 'Escape' && !event.isComposing && event.keyCode !== 229) {
    event.preventDefault(); event.stopPropagation(); close(true)
  }
}
watch(() => captured.value && props.current(captured.value), current => { if (captured.value && !current) close() }, { flush: 'sync' })
watch(() => props.selected, selected => { if (!selected) close() })
onMounted(() => document.addEventListener('keydown', popupKeyDown, true))
onUnmounted(() => document.removeEventListener('keydown', popupKeyDown, true))
onDeactivated(() => close())
defineExpose({ openKeyboard: () => open('context', true), close })
</script>

<template>
  <div class="chart-context-controls" :data-chart-menu-id="chartId" @keydown="keyboard" @contextmenu.capture="context">
    <Dropdown :trigger="['contextmenu']" :open="contextOpen" :menu="{ items, selectable: false }" auto-focus destroy-on-hidden
      @update:open="value => open('context', value)" @menu-click="choose">
      <slot />
    </Dropdown>
    <Dropdown v-if="selected" :trigger="['click']" :open="moreOpen" :menu="{ items, selectable: false }" auto-focus destroy-on-hidden
      @update:open="value => open('more', value)" @menu-click="choose">
      <Tooltip title="图表菜单">
        <Button class="chart-object-more" type="default" size="small" aria-label="图表菜单" :aria-expanded="moreOpen"
          @pointerdown.stop @click.stop>···</Button>
      </Tooltip>
    </Dropdown>
  </div>
</template>

<style scoped>
.chart-context-controls { position: absolute; inset: 0; pointer-events: none; }
.chart-object-more { position: absolute; right: 8px; top: 8px; min-width: 32px; pointer-events: auto; }
</style>
