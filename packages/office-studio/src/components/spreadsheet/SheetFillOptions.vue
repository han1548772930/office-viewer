<script setup lang="ts">
import { computed, h, nextTick, onActivated, onDeactivated, onMounted, onUnmounted, ref, useId, watch, type CSSProperties, type VNode } from 'vue'
import { Button, Dropdown, Tooltip } from 'antdv-next'
import { ChevronDown, Copy } from 'lucide-vue-next'
import type { SheetFillOptionsView } from '../../features/sheet-workbench/sheet-fill-store.ts'
import type { SheetFillApplyType } from '../../features/sheet-workbench/sheet-fill.ts'

const props = defineProps<{ options: SheetFillOptionsView; position: CSSProperties; revision?: number;
  selectionKey: string; label?: string; labels?: Partial<Record<SheetFillApplyType, string>> }>()
const emit = defineEmits<{
  (event: 'refill', value: { sessionId: string; applyType: SheetFillApplyType }): void
  (event: 'restore-focus'): void
}>()
const open = ref(false), owner = useId().replace(/:/g, '-')
const popupClass = `sheet-fill-options-popup-${owner}`
let active = true, focusRequest = 0, focusFrame = 0
const modes = ['COPY', 'SERIES', 'ONLY_FORMAT', 'NO_FORMAT'] as const
const defaultLabels: Record<SheetFillApplyType, string> = { COPY: '复制单元格', SERIES: '填充序列',
  ONLY_FORMAT: '仅填充格式', NO_FORMAT: '不带格式填充' }
const menu = computed(() => ({ items: modes.map(key => ({ key, label: props.labels?.[key] ?? defaultLabels[key],
  disabled: key === 'SERIES' && !props.options.supportsSeries })), selectable: true, selectedKeys: [props.options.applyType],
  onClick: ({ key }: { key: string }) => {
    const mode = modes.find(value => value === key)
    if (!mode || mode === 'SERIES' && !props.options.supportsSeries) return
    open.value = false
    emit('refill', { sessionId: props.options.id, applyType: mode })
  } }))
function owns(element: Element | null): boolean {
  return Boolean(element && (element.closest(`[data-sheet-fill-options-owner="${owner}"]`) || element.closest(`.${popupClass}`)))
}
function focusMenu(): void {
  const request = ++focusRequest, id = props.options.id, revision = props.revision, selection = props.selectionKey
  cancelAnimationFrame(focusFrame)
  void nextTick(() => {
    const attempt = (remaining: number) => {
      focusFrame = 0
      if (!active || request !== focusRequest || !open.value || props.options.id !== id
        || props.revision !== revision || props.selectionKey !== selection) return
      const focused = document.activeElement
      if (focused !== document.body && !owns(focused)) return
      const popupMenu = document.querySelector<HTMLElement>(`.${popupClass} [role="menu"]`)
      if (popupMenu) popupMenu.focus({ preventScroll: true })
      else if (remaining > 0) focusFrame = requestAnimationFrame(() => attempt(remaining - 1))
    }
    focusFrame = requestAnimationFrame(() => attempt(2))
  })
}
function setOpen(value: boolean): void { open.value = value; if (value) focusMenu() }
function closeAndRestore(): void {
  const id = props.options.id, revision = props.revision, selection = props.selectionKey
  open.value = false
  const request = ++focusRequest
  void nextTick(() => {
    if (active && request === focusRequest && !open.value && props.options.id === id
      && props.revision === revision && props.selectionKey === selection
      && (document.activeElement === document.body || owns(document.activeElement))) emit('restore-focus')
  })
}
function onKeydown(event: KeyboardEvent): void {
  if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); closeAndRestore() }
  else if (event.key === 'ArrowDown' && !open.value) { event.preventDefault(); setOpen(true) }
}
const popup = (node: VNode) => h('div', { onKeydownCapture: onKeydown }, [node])
function onFocusIn(event: FocusEvent): void {
  if (open.value && event.target instanceof Element && !owns(event.target)) { open.value = false; focusRequest++ }
}
watch(() => [props.options.id, props.revision, props.selectionKey, props.position.left, props.position.top], () => {
  open.value = false; focusRequest++
})
onMounted(() => document.addEventListener('focusin', onFocusIn, true))
onActivated(() => { active = true })
onDeactivated(() => { active = false; open.value = false; focusRequest++; cancelAnimationFrame(focusFrame) })
onUnmounted(() => { active = false; open.value = false; focusRequest++; cancelAnimationFrame(focusFrame); document.removeEventListener('focusin', onFocusIn, true) })
</script>

<template>
  <Dropdown :open="open" :trigger="['click']" :menu="menu" :popup-render="popup"
    :classes="{ root: `sheet-fill-options-popup ${popupClass}` }" :auto-focus="false"
    :destroy-on-hidden="true" :auto-adjust-overflow="true" placement="bottomLeft" @open-change="setOpen">
    <Tooltip :title="label ?? '自动填充选项'" :open="open ? false : undefined">
      <Button class="sheet-fill-options" :style="position" :data-sheet-fill-options-owner="owner" size="small"
        :aria-label="label ?? '自动填充选项'" :aria-expanded="open" aria-haspopup="menu"
        @pointerdown.stop @keydown.stop="onKeydown" @dblclick.stop @contextmenu.stop>
        <Copy :size="14" /><ChevronDown :size="12" />
      </Button>
    </Tooltip>
  </Dropdown>
</template>

<style scoped>
.sheet-fill-options { position: absolute; z-index: 4; display: inline-flex; align-items: center; justify-content: center;
  gap: 2px; width: 44px; height: 24px; padding: 0 4px; color: var(--sheet-canvas-text);
  background: var(--sheet-canvas-background); border-color: var(--sheet-canvas-border); }
.sheet-fill-options:focus-visible { outline: 2px solid var(--sheet-canvas-primary); outline-offset: 1px; }
</style>
