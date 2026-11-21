<script setup lang="ts">
import { computed, inject, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import { Button } from 'antdv-next'
import type { GridHit } from '../../rendering/spreadsheet/layout.ts'
import { createSheetContextMenuItems } from './sheet-context-menu.ts'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

const props = defineProps<{
  hit: GridHit | null
  x: number
  y: number
  open: boolean
  canDeleteRow: boolean
  canDeleteColumn: boolean
  frozen: boolean
}>()
const emit = defineEmits<{
  'update:open': [value: boolean]
  command: [commandId: string, params?: unknown]
}>()

const panel = ref<HTMLElement>()
const translate = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const activeIndex = ref(0)
const items = computed(() => props.hit ? createSheetContextMenuItems(props.hit, props).map(item => ({
  ...item,
  label: translate(`sheet.context-menu.${item.id}`, item.label),
})) : [])
const enabledItems = computed(() => items.value.filter(item => !item.disabled))
const position = computed(() => ({ left: `${Math.max(4, props.x)}px`, top: `${Math.max(4, props.y)}px` }))

function close() { emit('update:open', false) }
function execute(item: (typeof items.value)[number]) {
  if (item.disabled) return
  close()
  emit('command', item.commandId, item.params)
}
function focusActive() {
  const item = enabledItems.value[activeIndex.value]
  if (item) panel.value?.querySelector<HTMLButtonElement>(`[data-menu-item="${item.id}"]`)?.focus()
}
function onKeydown(event: KeyboardEvent) {
  if (event.key === 'Escape') { event.preventDefault(); close(); return }
  if (!enabledItems.value.length) return
  if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
    event.preventDefault()
    const delta = event.key === 'ArrowDown' ? 1 : -1
    activeIndex.value = (activeIndex.value + delta + enabledItems.value.length) % enabledItems.value.length
    void nextTick(focusActive)
  } else if (event.key === 'Home' || event.key === 'End') {
    event.preventDefault()
    activeIndex.value = event.key === 'Home' ? 0 : enabledItems.value.length - 1
    void nextTick(focusActive)
  }
}
function onOutsidePointer(event: PointerEvent) {
  if (props.open && !panel.value?.contains(event.target as Node)) close()
}
function onViewportChange() { if (props.open) close() }

watch(() => [props.open, props.x, props.y] as const, async ([open]) => {
  if (!open) return
  activeIndex.value = 0
  await nextTick()
  const bounds = panel.value?.getBoundingClientRect()
  if (bounds && (bounds.right > window.innerWidth - 4 || bounds.bottom > window.innerHeight - 4)) {
    panel.value!.style.left = `${Math.max(4, Math.min(props.x, window.innerWidth - bounds.width - 4))}px`
    panel.value!.style.top = `${Math.max(4, Math.min(props.y, window.innerHeight - bounds.height - 4))}px`
  }
  focusActive()
})
onMounted(() => {
  document.addEventListener('pointerdown', onOutsidePointer)
  window.addEventListener('resize', onViewportChange)
  window.addEventListener('scroll', onViewportChange, true)
})
onBeforeUnmount(() => {
  document.removeEventListener('pointerdown', onOutsidePointer)
  window.removeEventListener('resize', onViewportChange)
  window.removeEventListener('scroll', onViewportChange, true)
})
</script>

<template>
  <Teleport to="body">
    <div v-if="open && items.length" ref="panel" class="sheet-context-menu" role="menu" :aria-label="translate('sheet.context-menu.aria-label', '单元格上下文菜单')" :style="position" @keydown="onKeydown">
      <template v-for="item in items" :key="item.id">
        <div v-if="item.separatorBefore" class="sheet-context-menu__separator" role="separator" />
        <Button
          :data-menu-item="item.id"
          type="text"
          html-type="button"
          role="menuitem"
          class="sheet-context-menu__item"
          :class="{ 'is-disabled': item.disabled }"
          :disabled="item.disabled"
          tabindex="-1"
          @click="execute(item)"
        >{{ item.label }}</Button>
      </template>
    </div>
  </Teleport>
</template>

<style scoped>
.sheet-context-menu { position: fixed; z-index: 3000; box-sizing: border-box; display: flex; width: 224px; max-width: calc(100vw - 8px); max-height: min(520px, calc(100vh - 8px)); flex-direction: column; overflow: auto; padding: 4px; border: 1px solid #d8dde5; border-radius: 4px; background: #fff; box-shadow: 0 8px 24px #1720332b; color: #263244; }
.sheet-context-menu .sheet-context-menu__item.ant-btn { display: flex; width: 100%; min-height: 30px; height: auto; align-items: center; justify-content: flex-start; padding: 4px 9px; border: 0; border-radius: 2px; background: transparent; color: inherit; font: inherit; font-size: 13px; text-align: start; white-space: normal; }
.sheet-context-menu .sheet-context-menu__item.ant-btn:hover, .sheet-context-menu .sheet-context-menu__item.ant-btn:focus-visible { outline: none; background: #edf3fc; }
.sheet-context-menu .sheet-context-menu__item.ant-btn:disabled { color: #9aa3af; cursor: default; }
.sheet-context-menu__separator { height: 1px; flex: 0 0 1px; margin: 4px 3px; background: #e3e7ed; }
</style>
