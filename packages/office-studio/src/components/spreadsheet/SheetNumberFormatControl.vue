<script setup lang="ts">
import { computed, inject, nextTick, ref, useId, watch } from 'vue'
import { Button, Menu, Popover, Tooltip } from 'antdv-next'
import type { MenuInfo, MenuRef } from 'antdv-next'
import { ChevronDown } from 'lucide-vue-next'
import { formatSheetValue, sheetNumberFormatKey, sheetNumberFormatLabel, sheetNumberFormatMenu } from '../../features/sheet-workbench/sheet-number-format.ts'
import { numberFormatWithoutDecimals } from '../../features/sheet-workbench/number-format-pattern.ts'
import { useRibbonTheme } from '../../platform/ribbon/ribbon-theme.ts'
import type { OfficeTranslator } from '../../i18n/office-i18n.ts'
import { OfficeDirectionKey } from '../../i18n/office-i18n.ts'

const props = withDefaults(defineProps<{
  pattern: string | null; region?: string; title?: string; disabled?: boolean; measurement?: boolean
  grid?: boolean; overflow?: boolean; open?: boolean; translate?: OfficeTranslator
}>(), { region: 'zh-CN', title: '数字格式', disabled: false, measurement: false, grid: false, overflow: false })
const emit = defineEmits<{ insert: [value: { pattern: string | null }]; execute: [id: string, params?: unknown]; 'update:open': [open: boolean] }>()
const localOpen = ref(false), tooltipOpen = ref(false)
const root = ref<HTMLElement>(), menu = ref<MenuRef>()
const menuId = useId()
const themeStyle = useRibbonTheme()
const direction = inject(OfficeDirectionKey, ref<'ltr' | 'rtl'>('ltr'))
const open = computed({ get: () => props.open ?? localOpen.value, set: value => {
  if (value && (props.disabled || props.measurement)) return
  localOpen.value = value; emit('update:open', value); tooltipOpen.value = false
} })
const entries = computed(() => sheetNumberFormatMenu(props.region))
function t(key: string, fallback: string): string { return props.translate?.(key, fallback) ?? fallback }
const caption = computed(() => {
  const label = sheetNumberFormatLabel(props.pattern, props.region)
  return t(`sheet.number-format.${sheetNumberFormatKey(props.pattern, props.region)}`, label)
})
const items = computed(() => entries.value.map(entry => entry.type === 'divider' ? { key: entry.id, type: 'divider' as const }
  : { key: entry.id, label: t(`sheet.number-format.${entry.id}`, entry.label), disabled: props.disabled || props.measurement }))
const selected = computed(() => entries.value.flatMap(entry => entry.type === 'format'
  && props.pattern !== null && numberFormatWithoutDecimals(entry.pattern ?? 'General') === numberFormatWithoutDecimals(props.pattern) ? [entry.id] : []))
function preview(key: string): string {
  const entry = entries.value.find(entry => entry.id === key)
  return entry?.type === 'format' && entry.pattern ? formatSheetValue(1220, entry.pattern, false, props.region)?.trim() ?? '' : ''
}
function choose({ key }: MenuInfo): void {
  if (props.disabled || props.measurement) return
  const entry = entries.value.find(entry => entry.id === String(key))
  if (!entry || entry.type === 'divider') return
  open.value = false
  if (entry.type === 'more') emit('execute', 'sheet.operation.open-custom-number-format')
  else emit('insert', { pattern: entry.pattern })
}
async function openFromKeyboard(event: KeyboardEvent, last = false): Promise<void> {
  if (event.isComposing || event.keyCode === 229 || props.disabled || props.measurement) return
  event.preventDefault(); event.stopPropagation(); open.value = true
  await nextTick()
  const controls = menu.value?.menu?.list.querySelectorAll<HTMLElement>('[role="menuitem"]:not([aria-disabled="true"])')
  const target = last ? controls?.[controls.length - 1] : controls?.[0]
  if (target) target.focus()
  else menu.value?.focus()
}
function escape(event: KeyboardEvent): void {
  if (event.isComposing || event.keyCode === 229) return
  event.preventDefault(); event.stopPropagation(); open.value = false
  root.value?.querySelector<HTMLButtonElement>('button')?.focus({ preventScroll: true })
}
watch(() => props.disabled || props.measurement, unavailable => { if (unavailable) open.value = false })
</script>

<template>
  <span ref="root" class="sheet-number-format-control" :class="{ 'is-grid': grid, 'is-overflow': overflow }" :style="themeStyle">
    <span v-if="overflow" class="sheet-number-format-label">{{ title }}</span>
    <Popover v-model:open="open" trigger="click" :placement="direction === 'rtl' ? 'bottomRight' : 'bottomLeft'" :destroy-on-hidden="true" :align="{ overflow: { shiftX: true, adjustX: true, adjustY: true } }">
      <Tooltip :title="measurement || open ? undefined : title" :trigger="['hover', 'focus']" :mouse-enter-delay="0.6" :open="tooltipOpen && !open" @open-change="value => { tooltipOpen = !open && value }">
        <Button type="text" size="small" class="sheet-number-format-trigger" :styles="{ root: { border: 0 } }" :disabled="disabled || measurement" :aria-label="title" aria-haspopup="menu" :aria-expanded="open" :aria-controls="menuId"
          @mousedown.prevent @keydown.down="openFromKeyboard" @keydown.up="openFromKeyboard($event, true)" @keydown.esc="escape">
          <span>{{ caption }}</span><ChevronDown :size="12" aria-hidden="true" />
        </Button>
      </Tooltip>
      <template #content>
        <Menu :id="menuId" ref="menu" class="sheet-number-format-menu" :items="items" :selected-keys="selected" :aria-label="title" :dir="direction" :style="themeStyle" @click="choose" @keydown.esc.capture="escape">
          <template #labelRender="item">
            <span class="sheet-number-format-row"><span>{{ item.label }}</span><span class="sheet-number-format-preview">{{ preview(String(item.key)) }}</span></span>
          </template>
        </Menu>
      </template>
    </Popover>
  </span>
</template>

<style scoped>
.sheet-number-format-control { display: flex; width: 160px; max-width: 100%; min-width: 0; align-items: center; color: var(--ribbon-text); }
.sheet-number-format-trigger { display: flex; box-sizing: border-box; width: 100%; min-width: 0; height: 28px; align-items: center; justify-content: space-between; gap: 4px; padding-inline: 6px; color: inherit; font-size: 14px; line-height: 20px; }
.sheet-number-format-trigger > span { min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.sheet-number-format-trigger :deep(svg) { display: block; width: 12px; height: 12px; flex: 0 0 12px; }
.sheet-number-format-trigger:not(:disabled):hover { background: var(--ribbon-hover-background); }
.sheet-number-format-trigger[aria-expanded="true"]:not(:disabled),
.sheet-number-format-trigger[aria-expanded="true"]:not(:disabled):hover,
.sheet-number-format-trigger[aria-expanded="true"]:not(:disabled):active { background: var(--ribbon-active-background); color: var(--ribbon-primary); }
.sheet-number-format-trigger:disabled { color: var(--ribbon-disabled); background: transparent; cursor: not-allowed; }
.is-grid .sheet-number-format-trigger { height: 24px; font-size: 12px; line-height: 16px; }
.is-overflow { width: 100%; gap: 8px; }
.is-overflow .sheet-number-format-trigger { max-width: 160px; height: 28px; flex: 1 1 auto; }
.sheet-number-format-label { min-width: 56px; font-size: 12px; }
.sheet-number-format-menu { width: 260px; max-width: calc(100vw - 32px); border-inline-end: 0; background: var(--ribbon-background); color: var(--ribbon-text); }
.sheet-number-format-menu :deep(.ant-menu-item) { height: 28px; margin-block: 2px; line-height: 28px; }
.sheet-number-format-row { display: flex; align-items: center; justify-content: space-between; gap: 24px; }
.sheet-number-format-row > span:first-child { min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.sheet-number-format-preview { flex: 0 0 auto; color: var(--ribbon-text-secondary); font-size: 12px; }
</style>
