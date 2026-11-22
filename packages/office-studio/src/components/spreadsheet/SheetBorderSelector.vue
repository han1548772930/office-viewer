<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { Button, ColorPicker, Form, FormItem, Select, Tooltip, theme } from 'antdv-next'
import { sheetBorderLabels, sheetBorderStyleLabels, type SheetBorderOptions, type SheetBorderPreset, type SheetBorderStyle } from '../../features/sheet-workbench/sheet-toolbar-format.ts'
import type { OfficeTranslator } from '../../i18n/office-i18n.ts'

const presetOrder = Object.keys(sheetBorderLabels) as SheetBorderPreset[]
const styleOrder: SheetBorderStyle[] = [
  'thin', 'hair', 'dotted', 'dashed', 'dash-dot', 'dash-dot-dot', 'medium', 'medium-dashed', 'medium-dash-dot',
  'medium-dash-dot-dot', 'thick', 'double',
]

const props = defineProps<{
  value: SheetBorderOptions
  translate: OfficeTranslator
  preset: SheetBorderPreset
}>()
const emit = defineEmits<{
  'update:value': [value: SheetBorderOptions]
  insert: [preset: SheetBorderPreset]
}>()

const { token } = theme.useToken()
const selectedColor = ref(props.value.color)
const selectedStyle = ref<SheetBorderStyle>(props.value.style)
watch(() => props.value, value => {
  selectedColor.value = value.color
  selectedStyle.value = value.style
}, { deep: true })

const styleOptions = computed(() => styleOrder.map(value => ({ label: props.translate(`sheet.border.style.${value}`, sheetBorderStyleLabels[value]), value })))
const label = (preset: SheetBorderPreset) => props.translate(`sheet.items.sheet.menu.border.${preset}`, sheetBorderLabels[preset])

function updateColor(value: unknown): void {
  const color = typeof value === 'string'
    ? value
    : value && typeof value === 'object' && 'toHexString' in value && typeof value.toHexString === 'function'
      ? value.toHexString()
      : undefined
  if (typeof color !== 'string' || !/^#[0-9a-f]{6}$/i.test(color)) return
  selectedColor.value = color
  emit('update:value', { style: selectedStyle.value, color })
}

function updateStyle(value: unknown): void {
  const style = styleOrder.find(item => item === value)
  if (style) { selectedStyle.value = style; emit('update:value', { style, color: selectedColor.value }) }
}

function apply(preset: SheetBorderPreset): void {
  emit('insert', preset)
}
</script>

<template>
      <section class="border-selector" :style="{ '--edge-ink': token.colorText, '--edge-guide': token.colorBorder, '--edge-active': token.colorPrimaryBg }">
        <div class="border-selector-options">
          <Tooltip v-for="preset in presetOrder" :key="preset" :title="label(preset)">
            <Button type="text" class="border-preset" :aria-label="label(preset)" :aria-pressed="props.preset === preset" @mousedown.prevent @click="apply(preset)">
              <span class="border-preview" :class="`is-${preset}`"><svg v-if="preset.includes('pair') || preset.includes('triple')" viewBox="0 0 20 20" aria-hidden="true"><path :d="preset === 'diagonal-up-pair' ? 'M0 10 L20 0 L10 20' : preset === 'diagonal-down-triple' ? 'M20 10 L0 0 L10 20 M0 0 L20 20' : 'M20 10 L0 0 L10 20'" fill="none" stroke="currentColor" /></svg></span>
            </Button>
          </Tooltip>
        </div>
        <Form class="border-selector-fields" layout="vertical" @submit.prevent>
          <FormItem class="border-color-field" :label="translate('sheet.border.color', '边框颜色')">
            <ColorPicker :value="selectedColor" :disabled-alpha="true" :aria-label="translate('sheet.border.color', '边框颜色')" @update:value="updateColor" />
          </FormItem>
          <FormItem class="border-style-field" :label="translate('sheet.border.line-style', '边框样式')">
            <Select :value="selectedStyle" :options="styleOptions" :aria-label="translate('sheet.border.line-style', '边框样式')" size="small" @change="updateStyle" />
          </FormItem>
        </Form>
      </section>
</template>

<style scoped>
.border-selector { display: grid; width: min(220px, calc(100vw - 56px)); gap: 8px; color: var(--edge-ink); }
.border-selector-options { display: grid; grid-template-columns: repeat(5, 1fr); justify-items: center; gap: 8px; }
.border-preset[aria-pressed="true"] { background: var(--edge-active); }
.border-preset { display: grid; width: 24px; min-width: 24px; height: 24px; place-items: center; padding: 0; }
.border-preview { position: relative; box-sizing: border-box; display: block; width: 18px; height: 18px; border: 1px solid var(--edge-guide); color: var(--edge-ink); }
.border-preview.is-all { border-color: currentColor; }
.border-preview.is-outside { border: 2px solid currentColor; }
.border-preview.is-all, .border-preview.is-inside { background: linear-gradient(90deg, transparent 46%, currentColor 47%, currentColor 53%, transparent 54%), linear-gradient(transparent 46%, currentColor 47%, currentColor 53%, transparent 54%); }
.border-preview.is-top { border-top: 2px solid currentColor; }
.border-preview.is-bottom { border-bottom: 2px solid currentColor; }
.border-preview.is-left { border-left: 2px solid currentColor; }
.border-preview.is-right { border-right: 2px solid currentColor; }
.border-preview.is-horizontal { background: linear-gradient(transparent 46%, currentColor 47%, currentColor 53%, transparent 54%); }
.border-preview.is-vertical { background: linear-gradient(90deg, transparent 46%, currentColor 47%, currentColor 53%, transparent 54%); }
.border-preview.is-diagonal-down::after, .border-preview.is-diagonal-up::after { position: absolute; top: 8px; left: -2px; width: 20px; border-top: 1px solid currentColor; content: ''; }
.border-preview.is-diagonal-down::after { transform: rotate(45deg); }
.border-preview.is-diagonal-up::after { transform: rotate(-45deg); }
.border-preview.is-none { border: 1px dashed var(--edge-guide); }
.border-selector-fields { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; align-items: start; }
.border-selector-fields :deep(.ant-form-item) { min-width: 0; margin-bottom: 0; }
.border-selector-fields :deep(.ant-form-item-label) { padding-bottom: 4px; font-size: 12px; }
.border-color-field :deep(.ant-color-picker-trigger) { width: 100%; height: 28px; justify-content: flex-start; }
.border-style-field :deep(.ant-select) { width: 100%; }
</style>
