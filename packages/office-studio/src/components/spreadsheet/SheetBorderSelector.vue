<script setup lang="ts">
import { computed } from 'vue'
import { Button, Form, FormItem, Select, Tooltip, theme } from 'antdv-next'
import SheetColorControl from './SheetColorControl.vue'
import SheetBorderIcon from './SheetBorderIcon.vue'
import { sheetColorHex } from '../../features/sheet-workbench/sheet-colors.ts'
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
  selectedColor?: string | null
  mixed?: boolean
  disabled?: boolean
}>()
const emit = defineEmits<{
  'update:value': [value: SheetBorderOptions]
  insert: [preset: SheetBorderPreset]
}>()

const { token } = theme.useToken()
// Accepted command state owns the pen; rejected commands cannot leave local color/style behind.
const selectedColor = computed(() => props.value.color)
const selectedStyle = computed(() => props.value.style)

const styleOptions = computed(() => styleOrder.map(value => ({ label: props.translate(`sheet.border.style.${value}`, sheetBorderStyleLabels[value]), value })))
function renderStyleLabel(): string {
  const style = styleOrder.find(value => value === selectedStyle.value)
  return style ? props.translate(`sheet.border.style.${style}`, sheetBorderStyleLabels[style])
    : selectedStyle.value === 'none' ? props.translate('sheet.border.style.none', '无边框线')
      : props.translate('sheet.border.style.slant-dash-dot', '倾斜点划线')
}
const label = (preset: SheetBorderPreset) => props.translate(`sheet.items.sheet.menu.border.${preset}`, sheetBorderLabels[preset])

function updateColor(_commandId: string, value?: string): void {
  const color = sheetColorHex(value)
  if (props.disabled || !color) return
  emit('update:value', { style: selectedStyle.value, color })
}

function updateStyle(value: unknown): void {
  if (props.disabled) return
  const style = styleOrder.find(item => item === value)
  if (style) emit('update:value', { style, color: selectedColor.value })
}

function apply(preset: SheetBorderPreset): void {
  if (props.disabled) return
  emit('insert', preset)
}
</script>

<template>
      <section class="border-selector" :style="{ '--edge-ink': token.colorText, '--edge-guide': token.colorBorder, '--edge-active': token.colorPrimaryBg, '--edge-control-height': `${token.controlHeightSM}px` }">
        <div class="border-selector-options">
          <Tooltip v-for="preset in presetOrder" :key="preset" :title="label(preset)" :trigger="['hover', 'focus']">
            <Button type="text" class="border-preset" :disabled="disabled" :aria-label="label(preset)" :aria-pressed="props.preset === preset" @mousedown.prevent @click="apply(preset)">
              <SheetBorderIcon :preset="preset" />
            </Button>
          </Tooltip>
        </div>
        <Form class="border-selector-fields" layout="vertical" :disabled="disabled" @submit.prevent>
          <FormItem class="border-color-field" :label="translate('sheet.border.color', '边框颜色')">
            <SheetColorControl field="border" command-id="sheet.command.set-border-color" :title="translate('sheet.border.color', '边框颜色')" :last-color="selectedColor" :selected-color="props.selectedColor" :mixed="mixed" :disabled="disabled" :split="false" @execute="updateColor" />
          </FormItem>
          <FormItem class="border-style-field" :label="translate('sheet.border.line-style', '边框样式')">
            <Select :value="selectedStyle" :label-render="renderStyleLabel" :options="styleOptions" :disabled="disabled" :aria-label="translate('sheet.border.line-style', '边框样式')" size="small" @change="updateStyle" />
          </FormItem>
        </Form>
      </section>
</template>

<style scoped>
.border-selector { display: grid; width: min(220px, calc(100vw - 56px)); gap: 8px; color: var(--edge-ink); }
.border-selector-options { display: grid; grid-template-columns: repeat(5, 1fr); justify-items: center; gap: 8px; }
.border-preset[aria-pressed="true"] { background: var(--edge-active); }
.border-preset { display: grid; width: 24px; min-width: 24px; height: 24px; place-items: center; padding: 0; }
.border-selector-fields { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; align-items: start; }
.border-selector-fields :deep(.ant-form-item) { min-width: 0; margin-bottom: 0; }
.border-selector-fields :deep(.ant-form-item-label) { padding-bottom: 4px; font-size: 12px; }
.border-color-field :deep(.sheet-color-control) { width: 100%; height: var(--edge-control-height); }
.border-style-field :deep(.ant-select) { width: 100%; }
</style>
