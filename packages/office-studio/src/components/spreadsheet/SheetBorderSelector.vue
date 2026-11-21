<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { Button, ColorPicker, Form, FormItem, Popover, Select, Tooltip } from 'antdv-next'
import { ChevronDown, Grid3X3 } from 'lucide-vue-next'
import type { SheetBorderOptions, SheetBorderPreset, SheetBorderStyle } from '../../features/sheet-workbench/sheet-toolbar-format.ts'

const presetOrder: SheetBorderPreset[] = [
  'all', 'outside', 'inside', 'top', 'bottom', 'left', 'right', 'horizontal', 'vertical', 'diagonal-down', 'diagonal-up', 'none',
]
const styleOrder: SheetBorderStyle[] = [
  'thin', 'hair', 'dotted', 'dashed', 'dash-dot', 'dash-dot-dot', 'medium', 'medium-dashed', 'medium-dash-dot',
  'medium-dash-dot-dot', 'thick', 'double',
]

const props = defineProps<{
  value: SheetBorderOptions
  title: string
  colorLabel: string
  styleLabel: string
  presets: Record<SheetBorderPreset, string>
  styles: Record<SheetBorderStyle, string>
}>()
const emit = defineEmits<{
  'update:value': [value: SheetBorderOptions]
  apply: [preset: SheetBorderPreset, value: SheetBorderOptions]
}>()

const open = ref(false)
const selectedColor = ref(props.value.color)
const selectedStyle = ref<SheetBorderStyle>(props.value.style)
watch(() => props.value, value => {
  selectedColor.value = value.color
  selectedStyle.value = value.style
}, { deep: true })

const styleOptions = computed(() => styleOrder.map(value => ({ label: props.styles[value], value })))
const swatchStyle = computed(() => ({ borderColor: selectedColor.value }))

function updateColor(value: unknown): void {
  const color = typeof value === 'string'
    ? value
    : value && typeof value === 'object' && 'toHexString' in value && typeof value.toHexString === 'function'
      ? value.toHexString()
      : undefined
  if (!color) return
  selectedColor.value = color
  emit('update:value', { style: selectedStyle.value, color })
}

function updateStyle(style: SheetBorderStyle): void {
  selectedStyle.value = style
  emit('update:value', { style, color: selectedColor.value })
}

function apply(preset: SheetBorderPreset): void {
  emit('apply', preset, { style: selectedStyle.value, color: selectedColor.value })
  open.value = false
}
</script>

<template>
  <Popover v-model:open="open" trigger="click" placement="bottomLeft" :align="{ overflow: { shiftX: true, adjustX: true, adjustY: true } }" :destroy-on-hidden="true">
    <Button type="text" size="small" class="border-selector-trigger" :aria-label="title" aria-haspopup="dialog" :aria-expanded="open">
      <Grid3X3 :size="16" />
      <span class="border-selector-swatch" :style="swatchStyle" />
      <ChevronDown :size="12" />
    </Button>
    <template #content>
      <section class="border-selector" :aria-label="title">
        <div class="border-selector-options">
          <Tooltip v-for="preset in presetOrder" :key="preset" :title="presets[preset]">
            <Button type="text" class="border-preset" :aria-label="presets[preset]" @click="apply(preset)">
              <span class="border-preview" :class="`is-${preset}`" />
            </Button>
          </Tooltip>
        </div>
        <Form class="border-selector-fields" layout="vertical" @submit.prevent>
          <FormItem class="border-color-field" :label="colorLabel">
            <ColorPicker :value="selectedColor" value-format="hex" :aria-label="colorLabel" @update:value="updateColor" />
          </FormItem>
          <FormItem class="border-style-field" :label="styleLabel">
            <Select :value="selectedStyle" :options="styleOptions" :aria-label="styleLabel" size="small" @change="updateStyle" />
          </FormItem>
        </Form>
      </section>
    </template>
  </Popover>
</template>

<style scoped>
.border-selector-trigger { display: inline-flex; width: 54px; min-width: 54px; align-items: center; justify-content: space-between; padding: 0 5px; }
.border-selector-swatch { width: 13px; height: 3px; border-bottom: 2px solid; }
.border-selector { display: grid; width: 230px; gap: 10px; }
.border-selector-options { display: grid; grid-template-columns: repeat(6, 30px); justify-content: space-between; gap: 4px; }
.border-preset { display: grid; width: 30px; min-width: 30px; height: 30px; place-items: center; padding: 0; }
.border-preview { position: relative; box-sizing: border-box; display: block; width: 18px; height: 18px; border: 1px solid #c6cbd3; }
.border-preview.is-all { border-color: #475467; }
.border-preview.is-outside { border: 2px solid #475467; }
.border-preview.is-inside { background: linear-gradient(90deg, transparent 47%, #475467 48%, #475467 52%, transparent 53%), linear-gradient(transparent 47%, #475467 48%, #475467 52%, transparent 53%); }
.border-preview.is-top { border-top: 2px solid #475467; }
.border-preview.is-bottom { border-bottom: 2px solid #475467; }
.border-preview.is-left { border-left: 2px solid #475467; }
.border-preview.is-right { border-right: 2px solid #475467; }
.border-preview.is-horizontal { background: linear-gradient(transparent 47%, #475467 48%, #475467 52%, transparent 53%); }
.border-preview.is-vertical { background: linear-gradient(90deg, transparent 47%, #475467 48%, #475467 52%, transparent 53%); }
.border-preview.is-diagonal-down::after, .border-preview.is-diagonal-up::after { position: absolute; top: 8px; left: -2px; width: 20px; border-top: 1px solid #475467; content: ''; }
.border-preview.is-diagonal-down::after { transform: rotate(45deg); }
.border-preview.is-diagonal-up::after { transform: rotate(-45deg); }
.border-preview.is-none { border: 1px dashed #98a2b3; }
.border-selector-fields { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; align-items: start; }
.border-selector-fields :deep(.ant-form-item) { min-width: 0; margin-bottom: 0; }
.border-selector-fields :deep(.ant-form-item-label) { padding-bottom: 4px; color: #475467; font-size: 12px; }
.border-color-field :deep(.ant-color-picker-trigger) { width: 100%; height: 28px; justify-content: flex-start; }
.border-style-field :deep(.ant-select) { width: 100%; }
</style>
