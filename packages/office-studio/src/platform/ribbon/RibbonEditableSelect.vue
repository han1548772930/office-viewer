<script setup lang="ts">
import { computed, inject, ref, watch } from 'vue'
import { AutoComplete } from 'antdv-next'
import { ChevronDown } from 'lucide-vue-next'
import { confirmRibbonDraft } from './ribbon-fonts'
import type { RibbonItem } from './ribbon-schema'
import { RibbonEditorFocusScopeKey } from './ribbon-popup-scope'

const props = defineProps<{ item: RibbonItem; open: boolean; measurement: boolean; inputId: string }>()
const emit = defineEmits<{ execute: [value: string | number]; 'update:open': [value: boolean] }>()
const draft = ref<string>()
const navigating = ref(false)
const editorFocusScope = inject(RibbonEditorFocusScopeKey, undefined)
const options = computed(() => (props.item.options ?? []).map(option => ({ ...option, value: String(option.value), source: option.value })))
const displayValue = computed(() => draft.value ?? options.value.find(o => o.source === props.item.value)?.label ?? String(props.item.value ?? ''))
function reset() { draft.value = undefined; navigating.value = false }
watch(() => [props.item.value, props.item.disabled, props.item.id], reset)
function finish(value?: string | number) {
  reset(); emit('update:open', false)
  if (value !== undefined && !props.item.disabled && !props.measurement) emit('execute', value)
}
function onKey(event: KeyboardEvent) {
  if (event.isComposing || event.keyCode === 229) return
  if (event.key === 'ArrowDown' || event.key === 'ArrowUp') { navigating.value = true; return }
  if (event.key === 'Enter' && navigating.value && props.open) return // Let Ant select the active option.
  if (event.key !== 'Escape' && event.key !== 'Enter') return
  event.preventDefault(); event.stopPropagation()
  finish(event.key === 'Enter' ? confirmRibbonDraft(props.item, displayValue.value) : undefined)
}
</script>

<template>
  <span class="ribbon-edit" @keydown.capture="onKey" @keydown.stop>
    <AutoComplete :id="inputId" :aria-label="item.title" :placeholder="item.placeholder" :disabled="measurement || item.disabled"
      :value="displayValue" :open="open" :options="options" :default-active-first-option="false" :show-search="{ filterOption: false }"
      :popup-match-select-width="item.edit?.kind === 'font' ? 190 : 84" :list-height="288" size="small" variant="borderless"
      :style="{ fontFamily: item.edit?.kind === 'font' && typeof item.value === 'string' ? item.value : undefined }"
      @search="value => { draft = value; navigating = false }" @blur="reset" @open-change="value => emit('update:open', value)"
      @select="value => finish(options.find(option => option.value === value)?.source)">
      <template #optionRender="{ option }"><span :data-ribbon-editor-owner="editorFocusScope?.ownerId" :style="option.data.preview">{{ option.data.label }}</span></template>
      <template #suffix><ChevronDown :size="12" aria-hidden="true" /></template>
    </AutoComplete>
  </span>
</template>

<style scoped>
.ribbon-edit { display: inline-flex; min-width: 0; height: var(--ribbon-selector-height); align-items: center; }
.ribbon-edit :deep(.ant-select) { width: 100%; min-width: 0; height: 100%; font-size: inherit; }
.ribbon-edit :deep(.ant-select-suffix) { pointer-events: none; }
</style>
