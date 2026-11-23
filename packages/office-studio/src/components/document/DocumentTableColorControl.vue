<script setup lang="ts">
import { watch, type Component } from 'vue'
import DocumentColorControl from './DocumentColorControl.vue'
import type { DocumentColorParams } from '../../features/document/services/document-colors'
import type { DocumentTableCellRange } from '../../features/document/model/document-table-grid'

const props = withDefaults(defineProps<{
  title: string; field: 'background' | 'border'; icon: Component; defaultColor: string; lastColor?: string;
  selectedColor?: string; mixed?: boolean; tableId: string; range: DocumentTableCellRange;
  open?: boolean; disabled?: boolean; measurement?: boolean; grid?: boolean; overflow?: boolean; iconSize?: number;
}>(), { open: false, disabled: false, measurement: false, grid: false, overflow: false, iconSize: 16 })
const emit = defineEmits<{ insert: [params: { tableId: string; range: DocumentTableCellRange; color: string | null }]; 'update:open': [value: boolean] }>()
watch(() => [props.tableId, props.range.startRow, props.range.endRow, props.range.startColumn, props.range.endColumn], () => {
  if (props.open) emit('update:open', false)
})
function apply(params: DocumentColorParams) {
  if (props.disabled || props.measurement) return
  const field = props.field === 'background' ? 'highlight' : 'color', color = params.patch[field]
  if (typeof color !== 'string' && !params.clearFields?.includes(field)) return
  emit('insert', { tableId: props.tableId, range: { ...props.range }, color: typeof color === 'string' ? color : null })
}
</script>

<template>
  <DocumentColorControl :title="title" :field="field === 'background' ? 'highlight' : 'color'" :icon="icon" :default-color="defaultColor"
    :last-color="lastColor" :selected-color="selectedColor" :mixed="mixed" :open="open" :disabled="disabled" :measurement="measurement"
    :grid="grid" :overflow="overflow" :icon-size="iconSize" @update:open="value => emit('update:open', value)" @insert="apply" />
</template>
