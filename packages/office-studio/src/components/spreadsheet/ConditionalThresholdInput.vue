<script setup lang="ts">
import { computed } from 'vue'
import { FormItem, Input, Select } from 'antdv-next'
import type { SheetConditionalFormatValueConfig } from '../../domain/model.ts'

const props = defineProps<{ value: SheetConditionalFormatValueConfig; boundary?: 'min' | 'max'; automatic?: boolean; label: string }>()
const emit = defineEmits<{ 'update:value': [value: SheetConditionalFormatValueConfig] }>()
const options = computed(() => [
  ...(props.boundary ? [{ value: props.boundary, label: props.boundary === 'min' ? '最小值' : '最大值' }] : []),
  ...(props.automatic && props.boundary ? [{ value: props.boundary === 'min' ? 'autoMin' : 'autoMax', label: props.boundary === 'min' ? '自动最小值' : '自动最大值' }] : []),
  { value: 'num', label: '数字' }, { value: 'percent', label: '百分比' },
  { value: 'percentile', label: '百分位' }, { value: 'formula', label: '公式' },
])
function changeType(value: unknown) {
  const type = String(value)
  if (['min', 'max', 'autoMin', 'autoMax'].includes(type)) {
    emit('update:value', { type })
    return
  }
  emit('update:value', { type, value: type === 'formula' ? props.value.type === 'formula' ? props.value.value : '=' : props.value.type === 'formula' ? 0 : props.value.value ?? 0 })
}
function changeValue(value: string) {
  emit('update:value', { ...props.value, value })
}
</script>

<template>
  <FormItem :label="label" class="conditional-threshold-item">
    <div class="conditional-threshold-controls">
      <Select :value="value.type" :options="options" :aria-label="`${label}类型`" @update:value="changeType" />
      <Input v-if="!['min', 'max', 'autoMin', 'autoMax'].includes(value.type)" :value="String(value.value ?? '')" :aria-label="`${label}值`" :inputmode="value.type === 'formula' ? undefined : 'decimal'" @update:value="changeValue" />
    </div>
  </FormItem>
</template>

<style scoped>
.conditional-threshold-item { margin-bottom: 0; }
.conditional-threshold-controls { display: grid; grid-template-columns: minmax(80px, 1fr) minmax(0, 1fr); gap: 8px; }
</style>
