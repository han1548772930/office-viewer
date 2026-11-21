<script setup lang="ts">
import { computed } from 'vue'
import { Input, Select } from 'antdv-next'
import type { SheetConditionalFormatValueConfig } from '../../domain/model.ts'

const props = defineProps<{ value: SheetConditionalFormatValueConfig; boundary?: 'min' | 'max'; label: string }>()
const emit = defineEmits<{ 'update:value': [value: SheetConditionalFormatValueConfig] }>()
const options = computed(() => [
  ...(props.boundary ? [{ value: props.boundary, label: props.boundary === 'min' ? '最小值' : '最大值' }] : []),
  { value: 'num', label: '数字' }, { value: 'percent', label: '百分比' },
  { value: 'percentile', label: '百分位' }, { value: 'formula', label: '公式' },
])
function changeType(value: unknown) {
  const type = String(value)
  emit('update:value', { type, value: type === 'formula' ? props.value.type === 'formula' ? props.value.value : '=' : props.value.type === 'formula' ? 0 : props.value.value ?? 0 })
}
function changeValue(value: string) {
  emit('update:value', { ...props.value, value })
}
</script>

<template>
  <div class="conditional-threshold">
    <span>{{ label }}</span>
    <Select :value="value.type" :options="options" :aria-label="`${label}类型`" @update:value="changeType" />
    <Input v-if="value.type !== 'min' && value.type !== 'max'" :value="String(value.value ?? '')" :aria-label="`${label}值`" :inputmode="value.type === 'formula' ? undefined : 'decimal'" @update:value="changeValue" />
  </div>
</template>

<style scoped>
.conditional-threshold { display: grid; grid-template-columns: 52px minmax(80px, 1fr) minmax(0, 1fr); align-items: center; gap: 8px; font-size: 12px; }
</style>
