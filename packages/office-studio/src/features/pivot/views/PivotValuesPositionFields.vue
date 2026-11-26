<script setup lang="ts">
import { FormItem, Select } from 'antdv-next'
import { computed } from 'vue'
import type { PivotValuePosition } from '../../../domain/model.ts'
import { pivotValuesUiIndexOptions, type PivotValuesUiState } from '../services/pivot-values-position-ui.ts'

const props = defineProps<{ state: PivotValuesUiState; subjectLabel: '透视表' | '透视图'; disabled?: boolean }>()
const emit = defineEmits<{ change: [position: PivotValuePosition, index: number] }>()
const directionLabel = computed(() => props.subjectLabel === '透视表' ? '值字段排列' : '值字段放置方向')
const indexOptions = computed(() => pivotValuesUiIndexOptions(props.state))
const directionOptions = [{ value: 'column', label: '按列显示' }, { value: 'row', label: '按行显示' }, { value: 'none', label: '无 ΣValues 位置' }]
function direction(value: unknown): void {
  if (!props.disabled && (value === 'row' || value === 'column' || value === 'none')) emit('change', value, props.state.index)
}
function index(value: unknown): void {
  if (!props.disabled && props.state.position !== 'none' && typeof value === 'number' && indexOptions.value.some(option => option.value === value)) emit('change', props.state.position, value)
}
</script>

<template>
  <div class="pivot-values-position" :data-testid="subjectLabel === '透视表' ? 'pivot-table-values-position' : 'pivot-chart-values-position'">
    <FormItem :label="directionLabel"><Select :value="state.position" :options="directionOptions" :aria-label="directionLabel" :disabled="disabled" @change="direction" /></FormItem>
    <FormItem label="ΣValues 插入位置"><Select :value="state.index" :options="indexOptions" :aria-label="`${subjectLabel} ΣValues 插入位置`"
      :disabled="disabled || state.position === 'none'" @change="index" /></FormItem>
    <span v-if="state.position === 'none'" class="values-position-note">汇总字段按列显示。</span>
  </div>
</template>

<style scoped>
.pivot-values-position { display: flex; align-items: flex-start; flex-wrap: wrap; gap: 8px 12px; }
.pivot-values-position :deep(.ant-select) { min-width: 140px; }
.values-position-note { align-self: center; font-size: 12px; opacity: .75; }
</style>
