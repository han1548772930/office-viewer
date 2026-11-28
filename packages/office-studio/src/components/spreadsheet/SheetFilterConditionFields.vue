<script setup lang="ts">
import { Checkbox, FormItem, Input, Select } from 'antdv-next'
import { computed } from 'vue'
import { SHEET_FILTER_CONDITION_OPTIONS, type SheetFilterConditionDraft } from '../../features/sheet-sort-filter/sheet-filter-condition-draft.ts'

const props = defineProps<{ draft: SheetFilterConditionDraft; index: number; disabled?: boolean; blocked?: boolean }>()
const emit = defineEmits<{ 'update:draft': [draft: SheetFilterConditionDraft] }>()
const controlsDisabled = computed(() => props.disabled || props.blocked)
const operators = [{ label: '无条件', value: 'none' }, ...SHEET_FILTER_CONDITION_OPTIONS]
const needsValue = computed(() => !['none', 'isBlank', 'isNotBlank', 'isError'].includes(props.draft.operator))
const needsSecond = computed(() => props.draft.operator === 'between' || props.draft.operator === 'notBetween')
function update<K extends keyof SheetFilterConditionDraft>(key: K, value: SheetFilterConditionDraft[K]): void {
  if (!props.disabled) emit('update:draft', { ...props.draft, [key]: value })
}
function changeOperator(value: unknown): void {
  if (controlsDisabled.value) return
  if (value === 'none') update('operator', 'none')
  const found = SHEET_FILTER_CONDITION_OPTIONS.find(option => option.value === value)
  if (found) update('operator', found.value)
}
function changeValueType(value: unknown): void {
  if (!controlsDisabled.value && (value === 'text' || value === 'number' || value === 'boolean')) update('valueType', value)
}
</script>

<template>
  <div class="filter-condition-fields">
    <FormItem :label="`条件 ${index}`"><Select :value="draft.operator" :aria-label="`自定义条件 ${index}`"
      :options="operators" :disabled="controlsDisabled" @change="changeOperator" /></FormItem>
    <FormItem v-if="needsValue" :label="`比较值 ${index} 类型`"><Select :value="draft.valueType" :aria-label="`自定义比较值 ${index} 类型`"
      :options="[{ label: '文本', value: 'text' }, { label: '数值或日期序号', value: 'number' }, { label: '逻辑值', value: 'boolean' }]"
      :disabled="controlsDisabled" @change="changeValueType" /></FormItem>
    <FormItem v-if="needsValue" :label="`比较值 ${index}`"><Input :value="draft.value" :aria-label="`自定义比较值 ${index}`"
      :disabled="disabled" @update:value="update('value', $event)" /></FormItem>
    <FormItem v-if="needsSecond" :label="`第二比较值 ${index}`"><Input :value="draft.secondValue" :aria-label="`自定义第二比较值 ${index}`"
      :disabled="disabled" @update:value="update('secondValue', $event)" /></FormItem>
    <FormItem v-if="needsValue" :label="null"><Checkbox :checked="draft.caseSensitive" :disabled="controlsDisabled"
      :aria-label="`自定义条件 ${index} 区分大小写`" @update:checked="update('caseSensitive', $event)">区分大小写</Checkbox></FormItem>
  </div>
</template>

<style scoped>
.filter-condition-fields { display: grid; gap: 10px; min-width: 0; }
</style>
