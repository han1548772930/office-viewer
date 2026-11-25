<script setup lang="ts">
import { computed, ref } from 'vue'
import { Alert, FormItem, InputNumber, Select } from 'antdv-next'
import type { ChartChordMapping } from '../../../domain/model'
import type { ChartConfig } from '../chart-engine'
import { cloneChartDraftConfig, type ChartSettingsField } from '../chart-settings-draft'
import { chartChordSettingsMapping, chartChordSettingsMappingError, setChartChordSettingsField } from '../chart-chord-settings-controls'

const props = defineProps<{ config: ChartConfig; fields: ChartSettingsField[]; numericFields?: ChartSettingsField[]; fieldCount: number; disabled?: boolean }>()
const emit = defineEmits<{ 'update:config': [value: ChartConfig] }>()
const mapping = computed(() => chartChordSettingsMapping(props.config, props.fieldCount)), error = ref('')
const diagnostic = computed(() => chartChordSettingsMappingError(props.config, props.fieldCount))
function endChoices(role: 'sourceIndex' | 'targetIndex'): ChartSettingsField[] {
  const current = mapping.value[role], choices = [...props.fields]
  if (!choices.some(field => field.value === current)) choices.unshift({ value: current, label: `字段 ${current + 1}（当前数据中不存在）` })
  return choices
}
const weights = computed(() => {
  const choices: Array<{ value: number | 'count'; label: string; disabled?: boolean }> = [{ value: 'count', label: '计数' }]
  for (const field of props.fields) if (field.value !== mapping.value.sourceIndex && field.value !== mapping.value.targetIndex) choices.push(field)
  const current = mapping.value.valueIndex
  if (current !== undefined && !choices.some(field => field.value === current)) {
    choices.push({ value: current, label: `字段 ${current + 1}（当前权重字段无效）`, disabled: true })
  }
  return choices
})
function weightMode(value: unknown): void {
  if (value === 'count') { change('valueIndex', undefined); return }
  if (value !== 'field') return
  let field = 0
  while (field < props.fieldCount && (field === mapping.value.sourceIndex || field === mapping.value.targetIndex)) field++
  if (field < props.fieldCount) change('valueIndex', field)
}
function change(role: keyof ChartChordMapping, value: unknown, oneBased = false): void {
  if (props.disabled) return
  error.value = ''
  const field = role === 'valueIndex' && (value === 'count' || value === undefined || value === null) ? undefined
    : typeof value === 'number' ? value - (oneBased ? 1 : 0) : NaN
  try {
    const config = cloneChartDraftConfig(props.config)
    setChartChordSettingsField(config, role, field, props.fieldCount)
    emit('update:config', config)
  } catch (cause) { error.value = cause instanceof Error ? cause.message : '无法更新和弦字段。' }
}
</script>

<template>
  <div class="chart-chord-mapping">
    <FormItem label="来源字段">
      <Select v-if="fields.length" :value="mapping.sourceIndex" :options="endChoices('sourceIndex')" show-search option-filter-prop="label" aria-label="和弦来源字段" :disabled="disabled" @update:value="change('sourceIndex', $event)" />
      <InputNumber v-else :value="mapping.sourceIndex + 1" :min="1" :max="fieldCount" :precision="0" aria-label="和弦来源字段" :disabled="disabled" @update:value="change('sourceIndex', $event, true)" />
    </FormItem>
    <FormItem label="目标字段" :validate-status="mapping.sourceIndex === mapping.targetIndex ? 'error' : undefined" :help="mapping.sourceIndex === mapping.targetIndex ? '来源与目标须使用不同字段。' : undefined">
      <Select v-if="fields.length" :value="mapping.targetIndex" :options="endChoices('targetIndex')" show-search option-filter-prop="label" aria-label="和弦目标字段" :disabled="disabled" @update:value="change('targetIndex', $event)" />
      <InputNumber v-else :value="mapping.targetIndex + 1" :min="1" :max="fieldCount" :precision="0" aria-label="和弦目标字段" :disabled="disabled" @update:value="change('targetIndex', $event, true)" />
    </FormItem>
    <FormItem label="权重字段">
      <Select v-if="fields.length" :value="mapping.valueIndex ?? 'count'" :options="weights" show-search option-filter-prop="label" aria-label="和弦权重字段" :disabled="disabled" @update:value="change('valueIndex', $event)" />
      <template v-else>
        <Select :value="mapping.valueIndex === undefined ? 'count' : 'field'" :options="[{ value: 'count', label: '计数' }, { value: 'field', label: '指定权重字段', disabled: fieldCount < 3 }]" aria-label="和弦权重字段" :disabled="disabled" @update:value="weightMode" />
        <InputNumber v-if="mapping.valueIndex !== undefined" :value="mapping.valueIndex + 1" :min="1" :max="fieldCount" :precision="0" aria-label="和弦权重字段序号" :disabled="disabled" @update:value="change('valueIndex', $event, true)" />
      </template>
    </FormItem>
    <Alert v-if="error || diagnostic" type="warning" :message="error || diagnostic" show-icon />
  </div>
</template>

<style scoped>
.chart-chord-mapping { display: grid; gap: 8px; min-width: 0; }
.chart-chord-mapping :deep(.ant-select), .chart-chord-mapping :deep(.ant-input-number) { width: 100%; }
</style>
