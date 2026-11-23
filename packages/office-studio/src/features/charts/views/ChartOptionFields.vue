<script setup lang="ts">
import { ref } from 'vue'
import { Alert, ColorPicker, FormItem, Input, InputNumber, Select, Switch } from 'antdv-next'
import type { ChartConfig } from '../chart-engine'
import type { ChartSettingControl } from '../chart-settings-controls'
import { cloneChartDraftConfig, type ChartSettingsField } from '../chart-settings-draft'

const props = defineProps<{ config: ChartConfig; controls: readonly ChartSettingControl[]; fields?: ChartSettingsField[]; categories?: string[]; fieldCount?: number; disabled?: boolean }>()
const emit = defineEmits<{ 'update:config': [value: ChartConfig] }>()
const error = ref('')

function choices(control: ChartSettingControl): ChartSettingsField[] {
  return control.kind === 'points' ? (props.categories ?? []).slice(0, 16_384).map((label, value) => ({ label: `${value + 1} · ${label}`, value })) : props.fields ?? []
}
function numericValue(control: ChartSettingControl): number | undefined {
  if (control.kind !== 'number' && control.kind !== 'field') return undefined
  return control.read(props.config) ?? (control.kind === 'number' ? control.fallback : undefined)
}
function textValue(control: ChartSettingControl): string | undefined { return control.kind === 'text' || control.kind === 'color' || control.kind === 'select' ? control.read(props.config) : undefined }
function fieldList(control: ChartSettingControl): number[] { return control.kind === 'fields' || control.kind === 'points' ? control.read(props.config) ?? [] : [] }
function setValue(control: ChartSettingControl, value: unknown): void {
  const config = cloneChartDraftConfig(props.config)
  error.value = ''
  if (control.kind === 'switch') control.write(config, value === true)
  else if (control.kind === 'number' || control.kind === 'field') {
    const parsed = value === null || value === undefined || value === '' ? undefined : typeof value === 'number' ? value : typeof value === 'string' ? Number(value) : NaN
    if (parsed !== undefined && !Number.isFinite(parsed)) { error.value = '请输入有限数值。'; return }
    control.write(config, parsed)
  } else if (control.kind === 'fields' || control.kind === 'points') {
    const entries: unknown[] = Array.isArray(value) ? value : [], values: number[] = []
    for (const entry of entries) {
      if (typeof entry !== 'number' || !Number.isSafeInteger(entry) || entry < 0) { error.value = '请选择有效的字段或项目。'; return }
      values.push(entry)
    }
    control.write(config, values)
  } else control.write(config, typeof value === 'string' && value !== '' ? value : undefined)
  emit('update:config', config)
}
function setIndex(control: ChartSettingControl, value: unknown): void {
  setValue(control, typeof value === 'number' ? value - 1 : typeof value === 'string' && value !== '' ? Number(value) - 1 : undefined)
}
function setIndexes(control: ChartSettingControl, value: unknown): void {
  if (typeof value !== 'string') return
  const values = value.trim() ? value.split(/[,，\s]+/).filter(Boolean).map(index => Number(index) - 1) : []
  const limit = control.kind === 'points' ? props.categories?.length : props.fieldCount
  if (values.some(index => !Number.isSafeInteger(index) || index < 0 || limit !== undefined && index >= limit)) { error.value = '字段或项目序号从 1 开始，须位于当前数据范围内。'; return }
  setValue(control, values)
}
</script>

<template>
  <div class="chart-option-fields">
    <FormItem v-for="control in controls" :key="control.key" :label="control.label" :extra="control.extra">
      <Switch v-if="control.kind === 'switch'" :checked="control.read(config) ?? control.fallback ?? false" :aria-label="control.label" :disabled="disabled" @update:checked="setValue(control, $event)" />
      <InputNumber v-else-if="control.kind === 'number'" :value="numericValue(control)" :min="control.min" :max="control.max" :step="control.step ?? (control.max === 1 ? 0.05 : 1)" :aria-label="control.label" :disabled="disabled" @update:value="setValue(control, $event)" />
      <ColorPicker v-else-if="control.kind === 'color'" :value="textValue(control)" format="hex" value-format="hex" allow-clear show-text :aria-label="control.label" :disabled="disabled" @update:value="setValue(control, $event)" @clear="setValue(control, undefined)" />
      <Input v-else-if="control.kind === 'text'" :value="textValue(control)" allow-clear :aria-label="control.label" :disabled="disabled" @update:value="setValue(control, $event)" />
      <Select v-else-if="control.kind === 'select'" :value="textValue(control) ?? control.fallback" :options="control.choices" allow-clear :aria-label="control.label" :disabled="disabled" @update:value="setValue(control, $event)" />
      <Select v-else-if="control.kind === 'field' && choices(control).length" :value="numericValue(control)" :options="choices(control)" allow-clear show-search :aria-label="control.label" :disabled="disabled" @update:value="setValue(control, $event)" />
      <InputNumber v-else-if="control.kind === 'field'" :value="numericValue(control) === undefined ? undefined : numericValue(control)! + 1" :min="1" :max="fieldCount" :precision="0" :aria-label="control.label" placeholder="字段序号，从 1 开始" :disabled="disabled" @update:value="setIndex(control, $event)" />
      <Select v-else-if="choices(control).length" :value="fieldList(control)" :options="choices(control)" mode="multiple" show-search :aria-label="control.label" :disabled="disabled" @update:value="setValue(control, $event)" />
      <Input v-else :value="fieldList(control).map(index => index + 1).join(', ')" :aria-label="control.label" placeholder="序号从 1 开始，用逗号分隔" :disabled="disabled" @update:value="setIndexes(control, $event)" />
    </FormItem>
    <Alert v-if="error" type="error" :message="error" show-icon class="chart-field-error" />
  </div>
</template>

<style scoped>
.chart-option-fields { display: grid; grid-template-columns: repeat(auto-fit, minmax(175px, 1fr)); gap: 8px 14px; }
.chart-option-fields :deep(.ant-form-item) { min-width: 0; margin-bottom: 6px; }
.chart-option-fields :deep(.ant-input-number), .chart-option-fields :deep(.ant-select) { width: 100%; }
.chart-field-error { grid-column: 1 / -1; }
</style>
