<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { Alert, Button, FormItem, Select, Tooltip } from 'antdv-next'
import type { ChartConfig } from '../chart-engine'
import { cloneChartDraftConfig, type ChartSettingsField } from '../chart-settings-draft'
import { addChartHierarchyField, chartHierarchySettingsMapping, moveChartHierarchyField, removeChartHierarchyField,
  replaceChartHierarchyField, setChartHierarchyWeightField } from '../chart-hierarchy-settings'

const props = defineProps<{ config: ChartConfig; fields: ChartSettingsField[]; numericFields?: ChartSettingsField[]; fieldCount: number; disabled?: boolean }>()
const emit = defineEmits<{ 'update:config': [value: ChartConfig] }>()
const mapping = computed(() => chartHierarchySettingsMapping(props.config, props.fieldCount)), pendingField = ref<number>(), error = ref('')
const candidates = computed(() => props.fields.filter(field => field.value !== mapping.value.valueIndex && !mapping.value.hierarchyIndexes.includes(field.value)))
const weights = computed(() => (props.numericFields?.length ? props.numericFields : props.fields).filter(field => !mapping.value.hierarchyIndexes.includes(field.value)))
watch(candidates, choices => { if (!choices.some(choice => choice.value === pendingField.value)) pendingField.value = choices[0]?.value }, { immediate: true })
function rowChoices(position: number): ChartSettingsField[] {
  const current = mapping.value.hierarchyIndexes[position]
  const choices = props.fields.filter(field => field.value === current || field.value !== mapping.value.valueIndex && !mapping.value.hierarchyIndexes.includes(field.value))
  if (current !== undefined && !choices.some(choice => choice.value === current)) choices.unshift({ value: current, label: `字段 ${current + 1}（当前数据中不存在）` })
  return choices
}
function edit(change: (config: ChartConfig) => void): void {
  if (props.disabled) return
  error.value = ''
  try { const config = cloneChartDraftConfig(props.config); change(config); emit('update:config', config) }
  catch (cause) { error.value = cause instanceof Error ? cause.message : '无法更新层级字段。' }
}
function replace(position: number, value: unknown): void { if (typeof value === 'number') edit(config => replaceChartHierarchyField(config, position, value, props.fieldCount)) }
function weight(value: unknown): void { if (typeof value === 'number') edit(config => setChartHierarchyWeightField(config, value, props.fieldCount)) }
function pending(value: unknown): void { pendingField.value = typeof value === 'number' ? value : undefined }
function add(): void { if (pendingField.value !== undefined) edit(config => addChartHierarchyField(config, pendingField.value!, props.fieldCount)) }
function move(position: number, direction: -1 | 1): void { edit(config => moveChartHierarchyField(config, position, direction, props.fieldCount)) }
function remove(position: number): void { edit(config => removeChartHierarchyField(config, position, props.fieldCount)) }
</script>

<template>
  <div class="chart-hierarchy-mapping">
    <div v-for="(field, index) in mapping.hierarchyIndexes" :key="field" class="chart-hierarchy-field-row">
      <FormItem :label="`层级 ${index + 1}${index === 0 ? '（根）' : ''}`"><Select :value="field" :options="rowChoices(index)" :aria-label="`层级 ${index + 1} 字段`" :disabled="disabled" @update:value="replace(index, $event)" /></FormItem>
      <div class="chart-hierarchy-field-actions">
        <Tooltip title="上移一个层级"><Button size="small" :aria-label="`上移层级 ${index + 1}`" :disabled="disabled || index === 0" @click="move(index, -1)">上移</Button></Tooltip>
        <Tooltip title="下移一个层级"><Button size="small" :aria-label="`下移层级 ${index + 1}`" :disabled="disabled || index === mapping.hierarchyIndexes.length - 1" @click="move(index, 1)">下移</Button></Tooltip>
        <Tooltip :title="mapping.hierarchyIndexes.length <= 1 ? '至少保留一个层级字段' : '移除此层级字段'"><Button size="small" danger :aria-label="`移除层级 ${index + 1}`" :disabled="disabled || mapping.hierarchyIndexes.length <= 1" @click="remove(index)">移除</Button></Tooltip>
      </div>
    </div>
    <FormItem label="添加层级字段"><div class="chart-hierarchy-add-row"><Select :value="pendingField" :options="candidates" aria-label="添加层级字段" placeholder="选择下一层级" :disabled="disabled || candidates.length === 0" @update:value="pending" /><Button :disabled="disabled || pendingField === undefined" aria-label="添加层级" @click="add">添加层级</Button></div></FormItem>
    <FormItem label="权重字段"><Select :value="mapping.valueIndex" :options="weights" aria-label="权重字段" :disabled="disabled || weights.length === 0" @update:value="weight" /></FormItem>
    <Alert v-if="error" type="error" :message="error" show-icon />
  </div>
</template>

<style scoped>
.chart-hierarchy-mapping { display: grid; gap: 8px; min-width: 0; }
.chart-hierarchy-field-row { display: grid; grid-template-columns: minmax(130px, 1fr) auto; gap: 8px; align-items: end; }
.chart-hierarchy-field-row :deep(.ant-form-item) { margin-bottom: 0; }
.chart-hierarchy-field-actions, .chart-hierarchy-add-row { display: flex; gap: 6px; align-items: center; }
.chart-hierarchy-add-row > .ant-select { flex: 1; min-width: 0; }
.chart-hierarchy-mapping :deep(.ant-select) { width: 100%; }
@media (max-width: 600px) { .chart-hierarchy-field-row { grid-template-columns: 1fr; } }
</style>
