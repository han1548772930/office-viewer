<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { Button, Collapse, CollapsePanel, FormItem, InputNumber, Select } from 'antdv-next'
import type { ChartConfig, ChartData } from '../chart-engine'
import type { ChartComputedStatistics } from '../chart-derived'
import { cloneChartDraftConfig } from '../chart-settings-draft'
import { chartParetoSettingsCategories } from '../chart-pareto'
import { chartCategorySettingPoint, chartCategorySettingPoints } from '../chart-series-settings'
import { chartParetoBarControls, chartParetoDataPointControls, chartParetoLineControls, type ChartParetoStyleGroup } from '../chart-pareto-settings-controls'
import ChartOptionFields from './ChartOptionFields.vue'

const props = defineProps<{ config: ChartConfig; data?: ChartData; statistics?: ChartComputedStatistics; routeGroup?: ChartParetoStyleGroup; routeKey?: number; disabled?: boolean }>()
const emit = defineEmits<{ 'update:config': [value: ChartConfig] }>()
const openGroups = ref<string[]>([])
watch([() => props.routeGroup, () => props.routeKey], ([group]) => { if (group) openGroups.value = [group] }, { immediate: true })
function openKeys(value: unknown): string[] { return (Array.isArray(value) ? value : [value]).filter((key): key is string => typeof key === 'string') }
const selectedPoints = ref<Record<ChartParetoStyleGroup, number | undefined>>({ barStyle: undefined, cumulativeLineStyle: undefined })
const pointErrors = ref<Record<ChartParetoStyleGroup, string | undefined>>({ barStyle: undefined, cumulativeLineStyle: undefined })
const groups = [{ key: 'barStyle', label: '帕累托柱形样式', pointLabel: '选择帕累托柱形数据项目' },
  { key: 'cumulativeLineStyle', label: '帕累托累计线样式', pointLabel: '选择帕累托累计线数据项目' }] as const
const categories = computed(() => chartParetoSettingsCategories(props.statistics))
const pointOptions = computed(() => ({
  barStyle: chartCategorySettingPoints(categories.value, selectedPoints.value.barStyle),
  cumulativeLineStyle: chartCategorySettingPoints(categories.value, selectedPoints.value.cumulativeLineStyle),
}))
watch(categories, values => {
  for (const group of groups) {
    const selected = selectedPoints.value[group.key]
    if (selected !== undefined && !chartCategorySettingPoint(values, selected)) selectedPoints.value[group.key] = undefined
    pointErrors.value[group.key] = undefined
  }
})
const barControls = computed(() => chartParetoBarControls(props.data?.series[0]))
const lineControls = chartParetoLineControls()
function pointControls(group: ChartParetoStyleGroup) {
  const index = selectedPoints.value[group]
  return index === undefined ? [] : chartParetoDataPointControls(group, index, props.data?.series[0])
}
function selectPointNumber(group: ChartParetoStyleGroup, value: unknown): void {
  pointErrors.value[group] = undefined
  if (value === null || value === undefined || value === '') { selectedPoints.value[group] = undefined; return }
  const number = typeof value === 'number' ? value : typeof value === 'string' ? Number(value) : NaN, index = number - 1
  if (!chartCategorySettingPoint(categories.value, index)) {
    selectedPoints.value[group] = undefined
    pointErrors.value[group] = '请输入当前帕累托分类中的有效项目序号。'
    return
  }
  selectedPoints.value[group] = index
}
function resetStyle(group: ChartParetoStyleGroup): void {
  const config = cloneChartDraftConfig(props.config)
  ;((config.options ??= {}).pareto ??= {})[group] = undefined
  emit('update:config', config)
}
function resetPoint(group: ChartParetoStyleGroup): void {
  const index = selectedPoints.value[group]
  if (index === undefined) return
  const config = cloneChartDraftConfig(props.config), points = config.options?.pareto?.[group]?.dataPoints
  if (points) delete points[index]
  emit('update:config', config)
}
</script>

<template>
  <Collapse class="chart-pareto-settings" :active-key="openGroups" @update:active-key="openGroups = openKeys($event)">
    <CollapsePanel v-for="group in groups" :key="group.key" :header="group.label">
      <div :data-chart-pareto-group="group.key" :data-chart-edit-section="group.key === 'barStyle' ? 'paretoBarSeries' : 'paretoLineSeries'">
      <ChartOptionFields :config="config" :controls="group.key === 'barStyle' ? barControls : lineControls" :disabled="disabled" @update:config="emit('update:config', $event)" />
      <Button :disabled="disabled || config.options?.pareto?.[group.key] === undefined" :aria-label="`${group.label}恢复默认`" @click="resetStyle(group.key)">恢复{{ group.label }}默认值</Button>
      <FormItem :label="group.pointLabel" class="chart-pareto-point-selector"><Select v-model:value="selectedPoints[group.key]" :options="pointOptions[group.key]" show-search option-filter-prop="label" allow-clear :aria-label="group.pointLabel" :disabled="disabled" @update:value="pointErrors[group.key] = undefined" /></FormItem>
      <FormItem label="数据项目序号" extra="输入项目序号，可定位长列表中的数据项目。" :validate-status="pointErrors[group.key] ? 'error' : undefined" :help="pointErrors[group.key]"><InputNumber :value="selectedPoints[group.key] === undefined ? undefined : selectedPoints[group.key]! + 1" :min="1" :max="categories.length || undefined" :precision="0" :aria-label="`${group.pointLabel}序号`" :disabled="disabled || !categories.length" @update:value="selectPointNumber(group.key, $event)" /></FormItem>
      <ChartOptionFields :config="config" :controls="pointControls(group.key)" :disabled="disabled" @update:config="emit('update:config', $event)" />
      <Button v-if="selectedPoints[group.key] !== undefined" :disabled="disabled || config.options?.pareto?.[group.key]?.dataPoints?.[selectedPoints[group.key]!] === undefined" :aria-label="`${group.label}项目恢复默认`" @click="resetPoint(group.key)">恢复项目默认样式</Button>
      </div>
    </CollapsePanel>
  </Collapse>
</template>

<style scoped>
.chart-pareto-settings { margin-top: 12px; }
.chart-pareto-point-selector { margin-top: 14px; }
.chart-pareto-settings :deep(.ant-input-number) { width: 100%; }
</style>
