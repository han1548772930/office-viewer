<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { Alert, Button, Collapse, CollapsePanel, FormItem, Select } from 'antdv-next'
import type { ChartConfig, ChartData } from '../chart-engine'
import { cloneChartDraftConfig } from '../chart-settings-draft'
import { chartParetoSettingsCategories } from '../chart-pareto'
import { chartParetoBarControls, chartParetoDataPointControls, chartParetoLineControls, type ChartParetoStyleGroup } from '../chart-pareto-settings-controls'
import ChartOptionFields from './ChartOptionFields.vue'

const props = defineProps<{ config: ChartConfig; data?: ChartData; disabled?: boolean }>()
const emit = defineEmits<{ 'update:config': [value: ChartConfig] }>()
const selectedPoints = ref<Record<ChartParetoStyleGroup, number | undefined>>({ barStyle: undefined, cumulativeLineStyle: undefined })
const groups = [{ key: 'barStyle', label: '帕累托柱形样式', pointLabel: '选择帕累托柱形数据项目' },
  { key: 'cumulativeLineStyle', label: '帕累托累计线样式', pointLabel: '选择帕累托累计线数据项目' }] as const
const projection = computed(() => {
  try { return { categories: chartParetoSettingsCategories(props.data, props.config.options?.pareto?.includeZeroValues, props.config), error: '' } }
  catch (cause) { return { categories: [], error: cause instanceof Error ? cause.message : '无法生成帕累托分类。' } }
})
const pointOptions = computed(() => projection.value.categories.map((category, value) => ({ value, label: `${value + 1} · ${category}` })))
watch(pointOptions, options => {
  for (const group of groups) if (selectedPoints.value[group.key] !== undefined && !options.some(option => option.value === selectedPoints.value[group.key])) selectedPoints.value[group.key] = undefined
})
const barControls = computed(() => chartParetoBarControls(props.data?.series[0]))
const lineControls = chartParetoLineControls()
function pointControls(group: ChartParetoStyleGroup) {
  const index = selectedPoints.value[group]
  return index === undefined ? [] : chartParetoDataPointControls(group, index, props.data?.series[0])
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
  <Collapse class="chart-pareto-settings">
    <CollapsePanel v-for="group in groups" :key="group.key" :header="group.label">
      <ChartOptionFields :config="config" :controls="group.key === 'barStyle' ? barControls : lineControls" :disabled="disabled" @update:config="emit('update:config', $event)" />
      <Button :disabled="disabled || config.options?.pareto?.[group.key] === undefined" :aria-label="`${group.label}恢复默认`" @click="resetStyle(group.key)">恢复{{ group.label }}默认值</Button>
      <FormItem :label="group.pointLabel" class="chart-pareto-point-selector"><Select v-model:value="selectedPoints[group.key]" :options="pointOptions" show-search allow-clear :aria-label="group.pointLabel" :disabled="disabled" /></FormItem>
      <ChartOptionFields :config="config" :controls="pointControls(group.key)" :disabled="disabled" @update:config="emit('update:config', $event)" />
      <Button v-if="selectedPoints[group.key] !== undefined" :disabled="disabled || config.options?.pareto?.[group.key]?.dataPoints?.[selectedPoints[group.key]!] === undefined" :aria-label="`${group.label}项目恢复默认`" @click="resetPoint(group.key)">恢复项目默认样式</Button>
    </CollapsePanel>
  </Collapse>
  <Alert v-if="projection.error" type="warning" :message="projection.error" show-icon />
</template>

<style scoped>
.chart-pareto-settings { margin-top: 12px; }
.chart-pareto-point-selector { margin-top: 14px; }
</style>
