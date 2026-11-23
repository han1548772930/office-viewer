<script setup lang="ts">
import { Button, Collapse, CollapsePanel } from 'antdv-next'
import type { ChartConfig } from '../chart-engine'
import { cloneChartDraftConfig, type ChartSettingsField } from '../chart-settings-draft'
import { chartErrorBarControls, chartTrendlineControls } from '../chart-statistics-controls'
import ChartOptionFields from './ChartOptionFields.vue'

const props = defineProps<{ config: ChartConfig; seriesFields: ChartSettingsField[]; sourceFields: ChartSettingsField[]; numericX?: boolean; disabled?: boolean }>()
const emit = defineEmits<{ 'update:config': [value: ChartConfig] }>()
function change(value: ChartConfig): void { emit('update:config', value) }
function add(kind: 'trendlines' | 'errorBars'): void {
  const config = cloneChartDraftConfig(props.config), options = config.options ??= {}, field = props.seriesFields[0]?.value ?? config.valueFields?.[0] ?? 1
  if (kind === 'trendlines') (options.trendlines ??= []).push({ seriesIndex: field, type: 'linear' })
  else (options.errorBars ??= []).push({ seriesIndex: field, direction: 'y', type: 'both', valueType: 'fixedVal', value: 1 })
  change(config)
}
function remove(kind: 'trendlines' | 'errorBars', index: number): void {
  const config = cloneChartDraftConfig(props.config)
  config.options?.[kind]?.splice(index, 1)
  change(config)
}
</script>

<template>
  <div class="chart-statistics-settings">
    <div class="chart-statistics-actions">
      <Button :disabled="disabled" @click="add('trendlines')">添加趋势线</Button>
      <Button :disabled="disabled" @click="add('errorBars')">添加误差线</Button>
    </div>
    <Collapse>
      <CollapsePanel v-for="(_line, index) in config.options?.trendlines ?? []" :key="`trend-${index}`" :header="`趋势线 ${index + 1}`">
        <ChartOptionFields :config="config" :controls="chartTrendlineControls(config, index).filter(control => control.key.endsWith('seriesIndex'))" :fields="seriesFields" :disabled="disabled" @update:config="change" />
        <ChartOptionFields :config="config" :controls="chartTrendlineControls(config, index).filter(control => !control.key.endsWith('seriesIndex'))" :disabled="disabled" @update:config="change" />
        <Button danger :disabled="disabled" @click="remove('trendlines', index)">移除趋势线</Button>
      </CollapsePanel>
      <CollapsePanel v-for="(_error, index) in config.options?.errorBars ?? []" :key="`error-${index}`" :header="`误差线 ${index + 1}`">
        <ChartOptionFields :config="config" :controls="chartErrorBarControls(config, index, Boolean(numericX)).filter(control => control.key.endsWith('seriesIndex'))" :fields="seriesFields" :disabled="disabled" @update:config="change" />
        <ChartOptionFields :config="config" :controls="chartErrorBarControls(config, index, Boolean(numericX)).filter(control => !control.key.endsWith('seriesIndex'))" :fields="sourceFields" :disabled="disabled" @update:config="change" />
        <Button danger :disabled="disabled" @click="remove('errorBars', index)">移除误差线</Button>
      </CollapsePanel>
    </Collapse>
  </div>
</template>

<style scoped>
.chart-statistics-settings { display: grid; gap: 12px; }
.chart-statistics-actions { display: flex; gap: 8px; flex-wrap: wrap; }
</style>
