<script setup lang="ts">
import { computed, ref } from 'vue'
import { Alert, Button, Divider, FormItem, Select } from 'antdv-next'
import type { ChartColor } from '../../../domain/model'
import type { ChartConfig } from '../chart-engine'
import { cloneChartDraftConfig } from '../chart-settings-draft'
import { resolveChartSunburstLabel } from '../chart-sunburst'
import { chartSunburstBackgroundMode, chartSunburstCenterLabelControls, chartSunburstColorBackground, chartSunburstNodeLabelControls,
  chartSunburstPatternControls, clearChartSunburstSettings, setChartSunburstBackground, setChartSunburstBackgroundMode } from '../chart-sunburst-settings-controls'
import ChartColorInput from './ChartColorInput.vue'
import ChartOptionFields from './ChartOptionFields.vue'

const props = defineProps<{ config: ChartConfig; disabled?: boolean }>()
const emit = defineEmits<{ 'update:config': [value: ChartConfig] }>()
const error = ref(''), nodeControls = chartSunburstNodeLabelControls(), centerControls = chartSunburstCenterLabelControls()
const node = computed(() => resolveChartSunburstLabel(props.config)), backgroundMode = computed(() => chartSunburstBackgroundMode(props.config))
const colorBackground = computed(() => chartSunburstColorBackground(props.config)), patternControls = computed(() => chartSunburstPatternControls(props.config))
const backgroundModes = [{ value: 'color', label: '纯色或渐变' }, { value: 'pattern', label: '图案填充' }]
function edit(change: (config: ChartConfig) => void): void {
  if (props.disabled) return
  error.value = ''
  try { const config = cloneChartDraftConfig(props.config); change(config); emit('update:config', config) }
  catch (cause) { error.value = cause instanceof Error ? cause.message : '无法更新旭日图设置。' }
}
function backgroundType(value: unknown): void { if (value === 'color' || value === 'pattern') edit(config => setChartSunburstBackgroundMode(config, value)) }
function background(value: ChartColor | undefined): void { edit(config => setChartSunburstBackground(config, value)) }
function reset(target: 'label' | 'centerLabel' | 'all'): void { edit(config => clearChartSunburstSettings(config, target)) }
</script>

<template>
  <div class="chart-sunburst-settings">
    <ChartOptionFields :config="config" :controls="centerControls" :disabled="disabled" @update:config="emit('update:config', $event)" />
    <Button :disabled="disabled || config.options?.sunburst?.centerLabel === undefined" aria-label="恢复旭日图中心标签默认值" @click="reset('centerLabel')">恢复中心标签默认值</Button>
    <Divider />
    <ChartOptionFields :config="config" :controls="nodeControls" :disabled="disabled" @update:config="emit('update:config', $event)" />
    <template v-if="node.visible">
      <FormItem label="节点标签背景类型"><Select :value="backgroundMode" :options="backgroundModes" aria-label="节点标签背景类型" :disabled="disabled" @update:value="backgroundType" /></FormItem>
      <FormItem v-if="backgroundMode === 'color'" label="节点标签背景"><ChartColorInput :value="colorBackground" label="节点标签背景" :gradient-allowed="true" :disabled="disabled" @update:value="background" /></FormItem>
      <ChartOptionFields v-else :config="config" :controls="patternControls" :disabled="disabled" @update:config="emit('update:config', $event)" />
    </template>
    <div class="chart-sunburst-reset-row">
      <Button :disabled="disabled || config.options?.sunburst?.label === undefined" aria-label="恢复旭日图节点标签默认值" @click="reset('label')">恢复节点标签默认值</Button>
      <Button :disabled="disabled || config.options?.sunburst === undefined" aria-label="恢复全部旭日图默认设置" @click="reset('all')">恢复全部旭日图默认设置</Button>
    </div>
    <Alert v-if="error" type="error" :message="error" show-icon />
  </div>
</template>

<style scoped>
.chart-sunburst-settings { display: grid; gap: 12px; margin-top: 12px; min-width: 0; }
.chart-sunburst-settings :deep(.ant-select) { width: 100%; }
.chart-sunburst-settings :deep(.ant-divider) { margin-block: 4px; }
.chart-sunburst-reset-row { display: flex; flex-wrap: wrap; gap: 8px; }
</style>
