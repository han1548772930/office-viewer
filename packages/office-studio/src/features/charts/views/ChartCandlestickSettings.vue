<script setup lang="ts">
import { computed, ref } from 'vue'
import { Button, FormItem, Select } from 'antdv-next'
import type { ChartConfig, ChartData } from '../chart-engine'
import { chartCandlestickMode, CHART_CANDLESTICK_STATE_CHOICES, type ChartCandlestickState } from '../chart-candlestick'
import { chartCandlestickBodyControls, chartCandlestickBorderControls, chartCandlestickCloseMarkerControls,
  chartCandlestickHighLowControls, chartCandlestickWidthControls, clearChartCandlestickSettings, type ChartCandlestickResetTarget } from '../chart-candlestick-settings-controls'
import { cloneChartDraftConfig } from '../chart-settings-draft'
import ChartOptionFields from './ChartOptionFields.vue'

const props = defineProps<{ config: ChartConfig; data?: ChartData; disabled?: boolean }>()
const emit = defineEmits<{ 'update:config': [value: ChartConfig] }>()
// Choosing a state inspects one layer; it does not change the document/history.
const selectedState = ref<ChartCandlestickState>('rising')
const stateOptions = CHART_CANDLESTICK_STATE_CHOICES.map(([value, label]) => ({ value, label }))
const mode = computed(() => chartCandlestickMode(props.config, props.data))
const bodyControls = computed(() => chartCandlestickBodyControls(selectedState.value, props.data?.series[0]))
const borderControls = computed(() => chartCandlestickBorderControls(selectedState.value, props.data?.series[0]))
const widthControls = chartCandlestickWidthControls()
const highLowControls = computed(() => chartCandlestickHighLowControls(mode.value))
const closeMarkerControls = computed(() => chartCandlestickCloseMarkerControls(mode.value))
function selectState(value: unknown): void {
  if (value === 'rising' || value === 'falling' || value === 'doji') selectedState.value = value
}
function reset(target: ChartCandlestickResetTarget): void {
  if (props.disabled) return
  const config = cloneChartDraftConfig(props.config)
  clearChartCandlestickSettings(config, target)
  emit('update:config', config)
}
</script>

<template>
  <div class="chart-candlestick-settings">
    <div v-if="mode === 'ohlc'" class="chart-candlestick-state-settings">
      <div class="chart-candlestick-state-row">
        <FormItem label="K 线状态"><Select :value="selectedState" :options="stateOptions" aria-label="K 线状态" :disabled="disabled" @update:value="selectState" /></FormItem>
        <ChartOptionFields :config="config" :controls="widthControls" :disabled="disabled" @update:config="emit('update:config', $event)" />
      </div>
      <ChartOptionFields v-if="bodyControls.length" :config="config" :controls="bodyControls.slice(0, 2)" :disabled="disabled" @update:config="emit('update:config', $event)" />
      <ChartOptionFields v-if="bodyControls.length" :config="config" :controls="bodyControls.slice(2)" :disabled="disabled" @update:config="emit('update:config', $event)" />
      <ChartOptionFields :config="config" :controls="borderControls" :disabled="disabled" @update:config="emit('update:config', $event)" />
      <Button :disabled="disabled || config.options?.candlestick?.[selectedState] === undefined" aria-label="恢复所选 K 线状态继承样式" @click="reset(selectedState)">恢复所选状态继承样式</Button>
    </div>
    <div class="chart-candlestick-overlay-settings">
      <ChartOptionFields :config="config" :controls="highLowControls.slice(0, 1)" :disabled="disabled" @update:config="emit('update:config', $event)" />
      <ChartOptionFields :config="config" :controls="highLowControls.slice(1)" :disabled="disabled" @update:config="emit('update:config', $event)" />
      <Button :disabled="disabled || config.options?.candlestick?.highLowLine === undefined" aria-label="恢复 K 线高低线默认值" @click="reset('highLowLine')">恢复高低线默认值</Button>
    </div>
    <div class="chart-candlestick-overlay-settings">
      <ChartOptionFields :config="config" :controls="closeMarkerControls.slice(0, 1)" :disabled="disabled" @update:config="emit('update:config', $event)" />
      <ChartOptionFields :config="config" :controls="closeMarkerControls.slice(1)" :disabled="disabled" @update:config="emit('update:config', $event)" />
      <Button :disabled="disabled || config.options?.candlestick?.closeMarker === undefined" aria-label="恢复 K 线收盘标记默认值" @click="reset('closeMarker')">恢复收盘标记默认值</Button>
    </div>
    <Button :disabled="disabled || config.options?.candlestick === undefined" aria-label="恢复全部 K 线默认设置" @click="reset('all')">恢复全部 K 线默认设置</Button>
  </div>
</template>

<style scoped>
.chart-candlestick-settings { display: flex; flex-direction: column; gap: 14px; margin-top: 12px; }
.chart-candlestick-state-settings, .chart-candlestick-overlay-settings { display: flex; flex-direction: column; gap: 8px; }
.chart-candlestick-state-row { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 14px; align-items: start; }
.chart-candlestick-state-row :deep(.ant-form-item) { min-width: 0; margin-bottom: 6px; }
.chart-candlestick-state-row :deep(.ant-select) { width: 100%; }
.chart-candlestick-state-row :deep(.chart-option-fields) { grid-template-columns: minmax(0, 1fr); }
</style>
