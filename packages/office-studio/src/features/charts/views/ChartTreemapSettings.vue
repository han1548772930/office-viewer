<script setup lang="ts">
import { computed, ref } from 'vue'
import { Alert, Button, FormItem, Select } from 'antdv-next'
import type { ChartColor } from '../../../domain/model'
import type { ChartConfig } from '../chart-engine'
import { cloneChartDraftConfig } from '../chart-settings-draft'
import { resolveChartTreemapLabel } from '../chart-treemap'
import { chartTreemapBackgroundMode, chartTreemapColorBackground, chartTreemapLeafControls, chartTreemapParentControls,
  chartTreemapPatternControls, clearChartTreemapSettings, setChartTreemapBackground, setChartTreemapBackgroundMode } from '../chart-treemap-settings-controls'
import ChartColorInput from './ChartColorInput.vue'
import ChartOptionFields from './ChartOptionFields.vue'

const props = defineProps<{ config: ChartConfig; disabled?: boolean }>()
const emit = defineEmits<{ 'update:config': [value: ChartConfig] }>()
const error = ref('')
const parentControls = chartTreemapParentControls(), leafControls = chartTreemapLeafControls()
const leaf = computed(() => resolveChartTreemapLabel(props.config))
const backgroundMode = computed(() => chartTreemapBackgroundMode(props.config))
const colorBackground = computed(() => chartTreemapColorBackground(props.config))
const patternControls = computed(() => chartTreemapPatternControls(props.config))
const backgroundModes = [{ value: 'color', label: '纯色或渐变' }, { value: 'pattern', label: '图案填充' }]
function edit(change: (config: ChartConfig) => void): void {
  if (props.disabled) return
  error.value = ''
  try { const config = cloneChartDraftConfig(props.config); change(config); emit('update:config', config) }
  catch (cause) { error.value = cause instanceof Error ? cause.message : '无法更新树图设置。' }
}
function backgroundType(value: unknown): void {
  if (value === 'color' || value === 'pattern') edit(config => setChartTreemapBackgroundMode(config, value))
}
function background(value: ChartColor | undefined): void { edit(config => setChartTreemapBackground(config, value)) }
function reset(target: 'label' | 'all'): void { edit(config => clearChartTreemapSettings(config, target)) }
</script>

<template>
  <div class="chart-treemap-settings">
    <ChartOptionFields :config="config" :controls="parentControls" :disabled="disabled" @update:config="emit('update:config', $event)" />
    <ChartOptionFields :config="config" :controls="leafControls" :disabled="disabled" @update:config="emit('update:config', $event)" />
    <template v-if="leaf.visible">
      <FormItem label="叶级标签背景类型"><Select :value="backgroundMode" :options="backgroundModes" aria-label="叶级标签背景类型" :disabled="disabled" @update:value="backgroundType" /></FormItem>
      <FormItem v-if="backgroundMode === 'color'" label="叶级标签背景"><ChartColorInput :value="colorBackground" label="叶级标签背景" :gradient-allowed="true" :disabled="disabled" @update:value="background" /></FormItem>
      <ChartOptionFields v-else :config="config" :controls="patternControls" :disabled="disabled" @update:config="emit('update:config', $event)" />
    </template>
    <div class="chart-treemap-reset-row">
      <Button :disabled="disabled || config.options?.treemap?.label === undefined" aria-label="恢复树图叶级标签默认值" @click="reset('label')">恢复叶级标签默认值</Button>
      <Button :disabled="disabled || config.options?.treemap === undefined" aria-label="恢复全部树图默认设置" @click="reset('all')">恢复全部树图默认设置</Button>
    </div>
    <Alert v-if="error" type="error" :message="error" show-icon />
  </div>
</template>

<style scoped>
.chart-treemap-settings { display: grid; gap: 12px; margin-top: 12px; min-width: 0; }
.chart-treemap-settings :deep(.ant-select) { width: 100%; }
.chart-treemap-reset-row { display: flex; flex-wrap: wrap; gap: 8px; }
</style>
