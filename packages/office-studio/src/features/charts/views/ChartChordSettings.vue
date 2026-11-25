<script setup lang="ts">
import { computed, ref } from 'vue'
import { Alert, Button, Divider } from 'antdv-next'
import type { ChartConfig } from '../chart-engine'
import { cloneChartDraftConfig } from '../chart-settings-draft'
import { canResetChartChordSettings, chartChordEmphasisControls, chartChordLabelControls, chartChordLayoutControls,
  chartChordNodeControls, chartChordRibbonControls, chartChordSettingsRadiiError, clearChartChordSettings,
  resetChartChordLabelFont, type ChartChordSettingsGroup } from '../chart-chord-settings-controls'
import ChartOptionFields from './ChartOptionFields.vue'

const props = defineProps<{ config: ChartConfig; disabled?: boolean }>()
const emit = defineEmits<{ 'update:config': [value: ChartConfig] }>()
const layout = chartChordLayoutControls(), label = chartChordLabelControls(), ribbon = chartChordRibbonControls()
const node = chartChordNodeControls(), emphasis = chartChordEmphasisControls(), error = ref('')
const radiiError = computed(() => chartChordSettingsRadiiError(props.config))
const labelsVisible = computed(() => props.config.options?.chord?.label?.visible !== false)
function edit(change: (config: ChartConfig) => void): void {
  if (props.disabled) return
  error.value = ''
  try { const config = cloneChartDraftConfig(props.config); change(config); emit('update:config', config) }
  catch (cause) { error.value = cause instanceof Error ? cause.message : '无法更新和弦图设置。' }
}
function reset(target: ChartChordSettingsGroup): void { edit(config => clearChartChordSettings(config, target)) }
function fontReset(target: 'color' | 'font'): void { edit(config => resetChartChordLabelFont(config, target)) }
</script>

<template>
  <div class="chart-chord-settings">
    <Divider title-placement="start">布局</Divider>
    <ChartOptionFields class="chart-chord-two-columns" :config="config" :controls="layout.slice(0, 2)" :disabled="disabled" @update:config="emit('update:config', $event)" />
    <ChartOptionFields class="chart-chord-one-column" :config="config" :controls="layout.slice(2, 3)" :disabled="disabled" @update:config="emit('update:config', $event)" />
    <Alert v-if="radiiError" type="warning" :message="radiiError" show-icon />
    <ChartOptionFields :config="config" :controls="layout.slice(3)" :disabled="disabled" @update:config="emit('update:config', $event)" />
    <Button :disabled="disabled || !canResetChartChordSettings(config, 'layout')" aria-label="恢复和弦布局默认值" @click="reset('layout')">恢复布局默认值</Button>
    <Divider title-placement="start">标签</Divider>
    <ChartOptionFields :config="config" :controls="label" :disabled="disabled" @update:config="emit('update:config', $event)" />
    <div v-if="labelsVisible" class="chart-chord-reset-row">
      <Button type="link" :disabled="disabled || config.options?.chord?.label?.color === undefined" aria-label="恢复和弦标签颜色" @click="fontReset('color')">恢复标签颜色</Button>
      <Button type="link" :disabled="disabled || !['fontSize', 'bold', 'italic', 'family'].some(key => Reflect.get(config.options?.chord?.label ?? {}, key) !== undefined)" aria-label="恢复和弦标签字体" @click="fontReset('font')">恢复标签字体</Button>
    </div>
    <Button :disabled="disabled || !canResetChartChordSettings(config, 'label')" aria-label="恢复和弦标签默认值" @click="reset('label')">恢复标签默认值</Button>
    <Divider title-placement="start">弦带</Divider>
    <ChartOptionFields class="chart-chord-two-columns" :config="config" :controls="ribbon" :disabled="disabled" @update:config="emit('update:config', $event)" />
    <Button :disabled="disabled || !canResetChartChordSettings(config, 'ribbon')" aria-label="恢复和弦弦带默认值" @click="reset('ribbon')">恢复弦带默认值</Button>
    <Divider title-placement="start">节点</Divider>
    <ChartOptionFields :config="config" :controls="node" :disabled="disabled" @update:config="emit('update:config', $event)" />
    <Button :disabled="disabled || !canResetChartChordSettings(config, 'node')" aria-label="恢复和弦节点默认值" @click="reset('node')">恢复节点默认值</Button>
    <Divider title-placement="start">悬停强调</Divider>
    <ChartOptionFields :config="config" :controls="emphasis" :disabled="disabled" @update:config="emit('update:config', $event)" />
    <div class="chart-chord-reset-row">
      <Button :disabled="disabled || !canResetChartChordSettings(config, 'emphasis')" aria-label="恢复和弦强调默认值" @click="reset('emphasis')">恢复强调默认值</Button>
      <Button :disabled="disabled || !canResetChartChordSettings(config, 'all')" aria-label="恢复全部和弦默认设置" @click="reset('all')">恢复全部和弦默认设置</Button>
    </div>
    <Alert v-if="error" type="error" :message="error" show-icon />
  </div>
</template>

<style scoped>
.chart-chord-settings { display: grid; gap: 12px; margin-top: 12px; min-width: 0; }
.chart-chord-settings :deep(.ant-divider) { margin-block: 4px; }
.chart-chord-two-columns { grid-template-columns: minmax(0, 1fr) minmax(0, 1fr); }
.chart-chord-one-column { grid-template-columns: minmax(0, 1fr); }
.chart-chord-reset-row { display: flex; flex-wrap: wrap; gap: 8px; }
</style>
