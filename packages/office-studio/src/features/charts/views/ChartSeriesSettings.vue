<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { Button, Collapse, CollapsePanel, FormItem, Select } from 'antdv-next'
import type { ChartConfig, ChartData } from '../chart-engine'
import { chartColorControl, chartFillPatternControls, chartSeriesControls, colorControl, fieldControl, numberControl, selectControl, seriesStyleGroup, type ChartSettingControl } from '../chart-settings-controls'
import { cloneChartDraftConfig, type ChartSettingsField } from '../chart-settings-draft'
import { patchChartSeriesOverride } from '../chart-cartesian-api'
import { chartSeriesOverride } from '../chart-series-selector'
import { chartSeriesSettingPoints, chartSeriesSettingTargets } from '../chart-series-settings'
import { chartPointStyle } from '../chart-derived'
import { chartPalette, chartSeriesStyle } from '../chart-options'
import ChartOptionFields from './ChartOptionFields.vue'

const props = defineProps<{ config: ChartConfig; fields: ChartSettingsField[]; sourceFields: ChartSettingsField[]; categories: string[]; data?: ChartData; combination?: boolean; secondaryAxis?: boolean; horizontal?: boolean; scatter?: boolean; disabled?: boolean }>()
const emit = defineEmits<{ 'update:config': [value: ChartConfig] }>()
const selectedTarget = ref<string>(), selectedPoint = ref<number>()
const targets = computed(() => chartSeriesSettingTargets(props.config, props.fields, props.data))
watch(targets, items => { if (!items.some(item => item.value === selectedTarget.value)) selectedTarget.value = items.find(item => !item.disabled)?.value }, { immediate: true })
watch(selectedTarget, () => { selectedPoint.value = undefined })
const target = computed(() => targets.value.find(item => item.value === selectedTarget.value))
const controls = computed(() => target.value === undefined ? [] : chartSeriesControls(target.value.selector, props.combination, props.secondaryAxis, props.horizontal, target.value.source))
const pointOptions = computed(() => target.value === undefined ? [] : chartSeriesSettingPoints(props.config, target.value.selector, props.categories, props.fields, props.data))
watch(pointOptions, items => { if (selectedPoint.value !== undefined && !items.some(item => item.value === selectedPoint.value)) selectedPoint.value = undefined })
const pointControls = computed<ChartSettingControl[]>(() => {
  const selector = target.value?.selector, index = selectedPoint.value
  if (selector === undefined || index === undefined) return []
  const read = (config: ChartConfig) => chartSeriesOverride(config.options, selector)?.dataPoints?.[index]
  const write = (config: ChartConfig) => (seriesStyleGroup(config, selector).dataPoints ??= {})[index] ??= {}
  const effective = (config: ChartConfig) => chartPointStyle(chartSeriesStyle(config, typeof selector === 'number' ? selector : target.value?.source?.field ?? 0,
    target.value?.source?.position ?? 0, chartPalette(config), typeof selector === 'object' ? selector.name : undefined), index, config.type)
  return [
    chartColorControl('point.color', '项目颜色', config => read(config)?.color, (config, value) => { const point = write(config); point.color = value; point.fillPattern = null }, 'dataPoint', config => effective(config)?.color),
    ...chartFillPatternControls('point', read, write, effective, '项目'),
    numberControl('point.fillOpacity', '项目填充不透明度', config => read(config)?.fillOpacity, (config, value) => { write(config).fillOpacity = value }, 0, 1),
    numberControl('point.size', '项目标记大小', config => read(config)?.size, (config, value) => { write(config).size = value }, 0, 256),
    selectControl('point.shape', '项目标记形状', [['circle', '圆形'], ['rect', '方形'], ['triangle', '三角形'], ['diamond', '菱形'], ['none', '无标记']], config => read(config)?.shape, (config, value) => { write(config).shape = value }),
    colorControl('point.border.color', '项目边框颜色', config => read(config)?.border?.color, (config, value) => { (write(config).border ??= {}).color = value }),
    numberControl('point.border.width', '项目边框宽度', config => read(config)?.border?.width, (config, value) => { (write(config).border ??= {}).width = value }, 0, 30),
  ]
})
const scatterControl = computed(() => {
  const field = target.value?.selector
  if (typeof field !== 'number' || !props.scatter) return []
  return [fieldControl('seriesToCategory', '该系列的 X 字段', config => config.seriesToCategory?.[field]?.[0] ?? config.categoryField,
    (config, value) => { const mapping = config.seriesToCategory ??= {}; if (value === undefined) delete mapping[field]; else mapping[field] = [value] })]
})
function change(config: ChartConfig): void { emit('update:config', config) }
function resetSeries(): void {
  const config = cloneChartDraftConfig(props.config)
  if (target.value) config.options = { ...config.options, ...patchChartSeriesOverride(config.options, target.value.selector, undefined) }
  change(config)
}
function resetPoint(): void {
  const config = cloneChartDraftConfig(props.config)
  if (target.value && selectedPoint.value !== undefined) {
    const points = chartSeriesOverride(config.options, target.value.selector)?.dataPoints
    if (points) delete points[selectedPoint.value]
  }
  change(config)
}
</script>

<template>
  <FormItem label="选择系列" extra="字段样式应用到该字段的全部分组，名称样式可单独设置一组。"><Select v-model:value="selectedTarget" :options="targets" show-search option-filter-prop="label" aria-label="选择图表系列" :disabled="disabled" /></FormItem>
  <ChartOptionFields :config="config" :controls="scatterControl" :fields="sourceFields" :disabled="disabled" @update:config="change" />
  <ChartOptionFields :config="config" :controls="controls" :disabled="disabled || target?.ambiguous" @update:config="change" />
  <Button v-if="target" :disabled="disabled" @click="resetSeries">恢复该系列默认样式</Button>
  <Collapse class="chart-point-settings">
    <CollapsePanel key="point" header="单个数据项目">
      <FormItem label="选择数据项目"><Select v-model:value="selectedPoint" :options="pointOptions" show-search allow-clear aria-label="选择数据项目" :disabled="disabled || target?.ambiguous" /></FormItem>
      <ChartOptionFields :config="config" :controls="pointControls" :disabled="disabled || target?.ambiguous" @update:config="change" />
      <Button v-if="selectedPoint !== undefined" :disabled="disabled || target?.ambiguous" @click="resetPoint">恢复项目默认样式</Button>
    </CollapsePanel>
  </Collapse>
</template>

<style scoped>
.chart-point-settings { margin-top: 14px; }
</style>
