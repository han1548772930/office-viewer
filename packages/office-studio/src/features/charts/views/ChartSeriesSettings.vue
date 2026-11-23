<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { Button, Collapse, CollapsePanel, FormItem, Select } from 'antdv-next'
import type { ChartConfig } from '../chart-engine'
import { chartSeriesControls, colorControl, fieldControl, numberControl, seriesStyleGroup, type ChartSettingControl } from '../chart-settings-controls'
import { cloneChartDraftConfig, type ChartSettingsField } from '../chart-settings-draft'
import ChartOptionFields from './ChartOptionFields.vue'

const props = defineProps<{ config: ChartConfig; fields: ChartSettingsField[]; sourceFields: ChartSettingsField[]; categories: string[]; combination?: boolean; secondaryAxis?: boolean; scatter?: boolean; disabled?: boolean }>()
const emit = defineEmits<{ 'update:config': [value: ChartConfig] }>()
const selectedField = ref<number>(), selectedPoint = ref<number>()
watch(() => props.fields, fields => { if (!fields.some(field => field.value === selectedField.value)) selectedField.value = fields[0]?.value }, { immediate: true })
const controls = computed(() => selectedField.value === undefined ? [] : chartSeriesControls(selectedField.value, props.combination, props.secondaryAxis))
const pointOptions = computed(() => props.categories.slice(0, 16_384).map((category, value) => ({ value, label: `${value + 1} · ${category}` })))
const pointControls = computed<ChartSettingControl[]>(() => {
  const field = selectedField.value, index = selectedPoint.value
  if (field === undefined || index === undefined) return []
  const read = (config: ChartConfig) => config.options?.series?.[field]?.dataPoints?.[index]
  const write = (config: ChartConfig) => (seriesStyleGroup(config, field).dataPoints ??= {})[index] ??= {}
  return [
    colorControl('point.color', '项目颜色', config => read(config)?.color, (config, value) => { write(config).color = value }),
    numberControl('point.fillOpacity', '项目填充不透明度', config => read(config)?.fillOpacity, (config, value) => { write(config).fillOpacity = value }, 0, 1),
    numberControl('point.size', '项目标记大小', config => read(config)?.size, (config, value) => { write(config).size = value }, 0, 256),
    colorControl('point.border.color', '项目边框颜色', config => read(config)?.border?.color, (config, value) => { (write(config).border ??= {}).color = value }),
    numberControl('point.border.width', '项目边框宽度', config => read(config)?.border?.width, (config, value) => { (write(config).border ??= {}).width = value }, 0, 30),
  ]
})
const scatterControl = computed(() => {
  const field = selectedField.value
  if (field === undefined || !props.scatter) return []
  return [fieldControl('seriesToCategory', '该系列的 X 字段', config => config.seriesToCategory?.[field]?.[0] ?? config.categoryField,
    (config, value) => { const mapping = config.seriesToCategory ??= {}; if (value === undefined) delete mapping[field]; else mapping[field] = [value] })]
})
function change(config: ChartConfig): void { emit('update:config', config) }
function resetSeries(): void {
  const config = cloneChartDraftConfig(props.config)
  if (selectedField.value !== undefined && config.options?.series) delete config.options.series[selectedField.value]
  change(config)
}
function resetPoint(): void {
  const config = cloneChartDraftConfig(props.config)
  if (selectedField.value !== undefined && selectedPoint.value !== undefined) {
    const points = config.options?.series?.[selectedField.value]?.dataPoints
    if (points) delete points[selectedPoint.value]
  }
  change(config)
}
</script>

<template>
  <FormItem label="选择系列"><Select v-model:value="selectedField" :options="fields" show-search aria-label="选择图表系列" :disabled="disabled" /></FormItem>
  <ChartOptionFields :config="config" :controls="scatterControl" :fields="sourceFields" :disabled="disabled" @update:config="change" />
  <ChartOptionFields :config="config" :controls="controls" :disabled="disabled" @update:config="change" />
  <Button v-if="selectedField !== undefined" :disabled="disabled" @click="resetSeries">恢复该系列默认样式</Button>
  <Collapse class="chart-point-settings">
    <CollapsePanel key="point" header="单个数据项目">
      <FormItem label="选择数据项目"><Select v-model:value="selectedPoint" :options="pointOptions" show-search allow-clear aria-label="选择数据项目" :disabled="disabled" /></FormItem>
      <ChartOptionFields :config="config" :controls="pointControls" :disabled="disabled" @update:config="change" />
      <Button v-if="selectedPoint !== undefined" :disabled="disabled" @click="resetPoint">恢复项目默认样式</Button>
    </CollapsePanel>
  </Collapse>
</template>

<style scoped>
.chart-point-settings { margin-top: 14px; }
</style>
