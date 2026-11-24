<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { Alert, Button, Collapse, CollapsePanel, FormItem, InputNumber, Select } from 'antdv-next'
import type { ChartWaterfallStyleTarget } from '../../../domain/model'
import type { ChartConfig, ChartData } from '../chart-engine'
import { cloneChartDraftConfig, type ChartSettingsField } from '../chart-settings-draft'
import { chartSeriesSettingTargets } from '../chart-series-settings'
import { chartSeriesOverride } from '../chart-series-selector'
import { patchChartWaterfallStyle } from '../chart-waterfall'
import { CHART_WATERFALL_POINT_CHOICE_LIMIT, chartWaterfallConnectorControls, chartWaterfallLayoutControls, chartWaterfallPointControls,
  chartWaterfallPointRoleControls, chartWaterfallStyleControls, chartWaterfallStylePointIsActive, chartWaterfallStylePoints } from '../chart-waterfall-settings-controls'
import ChartOptionFields from './ChartOptionFields.vue'

const props = defineProps<{ config: ChartConfig; fields: ChartSettingsField[]; categories: string[]; data?: ChartData; disabled?: boolean }>()
const emit = defineEmits<{ 'update:config': [value: ChartConfig] }>()
const selectedTarget = ref('all'), selectedRole = ref<ChartWaterfallStyleTarget>('positive'), selectedPoint = ref<number>()
const selectedSourceField = ref<number>(), selectedSourcePoint = ref<number>()
const stacked = computed(() => props.config.options?.waterfall?.stackType === 'stacked')
const globalSubtotal = computed(() => stacked.value && selectedRole.value === 'subtotal')
const seriesTargets = computed(() => chartSeriesSettingTargets(props.config, props.fields, props.data))
const targets = computed(() => [{ value: 'all', label: globalSubtotal.value ? '全局合计共享样式' : '所有来源共享样式', disabled: false },
  ...seriesTargets.value.map(target => ({ ...target, disabled: globalSubtotal.value || target.disabled }))])
watch(targets, items => { if (!items.some(item => item.value === selectedTarget.value)) selectedTarget.value = 'all' }, { immediate: true })
watch(globalSubtotal, global => { if (global) selectedTarget.value = 'all' }, { immediate: true })
const target = computed(() => seriesTargets.value.find(item => item.value === selectedTarget.value))
const selector = computed(() => target.value?.selector)
const source = computed(() => {
  if (target.value?.source) return { ...target.value.source, sourceName: target.value.source.name }
  if (typeof selector.value !== 'number') return undefined
  const position = props.data?.series.findIndex(series => series.field === selector.value) ?? -1
  const series = position < 0 ? undefined : props.data?.series[position]
  return series ? { field: series.field, position, name: series.name, sourceName: series.sourceName ?? series.name } : undefined
})
const disabled = computed(() => props.disabled || target.value?.disabled || target.value?.ambiguous)
const controls = computed(() => chartWaterfallStyleControls(selectedRole.value, selector.value, source.value))
const points = computed(() => chartWaterfallStylePoints(props.config, selectedRole.value, selector.value, props.data))
const rowCount = computed(() => props.data?.categories.length ?? 0)
const largeSource = computed(() => rowCount.value > CHART_WATERFALL_POINT_CHOICE_LIMIT)
const pointActive = computed(() => selectedPoint.value !== undefined && chartWaterfallStylePointIsActive(props.config,
  selectedRole.value, selectedPoint.value, selector.value, props.data))
const pointControls = computed(() => selectedPoint.value === undefined || !pointActive.value ? []
  : chartWaterfallPointControls(selectedRole.value, selectedPoint.value, selector.value, source.value))
watch([selectedTarget, selectedRole], () => { selectedPoint.value = undefined })
watch(points, () => { if (selectedPoint.value !== undefined && !pointActive.value) selectedPoint.value = undefined })
const sourceTargets = computed(() => props.data?.series.map(series => ({ value: series.field, label: series.sourceName ?? series.name })) ?? [])
watch(sourceTargets, items => { if (!items.some(item => item.value === selectedSourceField.value)) selectedSourceField.value = items[0]?.value }, { immediate: true })
const sourcePoints = computed(() => (props.data?.categories ?? []).slice(0, CHART_WATERFALL_POINT_CHOICE_LIMIT).map((category, value) => ({ value, label: `${value + 1} · ${category}` })))
const sourcePointControls = computed(() => stacked.value || selectedSourceField.value === undefined || selectedSourcePoint.value === undefined
  ? [] : chartWaterfallPointRoleControls(selectedSourceField.value, selectedSourcePoint.value))
watch(selectedSourceField, () => { selectedSourcePoint.value = undefined })
watch(rowCount, count => { if (selectedSourcePoint.value !== undefined && selectedSourcePoint.value >= count) selectedSourcePoint.value = undefined })
const roles = [{ value: 'positive', label: '增加 / 正值' }, { value: 'negative', label: '减少 / 负值' }, { value: 'subtotal', label: '小计 / 合计' }]
function change(config: ChartConfig): void { emit('update:config', config) }
function role(value: unknown): void { if (value === 'positive' || value === 'negative' || value === 'subtotal') selectedRole.value = value }
function setPointIndex(value: unknown, sourcePoint: boolean): void {
  const maximum = rowCount.value + Number(!sourcePoint && selectedRole.value === 'subtotal' && props.config.options?.waterfall?.useSubtotal !== false)
  const point = typeof value === 'number' && Number.isSafeInteger(value) && value >= 1 && value <= maximum ? value - 1 : undefined
  if (sourcePoint) selectedSourcePoint.value = point
  else selectedPoint.value = point
}
function resetRole(): void {
  if (disabled.value) return
  const next = cloneChartDraftConfig(props.config), current = selector.value
  if (current === undefined) {
    const styles = next.options?.waterfall?.styles
    if (styles) delete styles[selectedRole.value]
  } else next.options = { ...next.options, ...patchChartWaterfallStyle(next.options, current, selectedRole.value, undefined) }
  change(next)
}
function resetPoint(): void {
  if (disabled.value || selectedPoint.value === undefined) return
  const next = cloneChartDraftConfig(props.config)
  const group = selector.value === undefined ? next.options?.waterfall?.styles?.[selectedRole.value]
    : chartSeriesOverride(next.options, selector.value)?.waterfallStyles?.[selectedRole.value]
  if (group?.dataPoints) delete group.dataPoints[selectedPoint.value]
  change(next)
}
</script>

<template>
  <div class="chart-waterfall-settings">
    <ChartOptionFields :config="config" :controls="[...chartWaterfallLayoutControls(), ...chartWaterfallConnectorControls()]" :disabled="props.disabled" @update:config="change" />
    <FormItem label="角色样式作用系列"><Select v-model:value="selectedTarget" :options="targets" show-search option-filter-prop="label" aria-label="瀑布图角色样式作用系列" :disabled="props.disabled" /></FormItem>
    <FormItem label="瀑布图角色"><Select :value="selectedRole" :options="roles" aria-label="瀑布图角色" :disabled="props.disabled" @update:value="role" /></FormItem>
    <ChartOptionFields :config="config" :controls="controls" :disabled="disabled" @update:config="change" />
    <Button :disabled="disabled" @click="resetRole">恢复该角色继承样式</Button>
    <Alert v-if="globalSubtotal" type="info" message="堆积瀑布图的自动合计使用全局共享样式。" show-icon />
    <Collapse>
      <CollapsePanel key="points" header="角色内的单个数据项目">
        <FormItem label="选择角色数据项目" extra="自动合计由当前数据生成，使用合计角色样式。">
          <InputNumber v-if="largeSource && !globalSubtotal" :value="selectedPoint === undefined ? undefined : selectedPoint + 1" :min="1" :max="rowCount + Number(selectedRole === 'subtotal' && config.options?.waterfall?.useSubtotal !== false)" :precision="0" aria-label="瀑布图角色项目序号" placeholder="项目序号，从 1 开始" :disabled="disabled" @update:value="setPointIndex($event, false)" />
          <Select v-else v-model:value="selectedPoint" :options="points" allow-clear show-search option-filter-prop="label" aria-label="选择瀑布图角色数据项目" :disabled="disabled" />
        </FormItem>
        <Alert v-if="selectedPoint !== undefined && !pointActive" type="info" message="此项目当前属于其他角色，请切换角色或调整源项目含义。" show-icon />
        <ChartOptionFields :config="config" :controls="pointControls" :disabled="disabled" @update:config="change" />
        <Button v-if="pointActive" :disabled="disabled" @click="resetPoint">恢复角色项目继承样式</Button>
      </CollapsePanel>
      <CollapsePanel v-if="!stacked" key="roles" header="源项目的变化与合计">
        <FormItem label="来源系列"><Select v-model:value="selectedSourceField" :options="sourceTargets" show-search option-filter-prop="label" aria-label="瀑布图项目含义来源系列" :disabled="props.disabled" /></FormItem>
        <FormItem label="来源数据项目">
          <InputNumber v-if="largeSource" :value="selectedSourcePoint === undefined ? undefined : selectedSourcePoint + 1" :min="1" :max="rowCount" :precision="0" aria-label="瀑布图源项目序号" placeholder="项目序号，从 1 开始" :disabled="props.disabled" @update:value="setPointIndex($event, true)" />
          <Select v-else v-model:value="selectedSourcePoint" :options="sourcePoints" allow-clear show-search option-filter-prop="label" aria-label="瀑布图项目含义来源项目" :disabled="props.disabled" />
        </FormItem>
        <ChartOptionFields :config="config" :controls="sourcePointControls" :disabled="props.disabled" @update:config="change" />
      </CollapsePanel>
    </Collapse>
  </div>
</template>

<style scoped>
.chart-waterfall-settings { display: grid; gap: 10px; min-width: 0; }
</style>
