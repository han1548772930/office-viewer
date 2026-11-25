<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { Alert, Button, FormItem, InputNumber, Select } from 'antdv-next'
import type { ChartConfig, ChartData } from '../chart-engine'
import { chartPieSliceControls } from '../chart-pie-settings-controls'
import { chartPieSliceAllowsExplosion } from '../chart-pie'
import { resetChartSpecialMarkDraft } from '../chart-special-mark-edit'
import ChartOptionFields from './ChartOptionFields.vue'

const props = defineProps<{ config: ChartConfig; data?: ChartData; routePointIndex?: number; routeKey?: number; disabled?: boolean }>()
const emit = defineEmits<{ 'update:config': [value: ChartConfig] }>()
const selectedPoint = ref<number>(props.routePointIndex ?? 0)
const count = computed(() => props.data?.categories.length ?? 0)
const pointReady = computed(() => !!props.data?.series.length && selectedPoint.value >= 0 && selectedPoint.value < count.value)
const controls = computed(() => pointReady.value ? chartPieSliceControls(selectedPoint.value, props.data!.series[0],
  chartPieSliceAllowsExplosion(props.config, selectedPoint.value, props.data)) : [])
const choices = computed(() => (props.data?.categories ?? []).slice(0, 1000).map((category, value) => ({ value, label: `${value + 1} · ${category}` })))
watch([() => props.routePointIndex, () => props.routeKey], ([point]) => { if (point !== undefined) selectedPoint.value = point })
watch(count, value => { if (props.routePointIndex === undefined && selectedPoint.value >= value) selectedPoint.value = 0 })
function change(config: ChartConfig): void { emit('update:config', config) }
function index(value: unknown): void {
  if (typeof value === 'number' && Number.isSafeInteger(value) && value >= 1 && value <= count.value) selectedPoint.value = value - 1
}
function reset(): void {
  if (props.disabled || !pointReady.value) return
  const source = props.data!.series[0]!
  change(resetChartSpecialMarkDraft(props.config, { kind: 'pieSliceStyle', sourceField: source.field,
    sourceName: source.sourceName ?? source.name, sourcePointIndex: selectedPoint.value,
    plot: chartPieSliceAllowsExplosion(props.config, selectedPoint.value, props.data) ? 'primary' : 'secondary' }))
}
</script>

<template>
  <div class="chart-pie-slice-settings" :data-chart-pie-source-point="selectedPoint">
    <FormItem label="来源扇区" extra="设置按来源项目保存，合计扇区使用独立样式。">
      <InputNumber v-if="count > 1000" :value="selectedPoint + 1" :min="1" :max="count" :precision="0" aria-label="饼图来源扇区序号" :disabled="disabled || !count" @update:value="index" />
      <Select v-else v-model:value="selectedPoint" :options="choices" show-search option-filter-prop="label" aria-label="饼图来源扇区" :disabled="disabled || !count" />
    </FormItem>
    <Alert v-if="!pointReady" type="info" message="来源扇区尚未就绪，请完成数据准备后选择。" show-icon />
    <ChartOptionFields :config="config" :controls="controls" :disabled="disabled || !pointReady" @update:config="change" />
    <Button :disabled="disabled || !pointReady" @click="reset">恢复所选扇区继承样式</Button>
  </div>
</template>

<style scoped>
.chart-pie-slice-settings { display: grid; gap: 10px; min-width: 0; }
</style>
