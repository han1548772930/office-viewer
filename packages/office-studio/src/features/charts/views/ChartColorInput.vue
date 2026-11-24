<script setup lang="ts">
import { computed, ref } from 'vue'
import { Alert, Button, ColorPicker, FormItem, InputNumber, Select } from 'antdv-next'
import type { ChartColor, ChartLinearGradient } from '../../../domain/model'
import { CHART_GRADIENT_MAX_STOPS, isChartLinearGradient, validateChartColor } from '../chart-color'
import { appendChartGradientStop, chartColorForMode, chartColorMode, patchChartGradientCoordinate, patchChartGradientStop, removeChartGradientStop, type ChartColorMode } from '../chart-color-editing'

const props = defineProps<{ value?: ChartColor; fallback?: ChartColor; label: string; gradientAllowed: boolean; disabled?: boolean }>()
const emit = defineEmits<{ 'update:value': [value: ChartColor | undefined] }>()
const error = ref('')
const mode = computed(() => chartColorMode(props.value))
const gradient = computed(() => isChartLinearGradient(props.value) ? props.value : undefined)
const modes = computed(() => [{ label: '继承颜色', value: 'inherit' }, { label: '纯色', value: 'solid' },
  { label: '线性渐变', value: 'gradient', disabled: !props.gradientAllowed }])
const gradientDisabled = computed(() => props.disabled || !props.gradientAllowed)
const endpoints = ['start', 'end'] as const
const coordinates = ['x', 'y'] as const

function update(edit: () => ChartColor | undefined): void {
  if (props.disabled) return
  error.value = ''
  try {
    const color = edit()
    if (color !== undefined) validateChartColor(color)
    if (isChartLinearGradient(color) && !props.gradientAllowed) throw new TypeError('该图型的此颜色只接受纯色。')
    emit('update:value', color)
  } catch (cause) { error.value = cause instanceof Error ? cause.message : '无法更新图表颜色。' }
}
function setMode(value: unknown): void {
  if (value !== 'inherit' && value !== 'solid' && value !== 'gradient') return
  update(() => chartColorForMode(props.value, value as ChartColorMode, props.fallback))
}
function setSolid(value: unknown): void {
  update(() => typeof value === 'string' && value !== '' ? value : undefined)
}
function editGradient(edit: (value: ChartLinearGradient) => ChartLinearGradient): void {
  if (gradientDisabled.value || !gradient.value) return
  update(() => edit(gradient.value!))
}
function finite(value: unknown): number {
  const parsed = typeof value === 'number' ? value : typeof value === 'string' && value.trim() ? Number(value) : NaN
  if (!Number.isFinite(parsed)) throw new RangeError('请输入有限数值。')
  return parsed
}
function setCoordinate(endpoint: 'start' | 'end', coordinate: 'x' | 'y', value: unknown): void {
  editGradient(color => patchChartGradientCoordinate(color, endpoint, coordinate, finite(value)))
}
function setStop(index: number, field: 'offset' | 'color' | 'opacity', value: unknown): void {
  editGradient(color => {
    if (field === 'color') {
      if (typeof value !== 'string' || !value) throw new TypeError('色标需要有效颜色。')
      return patchChartGradientStop(color, index, { color: value })
    }
    const number = field === 'opacity' && (value === null || value === undefined || value === '') ? undefined : finite(value)
    return patchChartGradientStop(color, index, field === 'offset' ? { offset: number! } : { opacity: number })
  })
}
</script>

<template>
  <div class="chart-color-input">
    <Select :value="mode" :options="modes" :aria-label="`${label}填充方式`" :disabled="disabled" @update:value="setMode" />
    <ColorPicker v-if="mode === 'solid'" :value="typeof value === 'string' ? value : undefined" value-format="rgb" format="rgb" allow-clear show-text :aria-label="label" :disabled="disabled" @update:value="setSolid" @clear="setSolid(undefined)" />
    <template v-if="gradient">
      <div class="chart-gradient-coordinates">
        <template v-for="endpoint in endpoints" :key="endpoint">
          <FormItem v-for="coordinate in coordinates" :key="coordinate" :label="`${endpoint === 'start' ? '起点' : '终点'} ${coordinate.toUpperCase()}`">
            <InputNumber :value="gradient[endpoint][coordinate]" :step="0.05" :aria-label="`${label}渐变${endpoint === 'start' ? '起点' : '终点'} ${coordinate.toUpperCase()}`" :disabled="gradientDisabled" @update:value="setCoordinate(endpoint, coordinate, $event)" />
          </FormItem>
        </template>
      </div>
      <div v-for="(stop, index) in gradient.stops" :key="index" class="chart-gradient-stop">
        <FormItem :label="`色标 ${index + 1} 颜色`">
          <ColorPicker :value="stop.color" value-format="rgb" format="rgb" show-text :aria-label="`${label}色标 ${index + 1} 颜色`" :disabled="gradientDisabled" @update:value="setStop(index, 'color', $event)" />
        </FormItem>
        <FormItem label="位置">
          <InputNumber :value="stop.offset" :min="0" :max="1" :step="0.05" :aria-label="`${label}色标 ${index + 1} 位置`" :disabled="gradientDisabled" @update:value="setStop(index, 'offset', $event)" />
        </FormItem>
        <FormItem label="不透明度">
          <InputNumber :value="stop.opacity" :min="0" :max="1" :step="0.05" placeholder="1" :aria-label="`${label}色标 ${index + 1} 不透明度`" :disabled="gradientDisabled" @update:value="setStop(index, 'opacity', $event)" />
        </FormItem>
        <Button :aria-label="`删除${label}渐变色标 ${index + 1}`" :disabled="gradientDisabled || gradient.stops.length <= 2" @click="editGradient(color => removeChartGradientStop(color, index))">删除</Button>
      </div>
      <Button :aria-label="`添加${label}渐变色标`" :disabled="gradientDisabled || gradient.stops.length >= CHART_GRADIENT_MAX_STOPS" @click="editGradient(appendChartGradientStop)">添加色标</Button>
      <Alert v-if="!gradientAllowed" type="info" message="该图型的此颜色只接受纯色。已有渐变会保留，切换填充方式可继续编辑。" show-icon />
    </template>
    <Alert v-if="error" type="error" :message="error" show-icon />
  </div>
</template>

<style scoped>
.chart-color-input { display: grid; gap: 8px; }
.chart-gradient-coordinates { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 6px 10px; }
.chart-gradient-stop { display: grid; grid-template-columns: minmax(120px, 1fr) repeat(2, minmax(85px, 1fr)) auto; align-items: end; gap: 8px; }
.chart-color-input :deep(.ant-input-number), .chart-color-input :deep(.ant-select) { width: 100%; }
.chart-gradient-stop :deep(.ant-form-item) { margin-bottom: 0; }
@media (max-width: 700px) { .chart-gradient-stop { grid-template-columns: repeat(2, minmax(0, 1fr)); } }
</style>
