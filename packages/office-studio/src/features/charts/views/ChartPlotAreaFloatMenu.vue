<script setup lang="ts">
import { computed, nextTick, ref, watch } from 'vue'
import { Alert, Button, Form, FormItem, InputNumber, Switch, Tooltip, theme } from 'antdv-next'
import type { InputNumberRef } from 'antdv-next'
import { RotateCcw, X } from 'lucide-vue-next'
import type { ChartModel, ChartPlotAreaOptions } from '../../../domain/model'
import { validateChartPlotArea } from '../chart-plot-area'
import { chartPlotAreasEqual, type ChartCanvasPlotAreaTarget, type ChartPlotAreaCoordinate } from '../chart-plot-area-interaction'

const props = defineProps<{
  chart: ChartModel
  target: ChartCanvasPlotAreaTarget
  disabled?: boolean
  apply: (area: ChartPlotAreaOptions | undefined) => Promise<void>
}>()
const emit = defineEmits<{ close: [] }>()
const root = ref<HTMLElement | null>(null)
const firstNumber = ref<InputNumberRef | null>(null)
const manual = ref(props.chart.options?.plotArea !== undefined)
const saving = ref(false), composing = ref(false), error = ref('')
const fields: readonly { key: ChartPlotAreaCoordinate; label: string }[] = [
  { key: 'left', label: '绘图区左侧位置 (%)' }, { key: 'top', label: '绘图区顶部位置 (%)' },
  { key: 'width', label: '绘图区宽度 (%)' }, { key: 'height', label: '绘图区高度 (%)' },
]
const percentages = (area: ChartPlotAreaOptions): Record<ChartPlotAreaCoordinate, number> => ({
  left: area.left * 100, top: area.top * 100, width: area.width * 100, height: area.height * 100,
})
const draft = ref<Record<ChartPlotAreaCoordinate, number | undefined>>(percentages(props.target.area))
const candidate = computed<ChartPlotAreaOptions | undefined>(() => manual.value ? {
  left: draft.value.left === undefined ? props.target.automatic.left : draft.value.left / 100,
  top: draft.value.top === undefined ? props.target.automatic.top : draft.value.top / 100,
  width: draft.value.width === undefined ? props.target.automatic.width : draft.value.width / 100,
  height: draft.value.height === undefined ? props.target.automatic.height : draft.value.height / 100,
} : undefined)
const dirty = computed(() => !chartPlotAreasEqual(candidate.value, props.chart.options?.plotArea))
const validationError = computed(() => {
  try { validateChartPlotArea(candidate.value); return '' }
  catch (cause) { return cause instanceof Error ? cause.message : '绘图区矩形无效。' }
})
const unavailable = computed(() => !!props.disabled || saving.value || composing.value)
const { token } = theme.useToken()
const menuStyle = computed(() => ({ background: token.value.colorBgElevated, color: token.value.colorText,
  borderColor: token.value.colorBorder, boxShadow: token.value.boxShadowSecondary, borderRadius: `${token.value.borderRadiusLG}px` }))

function syncDraft(): void {
  manual.value = props.chart.options?.plotArea !== undefined
  // The author can commit before the next Canvas frame. Use the authoritative
  // rectangle here rather than copying a still-retained preview bitmap.
  draft.value = percentages(props.chart.options?.plotArea ?? props.target.automatic)
}
watch([() => props.chart.options?.plotArea, () => props.target], () => { if (!dirty.value && !saving.value) syncDraft() }, { deep: true })
function setNumberRef(key: ChartPlotAreaCoordinate, value: unknown): void { if (key === 'left') firstNumber.value = value as InputNumberRef | null }
function setManual(value: unknown): void {
  if (unavailable.value) return
  manual.value = value === true
  draft.value = percentages(manual.value ? props.target.area : props.target.automatic)
  error.value = ''
}
function setCoordinate(key: ChartPlotAreaCoordinate, value: unknown): void {
  if (unavailable.value || !manual.value) return
  const number = value === null || value === undefined || value === '' ? undefined
    : typeof value === 'number' ? value : typeof value === 'string' ? Number(value) : NaN
  if (number !== undefined && !Number.isFinite(number)) { error.value = '请输入有限百分比。'; return }
  draft.value = { ...draft.value, [key]: number }
  error.value = ''
}
function resetCoordinates(keys: readonly ChartPlotAreaCoordinate[]): void {
  if (unavailable.value || !manual.value) return
  draft.value = { ...draft.value, ...Object.fromEntries(keys.map(key => [key, props.target.automatic[key] * 100])) }
  error.value = ''
}
function restoreAutomatic(): void { if (!unavailable.value) { manual.value = false; draft.value = percentages(props.target.automatic); error.value = '' } }
async function applyArea(): Promise<void> {
  if (unavailable.value || !dirty.value) return
  try { validateChartPlotArea(candidate.value) }
  catch (cause) { error.value = cause instanceof Error ? cause.message : '绘图区矩形无效。'; return }
  // All four coordinates (or the complete reset) share one author command.
  const area = candidate.value ? { ...candidate.value } : undefined
  saving.value = true; error.value = ''
  try { await props.apply(area); await nextTick(); syncDraft() }
  catch (cause) { error.value = cause instanceof Error ? cause.message : '绘图区修改失败。' }
  finally { saving.value = false }
}
function numberEnter(event: KeyboardEvent): void {
  if (composing.value || event.isComposing || event.keyCode === 229) return
  event.preventDefault(); event.stopPropagation(); void applyArea()
}
function menuKeyDown(event: KeyboardEvent): void {
  event.stopPropagation()
  if (event.key === 'Escape' && !event.defaultPrevented && !composing.value && !event.isComposing && event.keyCode !== 229) {
    event.preventDefault(); emit('close')
  }
}
async function focusCoordinate(): Promise<void> {
  await nextTick()
  if (props.disabled) return
  if (manual.value) firstNumber.value?.focus()
  else root.value?.querySelector<HTMLElement>('.ant-switch')?.focus({ preventScroll: true })
}
defineExpose({ focusCoordinate })
</script>

<template>
  <section ref="root" class="chart-plot-area-float-menu" :style="menuStyle" aria-label="图表绘图区浮动工具栏" @pointerdown.stop @click.stop @keydown="menuKeyDown">
    <Form layout="vertical" size="small" :disabled="disabled || saving" @finish="applyArea">
      <div class="chart-plot-area-heading">
        <FormItem label="手动设置绘图区" class="chart-plot-area-mode">
          <Switch :checked="manual" :disabled="unavailable" aria-label="绘图区手动布局" @update:checked="setManual" />
        </FormItem>
        <Tooltip title="关闭绘图区工具栏；Esc 取消未提交修改">
          <Button type="text" aria-label="关闭绘图区工具栏" @click="emit('close')"><X :size="14" /></Button>
        </Tooltip>
      </div>
      <div class="chart-plot-area-coordinates">
        <FormItem v-for="field in fields" :key="field.key" :label="field.label">
          <div class="chart-plot-area-value">
            <InputNumber :ref="value => setNumberRef(field.key, value)" :value="manual ? draft[field.key] : target.automatic[field.key] * 100"
              :min="field.key === 'width' || field.key === 'height' ? 0.0001 : 0" :max="100" :step="0.5" :precision="6" :aria-label="field.label"
              :placeholder="String(target.automatic[field.key] * 100)" :disabled="unavailable || !manual"
              @update:value="setCoordinate(field.key, $event)" @press-enter="numberEnter"
              @compositionstart="composing = true" @compositionend="composing = false" />
            <Tooltip title="只恢复此项的自动值；应用修改后保存">
              <Button type="text" size="small" :aria-label="`恢复${field.label}自动值`" :disabled="unavailable || !manual"
                @click="resetCoordinates([field.key])"><RotateCcw :size="12" /></Button>
            </Tooltip>
          </div>
        </FormItem>
      </div>
      <div class="chart-plot-area-actions">
        <Button size="small" aria-label="恢复绘图区自动位置" :disabled="unavailable || !manual" @click="resetCoordinates(['left', 'top'])">恢复自动位置</Button>
        <Button size="small" aria-label="恢复绘图区自动大小" :disabled="unavailable || !manual" @click="resetCoordinates(['width', 'height'])">恢复自动大小</Button>
        <Button size="small" aria-label="恢复绘图区整组自动布局" :disabled="unavailable || !manual" @click="restoreAutomatic">恢复自动布局</Button>
        <Button type="primary" size="small" html-type="submit" :disabled="unavailable || !dirty || !!validationError" :loading="saving">应用修改</Button>
      </div>
      <Alert v-if="validationError || error" type="error" show-icon :message="validationError || error" />
    </Form>
  </section>
</template>

<style scoped>
.chart-plot-area-float-menu { box-sizing: border-box; width: min(410px, 100%); padding: 8px; border: 1px solid; pointer-events: auto; }
.chart-plot-area-heading { display: flex; justify-content: space-between; align-items: flex-start; gap: 8px; }
.chart-plot-area-mode { flex: 1; }
.chart-plot-area-coordinates { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 6px 10px; }
.chart-plot-area-value { display: flex; align-items: center; gap: 4px; }
.chart-plot-area-value :deep(.ant-input-number) { flex: 1; min-width: 0; width: 100%; }
.chart-plot-area-actions { display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 6px; }
.chart-plot-area-float-menu :deep(.ant-form-item) { margin-bottom: 6px; }
.chart-plot-area-float-menu :deep(.ant-form-item-label) { padding-bottom: 2px; }
</style>
