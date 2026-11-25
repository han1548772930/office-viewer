<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, ref, watch } from 'vue'
import { Alert, Button, ConfigProvider, Form, Tooltip, theme } from 'antdv-next'
import { MoreHorizontal, RotateCcw, X } from 'lucide-vue-next'
import type { ChartModel } from '../../../domain/model'
import type { ChartConfig } from '../chart-engine'
import { chartCanvasElementAuthorSignature, chartCanvasElementControls, chartCanvasElementLabel, resetChartCanvasElementDraft } from '../chart-element-edit'
import type { ChartCanvasElementTarget } from '../chart-element-interaction'
import { cloneChartDraftConfig } from '../chart-settings-draft'
import { resolveChartType } from '../chart-catalog'
import { chartSpecialMarkScopeNote } from '../chart-special-mark-edit'
import ChartOptionFields from './ChartOptionFields.vue'

const props = defineProps<{
  chart: ChartModel
  target: ChartCanvasElementTarget
  disabled?: boolean
  moreAvailable?: boolean
  apply: (draft: ChartConfig) => Promise<void>
  more: () => Promise<boolean>
  promote: () => Promise<void>
}>()
const emit = defineEmits<{ close: [] }>()
const root = ref<HTMLElement | null>(null), saving = ref(false), composing = ref(false), error = ref('')
const draft = ref(cloneChartDraftConfig(props.chart))
let generation = 0, disposed = false
const { token } = theme.useToken()
const title = computed(() => chartCanvasElementLabel(props.target.element))
const controls = computed(() => chartCanvasElementControls(props.target.element, draft.value))
const dirty = computed(() => chartCanvasElementAuthorSignature(draft.value, props.target.element) !== chartCanvasElementAuthorSignature(props.chart, props.target.element))
const unavailable = computed(() => !!props.disabled || saving.value || composing.value)
const canPromote = computed(() => props.target.element.kind === 'dataPoint' || props.target.element.kind === 'dataLabel')
const scopeNote = computed(() => props.target.element.kind === 'mark' ? chartSpecialMarkScopeNote(props.target.element.directEdit)
  : props.target.element.kind !== 'dataLabel' ? '' : resolveChartType(props.chart.type).family === 'pie'
  ? '这些设置应用于全部扇区的数据标签。' : '这些设置应用于所选系列的数据标签。')
const menuStyle = computed(() => ({ background: token.value.colorBgElevated, color: token.value.colorText,
  borderColor: token.value.colorBorder, boxShadow: token.value.boxShadowSecondary, borderRadius: `${token.value.borderRadiusLG}px`,
  '--chart-element-secondary': token.value.colorTextSecondary }))
const popupContainer = (trigger?: HTMLElement): HTMLElement => root.value ?? trigger?.parentElement ?? document.body
function syncDraft(): void { draft.value = cloneChartDraftConfig(props.chart) }
watch(() => props.chart, () => { if (!saving.value && !dirty.value) syncDraft() }, { deep: true })
function updateDraft(value: ChartConfig): void {
  if (disposed || props.disabled || saving.value) return
  draft.value = value; error.value = ''
}
function reset(): void { if (!unavailable.value) { draft.value = resetChartCanvasElementDraft(draft.value, props.target.element); error.value = '' } }
async function applyDraft(): Promise<boolean> {
  if (unavailable.value || disposed) return false
  if (!dirty.value) return true
  const ticket = generation, candidate = cloneChartDraftConfig(draft.value)
  saving.value = true; error.value = ''
  try {
    await props.apply(candidate)
    await nextTick()
    if (!disposed && ticket === generation) syncDraft()
    return !disposed && ticket === generation
  } catch (cause) {
    if (!disposed && ticket === generation) error.value = cause instanceof Error ? cause.message : '图表元素修改失败。'
    return false
  } finally { if (!disposed && ticket === generation) saving.value = false }
}
async function openMore(): Promise<void> {
  if (unavailable.value || !props.moreAvailable || !await applyDraft() || disposed) return
  const ticket = generation
  saving.value = true; error.value = ''
  try { if (!await props.more() && !disposed && ticket === generation) error.value = '图表上下文已更新，请重新选择后打开更多设置。' }
  catch (cause) { if (!disposed && ticket === generation) error.value = cause instanceof Error ? cause.message : '无法打开图表设置。' }
  finally { if (!disposed && ticket === generation) saving.value = false }
}
async function selectSeries(): Promise<void> {
  if (unavailable.value || !canPromote.value || !await applyDraft() || disposed) return
  await props.promote()
}
function menuKeyDown(event: KeyboardEvent): void {
  event.stopPropagation()
  if (event.defaultPrevented || composing.value || event.isComposing || event.keyCode === 229) return
  if (event.key === 'Escape') { event.preventDefault(); emit('close'); return }
  if (event.key === 'Enter' && event.target instanceof HTMLInputElement
    && event.target.matches('.ant-input, .ant-input-number-input') && !event.target.closest('.ant-select, .ant-color-picker-panel')) {
    event.preventDefault(); void applyDraft()
  }
}
async function focusFirst(): Promise<void> {
  const ticket = generation
  await nextTick()
  if (disposed || ticket !== generation || props.disabled) return
  root.value?.querySelector<HTMLElement>('input:not(:disabled), textarea:not(:disabled), [role="combobox"]:not([aria-disabled="true"]), button:not(:disabled)')?.focus({ preventScroll: true })
}
onBeforeUnmount(() => { disposed = true; ++generation })
defineExpose({ focusFirst })
</script>

<template>
  <ConfigProvider :get-popup-container="popupContainer">
    <section ref="root" class="chart-element-float-menu" :style="menuStyle" aria-label="图表元素浮动工具栏"
      :data-chart-element-kind="target.element.kind" :data-chart-element-key="target.key"
      :data-chart-mark-kind="target.element.kind === 'mark' ? target.element.directEdit.kind : undefined"
      @pointerdown.stop @click.stop @keydown="menuKeyDown" @compositionstart.capture="composing = true" @compositionend.capture="composing = false">
      <div class="chart-element-heading">
        <span>{{ title }}</span>
        <Tooltip title="关闭元素工具栏；Esc 取消未提交修改">
          <Button type="text" size="small" aria-label="关闭图表元素工具栏" @click="emit('close')"><X :size="14" /></Button>
        </Tooltip>
      </div>
      <Form layout="vertical" size="small" :disabled="disabled || saving" @finish="applyDraft">
        <ChartOptionFields :config="draft" :controls="controls" :disabled="disabled || saving" @update:config="updateDraft" />
        <p v-if="scopeNote" class="chart-element-scope">{{ scopeNote }}</p>
        <div class="chart-element-actions">
          <Tooltip title="恢复所选元素的默认格式；应用修改后保存">
            <Button size="small" aria-label="恢复图表元素默认格式" :disabled="unavailable" @click="reset"><RotateCcw :size="13" />恢复格式</Button>
          </Tooltip>
          <Button v-if="canPromote" size="small" aria-label="选择整个图表系列" :disabled="unavailable" @click="selectSeries">选择系列</Button>
          <Button size="small" aria-label="图表元素更多设置" :disabled="unavailable || !moreAvailable" @click="openMore"><MoreHorizontal :size="14" />更多设置</Button>
          <Button type="primary" size="small" html-type="submit" aria-label="应用图表元素修改" :loading="saving" :disabled="unavailable || !dirty">应用修改</Button>
        </div>
        <Alert v-if="error" type="error" show-icon :message="error" />
      </Form>
    </section>
  </ConfigProvider>
</template>

<style scoped>
.chart-element-float-menu { box-sizing: border-box; width: 100%; max-height: min(var(--chart-element-max-height, 420px), 80vh); overflow: auto; padding: 8px; border: 1px solid; pointer-events: auto; }
.chart-element-heading { display: flex; justify-content: space-between; align-items: center; gap: 8px; margin-bottom: 8px; font-weight: 600; }
.chart-element-actions { display: flex; flex-wrap: wrap; gap: 6px; margin-block: 8px; }
.chart-element-actions :deep(.ant-btn) { display: inline-flex; align-items: center; gap: 4px; }
.chart-element-scope { margin: 6px 0; color: var(--chart-element-secondary); font-size: 12px; }
.chart-element-float-menu :deep(.ant-form-item-label) { padding-bottom: 2px; }
</style>
