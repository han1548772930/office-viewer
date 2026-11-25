<script setup lang="ts">
import { computed, nextTick, ref, watch } from 'vue'
import { Alert, Button, ColorPicker, Form, FormItem, Input, InputNumber, Select, Tooltip, theme } from 'antdv-next'
import { Bold, Italic, RotateCcw, X } from 'lucide-vue-next'
import type { InputTextAreaRef } from 'antdv-next'
import type { ChartModel } from '../../../domain/model'
import { chartTitleFont, type ChartTitleKind, type ChartTitleObjectSpec } from '../chart-titles'

const props = defineProps<{
  chart: ChartModel
  kind: ChartTitleKind
  disabled?: boolean
  apply: (patch: ChartTitleObjectSpec) => Promise<void>
}>()
const emit = defineEmits<{ close: [] }>()
const root = ref<HTMLElement | null>(null)
const input = ref<InputTextAreaRef | null>(null)
const text = ref(props.chart[props.kind] ?? '')
const composing = ref(false), saving = ref(false), error = ref(''), blurRequested = ref(false)
const { token } = theme.useToken()
const style = computed(() => props.chart.options?.[props.kind === 'title' ? 'titleStyle' : 'subtitleStyle'])
const font = computed(() => chartTitleFont(props.chart, props.kind))
const numericDraft = ref<{ fontSize?: number; rotate?: number }>({ fontSize: font.value.fontSize, rotate: style.value?.rotate ?? 0 })
const numericDirty = ref({ fontSize: false, rotate: false })
const colorDraft = ref<string | undefined>(font.value.color ?? token.value.colorText), colorDirty = ref(false)
const textDirty = computed(() => text.value !== (props.chart[props.kind] ?? ''))
const dirty = computed(() => textDirty.value || numericDirty.value.fontSize || numericDirty.value.rotate || colorDirty.value)
const unavailable = computed(() => !!props.disabled || saving.value || composing.value)
const positions = [{ value: 'top', label: '上方' }, { value: 'bottom', label: '下方' }, { value: 'left', label: '左侧' }, { value: 'right', label: '右侧' }, { value: 'hide', label: '隐藏' }]
const alignments = [{ value: 'left', label: '靠左' }, { value: 'center', label: '居中' }, { value: 'right', label: '靠右' }]
const position = computed(() => props.chart.options?.titleStyle?.visible === false ? 'hide' : props.chart.options?.titleStyle?.position ?? 'top')
const menuStyle = computed(() => ({ background: token.value.colorBgElevated, color: token.value.colorText,
  borderColor: token.value.colorBorder, boxShadow: token.value.boxShadowSecondary, borderRadius: `${token.value.borderRadiusLG}px` }))
const popupContainer = () => root.value ?? document.body

watch(() => props.chart[props.kind], (value, previous) => {
  if (text.value === (previous ?? '') || text.value === (value ?? '')) text.value = value ?? ''
})
watch([() => font.value.fontSize, () => style.value?.rotate, () => font.value.color, () => token.value.colorText], () => {
  if (!numericDirty.value.fontSize) numericDraft.value.fontSize = font.value.fontSize
  if (!numericDirty.value.rotate) numericDraft.value.rotate = style.value?.rotate ?? 0
  if (!colorDirty.value) colorDraft.value = font.value.color ?? token.value.colorText
})

function acknowledgeDraft(): void {
  numericDirty.value = { fontSize: false, rotate: false }; colorDirty.value = false
  numericDraft.value = { fontSize: font.value.fontSize, rotate: style.value?.rotate ?? 0 }
  colorDraft.value = font.value.color ?? token.value.colorText
}

async function applyStyle(patch: ChartTitleObjectSpec = {}): Promise<void> {
  if (unavailable.value) return
  const authored = style.value
  const requested = { ...(numericDirty.value.fontSize ? { fontSize: numericDraft.value.fontSize } : {}),
    ...(numericDirty.value.rotate ? { rotate: numericDraft.value.rotate } : {}), ...(colorDirty.value ? { color: colorDraft.value } : {}), ...patch }
  const changedStyle = { ...requested }
  for (const [key, value] of Object.entries(requested)) if (Object.is(authored && Reflect.get(authored, key), value)) Reflect.deleteProperty(changedStyle, key)
  const candidate: ChartTitleObjectSpec = { ...(textDirty.value ? { text: text.value } : {}), ...changedStyle }
  if (!Object.keys(candidate).length) { acknowledgeDraft(); return }
  saving.value = true; error.value = ''
  try { await props.apply(candidate); acknowledgeDraft() }
  catch (cause) { error.value = cause instanceof Error ? cause.message : '标题修改失败。' }
  finally { saving.value = false }
}

function numberValue(value: unknown): number | undefined {
  if (value === null || value === undefined || value === '') return undefined
  if (typeof value !== 'number' || !Number.isFinite(value)) throw new TypeError('标题数值无效。')
  return value
}
function setNumber(field: 'fontSize' | 'rotate', value: unknown): void {
  try {
    numericDraft.value[field] = numberValue(value)
    numericDirty.value[field] = !Object.is(style.value?.[field], numericDraft.value[field])
  }
  catch (cause) { error.value = cause instanceof Error ? cause.message : '标题数值无效。' }
}
function setColor(value: unknown): void {
  if (value === null) value = undefined
  if (value !== undefined && typeof value !== 'string') { error.value = '标题颜色无效。'; return }
  colorDraft.value = value
  colorDirty.value = !Object.is(style.value?.color, value)
}
function clearColor(): void { setColor(undefined); void applyStyle({ color: undefined }) }
function numberEnter(event: KeyboardEvent): void { event.preventDefault(); event.stopPropagation(); void applyStyle() }
function numberStep(field: 'fontSize' | 'rotate', value: unknown): void { setNumber(field, value); void applyStyle() }
function setPosition(value: unknown): void {
  if (!['top', 'bottom', 'left', 'right', 'hide'].includes(String(value))) return
  const selected = positions.find(item => item.value === value)
  if (!selected) return
  // Clearing manualLayout is part of this one author patch, not a second edit.
  if (selected.value === 'hide') void applyStyle({ visible: false })
  else if (selected.value === 'top' || selected.value === 'bottom' || selected.value === 'left' || selected.value === 'right')
    void applyStyle({ visible: true, position: selected.value, manualLayout: undefined })
}
function setAlignment(value: unknown): void {
  if (value === 'left' || value === 'center' || value === 'right') void applyStyle({ alignment: value })
}
function compositionEnd(event: CompositionEvent): void {
  if (event.target instanceof HTMLTextAreaElement) text.value = event.target.value
  composing.value = false
  if (blurRequested.value) {
    blurRequested.value = false
    void nextTick(() => { if (root.value && !root.value.contains(document.activeElement)) void applyStyle() })
  }
}
function textKeyDown(event: KeyboardEvent): void {
  if (composing.value || event.isComposing || event.keyCode === 229) return
  if (event.key === 'Enter' && !event.shiftKey) { event.preventDefault(); event.stopPropagation(); void applyStyle() }
  if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); emit('close') }
}
async function textBlur(event: FocusEvent): Promise<void> {
  if (event.relatedTarget instanceof Node && root.value?.contains(event.relatedTarget)) return
  await nextTick()
  if (composing.value) { blurRequested.value = true; return }
  if (!composing.value && root.value && !root.value.contains(document.activeElement)) await applyStyle()
}
function menuKeyDown(event: KeyboardEvent): void {
  event.stopPropagation()
  if (event.key === 'Escape' && !event.defaultPrevented && !composing.value && !event.isComposing && event.keyCode !== 229) emit('close')
}
async function focusText(selectAll = false): Promise<void> {
  await nextTick()
  if (props.disabled || !input.value) return
  input.value.focus?.()
  const node = antTextArea(input.value.nativeElement)
  // Selection uses the existing Ant TextArea node. No separate native input
  // or second author model is introduced for the Canvas title input bridge.
  if (selectAll && node instanceof HTMLTextAreaElement) node.setSelectionRange(0, node.value.length)
}
function antTextArea(node: HTMLElement | null): HTMLTextAreaElement | null {
  return node instanceof HTMLTextAreaElement ? node : node?.querySelector<HTMLTextAreaElement>('textarea') ?? null
}
defineExpose({ focusText })
</script>

<template>
  <section ref="root" class="chart-title-float-menu" :style="menuStyle" :aria-label="kind === 'title' ? '图表标题浮动工具栏' : '图表副标题浮动工具栏'"
    :data-title-kind="kind" @pointerdown.stop @click.stop @keydown="menuKeyDown">
    <Form layout="vertical" size="small" :disabled="disabled || saving" @finish="applyStyle()">
      <div class="chart-title-text-row">
        <FormItem :label="kind === 'title' ? '图表标题' : '副标题'" class="chart-title-text-field">
          <Input.TextArea ref="input" v-model:value="text" :maxlength="100000" :auto-size="{ minRows: 1, maxRows: 3 }" :change-on-composing="true"
            :disabled="disabled || saving" :aria-label="kind === 'title' ? '编辑图表标题' : '编辑图表副标题'" @compositionstart="composing = true" @compositionend="compositionEnd"
            @keydown="textKeyDown" @blur="textBlur" />
        </FormItem>
        <Tooltip title="关闭标题工具栏；Esc 取消未提交文字">
          <Button type="text" aria-label="关闭标题工具栏" :disabled="false" @click="emit('close')"><X :size="14" /></Button>
        </Tooltip>
      </div>
      <div class="chart-title-format-row">
        <FormItem v-if="kind === 'title'" label="位置">
          <Select :value="position" :options="positions" :disabled="unavailable" :get-popup-container="popupContainer" aria-label="标题位置" @update:value="setPosition" />
        </FormItem>
        <FormItem label="字号">
          <InputNumber :value="numericDraft.fontSize" :min="1" :max="256" :disabled="unavailable" aria-label="标题字号" @update:value="setNumber('fontSize', $event)"
            @press-enter="numberEnter" @blur="textBlur" @step="numberStep('fontSize', $event)" />
          <Tooltip title="恢复继承字号"><Button type="text" size="small" aria-label="恢复标题字号" :disabled="unavailable || style?.fontSize === undefined && !numericDirty.fontSize" @click="applyStyle({ fontSize: undefined })"><RotateCcw :size="12" /></Button></Tooltip>
        </FormItem>
        <FormItem label="颜色">
          <ColorPicker :value="colorDraft" value-format="rgb" format="hex" allow-clear :disabled="unavailable" :get-popup-container="popupContainer"
            aria-label="标题文字颜色" @update:value="setColor" @change-complete="applyStyle()" @clear="clearColor" />
        </FormItem>
        <FormItem v-if="kind === 'title'" label="对齐">
          <Select :value="style?.alignment ?? 'center'" :options="alignments" :disabled="unavailable" :get-popup-container="popupContainer" aria-label="标题对齐" @update:value="setAlignment" />
        </FormItem>
        <FormItem label="字体">
          <Tooltip title="粗体"><Button :type="font.bold ? 'primary' : 'text'" :aria-pressed="!!font.bold" aria-label="标题粗体" :disabled="unavailable" @click="applyStyle({ bold: !font.bold })"><Bold :size="14" /></Button></Tooltip>
          <Tooltip title="斜体"><Button :type="font.italic ? 'primary' : 'text'" :aria-pressed="!!font.italic" aria-label="标题斜体" :disabled="unavailable" @click="applyStyle({ italic: !font.italic })"><Italic :size="14" /></Button></Tooltip>
        </FormItem>
      </div>
      <div class="chart-title-action-row">
        <FormItem label="旋转角度">
          <InputNumber :value="numericDraft.rotate" :min="-360" :max="360" :disabled="unavailable" aria-label="标题旋转角度" @update:value="setNumber('rotate', $event)"
            @press-enter="numberEnter" @blur="textBlur" @step="numberStep('rotate', $event)" />
        </FormItem>
        <Tooltip title="清除手动锚点，恢复自动布局">
          <Button size="small" :disabled="unavailable || !style?.manualLayout" aria-label="恢复标题自动位置" @click="applyStyle({ manualLayout: undefined })">恢复自动位置</Button>
        </Tooltip>
        <Button size="small" type="primary" html-type="submit" :disabled="unavailable || !dirty" :loading="saving">应用修改</Button>
      </div>
      <Alert v-if="error" type="error" show-icon :message="error" />
    </Form>
  </section>
</template>

<style scoped>
.chart-title-float-menu { box-sizing: border-box; width: min(440px, 100%); padding: 8px; border: 1px solid; pointer-events: auto; }
.chart-title-text-row, .chart-title-format-row, .chart-title-action-row { display: flex; align-items: flex-end; gap: 6px; }
.chart-title-text-row { align-items: flex-start; }
.chart-title-text-field { flex: 1; min-width: 0; }
.chart-title-format-row { flex-wrap: wrap; }
.chart-title-action-row { flex-wrap: wrap; align-items: center; }
.chart-title-float-menu :deep(.ant-form-item) { margin-bottom: 6px; }
.chart-title-float-menu :deep(.ant-form-item-label) { padding-bottom: 2px; }
.chart-title-format-row :deep(.ant-select) { min-width: 76px; }
.chart-title-format-row :deep(.ant-input-number), .chart-title-action-row :deep(.ant-input-number) { width: 72px; }
</style>
