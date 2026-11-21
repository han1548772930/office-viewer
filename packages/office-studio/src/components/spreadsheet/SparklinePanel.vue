<script setup lang="ts">
import { computed, inject, ref, watch } from 'vue'
import { Button, ColorPicker, Form, FormItem, Input, InputNumber, Modal, Select } from 'antdv-next'
import { FOffice } from '../../api/facade'
import type { OfficeStudio } from '../../api/office-studio'
import type { SparklineEmptyCells, SparklineType } from '../../domain/model'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

const props = defineProps<{ open: boolean; studio: OfficeStudio; workbookId: string; sheetId: string; targetRange: string }>()
const emit = defineEmits<{ 'update:open': [value: boolean]; applied: [] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const types = computed<Array<{ label: string; value: SparklineType }>>(() => [
  { label: t('sheet.sparkline.type.line', '折线图'), value: 'line' }, { label: t('sheet.sparkline.type.column', '柱形图'), value: 'column' }, { label: t('sheet.sparkline.type.winloss', '胜负图'), value: 'winloss' },
])
const emptyModes = computed<Array<{ label: string; value: SparklineEmptyCells }>>(() => [
  { label: t('sheet.sparkline.empty-gap', '留空隙'), value: 'gap' }, { label: t('sheet.sparkline.empty-zero', '按零值绘制'), value: 'zero' }, { label: t('sheet.sparkline.empty-connect', '连接数据点'), value: 'connect' },
])
const type = ref<SparklineType>('line')
const sourceRange = ref('B1:E1')
const targetRange = ref(props.targetRange)
const color = ref('#167a55')
const negativeColor = ref('#c43b3b')
const highPointColor = ref('')
const lowPointColor = ref('')
const emptyCells = ref<SparklineEmptyCells>('gap')
const axis = ref<'individual' | 'custom'>('individual')
const minimum = ref<number | null>(0)
const maximum = ref<number | null>(100)
const message = ref('')
const busy = ref(false)
const existing = ref(false)
const isCustomAxis = computed(() => axis.value === 'custom')

watch(() => [props.open, props.targetRange, props.sheetId] as const, ([open]) => {
  if (!open) return
  targetRange.value = props.targetRange
  message.value = ''
  const cellAddress = props.targetRange.trim().split(':')[0] ?? ''
  try {
    const spec = FOffice.newAPI(props.studio).getWorkbook(props.workbookId)?.getSheetBySheetId(props.sheetId)?.getSparkline(cellAddress) ?? null
    existing.value = Boolean(spec)
    if (!spec) return resetDraft()
    type.value = spec.type
    sourceRange.value = spec.sourceRange
    color.value = spec.color
    negativeColor.value = spec.negativeColor ?? '#c43b3b'
    highPointColor.value = spec.highPointColor ?? ''
    lowPointColor.value = spec.lowPointColor ?? ''
    emptyCells.value = spec.emptyCells
    axis.value = spec.axis
    minimum.value = spec.min ?? 0
    maximum.value = spec.max ?? 100
  } catch (error) {
    message.value = error instanceof Error ? error.message : t('sheet.sparkline.error.read', '无法读取迷你图配置。')
  }
}, { immediate: true })

function resetDraft() {
  type.value = 'line'
  sourceRange.value = 'B1:E1'
  color.value = '#167a55'
  negativeColor.value = '#c43b3b'
  highPointColor.value = ''
  lowPointColor.value = ''
  emptyCells.value = 'gap'
  axis.value = 'individual'
  minimum.value = 0
  maximum.value = 100
}

function apply() {
  message.value = ''
  const sheet = FOffice.newAPI(props.studio).getWorkbook(props.workbookId)?.getSheetBySheetId(props.sheetId)
  if (!sheet) { message.value = t('sheet.sparkline.error.sheet-closed', '工作表已关闭，请重新打开迷你图配置。'); return }
  const min = minimum.value ?? 0
  const max = maximum.value ?? 100
  busy.value = true
  try {
    sheet.setSparklines({
      targetRange: targetRange.value,
      sourceRange: sourceRange.value,
      type: type.value,
      color: color.value,
      ...(negativeColor.value ? { negativeColor: negativeColor.value } : {}),
      ...(highPointColor.value ? { highPointColor: highPointColor.value } : {}),
      ...(lowPointColor.value ? { lowPointColor: lowPointColor.value } : {}),
      emptyCells: emptyCells.value,
      axis: axis.value,
      ...(axis.value === 'custom' ? { min, max } : {}),
    })
    existing.value = true
    emit('applied')
    emit('update:open', false)
  } catch (error) {
    message.value = error instanceof Error ? error.message : t('sheet.sparkline.error.save', '迷你图配置未能保存。')
  } finally { busy.value = false }
}

function remove() {
  const sheet = FOffice.newAPI(props.studio).getWorkbook(props.workbookId)?.getSheetBySheetId(props.sheetId)
  if (!sheet) return
  message.value = ''
  busy.value = true
  try {
    sheet.clearSparklines(targetRange.value)
    existing.value = false
    emit('applied')
    emit('update:open', false)
  } catch (error) {
    message.value = error instanceof Error ? error.message : t('sheet.sparkline.error.clear', '无法清除迷你图。')
  } finally { busy.value = false }
}
</script>

<template>
  <Modal :open="open" :title="t(existing ? 'sheet.sparkline.title.edit' : 'sheet.sparkline.title.insert', existing ? '编辑迷你图' : '插入迷你图')" :ok-text="t('common.apply', '应用')" :cancel-text="t('common.cancel', '取消')" :confirm-loading="busy" :width="560" @ok="apply" @cancel="emit('update:open', false)">
    <Form layout="vertical" class="sparkline-form">
      <FormItem name="type" :label="t('sheet.sparkline.type', '类型')"><Select v-model:value="type" :options="types" :aria-label="t('sheet.sparkline.type', '迷你图类型')" /></FormItem>
      <FormItem name="sourceRange" :label="t('sheet.sparkline.source-range', '源数据区域')"><Input v-model:value="sourceRange" :aria-label="t('sheet.sparkline.source-range', '迷你图源数据区域')" placeholder="B1:E1" /></FormItem>
      <FormItem name="targetRange" :label="t('sheet.sparkline.target-range', '放置区域')"><Input v-model:value="targetRange" :aria-label="t('sheet.sparkline.target-range', '迷你图放置区域')" placeholder="F1" /></FormItem>
      <FormItem name="color" :label="t('sheet.sparkline.positive-color', '正值颜色')"><ColorPicker v-model:value="color" value-format="hex" :aria-label="t('sheet.sparkline.positive-color', '迷你图正值颜色')" /></FormItem>
      <FormItem name="negativeColor" :label="t('sheet.sparkline.negative-color', '负值颜色')"><ColorPicker v-model:value="negativeColor" value-format="hex" :aria-label="t('sheet.sparkline.negative-color', '迷你图负值颜色')" /></FormItem>
      <FormItem name="highPointColor" :label="t('sheet.sparkline.high-point-color', '高值点颜色')"><ColorPicker v-model:value="highPointColor" value-format="hex" :aria-label="t('sheet.sparkline.high-point-color', '迷你图高值点颜色')" /></FormItem>
      <FormItem name="lowPointColor" :label="t('sheet.sparkline.low-point-color', '低值点颜色')"><ColorPicker v-model:value="lowPointColor" value-format="hex" :aria-label="t('sheet.sparkline.low-point-color', '迷你图低值点颜色')" /></FormItem>
      <FormItem name="emptyCells" :label="t('sheet.sparkline.empty-cells', '空单元格')"><Select v-model:value="emptyCells" :options="emptyModes" :aria-label="t('sheet.sparkline.empty-cells', '迷你图空单元格处理')" /></FormItem>
      <FormItem name="axis" :label="t('sheet.sparkline.axis', '坐标轴')"><Select v-model:value="axis" :options="[{ label: t('sheet.sparkline.axis-individual', '每个迷你图单独缩放'), value: 'individual' }, { label: t('sheet.sparkline.axis-custom', '自定义范围'), value: 'custom' }]" :aria-label="t('sheet.sparkline.axis', '迷你图坐标轴')" /></FormItem>
      <template v-if="isCustomAxis">
        <FormItem name="minimum" :label="t('sheet.sparkline.minimum', '最小值')"><InputNumber v-model:value="minimum" :aria-label="t('sheet.sparkline.minimum', '迷你图坐标轴最小值')" /></FormItem>
        <FormItem name="maximum" :label="t('sheet.sparkline.maximum', '最大值')"><InputNumber v-model:value="maximum" :aria-label="t('sheet.sparkline.maximum', '迷你图坐标轴最大值')" /></FormItem>
      </template>
      <p v-if="message" class="sparkline-error" role="alert">{{ message }}</p>
      <Button v-if="existing" type="text" danger :disabled="busy" @click="remove">{{ t('sheet.sparkline.remove', '清除目标区域迷你图') }}</Button>
    </Form>
  </Modal>
</template>

<style scoped>
.sparkline-form { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px; }
.sparkline-form :deep(.ant-form-item) { min-width: 0; margin-bottom: 0; }
.sparkline-form :deep(.ant-form-item-label) { padding: 0 0 4px; text-align: start; }
.sparkline-form :deep(.ant-color-picker-trigger), .sparkline-form :deep(.ant-input-number) { width: 100%; }
.sparkline-error { grid-column: 1 / -1; margin: 0; color: #b42318; }
@media (max-width: 560px) { .sparkline-form { grid-template-columns: minmax(0, 1fr); } }
</style>
