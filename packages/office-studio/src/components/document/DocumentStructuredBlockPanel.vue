<script setup lang="ts">
import { computed, reactive, ref, watch } from 'vue'
import { Alert, Button, Checkbox, ColorPicker, Form, FormItem, Input, InputNumber, Modal, Select, Switch, Tooltip, theme } from 'antdv-next'
import type { DocumentBlock, DocumentColumnGroup } from '../../domain/model'
import { defaultDocumentCalloutConfig, validateDocumentCalloutConfig, validateDocumentColumnGroup,
  DOCUMENT_COLUMN_GROUP_MIN_COLUMNS, DOCUMENT_COLUMN_GROUP_MAX_COLUMNS, DOCUMENT_COLUMN_HORIZONTAL_PADDING, DOCUMENT_COLUMN_MIN_HEIGHT } from '../../features/document/model/document-structural-blocks'
import type { DocumentStructureCommand } from '../../features/document/commands/document-structure-commands'
import type { DocumentStructurePanelValue } from './document-structure-panel-state'
import { documentQuoteStyle, documentStructuredTextColor } from '../../features/document/model/document-structured-appearance'
import { validateDocumentRunPatch } from '../../features/document/model/document-run-range'

const props = defineProps<{ open: boolean; kind: 'quote' | 'callout' | 'columns'; block?: DocumentBlock; readOnly?: boolean }>()
const emit = defineEmits<{ 'update:open': [open: boolean]; submit: [value: DocumentStructurePanelValue]; command: [command: DocumentStructureCommand]; closed: [] }>()
const { token } = theme.useToken()
const callout = reactive(defaultDocumentCalloutConfig())
const quote = reactive({ lineColor: '#8796aa' })
const textAppearance = reactive({ color: '#1f2937', reset: false, changed: false })
const columns = reactive({ count: 2, gap: 16, horizontalPadding: DOCUMENT_COLUMN_HORIZONTAL_PADDING, minHeight: DOCUMENT_COLUMN_MIN_HEIGHT,
  layout: 'fixed' as DocumentColumnGroup['layout'], responsive: 'shrink' as DocumentColumnGroup['responsive'], ratios: [1, 1], minimums: [null, null] as Array<number | null> })
const columnIds = ref<string[]>([]), keepContents = ref(true), error = ref('')
const title = computed(() => `${props.block ? '设置' : '插入'}${{ quote: '引用块', callout: '标注块', columns: '多栏组' }[props.kind]}`)
const sizes = [
  { key: 'paddingTop' as const, label: '上内边距' }, { key: 'paddingRight' as const, label: '右内边距' },
  { key: 'paddingBottom' as const, label: '下内边距' }, { key: 'paddingLeft' as const, label: '左内边距' },
  { key: 'iconSize' as const, label: '图标尺寸' }, { key: 'iconGap' as const, label: '图标间距' },
]
watch(() => [props.open, props.kind, props.block] as const, ([open]) => {
  if (!open) return
  error.value = ''; keepContents.value = true
  Object.assign(callout, defaultDocumentCalloutConfig(), props.block?.callout)
  quote.lineColor = props.block?.type === 'quote' ? documentQuoteStyle(props.block).lineColor : '#8796aa'
  textAppearance.color = props.block ? documentStructuredTextColor(props.block) : '#1f2937'; textAppearance.reset = false; textAppearance.changed = false
  const group = props.block?.columns
  columns.count = group?.columns.length ?? 2; columns.gap = group?.gap ?? 16
  columns.horizontalPadding = group?.horizontalPadding ?? DOCUMENT_COLUMN_HORIZONTAL_PADDING
  columns.minHeight = group?.minHeight ?? DOCUMENT_COLUMN_MIN_HEIGHT
  columns.layout = group?.layout ?? 'fixed'; columns.responsive = group?.responsive ?? 'shrink'
  columns.ratios = group?.columns.map(column => column.widthRatio) ?? [1, 1]
  columns.minimums = group?.columns.map(column => column.minWidth ?? null) ?? [null, null]
  columnIds.value = group?.columns.map(column => column.id) ?? []
}, { immediate: true })
watch(() => columns.count, count => {
  if (props.block || !Number.isInteger(count)) return
  columns.ratios = Array.from({ length: count }, (_, index) => columns.ratios[index] ?? 1)
  columns.minimums = Array.from({ length: count }, (_, index) => columns.minimums[index] ?? null)
})

function submit() {
  if (props.readOnly) return
  try {
    const textColor = textAppearance.changed ? textAppearance.reset ? null : textAppearance.color : undefined
    if (textColor !== undefined && textColor !== null) validateDocumentRunPatch({ color: textColor })
    if (props.kind === 'quote') {
      validateDocumentRunPatch({ color: quote.lineColor })
      emit('submit', { kind: 'quote', style: { lineColor: quote.lineColor, ...(textColor === undefined ? {} : { textColor: textColor ?? 'auto' }) } })
    } else if (props.kind === 'callout') {
      const config = { ...callout }; validateDocumentCalloutConfig(config); emit('submit', { kind: 'callout', config, ...(textColor === undefined ? {} : { textColor }) })
    } else {
      const config: DocumentColumnGroup = { version: 1, gap: columns.gap, horizontalPadding: columns.horizontalPadding, minHeight: columns.minHeight,
        layout: columns.layout, responsive: columns.responsive,
        columns: columns.ratios.map((widthRatio, index) => ({ id: columnIds.value[index] ?? `draft-column-${index}`, widthRatio,
          ...(columns.minimums[index] == null ? {} : { minWidth: columns.minimums[index]! }), blocks: [] })) }
      validateDocumentColumnGroup(config)
      emit('submit', { kind: 'columns', columnCount: columns.count, gap: columns.gap, horizontalPadding: columns.horizontalPadding,
        minHeight: columns.minHeight, layout: columns.layout,
        responsive: columns.responsive, widthRatios: [...columns.ratios], minWidths: [...columns.minimums] })
    }
    error.value = ''
  } catch (failure) { error.value = failure instanceof Error ? failure.message : String(failure) }
}
function setTextColor(value: unknown) {
  if (typeof value !== 'string') return
  textAppearance.color = value; textAppearance.changed = true; textAppearance.reset = false
}
function resetTextColor() { textAppearance.color = '#1f2937'; textAppearance.reset = true; textAppearance.changed = true }
function command(value: DocumentStructureCommand) { if (!props.readOnly) emit('command', value) }
</script>

<template>
  <Modal :open="open" :title="title" :width="600" ok-text="应用" cancel-text="取消" :ok-button-props="{ disabled: readOnly }"
    destroy-on-hidden @ok="submit" @cancel="emit('update:open', false)" @after-close="emit('closed')">
    <div class="structured-panel" :style="{ '--structure-border': token.colorBorderSecondary, '--structure-text': token.colorTextSecondary }">
      <Alert v-if="error" :title="error" type="error" show-icon />
      <Form v-if="kind === 'quote'" :model="quote" :disabled="readOnly" layout="vertical">
        <FormItem name="lineColor" label="左侧线条颜色"><ColorPicker v-model:value="quote.lineColor" value-format="hex" show-text aria-label="引用块线条颜色" /></FormItem>
        <FormItem label="文字颜色"><ColorPicker :value="textAppearance.color" value-format="hex" show-text aria-label="引用块文字颜色" @update:value="setTextColor" /> <Button type="link" @click="resetTextColor">恢复继承颜色</Button></FormItem>
      </Form>
      <Form v-else-if="kind === 'callout'" :model="callout" :disabled="readOnly" layout="vertical">
        <div class="structured-panel__fields">
          <FormItem name="icon" label="图标"><Input v-model:value="callout.icon" :maxlength="256" aria-label="标注块图标" /></FormItem>
          <FormItem name="showIcon" label="显示图标"><Switch v-model:checked="callout.showIcon" aria-label="显示标注块图标" /></FormItem>
          <FormItem name="backgroundColor" label="背景颜色"><ColorPicker v-model:value="callout.backgroundColor" value-format="hex" show-text aria-label="标注块背景颜色" /></FormItem>
          <FormItem name="borderColor" label="边框颜色"><ColorPicker v-model:value="callout.borderColor" value-format="hex" show-text aria-label="标注块边框颜色" /></FormItem>
          <FormItem label="文字颜色"><ColorPicker :value="textAppearance.color" value-format="hex" show-text aria-label="标注块文字颜色" @update:value="setTextColor" /> <Button type="link" @click="resetTextColor">恢复继承颜色</Button></FormItem>
          <FormItem name="borderStyle" label="边框线型"><Select v-model:value="callout.borderStyle" :options="[{ value: 'solid', label: '实线' }, { value: 'dashed', label: '虚线' }, { value: 'dotted', label: '点线' }]" /></FormItem>
          <FormItem name="borderWidth" label="边框宽度"><InputNumber v-model:value="callout.borderWidth" :min="0" :max="4096" /></FormItem>
          <FormItem name="borderOpacity" label="边框不透明度"><InputNumber v-model:value="callout.borderOpacity" :min="0" :max="1" :step="0.05" /></FormItem>
          <FormItem name="borderRadius" label="圆角半径"><InputNumber v-model:value="callout.borderRadius" :min="0" :max="4096" /></FormItem>
          <FormItem v-for="size in sizes" :key="size.key" :name="size.key" :label="size.label"><InputNumber v-model:value="callout[size.key]" :min="0" :max="4096" /></FormItem>
        </div>
      </Form>
      <Form v-else :model="columns" :disabled="readOnly" layout="vertical">
        <div class="structured-panel__fields">
          <FormItem name="count" label="栏数"><InputNumber v-model:value="columns.count" :min="DOCUMENT_COLUMN_GROUP_MIN_COLUMNS" :max="DOCUMENT_COLUMN_GROUP_MAX_COLUMNS" :precision="0" :disabled="!!block" /></FormItem>
          <FormItem name="gap" label="栏间距"><InputNumber v-model:value="columns.gap" :min="0" :max="4096" /></FormItem>
          <FormItem name="horizontalPadding" label="栏水平内边距（px）"><InputNumber v-model:value="columns.horizontalPadding" :min="0" :max="4096" aria-label="栏水平内边距" /></FormItem>
          <FormItem name="minHeight" label="栏最小高度（px）"><InputNumber v-model:value="columns.minHeight" :min="0" :max="1000000" aria-label="栏最小高度" /></FormItem>
          <FormItem name="layout" label="布局"><Select v-model:value="columns.layout" :options="[{ value: 'fixed', label: '按比例分配' }, { value: 'auto', label: '自动布局' }]" /></FormItem>
          <FormItem name="responsive" label="窄屏布局"><Select v-model:value="columns.responsive" :options="[{ value: 'shrink', label: '收紧栏宽' }, { value: 'stack', label: '纵向排列' }]" /></FormItem>
        </div>
        <div v-for="(_, index) in columns.ratios" :key="columnIds[index] ?? index" class="structured-panel__column">
          <FormItem :name="['ratios', index]" :label="`第 ${index + 1} 栏宽度比例`"><InputNumber v-model:value="columns.ratios[index]" :min="0.01" :max="1000000" :step="0.1" /></FormItem>
          <FormItem :name="['minimums', index]" label="最小宽度"><InputNumber v-model:value="columns.minimums[index]" :min="0" :max="1000000" placeholder="自动" /></FormItem>
          <div v-if="block" class="structured-panel__actions">
            <Tooltip title="向左移动"><Button size="small" :disabled="readOnly || index === 0" @click="command({ id: 'doc.command.move-column', blockId: block.id, columnId: columnIds[index]!, index: index - 1 })">左移</Button></Tooltip>
            <Tooltip title="向右移动"><Button size="small" :disabled="readOnly || index === columns.count - 1" @click="command({ id: 'doc.command.move-column', blockId: block.id, columnId: columnIds[index]!, index: index + 1 })">右移</Button></Tooltip>
            <Tooltip title="在右侧增加一栏"><Button size="small" :disabled="readOnly || columns.count >= DOCUMENT_COLUMN_GROUP_MAX_COLUMNS" @click="command({ id: 'doc.command.add-column', blockId: block.id, columnId: columnIds[index]!, position: 'after' })">添加</Button></Tooltip>
            <Tooltip title="删除这一栏"><Button size="small" danger :disabled="readOnly || columns.count <= DOCUMENT_COLUMN_GROUP_MIN_COLUMNS" @click="command({ id: 'doc.command.remove-column', blockId: block.id, columnId: columnIds[index]!, keepContents })">删除</Button></Tooltip>
          </div>
        </div>
        <FormItem v-if="block" label="删除栏的内容"><Checkbox v-model:checked="keepContents">保留内容并合入相邻栏</Checkbox></FormItem>
      </Form>
    </div>
  </Modal>
</template>

<style scoped>
.structured-panel { display: grid; gap: 12px; color: var(--structure-text); }
.structured-panel__fields { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 0 16px; }
.structured-panel__column { display: grid; grid-template-columns: 1fr 1fr; gap: 0 16px; padding-top: 12px; border-top: 1px solid var(--structure-border); }
.structured-panel__actions { grid-column: 1 / -1; display: flex; gap: 8px; margin-bottom: 12px; }
.structured-panel :deep(.ant-input-number) { width: 100%; }
@media (max-width: 520px) { .structured-panel__fields, .structured-panel__column { grid-template-columns: 1fr; } }
</style>
