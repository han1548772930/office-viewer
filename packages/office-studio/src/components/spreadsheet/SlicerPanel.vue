<script setup lang="ts">
import { Button, Checkbox, Drawer, Form, FormItem, Input, InputNumber, Modal, Popover, Segmented, Select, Tooltip } from 'antdv-next'
import { computed, onBeforeUnmount, ref, shallowRef, watch } from 'vue'
import { ArrowDown, ArrowDownToLine, ArrowUp, ArrowUpToLine, ChevronDown } from 'lucide-vue-next'
import type { OfficeStudio } from '../../api/office-studio'
import type { CellRange, SlicerModel, SlicerType, TimelineLevel } from '../../domain/model'
import { slicerPalette } from '../../features/slicer/slicer-renderer'
import {
  applyPanelSlicerPresentation, insertPanelSlicer, readSlicerPanelState,
  type SlicerPanelSnapshot,
} from '../../features/slicer/slicer-panel-state'
import type { SlicerPresentationPatch } from '../../features/slicer/slicer-engine'

const props = withDefaults(defineProps<{
  open: boolean
  studio: OfficeStudio
  workbookId: string
  sheetId: string
  sourceRange?: CellRange
  initialKind?: SlicerType
  selectedSlicerId?: string
}>(), { initialKind: 'slicer' })
const emit = defineEmits<{
  'update:open': [open: boolean]
  changed: [slicerId: string]
  error: [message: string]
}>()

const snapshot = shallowRef<SlicerPanelSnapshot | null>(null)
const fieldId = ref('')
const busy = ref(false)
const error = ref('')
const settingsId = ref('')
const draft = ref<SlicerModel | null>(null)
const originalDraft = ref<SlicerModel | null>(null)
const settingsPage = ref('slicer')
const styleGalleryOpen = ref(false)
let unsubscribe: (() => void) | undefined
const target = computed(() => ({ workbookId: props.workbookId, sheetId: props.sheetId }))
const setting = computed(() => snapshot.value?.controls.find(({ info }) => info.id === settingsId.value))
const arrangement = computed(() => {
  const controls = snapshot.value?.controls ?? []
  const ordered = controls.map(({ info }, index) => ({ id: info.id, zIndex: info.drawing?.zIndex ?? index }))
    .sort((first, second) => first.zIndex - second.zIndex)
  const index = ordered.findIndex(({ id }) => id === settingsId.value)
  return {
    forward: index >= 0 && index < ordered.length - 1,
    backward: index > 0,
    front: index >= 0 && index < ordered.length - 1,
    back: index > 0,
  }
})
const insertionSource = computed(() => {
  const range = props.sourceRange
  if (!range) return undefined
  return snapshot.value?.pivots.find((source) => source.sheetId === props.sheetId
    && range.startRow >= source.range.startRow && range.endRow <= source.range.endRow
    && range.startCol >= source.range.startCol && range.endCol <= source.range.endCol
    && (props.initialKind !== 'timeline' || source.kind === 'pivot'))
})
const fieldOptions = computed(() => insertionSource.value?.fields.flatMap(({ id, name, available, controls, reason }) => {
  if (!available || !controls.includes(props.initialKind)) return []
  return [{ label: name, value: id, title: reason }]
}) ?? [])
const styleNames = {
  slicer: ['SlicerStyleLight1', 'SlicerStyleLight2', 'SlicerStyleLight3', 'SlicerStyleLight4', 'SlicerStyleLight5', 'SlicerStyleLight6', 'SlicerStyleOther1', 'SlicerStyleOther2', 'SlicerStyleDark1', 'SlicerStyleDark2', 'SlicerStyleDark3', 'SlicerStyleDark4', 'SlicerStyleDark5', 'SlicerStyleDark6'],
  timeline: ['TimelineStyleLight1', 'TimelineStyleLight2', 'TimelineStyleLight3', 'TimelineStyleLight4', 'TimelineStyleLight5', 'TimelineStyleLight6', 'TimelineStyleDark1', 'TimelineStyleDark2', 'TimelineStyleDark3', 'TimelineStyleDark4', 'TimelineStyleDark5', 'TimelineStyleDark6'],
}
const styleOptions = computed(() => [
  { label: '文档默认值', value: '' },
  ...(draft.value?.type === 'timeline' ? styleNames.timeline : styleNames.slicer)
    .map((value, index) => ({ label: `样式 ${index + 1}`, value })),
])
const selectedStyleLabel = computed(() => {
  const style = draft.value?.style ?? ''
  return styleOptions.value.find(option => option.value === style)?.label ?? '样式'
})
const settingsDirty = computed(() => Boolean(draft.value && originalDraft.value
  && JSON.stringify(presentation(draft.value)) !== JSON.stringify(presentation(originalDraft.value))))
const settingsValid = computed(() => {
  const current = draft.value
  const layout = current?.layout
  if (!current || !layout) return false
  return Boolean(current.caption.trim()
    && Number.isFinite(layout.left) && Number.isFinite(layout.top)
    && layout.width >= 80 && layout.height >= 48
    && (current.type !== 'slicer' || Number(current.columnCount) >= 1 && Number(current.rowHeight) > 0))
})
const timelineLevels: Array<{ label: string; value: TimelineLevel }> = [
  { label: '年', value: 'year' }, { label: '季度', value: 'quarter' }, { label: '月', value: 'month' }, { label: '日', value: 'day' },
]

function presentation(model: SlicerModel): SlicerPresentationPatch {
  const { caption, style, columnCount, rowHeight, showCaption, lockedPosition, level, selectionLevel, showHeader, showSelectionLabel, showTimeLevel, showHorizontalScrollbar, layout } = model
  return {
    caption, style, columnCount, rowHeight, showCaption, lockedPosition, level, selectionLevel,
    showHeader, showSelectionLabel, showTimeLevel, showHorizontalScrollbar, layout,
    anchorType: model.drawing?.anchorType ?? 'position',
    keepRatio: model.drawing?.keepRatio ?? false,
  }
}

function refresh(): void {
  if (!props.open) return
  try {
    snapshot.value = readSlicerPanelState(props.studio, target.value)
    const current = snapshot.value.controls.find(({ info }) => info.id === settingsId.value)?.info
    if (!settingsId.value) return
    if (!current) {
      settingsId.value = ''
      draft.value = null
      originalDraft.value = null
    } else if (!settingsDirty.value) {
      loadDraft(current)
    }
  } catch (failure) {
    snapshot.value = null
    error.value = failure instanceof Error ? failure.message : String(failure)
  }
}

function loadDraft(info: SlicerModel): void {
  const layout = info.layout ? { ...info.layout } : { left: 200, top: 200, width: info.type === 'timeline' ? 420 : 220, height: info.type === 'timeline' ? 140 : 280 }
  draft.value = { ...info, layout }
  originalDraft.value = { ...info, layout: { ...layout } }
}

watch(() => props.studio, (studio) => {
  unsubscribe?.()
  const subscriptions = [
    studio.events.addEvent('CommandExecuted', (event) => { if (event.unitId === props.workbookId) refresh() }),
    studio.events.addEvent('Undo', (event) => { if (event.unitId === props.workbookId) refresh() }),
    studio.events.addEvent('Redo', (event) => { if (event.unitId === props.workbookId) refresh() }),
    studio.events.addEvent('WorkbookDisposed', (event) => { if (event.workbookId === props.workbookId) refresh() }),
  ]
  const stop = watch(() => studio.editor.workbook.model, refresh, { flush: 'post' })
  unsubscribe = () => { subscriptions.forEach((subscription) => subscription.dispose()); stop() }
  refresh()
}, { immediate: true })
watch([() => props.open, () => props.workbookId, () => props.sheetId, () => props.sourceRange, () => props.initialKind], () => {
  fieldId.value = ''
  settingsId.value = props.selectedSlicerId ?? ''
  draft.value = null
  originalDraft.value = null
  settingsPage.value = props.initialKind === 'timeline' ? 'timeline' : 'slicer'
  error.value = ''
  refresh()
}, { immediate: true })
watch(() => props.selectedSlicerId, (id) => {
  settingsId.value = id ?? ''
  if (settingsId.value) refresh()
})
watch(insertionSource, (source) => {
  if (!source?.fields.some(({ id, available, controls }) => id === fieldId.value && available && controls.includes(props.initialKind))) {
    fieldId.value = source?.fields.find(({ available, controls }) => available && controls.includes(props.initialKind))?.id ?? ''
  }
})
watch(setting, (value) => {
  if (!value || !settingsId.value || settingsDirty.value) return
  loadDraft(value.info)
})
onBeforeUnmount(() => unsubscribe?.())

async function perform(action: () => Promise<string | void>): Promise<void> {
  if (busy.value) return
  busy.value = true
  error.value = ''
  try {
    const id = await action()
    refresh()
    if (id) emit('changed', id)
  } catch (failure) {
    error.value = failure instanceof Error ? failure.message : String(failure)
    emit('error', error.value)
  } finally { busy.value = false }
}

function insert(): void {
  const source = insertionSource.value
  if (!source) return
  const input = { ...(source.kind === 'pivot' ? { pivotTableId: source.id } : { tableId: source.id }), sourceFieldId: fieldId.value, type: props.initialKind }
  void perform(async () => {
    const id = await insertPanelSlicer(props.studio, target.value, input)
    emit('update:open', false)
    return id
  })
}

function saveSettings(): void {
  if (!draft.value) return
  const current = draft.value
  const patch = presentation(current)
  void perform(async () => {
    await applyPanelSlicerPresentation(props.studio, target.value, current.id, patch)
    originalDraft.value = { ...current, layout: current.layout ? { ...current.layout } : undefined }
    return current.id
  })
}

function close(): void { settingsId.value = ''; emit('update:open', false) }
function stylePreview(style: string): Record<string, string> {
  const palette = slicerPalette(style || (draft.value?.type === 'timeline' ? 'TimelineStyleLight1' : 'SlicerStyleLight1'), draft.value?.type === 'timeline')
  return { '--preview-header': palette.header, '--preview-item': palette.item, '--preview-border': palette.itemBorder, '--preview-text': palette.text }
}
function chooseStyle(style: string): void {
  if (draft.value) draft.value.style = style || undefined
  styleGalleryOpen.value = false
}

function setAnchorType(anchorType: unknown): void {
  if (!draft.value || (anchorType !== 'both' && anchorType !== 'position' && anchorType !== 'none')) return
  draft.value.drawing = {
    ...(draft.value.drawing ?? { from: { row: 0, column: 0, offsetX: 0, offsetY: 0 }, to: { row: 0, column: 0, offsetX: 0, offsetY: 0 } }),
    anchorType,
  }
}

function setKeepRatio(value: boolean | { target?: { checked?: boolean } }): void {
  if (!draft.value) return
  draft.value.drawing = {
    ...(draft.value.drawing ?? { anchorType: 'position', from: { row: 0, column: 0, offsetX: 0, offsetY: 0 }, to: { row: 0, column: 0, offsetX: 0, offsetY: 0 } }),
    keepRatio: typeof value === 'boolean' ? value : Boolean(value.target?.checked),
  }
}

function updateSizeDimension(dimension: 'width' | 'height', value: unknown): void {
  const current = draft.value
  const layout = current?.layout
  if (!current || !layout || !current.drawing?.keepRatio || typeof value !== 'number' || !Number.isFinite(value)) return
  const original = originalDraft.value?.layout ?? layout
  if (original.width <= 0 || original.height <= 0) return
  if (dimension === 'width') layout.height = Math.min(100000, Math.max(48, Math.round(value * original.height / original.width)))
  else layout.width = Math.min(100000, Math.max(80, Math.round(value * original.width / original.height)))
}

function arrange(action: 'forward' | 'backward' | 'front' | 'back'): void {
  if (!draft.value) return
  const slicerId = draft.value.id
  void perform(async () => {
    await applyPanelSlicerPresentation(props.studio, target.value, slicerId, { arrange: action })
    return slicerId
  })
}
</script>

<template>
  <Modal :open="open && !settingsId" :title="initialKind === 'timeline' ? '插入时间线' : '插入切片器'" :width="400" :footer="null" :mask-closable="false" destroy-on-hidden @cancel="close">
    <Form class="slicer-insert" layout="vertical" @submit.prevent="insert">
      <FormItem label="字段"><Select v-model:value="fieldId" :options="fieldOptions" aria-label="字段" placeholder="选择字段" :disabled="busy || !insertionSource" /></FormItem>
      <p v-if="!insertionSource" class="slicer-muted">请选择数据透视表或表格中的单元格</p>
      <p v-else-if="!fieldOptions.length" class="slicer-muted">当前选区没有可用字段</p>
      <p v-if="error" role="alert" class="slicer-error">{{ error }}</p>
      <div class="slicer-actions"><Button @click="close">取消</Button><Button type="primary" :loading="busy" :disabled="!fieldId || busy" @click="insert">确定</Button></div>
    </Form>
  </Modal>

  <Drawer :open="open && !!settingsId" :title="draft?.type === 'timeline' ? '时间线设置' : '切片器设置'" :size="388" placement="right" :mask="false" destroy-on-hidden @close="close">
    <Form v-if="draft" class="slicer-settings" layout="vertical" @submit.prevent="saveSettings">
      <Segmented v-model:value="settingsPage" :options="[{ value: draft.type, label: draft.type === 'timeline' ? '时间线' : '切片器' }, { value: 'size', label: '大小与属性' }]" block />
      <template v-if="settingsPage !== 'size'">
        <section class="slicer-group">
          <header>切片器</header>
          <FormItem label="标题"><Input v-model:value="draft.caption" aria-label="标题" :disabled="busy" /></FormItem>
          <FormItem :label="null"><Checkbox v-if="draft.type === 'slicer'" v-model:checked="draft.showCaption" :disabled="busy">显示标题</Checkbox><Checkbox v-else v-model:checked="draft.showHeader" :disabled="busy">显示标题</Checkbox></FormItem>
          <template v-if="draft.type === 'slicer'">
            <FormItem class="slicer-number" layout="horizontal" label="列数"><InputNumber v-model:value="draft.columnCount" :min="1" :max="20000" :precision="0" aria-label="列数" :disabled="busy" /></FormItem>
            <FormItem class="slicer-number" layout="horizontal" label="按钮高度"><InputNumber v-model:value="draft.rowHeight" :min="0.01" :max="450678" :precision="2" aria-label="按钮高度" :disabled="busy" /></FormItem>
            <FormItem :label="null"><Checkbox v-model:checked="draft.lockedPosition" :disabled="busy">锁定位置</Checkbox></FormItem>
          </template>
          <template v-else>
            <FormItem label="时间级别"><Select v-model:value="draft.level" :options="timelineLevels" aria-label="时间级别" :disabled="busy" /></FormItem>
            <FormItem label="选择级别"><Select v-model:value="draft.selectionLevel" :options="timelineLevels" aria-label="选择级别" :disabled="busy" /></FormItem>
            <FormItem :label="null"><Checkbox v-model:checked="draft.showSelectionLabel" :disabled="busy">显示选择标签</Checkbox></FormItem>
            <FormItem :label="null"><Checkbox v-model:checked="draft.showTimeLevel" :disabled="busy">显示时间级别</Checkbox></FormItem>
            <FormItem :label="null"><Checkbox v-model:checked="draft.showHorizontalScrollbar" :disabled="busy">显示滚动条</Checkbox></FormItem>
          </template>
        </section>
        <section class="slicer-group">
          <header>样式</header>
          <Popover v-model:open="styleGalleryOpen" trigger="click" placement="bottomRight">
            <template #content>
              <div class="slicer-gallery" role="group" aria-label="样式图库">
                <Button v-for="option in styleOptions" :key="option.value || 'default'" html-type="button" class="slicer-style-tile" :class="{ 'is-selected': (draft.style ?? '') === option.value }" :style="stylePreview(option.value)" :title="option.label" :aria-label="option.label" :aria-pressed="(draft.style ?? '') === option.value" :disabled="busy" @click="chooseStyle(option.value)">
                  <span v-if="option.value" class="slicer-style-sample"><i /><i /><i /></span><span v-else class="slicer-default-sample">Aa</span>
                </Button>
              </div>
            </template>
            <Button html-type="button" class="slicer-style-trigger" :style="stylePreview(draft.style ?? '')" aria-label="选择样式" :disabled="busy">
              <span class="slicer-style-sample"><i /><i /><i /></span><span>{{ selectedStyleLabel }}</span><ChevronDown :size="14" />
            </Button>
          </Popover>
        </section>
      </template>
      <template v-else>
        <section class="slicer-group">
          <header>大小与属性</header>
          <FormItem class="slicer-number" layout="horizontal" label="水平位置"><InputNumber v-model:value="draft.layout!.left" :min="0" :max="1000000" :precision="0" aria-label="水平位置" :disabled="busy || draft.lockedPosition" /></FormItem>
          <FormItem class="slicer-number" layout="horizontal" label="垂直位置"><InputNumber v-model:value="draft.layout!.top" :min="0" :max="1000000" :precision="0" aria-label="垂直位置" :disabled="busy || draft.lockedPosition" /></FormItem>
          <FormItem class="slicer-number" layout="horizontal" label="宽度"><InputNumber v-model:value="draft.layout!.width" :min="80" :max="100000" :precision="0" aria-label="宽度" :disabled="busy || draft.lockedPosition" @change="updateSizeDimension('width', $event)" /></FormItem>
          <FormItem class="slicer-number" layout="horizontal" label="高度"><InputNumber v-model:value="draft.layout!.height" :min="48" :max="100000" :precision="0" aria-label="高度" :disabled="busy || draft.lockedPosition" @change="updateSizeDimension('height', $event)" /></FormItem>
          <FormItem label="锚定方式"><Select :value="draft.drawing?.anchorType ?? 'position'" :options="[
            { value: 'both', label: '随单元格移动并调整大小' },
            { value: 'position', label: '随单元格移动，不调整大小' },
            { value: 'none', label: '不随单元格移动或调整大小' },
          ]" aria-label="锚定方式" :disabled="busy" @change="setAnchorType" /></FormItem>
          <FormItem :label="null"><Checkbox :checked="draft.drawing?.keepRatio ?? false" :disabled="busy" @change="setKeepRatio">保持纵横比</Checkbox></FormItem>
          <FormItem :label="null"><Checkbox v-model:checked="draft.lockedPosition" :disabled="busy">锁定位置</Checkbox></FormItem>
          <div class="slicer-arrange" role="group" aria-label="对象排列">
            <Tooltip title="置于顶层"><Button html-type="button" aria-label="置于顶层" :disabled="busy || !arrangement.front" @click="arrange('front')"><ArrowUpToLine :size="15" /></Button></Tooltip>
            <Tooltip title="上移一层"><Button html-type="button" aria-label="上移一层" :disabled="busy || !arrangement.forward" @click="arrange('forward')"><ArrowUp :size="15" /></Button></Tooltip>
            <Tooltip title="下移一层"><Button html-type="button" aria-label="下移一层" :disabled="busy || !arrangement.backward" @click="arrange('backward')"><ArrowDown :size="15" /></Button></Tooltip>
            <Tooltip title="置于底层"><Button html-type="button" aria-label="置于底层" :disabled="busy || !arrangement.back" @click="arrange('back')"><ArrowDownToLine :size="15" /></Button></Tooltip>
          </div>
        </section>
      </template>
      <p v-if="error" role="alert" class="slicer-error">{{ error }}</p>
      <div class="slicer-actions"><span /><Button @click="close">取消</Button><Button type="primary" :loading="busy" :disabled="busy || !settingsValid || !settingsDirty" @click="saveSettings">确定</Button></div>
    </Form>
  </Drawer>
</template>

<style scoped>
.slicer-insert, .slicer-settings { display: flex; flex-direction: column; gap: 16px; }
.slicer-insert :deep(.ant-form-item), .slicer-settings :deep(.ant-form-item) { margin-bottom: 0; }
.slicer-settings :deep(.ant-form-item-label) { color: #596779; font-size: 13px; }
.slicer-settings :deep(.slicer-number .ant-form-item-row) { display: grid; grid-template-columns: minmax(0, 1fr) 128px; align-items: center; gap: 8px; }
.slicer-settings :deep(.slicer-number .ant-form-item-label) { text-align: start; }
.slicer-insert :deep(.ant-select), .slicer-settings :deep(.ant-select), .slicer-settings :deep(.ant-input-number) { width: 100%; }
.slicer-group { display: flex; flex-direction: column; gap: 12px; padding: 0 0 16px; border-bottom: 1px solid #e5e8ec; }
.slicer-group header { color: #252a31; font-size: 14px; font-weight: 600; }
.slicer-number :deep(.ant-input-number) { width: 128px; }
.slicer-arrange { display: flex; gap: 4px; }
.slicer-arrange .ant-btn { flex: 1; }
.slicer-gallery { width: 320px; max-height: 70vh; overflow: auto; display: grid; grid-template-columns: repeat(3, 72px); justify-content: space-between; gap: 8px; }
.slicer-style-tile { width: 72px; height: 72px; padding: 9px; border: 1px solid #d9dde3; background: #fff; }
.slicer-style-tile.is-selected { outline: 2px solid #1677ff; outline-offset: 1px; }
.slicer-style-sample { width: 100%; height: 100%; display: grid; grid-template-rows: 1fr 1fr 1fr; gap: 3px; padding: 4px; background: var(--preview-header); border: 1px solid var(--preview-border); }
.slicer-style-sample i { display: block; background: var(--preview-item); border: 1px solid var(--preview-border); }
.slicer-style-trigger { width: 100%; height: 48px; display: grid; grid-template-columns: 36px 1fr auto; align-items: center; gap: 10px; padding: 5px 9px; text-align: left; color: var(--preview-text); border: 1px solid #d9dde3; background: #fff; }
.slicer-style-trigger .slicer-style-sample { width: 36px; height: 36px; }
.slicer-default-sample { width: 100%; height: 100%; display: grid; place-items: center; font-weight: 600; color: var(--preview-text); background: var(--preview-item); }
.slicer-style-tile .slicer-default-sample { font-size: 20px; }
.slicer-actions { display: flex; justify-content: flex-end; gap: 8px; margin-top: auto; }
.slicer-actions > span { flex: 1; }
.slicer-muted, .slicer-error { margin: 0; font-size: 12px; line-height: 1.6; }
.slicer-muted { color: #7d8896; }
.slicer-error { color: #c43737; }
</style>
