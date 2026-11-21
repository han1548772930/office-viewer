<script setup lang="ts">
import { Button, Drawer, Input, InputNumber, Modal, Select, Switch, Tabs } from 'antdv-next'
import { computed, onBeforeUnmount, ref, shallowRef, watch } from 'vue'
import type { OfficeStudio } from '../../api/office-studio'
import type { SlicerModel, SlicerType, TimelineLevel } from '../../domain/model'
import { slicerPalette } from '../../features/slicer/slicer-renderer'
import {
  applyPanelSlicerPresentation, insertPanelSlicer, readSlicerPanelState, removePanelSlicer,
  type SlicerPanelSnapshot,
} from '../../features/slicer/slicer-panel-state'

const props = withDefaults(defineProps<{
  open: boolean
  studio: OfficeStudio
  workbookId: string
  sheetId: string
  initialKind?: SlicerType
  selectedSlicerId?: string
}>(), { initialKind: 'slicer' })
const emit = defineEmits<{
  'update:open': [open: boolean]
  changed: [slicerId: string]
  error: [message: string]
}>()

const snapshot = shallowRef<SlicerPanelSnapshot | null>(null)
const pivotId = ref('')
const fieldId = ref('')
const busy = ref(false)
const error = ref('')
const settingsId = ref('')
const draft = ref<SlicerModel | null>(null)
const originalDraft = ref<SlicerModel | null>(null)
const settingsTab = ref('control')
let unsubscribe: (() => void) | undefined
const target = computed(() => ({ workbookId: props.workbookId, sheetId: props.sheetId }))
const setting = computed(() => snapshot.value?.controls.find(({ info }) => info.id === settingsId.value))
const pivot = computed(() => snapshot.value?.pivots.find(({ id }) => id === pivotId.value))
const pivotOptions = computed(() => snapshot.value?.pivots.map(({ id, label }) => ({ label, value: id })) ?? [])
const fieldOptions = computed(() => pivot.value?.fields.flatMap(({ id, name, available, controls, reason }) => {
  if (!available || !controls.includes(props.initialKind)) return []
  return [{ label: name, value: id, title: reason }]
}) ?? [])
const slicerStyles = ['SlicerStyleLight1', 'SlicerStyleLight2', 'SlicerStyleLight3', 'SlicerStyleLight4', 'SlicerStyleLight5', 'SlicerStyleLight6', 'SlicerStyleOther1', 'SlicerStyleOther2', 'SlicerStyleDark1', 'SlicerStyleDark2', 'SlicerStyleDark3', 'SlicerStyleDark4', 'SlicerStyleDark5', 'SlicerStyleDark6']
const timelineStyles = ['TimelineStyleLight1', 'TimelineStyleLight2', 'TimelineStyleLight3', 'TimelineStyleLight4', 'TimelineStyleLight5', 'TimelineStyleLight6', 'TimelineStyleDark1', 'TimelineStyleDark2', 'TimelineStyleDark3', 'TimelineStyleDark4', 'TimelineStyleDark5', 'TimelineStyleDark6']
const styleOptions = computed(() => (draft.value?.type === 'timeline' ? timelineStyles : slicerStyles).map((value, index) => ({ label: `快速样式 ${index + 1}`, value })))
const settingsTabs = computed(() => [
  { key: 'control', label: draft.value?.type === 'timeline' ? '日程表' : '切片器' },
  { key: 'position', label: '大小与属性' },
])
const settingsDirty = computed(() => Boolean(draft.value && originalDraft.value && JSON.stringify(draft.value) !== JSON.stringify(originalDraft.value)))
const settingsValid = computed(() => Boolean(draft.value?.caption.trim() && draft.value.style && draft.value.layout && Number.isFinite(draft.value.layout.left) && Number.isFinite(draft.value.layout.top) && draft.value.layout.width >= 80 && draft.value.layout.height >= 48 && (draft.value.type !== 'slicer' || Number(draft.value.columnCount) >= 1 && Number(draft.value.rowHeight) > 0)))
const timelineLevels: Array<{ label: string; value: TimelineLevel }> = [
  { label: '年', value: 'year' }, { label: '季度', value: 'quarter' }, { label: '月', value: 'month' }, { label: '日', value: 'day' },
]

function refresh(): void {
  if (!props.open) return
  try {
    snapshot.value = readSlicerPanelState(props.studio, target.value)
    if (!snapshot.value.pivots.some(({ id }) => id === pivotId.value)) pivotId.value = snapshot.value.pivots[0]?.id ?? ''
    if (settingsId.value) {
      const current = snapshot.value.controls.find(({ info }) => info.id === settingsId.value)?.info
      if (current) {
        const layout = current.layout ? { ...current.layout } : { left: 200, top: 200, width: current.type === 'timeline' ? 420 : 220, height: current.type === 'timeline' ? 140 : 280 }
        draft.value = { ...current, layout }; originalDraft.value = { ...current, layout: { ...layout } }
      }
      else { settingsId.value = ''; draft.value = null; originalDraft.value = null }
    }
  } catch (failure) {
    snapshot.value = null
    error.value = failure instanceof Error ? failure.message : String(failure)
  }
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
watch([() => props.open, () => props.workbookId, () => props.sheetId], () => {
  fieldId.value = ''
  settingsId.value = props.selectedSlicerId ?? ''
  draft.value = null
  originalDraft.value = null
  settingsTab.value = 'control'
  error.value = ''
  refresh()
}, { immediate: true })
watch(() => props.selectedSlicerId, (id) => {
  settingsId.value = id ?? ''
  if (settingsId.value) refresh()
})
watch(pivot, (value) => {
  if (!value?.fields.some(({ id, available, controls }) => id === fieldId.value && available && controls.includes(props.initialKind))) {
    fieldId.value = value?.fields.find(({ available, controls }) => available && controls.includes(props.initialKind))?.id ?? ''
  }
})
watch(setting, (value) => {
  if (!value) { draft.value = null; return }
  draft.value = {
    ...value.info,
    layout: value.info.layout ? { ...value.info.layout } : { left: 200, top: 200, width: value.info.type === 'timeline' ? 420 : 220, height: value.info.type === 'timeline' ? 140 : 280 },
    style: value.info.style ?? (value.info.type === 'timeline' ? timelineStyles[0] : slicerStyles[0]),
  }
  originalDraft.value = { ...draft.value, layout: draft.value.layout ? { ...draft.value.layout } : undefined }
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
  const currentTarget = target.value
  const input = { pivotTableId: pivotId.value, sourceFieldId: fieldId.value, type: props.initialKind }
  void perform(async () => {
    const id = await insertPanelSlicer(props.studio, currentTarget, input)
    emit('update:open', false)
    return id
  })
}

function saveSettings(): void {
  if (!draft.value) return
  const currentTarget = target.value
  const { id, type: kind, caption: title, style, columnCount, rowHeight, showCaption, lockedPosition, level, selectionLevel, showHeader, showSelectionLabel, showTimeLevel, showHorizontalScrollbar, layout } = draft.value
  const patch = kind === 'slicer'
    ? { caption: title, style, columnCount, rowHeight, showCaption, lockedPosition, ...(layout ? { layout } : {}) }
    : { caption: title, style, level, selectionLevel, showHeader, showSelectionLabel, showTimeLevel, showHorizontalScrollbar, ...(layout ? { layout } : {}) }
  void perform(async () => { await applyPanelSlicerPresentation(props.studio, currentTarget, id, patch); emit('update:open', false); return id })
}

function close(): void { settingsId.value = ''; emit('update:open', false) }
function stylePreview(style: string): Record<string, string> {
  const palette = slicerPalette(style, style.startsWith('Timeline'))
  return { '--slicer-preview-header': palette.header, '--slicer-preview-item': palette.item, '--slicer-preview-border': palette.itemBorder }
}
function remove(): void {
  if (!settingsId.value) return
  const currentTarget = target.value
  const id = settingsId.value
  void perform(async () => { await removePanelSlicer(props.studio, currentTarget, id); settingsId.value = ''; emit('update:open', false); emit('changed', ''); return undefined })
}
</script>

<template>
  <Modal :open="open && !settingsId" :title="initialKind === 'timeline' ? '插入日程表' : '插入切片器'" :width="400" :footer="null" :mask-closable="false" destroy-on-hidden @cancel="close">
    <div class="slicer-insert">
      <label>字段<Select v-model:value="fieldId" :options="fieldOptions" aria-label="字段" placeholder="选择字段" :disabled="busy || !pivot" /></label>
      <p v-if="!snapshot?.pivots.length" class="slicer-muted">沒有可用的数据透视表。</p>
      <label v-else-if="snapshot.pivots.length > 1">目标透视表<Select v-model:value="pivotId" :options="pivotOptions" aria-label="目标透视表" :disabled="busy" /></label>
      <p v-if="error" role="alert" class="slicer-error">{{ error }}</p>
      <div class="slicer-actions"><Button @click="close">取消</Button><Button type="primary" :loading="busy" :disabled="!fieldId || busy" @click="insert">应用</Button></div>
    </div>
  </Modal>

  <Drawer :open="open && !!settingsId" :title="draft?.type === 'timeline' ? '日程表设置' : '切片器设置'" :width="388" placement="right" destroy-on-hidden @close="close">
    <div v-if="draft" class="slicer-settings">
      <Tabs v-model:active-key="settingsTab" :items="settingsTabs" />
      <template v-if="settingsTab === 'control'">
        <section class="slicer-settings__group">
          <label>标题<Input v-model:value="draft.caption" aria-label="标题" :disabled="busy" /></label>
          <label class="slicer-toggle"><Switch v-if="draft.type === 'slicer'" v-model:checked="draft.showCaption" aria-label="显示页眉" :disabled="busy" /><Switch v-else v-model:checked="draft.showHeader" aria-label="显示页眉" :disabled="busy" />显示页眉</label>
        </section>
        <section v-if="draft.type === 'slicer'" class="slicer-settings__group">
          <label>列<InputNumber v-model:value="draft.columnCount" :min="1" :max="20000" :precision="0" aria-label="列数" :disabled="busy" /></label>
          <label>按钮高度<InputNumber v-model:value="draft.rowHeight" :min="0.01" :max="450678" :precision="2" aria-label="按钮高度" :disabled="busy" /></label>
          <label class="slicer-toggle"><Switch v-model:checked="draft.lockedPosition" aria-label="禁止移动和调整大小" :disabled="busy" />禁止移动和调整大小</label>
        </section>
        <section v-else class="slicer-settings__group">
          <label>时间级别<Select v-model:value="draft.level" :options="timelineLevels" aria-label="时间级别" :disabled="busy" /></label>
          <label>选择级别<Select v-model:value="draft.selectionLevel" :options="timelineLevels" aria-label="选择级别" :disabled="busy" /></label>
          <label class="slicer-toggle"><Switch v-model:checked="draft.showSelectionLabel" aria-label="显示选择标签" :disabled="busy" />显示选择标签</label>
          <label class="slicer-toggle"><Switch v-model:checked="draft.showTimeLevel" aria-label="显示时间级别" :disabled="busy" />显示时间级别</label>
          <label class="slicer-toggle"><Switch v-model:checked="draft.showHorizontalScrollbar" aria-label="显示滚动条" :disabled="busy" />显示滚动条</label>
        </section>
        <section class="slicer-settings__group">
          <label>快速样式</label>
          <div class="slicer-style-gallery" role="group" aria-label="快速样式">
            <Button v-for="option in styleOptions" :key="option.value" class="slicer-style-option" :class="{ 'is-selected': draft.style === option.value }" :aria-pressed="draft.style === option.value" :disabled="busy" @click="draft.style = option.value">
              <span class="slicer-style-preview" :style="stylePreview(option.value)"><i /><i /><i /></span>
              <span>{{ option.label }}</span>
            </Button>
          </div>
        </section>
      </template>
      <template v-else>
        <section class="slicer-settings__group">
          <label>水平位置<InputNumber v-model:value="draft.layout!.left" :min="0" :max="1000000" :precision="0" aria-label="水平位置" :disabled="busy || draft.lockedPosition" /></label>
          <label>垂直位置<InputNumber v-model:value="draft.layout!.top" :min="0" :max="1000000" :precision="0" aria-label="垂直位置" :disabled="busy || draft.lockedPosition" /></label>
          <label>宽度<InputNumber v-model:value="draft.layout!.width" :min="80" :max="100000" :precision="0" aria-label="宽度" :disabled="busy || draft.lockedPosition" /></label>
          <label>高度<InputNumber v-model:value="draft.layout!.height" :min="48" :max="100000" :precision="0" aria-label="高度" :disabled="busy || draft.lockedPosition" /></label>
          <label v-if="draft.type === 'slicer'" class="slicer-toggle"><Switch v-model:checked="draft.lockedPosition" aria-label="禁止移动和调整大小" :disabled="busy" />禁止移动和调整大小</label>
        </section>
      </template>
      <p v-if="error" role="alert" class="slicer-error">{{ error }}</p>
      <div class="slicer-actions"><Button @click="remove">删除</Button><span /><Button @click="close">取消</Button><Button type="primary" :loading="busy" :disabled="busy || !settingsValid || !settingsDirty" @click="saveSettings">应用</Button></div>
    </div>
  </Drawer>
</template>

<style scoped>
.slicer-insert, .slicer-settings { display: flex; flex-direction: column; gap: 16px; }
.slicer-insert label, .slicer-settings label:not(.slicer-toggle) { display: flex; flex-direction: column; gap: 7px; color: #596779; font-size: 13px; }
.slicer-insert :deep(.ant-select), .slicer-settings :deep(.ant-select), .slicer-settings :deep(.ant-input-number) { width: 100%; }
.slicer-settings__group { display: flex; flex-direction: column; gap: 13px; padding: 14px; border: 1px solid #e3e7ec; border-radius: 8px; }
.slicer-style-gallery { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 8px; }
.slicer-style-option { height: auto; min-width: 0; min-height: 62px; display: flex; flex-direction: column; align-items: stretch; gap: 5px; padding: 5px; font-size: 11px; }
.slicer-style-option.is-selected { border-color: #246bfd; box-shadow: inset 0 0 0 1px #246bfd; }
.slicer-style-preview { display: grid; grid-template-columns: repeat(3, 1fr); gap: 2px; height: 25px; padding: 3px; background: var(--slicer-preview-header); }
.slicer-style-preview i { display: block; background: var(--slicer-preview-item); border: 1px solid var(--slicer-preview-border); }
.slicer-actions > span { flex: 1; }
.slicer-toggle { display: flex; align-items: center; gap: 9px; min-height: 24px; color: #344152; font-size: 13px; }
.slicer-actions { display: flex; justify-content: flex-end; gap: 8px; }
.slicer-muted, .slicer-error { margin: 0; font-size: 12px; line-height: 1.6; }
.slicer-muted { color: #7d8896; }
.slicer-error { color: #c43737; }
</style>
