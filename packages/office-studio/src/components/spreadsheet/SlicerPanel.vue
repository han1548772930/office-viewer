<script setup lang="ts">
import { Button, Checkbox, Input, Modal, Select } from 'antdv-next'
import { computed, onBeforeUnmount, ref, shallowRef, watch } from 'vue'
import type { OfficeStudio } from '../../api/office-studio'
import type { SlicerType, TimelineLevel } from '../../domain/model'
import {
  applyPanelSlicerItems, applyPanelTimelineRange, clearPanelSlicer, insertPanelSlicer,
  readSlicerPanelState, removePanelSlicer, timelineValueToInput,
  type SlicerPanelSnapshot,
} from '../../features/slicer/slicer-panel-state'

const props = withDefaults(defineProps<{
  open: boolean
  studio: OfficeStudio
  workbookId: string
  sheetId: string
  initialKind?: SlicerType
}>(), { initialKind: 'slicer' })
const emit = defineEmits<{
  'update:open': [open: boolean]
  changed: [slicerId: string]
  error: [message: string]
}>()

const snapshot = shallowRef<SlicerPanelSnapshot | null>(null)
const selectedId = ref('')
const adding = ref(true)
const type = ref<SlicerType>(props.initialKind)
const pivotId = ref('')
const fieldId = ref('')
const caption = ref('')
const selectedValues = ref<string[]>([])
const level = ref<TimelineLevel>('month')
const start = ref('')
const end = ref('')
const busy = ref(false)
const error = ref('')
let generation = 0
let unsubscribe: (() => void) | undefined
const target = computed(() => ({ workbookId: props.workbookId, sheetId: props.sheetId }))
const selected = computed(() => snapshot.value?.controls.find((control) => control.info.id === selectedId.value))
const pivot = computed(() => snapshot.value?.pivots.find((candidate) => candidate.id === pivotId.value))
const field = computed(() => pivot.value?.fields.find((candidate) => candidate.id === fieldId.value))
const pivotOptions = computed(() => snapshot.value?.pivots.map((candidate) => ({ label: candidate.label, value: candidate.id })) ?? [])
const fieldOptions = computed(() => pivot.value?.fields.map((candidate) => ({ label: `${candidate.name}${candidate.reason ? `（${candidate.reason}）` : ''}`, value: candidate.id, disabled: !candidate.available })) ?? [])
const levels = [{ label: '日', value: 'day' }, { label: '月', value: 'month' }, { label: '季度', value: 'quarter' }, { label: '年', value: 'year' }]
const kinds = [{ label: '切片器', value: 'slicer' }, { label: '时间线', value: 'timeline' }]

function messageFor(failure: unknown): string {
  const message = failure instanceof Error ? failure.message : String(failure)
  if (/Timeline requires a date field/.test(message)) return '所选字段没有有效日期。请选择包含日期的源字段。'
  if (/Slicer change was canceled|Command was canceled/.test(message)) return '当前权限或命令规则拒绝了这次操作。'
  return message
}

function refresh(): void {
  if (!props.open) return
  try {
    snapshot.value = readSlicerPanelState(props.studio, target.value)
    if (!snapshot.value.pivots.some((candidate) => candidate.id === pivotId.value)) pivotId.value = snapshot.value.pivots[0]?.id ?? ''
    if (selectedId.value && !snapshot.value.controls.some((control) => control.info.id === selectedId.value)) {
      selectedId.value = ''
      adding.value = true
    }
  } catch (failure) {
    snapshot.value = null
    error.value = messageFor(failure)
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
  generation += 1
  selectedId.value = ''
  adding.value = true
  error.value = ''
  caption.value = ''
  type.value = props.initialKind
  refresh()
}, { immediate: true })
watch(() => props.initialKind, (value) => { type.value = value; adding.value = true })
watch(pivot, (value) => {
  if (!value?.fields.some((candidate) => candidate.id === fieldId.value && candidate.available)) fieldId.value = value?.fields.find((candidate) => candidate.available)?.id ?? ''
})
watch(selected, (value) => {
  if (!value) return
  selectedValues.value = value.selectedItems ? [...value.selectedItems] : [...value.items]
  level.value = value.info.timelineRange?.level ?? 'month'
  start.value = timelineValueToInput(value.info.timelineRange?.start, snapshot.value?.date1904 ?? false)
  end.value = timelineValueToInput(value.info.timelineRange?.end, snapshot.value?.date1904 ?? false)
})
onBeforeUnmount(() => { generation += 1; unsubscribe?.() })

async function perform(action: () => Promise<string | void>): Promise<void> {
  if (busy.value) return
  const version = generation
  busy.value = true
  error.value = ''
  try {
    const changedId = await action()
    if (version !== generation) return
    refresh()
    if (changedId) selectedId.value = changedId
    adding.value = false
    emit('changed', changedId || selectedId.value)
  } catch (failure) {
    if (version !== generation) return
    error.value = messageFor(failure)
    emit('error', error.value)
  } finally {
    busy.value = false
  }
}

function insert(): void {
  const currentTarget = target.value
  void perform(() => insertPanelSlicer(props.studio, currentTarget, { pivotTableId: pivotId.value, sourceFieldId: fieldId.value, type: type.value, caption: caption.value }))
}
function toggleValue(item: string): void {
  const values = new Set(selectedValues.value)
  if (values.has(item)) values.delete(item)
  else values.add(item)
  selectedValues.value = [...values]
}
function applyValues(): void {
  const currentTarget = target.value
  const id = selectedId.value
  const values = [...selectedValues.value]
  void perform(() => applyPanelSlicerItems(props.studio, currentTarget, id, values))
}
function applyTimeline(): void {
  const currentTarget = target.value
  const id = selectedId.value
  const range = { start: start.value, end: end.value, level: level.value }
  void perform(() => applyPanelTimelineRange(props.studio, currentTarget, id, range))
}
function clear(): void {
  const currentTarget = target.value
  const id = selectedId.value
  void perform(() => clearPanelSlicer(props.studio, currentTarget, id))
}
function remove(): void {
  const currentTarget = target.value
  const id = selectedId.value
  void perform(async () => { await removePanelSlicer(props.studio, currentTarget, id); adding.value = true })
}
function selectControl(id: string): void { selectedId.value = id; adding.value = false; error.value = '' }
function close(): void { emit('update:open', false) }
</script>

<template>
  <Modal :open="open" title="切片器与时间线" :width="760" :footer="null" destroy-on-hidden @cancel="close">
    <div class="slicer-panel" :aria-busy="busy">
      <section class="slicer-panel__list" aria-label="当前工作表切片器列表">
        <div class="slicer-panel__list-header"><span>当前工作表（{{ snapshot?.controls.length ?? 0 }}）</span><Button size="small" :disabled="busy" @click="adding = true; error = ''">新增</Button></div>
        <Button v-for="control in snapshot?.controls ?? []" :key="control.info.id" type="text" block class="slicer-panel__control" :class="{ 'is-selected': selectedId === control.info.id && !adding }" :aria-pressed="selectedId === control.info.id && !adding" :disabled="busy" @click="selectControl(control.info.id)">
          <span>{{ control.info.caption }}</span><small>{{ control.info.type === 'timeline' ? '时间线' : '切片器' }} · {{ control.selectedItems ? `已选 ${control.selectedItems.length} 项` : '全部数据' }}</small>
        </Button>
        <p v-if="!snapshot?.controls.length" class="slicer-panel__muted">此工作表还没有切片器或时间线。</p>
      </section>
      <section v-if="adding" class="slicer-panel__editor" aria-label="插入切片器配置">
        <h3>插入筛选控件</h3>
        <label>类型<Select v-model:value="type" :options="kinds" aria-label="筛选控件类型" :disabled="busy" /></label>
        <label>目标透视表<Select v-model:value="pivotId" :options="pivotOptions" aria-label="目标透视表" placeholder="选择数据透视表" :disabled="busy" /></label>
        <label>源字段<Select v-model:value="fieldId" :options="fieldOptions" aria-label="切片器源字段" placeholder="选择行、列或筛选字段" :disabled="busy || !pivot" /></label>
        <label>标题<Input v-model:value="caption" aria-label="切片器标题" :placeholder="field?.name || '使用源字段名称'" :disabled="busy" /></label>
        <p v-if="!snapshot?.pivots.length" class="slicer-panel__muted">工作簿中还没有数据透视表。请先从“数据透视表”入口创建，并将要筛选的字段放入行、列或筛选区域，再来插入切片器或时间线。</p>
        <p v-else-if="pivot && !pivot.fields.some((candidate) => candidate.available)" class="slicer-panel__muted">此透视表没有可绑定的新字段。请在透视表中配置标签字段，或选择已有控件进行筛选。</p>
        <p v-if="type === 'timeline'" class="slicer-panel__muted">时间线使用日期字段，支持日、月、季度和年。日期范围将按工作簿日期系统筛选。</p>
        <Button type="primary" :loading="busy" :disabled="!field?.available || busy" @click="insert">插入{{ type === 'timeline' ? '时间线' : '切片器' }}</Button>
      </section>
      <section v-else-if="selected" class="slicer-panel__editor" aria-label="切片器筛选设置">
        <h3>{{ selected.info.caption }}</h3>
        <p v-if="selected.error" role="alert" class="slicer-panel__error">{{ selected.error }}</p>
        <template v-else-if="selected.info.type === 'slicer'">
          <div class="slicer-panel__actions"><Button size="small" :disabled="busy" @click="selectedValues = [...selected.items]">全选</Button><Button size="small" :disabled="busy" @click="selectedValues = []">全不选</Button><span>已选 {{ selectedValues.length }} 项</span></div>
          <div class="slicer-panel__items" role="group" :aria-label="`${selected.info.caption}筛选值`">
            <Checkbox v-for="item in selected.items" :key="item" :checked="selectedValues.includes(item)" :disabled="busy" @change="toggleValue(item)">{{ item }}</Checkbox>
          </div>
          <p v-if="!selected.items.length" class="slicer-panel__muted">源字段中没有可用的值。</p>
          <Button type="primary" :loading="busy" :disabled="busy" @click="applyValues">应用筛选</Button>
        </template>
        <template v-else>
          <label>日期粒度<Select v-model:value="level" :options="levels" aria-label="时间线粒度" :disabled="busy" /></label>
          <label>开始日期<Input v-model:value="start" type="date" aria-label="时间线开始日期" :disabled="busy" /></label>
          <label>结束日期<Input v-model:value="end" type="date" aria-label="时间线结束日期" :disabled="busy" /></label>
          <p class="slicer-panel__muted">边界按所选粒度包含整个日、月、季度或年。留空的一端不限范围。</p>
          <Button type="primary" :loading="busy" :disabled="busy" @click="applyTimeline">应用日期范围</Button>
        </template>
        <div class="slicer-panel__footer"><Button :disabled="busy || !!selected.error" @click="clear">清除筛选</Button><Button danger :disabled="busy" @click="remove">删除控件</Button></div>
      </section>
      <section v-else class="slicer-panel__editor"><p class="slicer-panel__muted">选择当前工作表中的筛选控件，或新增一个控件。</p></section>
    </div>
    <p v-if="error" role="alert" class="slicer-panel__error">{{ error }}</p>
  </Modal>
</template>

<style scoped>
.slicer-panel { display: grid; grid-template-columns: minmax(190px, .8fr) minmax(250px, 1.2fr); gap: 20px; min-height: 340px; }
.slicer-panel__list { min-width: 0; padding-inline-end: 14px; border-inline-end: 1px solid #e7ebf0; }
.slicer-panel__list-header, .slicer-panel__actions, .slicer-panel__footer { display: flex; align-items: center; justify-content: space-between; gap: 8px; }
.slicer-panel__list-header { margin-bottom: 10px; color: #5c6979; font-size: 12px; }
.slicer-panel__control.ant-btn { display: flex; flex-direction: column; align-items: flex-start; justify-content: center; height: auto; min-height: 46px; gap: 2px; padding: 8px; white-space: normal; text-align: start; }
.slicer-panel__control.is-selected { color: #2563ba; background: #eaf2ff; }
.slicer-panel__control small { color: #7f8b99; font-size: 11px; }
.slicer-panel__editor { display: flex; flex-direction: column; min-width: 0; gap: 12px; }
.slicer-panel__editor h3 { margin: 0; font-size: 14px; font-weight: 600; }
.slicer-panel__editor label { display: flex; flex-direction: column; gap: 5px; color: #596779; font-size: 12px; }
.slicer-panel__muted { margin: 0; color: #7d8896; font-size: 12px; line-height: 1.8; }
.slicer-panel__error { margin-bottom: 0; color: #c43737; font-size: 12px; line-height: 1.7; }
.slicer-panel__items { display: flex; flex-direction: column; min-height: 80px; max-height: 260px; overflow: auto; gap: 8px; padding: 8px; border: 1px solid #e4e9ef; border-radius: 4px; }
.slicer-panel__actions { justify-content: flex-start; color: #788595; font-size: 12px; }
.slicer-panel__footer { justify-content: flex-start; margin-top: auto; padding-top: 12px; border-top: 1px solid #e9edf2; }
@media (max-width: 640px) { .slicer-panel { grid-template-columns: 1fr; } .slicer-panel__list { padding: 0 0 12px; border-inline-end: 0; border-bottom: 1px solid #e7ebf0; } }
</style>
