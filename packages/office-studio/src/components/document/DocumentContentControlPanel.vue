<script setup lang="ts">
import { Button, Input, Modal, Select, Switch } from 'antdv-next'
import { computed, ref, watch } from 'vue'
import type { DocumentContentControl, DocumentContentControlType, DocumentModel } from '../../domain/model'
import type { FDocument } from '../../api/facade'
import { runText, runsForTarget, selectedPlainText, type DocumentSelection } from '../../features/document/model/document-model'

const props = defineProps<{
  open: boolean
  controls: DocumentContentControl[]
  selection: DocumentSelection
  model: DocumentModel
  documentApi?: FDocument
}>()
const emit = defineEmits<{
  'update:open': [value: boolean]
  change: [id: string]
  changed: [id: string]
  error: [message: string]
}>()

const selectedId = ref('')
const adding = ref(false)
const type = ref<DocumentContentControlType>('text')
const title = ref('')
const placeholder = ref('在此处输入')
const listText = ref('草稿|draft\n已完成|done')
const value = ref('')
const checked = ref(false)
const lock = ref<DocumentContentControl['lock']>('unlocked')
const visibleTypes: Array<{ label: string; value: DocumentContentControlType }> = [
  { label: '纯文本', value: 'text' },
  { label: '富文本', value: 'richText' },
  { label: '复选框', value: 'checkbox' },
  { label: '日期', value: 'date' },
  { label: '下拉列表', value: 'dropDownList' },
  { label: '组合框', value: 'comboBox' },
]
const lockOptions = [
  { label: '不锁定', value: 'unlocked' },
  { label: '锁定内容', value: 'contentLocked' },
  { label: '锁定控件', value: 'sdtLocked' },
  { label: '内容和控件都锁定', value: 'sdtContentLocked' },
]
const selected = computed(() => props.controls.find((control) => control.id === selectedId.value))
const rangeLabel = computed(() => {
  if (!sameTarget(props.selection.anchor, props.selection.focus)) return '请选择同一段落或表格单元格内的一段文字。'
  const start = Math.min(props.selection.anchor.offset, props.selection.focus.offset)
  const end = Math.max(props.selection.anchor.offset, props.selection.focus.offset)
  return start === end ? `插入点：${start}（可创建空控件并稍后填写）` : `选中 ${end - start} 个字符（${start}–${end}）`
})

watch(() => [props.open, props.controls.map((control) => control.id).join('|')], () => {
  if (!props.controls.some((control) => control.id === selectedId.value)) selectedId.value = props.controls[0]?.id ?? ''
}, { immediate: true })
watch(selected, (control) => {
  if (!control) { value.value = ''; checked.value = false; title.value = ''; placeholder.value = ''; lock.value = 'unlocked'; return }
  const api = props.documentApi?.getContentControl(control.id)
  value.value = api?.getText() ?? control.value ?? ''
  checked.value = control.checked ?? false
  title.value = control.title ?? ''
  placeholder.value = control.placeholder ?? ''
  lock.value = control.lock ?? 'unlocked'
}, { immediate: true, deep: true })

function sameTarget(a: DocumentSelection['anchor'], b: DocumentSelection['focus']): boolean {
  return a.blockId === b.blockId && a.row === b.row && a.col === b.col && a.cellBlockId === b.cellBlockId
}

function selectedText(): string {
  if (!sameTarget(props.selection.anchor, props.selection.focus)) return selectedPlainText(props.model, props.selection)
  const point = props.selection.focus
  return runText(runsForTarget(props.model, point))
    .slice(Math.min(props.selection.anchor.offset, point.offset), Math.max(props.selection.anchor.offset, point.offset))
}

function parseItems(): Array<{ displayText: string; value: string }> {
  return listText.value.split(/\r?\n/).map((line) => line.trim()).filter(Boolean).map((line) => {
    const split = line.indexOf('|')
    const displayText = (split < 0 ? line : line.slice(0, split)).trim()
    const itemValue = (split < 0 ? line : line.slice(split + 1)).trim()
    if (!displayText || !itemValue) throw new TypeError('列表项需要填写显示名称和保存值')
    return { displayText, value: itemValue }
  })
}

function createControl(): void {
  const api = props.documentApi
  if (!api) { emit('error', '内容控件命令接口尚未连接'); return }
  if (!sameTarget(props.selection.anchor, props.selection.focus)) { emit('error', '内容控件范围必须位于同一段落或单元格'); return }
  const focus = props.selection.focus
  const startOffset = Math.min(props.selection.anchor.offset, focus.offset)
  const endOffset = Math.max(props.selection.anchor.offset, focus.offset)
  const blockId = focus.cellBlockId ?? focus.blockId
  try {
    const initial = selectedText()
    const control = api.addContentControl({
      type: type.value,
      blockId,
      startOffset,
      endOffset,
      title: title.value.trim() || undefined,
      placeholder: placeholder.value.trim() || undefined,
      showingPlaceholder: !initial && Boolean(placeholder.value.trim()),
      lock: lock.value,
      value: ['text', 'richText', 'comboBox'].includes(type.value) ? initial || undefined : undefined,
      checked: type.value === 'checkbox' ? false : undefined,
      dateValue: type.value === 'date' && /^\d{4}-\d{2}-\d{2}$/.test(initial) ? initial : undefined,
      items: type.value === 'dropDownList' || type.value === 'comboBox' ? parseItems() : undefined,
    })
    selectedId.value = control.getId()
    adding.value = false
    emit('change', selectedId.value)
    emit('changed', selectedId.value)
  } catch (error) { emit('error', error instanceof Error ? error.message : String(error)) }
}

function applyProperties(): void {
  if (!selected.value || !props.documentApi) return
  try {
    props.documentApi.getContentControl(selected.value.id)?.update({ title: title.value || undefined, placeholder: placeholder.value || undefined, lock: lock.value })
    emit('changed', selected.value.id)
  } catch (error) { emit('error', error instanceof Error ? error.message : String(error)) }
}

function applyValue(): void {
  if (!selected.value || !props.documentApi) return
  try {
    props.documentApi.getContentControl(selected.value.id)?.setValue(value.value)
    emit('changed', selected.value.id)
  } catch (error) { emit('error', error instanceof Error ? error.message : String(error)) }
}

function setChecked(next: boolean): void {
  if (!selected.value || !props.documentApi) return
  try {
    props.documentApi.getContentControl(selected.value.id)?.setChecked(next)
    emit('changed', selected.value.id)
  } catch (error) { checked.value = !next; emit('error', error instanceof Error ? error.message : String(error)) }
}

function selectValue(next: string | number): void {
  value.value = String(next)
  if (!selected.value || !props.documentApi) return
  try {
    props.documentApi.getContentControl(selected.value.id)?.selectItem(value.value)
    emit('changed', selected.value.id)
  } catch (error) { emit('error', error instanceof Error ? error.message : String(error)) }
}

function removeControl(): void {
  if (!selected.value || !props.documentApi) return
  try {
    const id = selected.value.id
    props.documentApi.getContentControl(id)?.remove()
    selectedId.value = ''
    emit('changed', id)
  } catch (error) { emit('error', error instanceof Error ? error.message : String(error)) }
}

function close(): void { emit('update:open', false); adding.value = false }
</script>

<template>
  <Modal :open="open" title="内容控件" :width="660" :footer="null" destroy-on-hidden @cancel="close">
    <div class="content-controls-panel">
      <section class="controls-list" aria-label="文档内容控件列表">
        <div class="controls-list-header">
          <span>文档控件（{{ controls.length }}）</span>
          <Button type="primary" size="small" :disabled="!documentApi" @click="adding = !adding">添加控件</Button>
        </div>
        <Button v-for="control in controls" :key="control.id" class="control-row" :class="{ selected: selectedId === control.id }" type="text" block @click="selectedId = control.id; emit('change', control.id)">
          <span>{{ control.title || control.tag || control.type }}</span>
          <span class="control-row-value">{{ control.type === 'checkbox' ? (control.checked ? '已勾选' : '未勾选') : control.value || control.placeholder || '空' }}</span>
        </Button>
        <p v-if="!controls.length" class="controls-empty">还没有内容控件。先在正文中选择范围，或将光标放到插入位置，然后添加控件。</p>
      </section>

      <section v-if="adding" class="control-editor" aria-label="新增内容控件">
        <h3>插入内容控件</h3>
        <p class="control-range">{{ rangeLabel }}</p>
        <label>类型<Select v-model:value="type" :options="visibleTypes" aria-label="控件类型" /></label>
        <label>标题<Input v-model:value="title" aria-label="控件标题" placeholder="例如：客户名称" /></label>
        <label>占位文本<Input v-model:value="placeholder" aria-label="占位文本" /></label>
        <label v-if="type === 'dropDownList' || type === 'comboBox'">列表项<Input.TextArea v-model:value="listText" aria-label="列表项" :auto-size="{ minRows: 2, maxRows: 5 }" /></label>
        <label>锁定<Select v-model:value="lock" :options="lockOptions" aria-label="控件锁定状态" /></label>
        <Button type="primary" block @click="createControl">插入到当前范围</Button>
      </section>

      <section v-else-if="selected" class="control-editor" aria-label="内容控件属性">
        <h3>{{ selected.title || '控件属性' }}</h3>
        <p class="control-range">{{ selected.type }} · {{ selected.blockId }} · {{ selected.startOffset }}–{{ selected.endOffset }}</p>
        <template v-if="selected.type === 'checkbox'">
          <label class="check-setting"><span>复选框状态</span><Switch :checked="checked" checked-children="已选" un-checked-children="未选" @change="setChecked" /></label>
        </template>
        <template v-else-if="selected.type === 'date'">
          <label>日期<Input v-model:value="value" type="date" aria-label="日期值" /></label>
          <Button type="primary" @click="applyValue">应用日期</Button>
        </template>
        <template v-else-if="selected.type === 'dropDownList'">
          <label>列表值<Select :value="value" :options="(selected.items ?? []).map((item) => ({ label: item.displayText ?? item.value, value: item.value }))" aria-label="列表值" @change="selectValue" /></label>
        </template>
        <template v-else>
          <label>控件文本<Input v-model:value="value" aria-label="控件文本" :disabled="selected.lock === 'contentLocked' || selected.lock === 'sdtContentLocked'" /></label>
          <Button type="primary" @click="applyValue">应用文本</Button>
        </template>
        <label>标题<Input v-model:value="title" aria-label="属性标题" /></label>
        <label>占位文本<Input v-model:value="placeholder" aria-label="属性占位文本" /></label>
        <label>锁定<Select v-model:value="lock" :options="lockOptions" aria-label="属性锁定状态" /></label>
        <div class="control-actions">
          <Button @click="applyProperties">保存属性</Button>
          <Button danger @click="removeControl">删除控件</Button>
        </div>
      </section>
      <section v-else class="control-editor control-empty-selection">
        <p>选择一个内容控件以编辑其值和属性。</p>
        <p v-if="!documentApi" role="alert">内容控件命令接口尚未连接。</p>
      </section>
    </div>
  </Modal>
</template>

<style scoped>
.content-controls-panel { display: grid; grid-template-columns: minmax(200px, .9fr) minmax(250px, 1.1fr); min-height: 330px; gap: 16px; }
.controls-list { min-width: 0; padding-right: 12px; border-right: 1px solid #e5e8ed; }
.controls-list-header, .control-actions, .check-setting { display: flex; align-items: center; justify-content: space-between; gap: 10px; }
.controls-list-header { margin-bottom: 10px; color: #596372; font-size: 12px; }
.control-row.ant-btn { height: auto; min-height: 36px; display: flex; flex-direction: column; align-items: flex-start; justify-content: center; gap: 2px; padding: 6px 8px; text-align: left; white-space: normal; }
.control-row.selected { background: #edf5ff; color: #276ac5; }
.control-row-value { max-width: 100%; overflow: hidden; color: #86909c; font-size: 11px; text-overflow: ellipsis; white-space: nowrap; }
.controls-empty, .control-empty-selection { color: #77818d; font-size: 12px; line-height: 1.7; }
.control-editor { min-width: 0; display: flex; flex-direction: column; gap: 10px; }
.control-editor h3 { margin: 0; color: #303742; font-size: 14px; font-weight: 600; }
.control-editor label { display: flex; flex-direction: column; gap: 5px; color: #66717e; font-size: 12px; }
.control-range { margin: -5px 0 0; color: #8893a0; font-size: 11px; overflow-wrap: anywhere; }
.control-actions { justify-content: flex-start; margin-top: auto; padding-top: 10px; border-top: 1px solid #edf0f3; }
@media (max-width: 640px) { .content-controls-panel { grid-template-columns: 1fr; } .controls-list { padding: 0 0 12px; border-right: 0; border-bottom: 1px solid #e5e8ed; } }
</style>
