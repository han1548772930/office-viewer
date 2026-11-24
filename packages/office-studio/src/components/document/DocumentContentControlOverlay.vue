<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, reactive, ref, shallowRef, watch } from 'vue'
import { Alert, AutoComplete, Button, Checkbox, DatePicker, Form, FormItem, Popover, Select, Tooltip, theme } from 'antdv-next'
import type { FormInstance } from 'antdv-next'
import dayjs from 'dayjs'
import { CalendarDays, ChevronDown } from 'lucide-vue-next'
import type { DocumentContentControl, DocumentModel } from '../../domain/model'
import type { DocumentAuxiliaryBlockHost } from '../../features/document/model/document-block-host'
import { sameDocumentBlockHost } from '../../features/document/model/document-block-host'
import type { DocumentInputSession } from '../../features/document/model/document-input-session'
import { sameTarget, type DocumentSelection } from '../../features/document/model/document-model'
import type { DocumentLayout } from '../../features/document/layout/document-layout-types'
import { contentControlFragmentVisibleRect, contentControlSelectionAtFragment, type DocumentContentControlGeometry, type LayoutContentControlFragment } from '../../features/document/layout/content-control-geometry'
import { sceneOffset, type DocumentViewport } from '../../features/document/rendering/document-scene'
import { contentControlValueCommandState } from '../../features/document/commands/content-control-command-state'
import type { DocumentContentControlCommandRequest, DocumentContentControlSelectionRequest, DocumentContentControlValueCommand } from '../../features/document/commands/content-control-command-types'
import { formatContentControlDate, validContentControlIsoDate } from '../../features/document/services/content-control-presentation'
import type { DocumentPermissionContext } from '../../features/document/permissions/document-permissions'

const props = defineProps<{
  source: DocumentModel
  model: DocumentModel
  layout: DocumentLayout
  geometries: DocumentContentControlGeometry[]
  selection: DocumentSelection
  viewport: DocumentViewport
  activeControlId?: string
  host?: DocumentAuxiliaryBlockHost
  permission?: DocumentPermissionContext
  disabled?: boolean
  inactive?: boolean
  captureInputSession?: (sourceModel?: DocumentModel) => DocumentInputSession | undefined
}>()
const emit = defineEmits<{
  selection: [request: DocumentContentControlSelectionRequest]
  command: [request: DocumentContentControlCommandRequest]
  focus: [request: DocumentContentControlSelectionRequest]
  notify: [message: string]
}>()
const { token } = theme.useToken()
const form = ref<Pick<FormInstance, 'validateFields'>>()
const editorInput = ref<{ focus: () => void }>()
const editor = shallowRef<{ request: DocumentContentControlSelectionRequest; control: DocumentContentControl; anchorKey: string }>()
const draft = reactive({ value: '' })
const validationError = ref('')
const saving = ref(false)
const editorComposing = ref(false)
function isValueForm(value: unknown): value is Pick<FormInstance, 'validateFields'> {
  return value !== null && typeof value === 'object' && 'validateFields' in value && typeof value.validateFields === 'function'
}
function isFocusable(value: unknown): value is { focus: () => void } {
  return value !== null && typeof value === 'object' && 'focus' in value && typeof value.focus === 'function'
}
function setForm(value: unknown) { form.value = isValueForm(value) ? value : undefined }
function setEditorInput(value: unknown) { editorInput.value = isFocusable(value) ? value : undefined }
const themeStyle = computed(() => ({ '--sdt-primary': token.value.colorPrimary, '--sdt-background': token.value.colorBgElevated,
  '--sdt-border': token.value.colorBorder, '--sdt-text': token.value.colorText, '--sdt-secondary': token.value.colorTextSecondary }))
const offset = computed(() => sceneOffset(props.layout, props.viewport))
const scale = computed(() => props.viewport.zoom / 100)
const label = (control: DocumentContentControl) => control.alias || control.title || control.tag || ({ richText: '富文本控件', text: '纯文本控件',
  checkbox: '复选框控件', date: '日期控件', dropDownList: '下拉列表控件', comboBox: '组合框控件', picture: '图片控件', group: '分组控件',
  repeatingSection: '重复节', repeatingSectionItem: '重复节项', docPartObj: '文档部件控件' }[control.type])
const keyFor = (geometry: DocumentContentControlGeometry, fragment: LayoutContentControlFragment) => `${geometry.control.id}:${fragment.page}:${fragment.startOffset}:${fragment.y}`
function visible(fragment: LayoutContentControlFragment) {
  const rect = contentControlFragmentVisibleRect(fragment), k = scale.value, off = offset.value
  return rect && off.x + (rect.x + rect.width) * k > 0 && off.x + rect.x * k < props.viewport.width
    && off.y + (rect.y + rect.height) * k > 0 && off.y + rect.y * k < props.viewport.height
}
const anchors = computed(() => props.inactive ? [] : props.geometries.flatMap(geometry => {
  const fragments = geometry.fragments.filter(visible)
  const fragment = fragments.find(fragment => fragment.page === props.selection.focus.pageAffinity) ?? fragments[0]
  return fragment ? [{ geometry, fragment, key: keyFor(geometry, fragment) }] : []
}))
const checkboxes = computed(() => props.inactive ? [] : props.geometries.filter(geometry => geometry.control.type === 'checkbox')
  .flatMap(geometry => geometry.fragments.filter(visible).map(fragment => ({ geometry, fragment, key: keyFor(geometry, fragment) }))))

function rectStyle(fragment: LayoutContentControlFragment) {
  const rect = contentControlFragmentVisibleRect(fragment)!, k = scale.value, off = offset.value
  return { left: `${off.x + rect.x * k}px`, top: `${off.y + rect.y * k}px`, width: `${rect.width * k}px`, height: `${rect.height * k}px` }
}
function buttonStyle(fragment: LayoutContentControlFragment, tag = false) {
  const rect = contentControlFragmentVisibleRect(fragment)!, k = scale.value, off = offset.value
  return { left: `${Math.max(0, Math.min(props.viewport.width - (tag ? 112 : 24), off.x + (tag ? rect.x : rect.x + rect.width) * k))}px`,
    top: `${Math.max(0, Math.min(props.viewport.height - 24, off.y + rect.y * k + (tag ? -24 : (rect.height * k - 24) / 2)))}px` }
}
function commandFor(control: DocumentContentControl, value: string | boolean): DocumentContentControlValueCommand {
  return control.type === 'checkbox'
    ? { id: 'doc.command.set-content-control-checked', params: { controlId: control.id, checked: Boolean(value) } }
    : control.type === 'dropDownList' || control.type === 'comboBox'
      ? { id: 'doc.command.select-content-control-item', params: { controlId: control.id, value: String(value) } }
      : { id: 'doc.command.set-content-control-value', params: { controlId: control.id, value: String(value) } }
}
function state(control: DocumentContentControl) {
  return contentControlValueCommandState(props.source, commandFor(control, control.checked ?? control.value ?? ''),
    { host: props.host, permission: props.permission, disabled: props.disabled || props.inactive })
}
function requestFor(geometry: DocumentContentControlGeometry, fragment: LayoutContentControlFragment): DocumentContentControlSelectionRequest {
  return { controlId: geometry.control.id, selection: contentControlSelectionAtFragment(geometry, fragment), source: props.source,
    host: props.host ? { ...props.host } : undefined, inputSession: props.captureInputSession?.(props.source) }
}
function requestCurrent(request: DocumentContentControlSelectionRequest) {
  return !props.disabled && !props.inactive && request.source === props.source && sameDocumentBlockHost(request.host, props.host)
    && props.model.contentControls?.some(control => control.id === request.controlId)
    && requestSelectionCurrent(request.selection)
    && (!props.captureInputSession || request.inputSession?.valid)
}
function requestSelectionCurrent(selection: DocumentSelection) {
  return sameTarget(props.selection.anchor, selection.anchor) && sameTarget(props.selection.focus, selection.focus)
    && props.selection.anchor.offset === selection.anchor.offset && props.selection.focus.offset === selection.focus.offset
}
function selectControl(geometry: DocumentContentControlGeometry, fragment: LayoutContentControlFragment) {
  if (props.inactive || props.disabled) return
  emit('selection', requestFor(geometry, fragment))
}
async function checkboxChanged(geometry: DocumentContentControlGeometry, fragment: LayoutContentControlFragment, event: boolean | { target?: { checked?: boolean } }) {
  if (!state(geometry.control).available) return
  const checked = typeof event === 'boolean' ? event : Boolean(event.target?.checked)
  const source = props.source, host = props.host ? { ...props.host } : undefined
  selectControl(geometry, fragment)
  await nextTick()
  if (source !== props.source || !sameDocumentBlockHost(host, props.host) || !state(geometry.control).available) return
  const request = requestFor(geometry, fragment)
  if (!requestCurrent(request)) return
  emit('command', { ...request, command: commandFor(geometry.control, checked) })
  void nextTick(() => emit('focus', request))
}
async function openEditor(geometry: DocumentContentControlGeometry, fragment: LayoutContentControlFragment, anchorKey: string) {
  if (!state(geometry.control).available) return
  const source = props.source, host = props.host ? { ...props.host } : undefined
  selectControl(geometry, fragment)
  await nextTick()
  if (source !== props.source || !sameDocumentBlockHost(host, props.host) || !state(geometry.control).available) return
  const request = requestFor(geometry, fragment)
  if (!requestCurrent(request)) return
  editor.value = { request, control: geometry.control, anchorKey }
  draft.value = geometry.control.showingPlaceholder ? '' : geometry.control.type === 'date'
    ? geometry.control.dateValue ?? geometry.control.value ?? '' : geometry.control.value ?? ''
  validationError.value = ''
  await nextTick()
  if (editor.value?.request === request && requestCurrent(request)) editorInput.value?.focus()
}
function cancelInteraction() { editor.value = undefined; validationError.value = ''; editorComposing.value = false; saving.value = false }
function closeEditor(refocus = true) {
  const request = editor.value?.request
  cancelInteraction()
  if (refocus && request) void nextTick(() => emit('focus', request))
}
function updateDraft(value: unknown) { if (typeof value === 'string') draft.value = value }
const pickerDate = computed(() => validContentControlIsoDate(draft.value)
  ? dayjs(draft.value).year(Number(draft.value.slice(0, 4))) : null)
function updateDate(_value: unknown, dateString: string | string[]) { draft.value = typeof dateString === 'string' ? dateString : '' }
const options = computed(() => (editor.value?.control.items ?? []).map(item => ({ value: item.value, label: item.displayText ?? item.value })))
const datePreview = computed(() => {
  const control = editor.value?.control
  if (control?.type !== 'date' || !draft.value) return ''
  try { return formatContentControlDate(draft.value, control.dateFormat, control.dateLocale, control.dateCalendar) }
  catch { return draft.value }
})
const valueRules = computed(() => [{ validator: async (_rule: unknown, value: unknown) => {
  const control = editor.value?.control
  if (!control || typeof value !== 'string') throw new Error('控件值无效。')
  if (control.type === 'date' && value && !validContentControlIsoDate(value)) throw new Error('请选择有效日期。')
  if (control.type === 'dropDownList' && !control.items?.some(item => item.value === value)) throw new Error('请选择列表中的选项。')
} }])
async function submitValue() {
  const current = editor.value
  if (!current || saving.value || editorComposing.value) return
  saving.value = true; validationError.value = ''
  try {
    await form.value?.validateFields()
    if (editor.value !== current || !requestCurrent(current.request)) return
    const command = commandFor(current.control, draft.value)
    const availability = contentControlValueCommandState(props.source, command, { host: props.host, permission: props.permission, disabled: props.disabled })
    if (!availability.available) { validationError.value = availability.reason ?? '无法修改控件值。'; return }
    emit('command', { ...current.request, command })
    closeEditor()
  } catch (reason) {
    // Form owns field validation messages. Other errors remain visible beside the editor.
    if (reason instanceof Error) validationError.value = reason.message
  } finally { saving.value = false }
}
function editorKeys(event: KeyboardEvent) {
  if (event.isComposing || editorComposing.value || event.keyCode === 229) return
  if (event.key === 'Escape') { event.preventDefault(); closeEditor() }
}
watch(() => props.source, cancelInteraction)
watch(() => props.host, cancelInteraction, { deep: true })
watch(() => props.permission, () => { if (editor.value && !state(editor.value.control).available) cancelInteraction() }, { deep: true })
watch(() => [props.disabled, props.inactive], ([disabled, inactive]) => { if (disabled || inactive) cancelInteraction() })
watch(() => props.selection, selection => {
  const current = editor.value
  if (current && (!sameTarget(selection.anchor, current.request.selection.anchor) || !sameTarget(selection.focus, current.request.selection.focus)
    || selection.anchor.offset !== current.request.selection.anchor.offset || selection.focus.offset !== current.request.selection.focus.offset)) cancelInteraction()
}, { deep: true })
watch(anchors, anchors => { if (editor.value && !anchors.some(anchor => anchor.key === editor.value!.anchorKey)) cancelInteraction() })
onBeforeUnmount(cancelInteraction)
defineExpose({ cancelInteraction })
</script>

<template>
  <div class="content-control-overlay" :style="themeStyle">
    <Tooltip v-for="item in checkboxes" :key="item.key" :title="state(item.geometry.control).reason ?? label(item.geometry.control)">
      <Checkbox class="content-control-checkbox" :style="rectStyle(item.fragment)" :data-content-control-id="item.geometry.control.id"
        :checked="item.geometry.control.checked ?? item.geometry.control.checkbox?.checked ?? false" :disabled="!state(item.geometry.control).available"
        :aria-label="label(item.geometry.control)" @pointerdown.stop @dblclick.stop
        @focus="selectControl(item.geometry, item.fragment)" @change="event => checkboxChanged(item.geometry, item.fragment, event)" />
    </Tooltip>
    <template v-for="item in anchors" :key="item.key">
      <Popover v-if="['date', 'dropDownList', 'comboBox'].includes(item.geometry.control.type)" :open="editor?.anchorKey === item.key"
        placement="bottomLeft" :trigger="['click']" :title="label(item.geometry.control)" destroy-on-hidden
        @open-change="open => { if (!open && editor?.anchorKey === item.key) closeEditor(false) }">
        <template #content>
          <Form v-if="editor?.anchorKey === item.key" :ref="setForm" :model="draft" class="content-control-editor" :style="themeStyle" layout="vertical"
            @submit.prevent="submitValue" @keydown.capture="editorKeys" @keydown.stop @pointerdown.stop @compositionstart="editorComposing = true" @compositionend="editorComposing = false">
            <FormItem name="value" :label="item.geometry.control.type === 'date' ? '选择日期' : item.geometry.control.type === 'comboBox' ? '选择或输入内容' : '选择选项'" :rules="valueRules">
              <DatePicker v-if="item.geometry.control.type === 'date'" :ref="setEditorInput" :value="pickerDate" format="YYYY-MM-DD"
                class="content-control-value-input" :aria-label="label(item.geometry.control)" @change="updateDate" />
              <Select v-else-if="item.geometry.control.type === 'dropDownList'" :ref="setEditorInput" :value="draft.value || undefined" :options="options" :show-search="true"
                option-filter-prop="label" class="content-control-value-input" :aria-label="label(item.geometry.control)" @update:value="updateDraft" />
              <AutoComplete v-else :ref="setEditorInput" :value="draft.value" :options="options" class="content-control-value-input" :aria-label="label(item.geometry.control)"
                @update:value="updateDraft" @select="updateDraft" />
            </FormItem>
            <p v-if="datePreview" class="content-control-preview">显示：{{ datePreview }}</p>
            <Alert v-if="validationError" type="error" :message="validationError" show-icon />
            <div class="content-control-editor-actions">
              <Button html-type="button" @click="closeEditor()">取消</Button>
              <Button type="primary" html-type="submit" :loading="saving" :disabled="editorComposing">确定</Button>
            </div>
          </Form>
        </template>
        <Tooltip :title="state(item.geometry.control).reason ?? `编辑${label(item.geometry.control)}`">
          <Button type="text" size="small" class="content-control-value-button" :style="buttonStyle(item.fragment)" :disabled="!state(item.geometry.control).available"
            :data-content-control-id="item.geometry.control.id" :aria-label="`编辑${label(item.geometry.control)}`" @pointerdown.stop
            @click="openEditor(item.geometry, item.fragment, item.key)">
            <CalendarDays v-if="item.geometry.control.type === 'date'" :size="14" /><ChevronDown v-else :size="14" />
          </Button>
        </Tooltip>
      </Popover>
      <Tooltip v-if="item.geometry.control.id === activeControlId && item.geometry.control.appearance !== 'hidden'" :title="`选择${label(item.geometry.control)}`">
        <Button type="text" size="small" class="content-control-label" :style="buttonStyle(item.fragment, true)" :aria-label="`选择${label(item.geometry.control)}`"
          @pointerdown.stop @click="selectControl(item.geometry, item.fragment)">{{ label(item.geometry.control) }}</Button>
      </Tooltip>
    </template>
  </div>
</template>

<style scoped>
.content-control-overlay { position: absolute; z-index: 4; inset: 0; overflow: visible; pointer-events: none; color: var(--sdt-text); }
.content-control-checkbox, .content-control-value-button, .content-control-label { position: absolute; pointer-events: auto; }
/* The Canvas paints the persisted Unicode style; antd Checkbox owns focus, keyboard and boolean input. */
.content-control-checkbox { display: block; margin: 0; padding: 0; border-radius: 2px; }
.content-control-checkbox :deep(.ant-checkbox) { position: absolute; inset: 0; width: 100%; height: 100%; }
.content-control-checkbox :deep(.ant-checkbox-inner) { opacity: 0; }
.content-control-checkbox :deep(.ant-checkbox::after) { display: none; }
.content-control-checkbox:focus-within { outline: 2px solid var(--sdt-primary); outline-offset: 2px; }
.content-control-value-button { width: 24px; height: 24px; padding: 0; border-radius: 3px; color: var(--sdt-primary); background: var(--sdt-background); }
.content-control-label { max-width: 112px; height: 24px; overflow: hidden; padding: 0 6px; text-overflow: ellipsis; white-space: nowrap; color: var(--sdt-primary); background: var(--sdt-background); border: 1px solid var(--sdt-border); }
.content-control-editor { width: min(280px, calc(100vw - 48px)); color: var(--sdt-text); }
.content-control-editor :deep(.ant-form-item) { margin-bottom: 8px; }
.content-control-value-input { width: 100%; }
.content-control-preview { margin: 0 0 8px; color: var(--sdt-secondary); overflow-wrap: anywhere; }
.content-control-editor-actions { display: flex; justify-content: flex-end; gap: 8px; margin-top: 12px; }
</style>
