<script setup lang="ts">
import { computed, reactive, ref, shallowRef, watch } from 'vue'
import { Alert, Button, Checkbox, Form, FormItem, Input, InputNumber, Modal, Select, Tooltip, theme } from 'antdv-next'
import type { DocumentModel } from '../../domain/model'
import type { DocumentSelection } from '../../features/document/model/document-model'
import type { DocumentNoteDecorationHost } from '../../features/document/model/document-block-host'
import { DOCUMENT_NOTE_NUMBER_FORMATS } from '../../features/document/notes/document-note-numbering'
import { documentSectionSuppressesEndnotes, effectiveDocumentNoteSettings, getDocumentNoteSettings, type DocumentNoteSettingsValue } from '../../features/document/notes/document-note-settings'
import type { DocumentNoteKind, DocumentNoteMutationRequest } from '../../features/document/notes/document-note-types'
import { documentNoteInsertState, documentNoteSettingsState } from '../../features/document/commands/document-note-command-state'
import type { DocumentPermissionContext } from '../../features/document/permissions/document-permissions'
import { documentNoteSettingsFromForm, type DocumentNoteSettingsForm } from '../../features/document/controllers/document-note-settings-draft'
import { documentNoteDecorationEntries, documentNoteDecorationInspectState, documentNoteDecorationResetState } from '../../features/document/controllers/document-note-settings-state'

const props = defineProps<{ open: boolean; model: DocumentModel; kind: DocumentNoteKind; sectionId: string; selection: DocumentSelection;
  permission: DocumentPermissionContext; disabled?: boolean; insertDisabled?: boolean; error?: string }>()
const emit = defineEmits<{ 'update:open': [open: boolean]; submit: [source: DocumentModel, request: DocumentNoteMutationRequest];
  inspect: [source: DocumentModel, host: DocumentNoteDecorationHost]; closed: [] }>()
const { token } = theme.useToken()
const scope = ref<'document' | 'section'>('document'), source = shallowRef<DocumentModel>(), base = shallowRef<DocumentNoteSettingsValue>()
const form = reactive<DocumentNoteSettingsForm>({}), customMark = ref(''), useCustomMark = ref(false), suppressEndnotes = ref(false)
const draftError = ref(''), markComposing = ref(false), markInput = ref<{ blur: () => void }>()
const title = computed(() => props.kind === 'footnote' ? '脚注设置' : '尾注设置')
const effective = computed(() => effectiveDocumentNoteSettings(source.value ?? props.model, props.kind, props.sectionId))
const inherited = computed(() => scope.value === 'section'
  ? effectiveDocumentNoteSettings({ ...(source.value ?? props.model), sectionSetup: undefined,
    blocks: (source.value ?? props.model).blocks.map(block => block.type === 'sectionBreak' ? { ...block, sectionBreak: { ...block.sectionBreak!, setup: undefined } } : block) }, props.kind, props.sectionId)
  : effectiveDocumentNoteSettings({ ...(source.value ?? props.model), noteSettings: undefined, sectionSetup: undefined,
    blocks: (source.value ?? props.model).blocks.map(block => block.type === 'sectionBreak' ? { ...block, sectionBreak: { ...block.sectionBreak!, setup: undefined } } : block) }, props.kind, props.sectionId))
const settingsState = computed(() => documentNoteSettingsState(props.model, { permission: props.permission, disabled: props.disabled }, scope.value === 'section' ? props.sectionId : undefined))
const insertState = computed(() => documentNoteInsertState(props.model, props.selection, { permission: props.permission, disabled: props.disabled || props.insertDisabled }))
const positionOptions = computed(() => (props.kind === 'footnote' ? [['pageBottom', '页底'], ['beneathText', '正文下方']] : [['docEnd', '文档末尾'], ['sectEnd', '节末尾']]).map(([value, label]) => ({ value, label })))
const restartOptions = computed(() => [['continuous', '连续编号'], ['eachSect', '每节重新编号'], ...(props.kind === 'footnote' ? [['eachPage', '每页重新编号']] : [])].map(([value, label]) => ({ value, label })))
const numberOptions = computed(() => [...DOCUMENT_NOTE_NUMBER_FORMATS.map(([value, label]) => ({ value: String(value), label: String(label) })),
  { value: 'none', label: '无标记' }, ...(form.numberFormat && !DOCUMENT_NOTE_NUMBER_FORMATS.some(([value]) => value === form.numberFormat) && form.numberFormat !== 'none'
    ? [{ value: form.numberFormat, label: form.numberFormat + '（导入格式）' }] : [])])
const stale = computed(() => source.value !== props.model)
const actionBlocked = computed(() => stale.value || props.disabled === true || markComposing.value)
const pendingMark = computed(() => useCustomMark.value || customMark.value.length > 0)
const settingsChanged = computed(() => {
  if (!source.value) return false
  try {
    const sectionId = scope.value === 'section' ? props.sectionId : undefined
    const settings = documentNoteSettingsFromForm(base.value, form, props.kind, sectionId)
    return JSON.stringify(settings) !== JSON.stringify(getDocumentNoteSettings(source.value, props.kind, sectionId) ?? {})
      || props.kind === 'endnote' && sectionId !== undefined && suppressEndnotes.value !== documentSectionSuppressesEndnotes(source.value, sectionId)
  } catch { return true }
})
const decorations = computed(() => {
  if (!props.open || scope.value !== 'document') return []
  const current = source.value ?? props.model, context = { permission: props.permission, disabled: actionBlocked.value,
    pendingSettings: settingsChanged.value, pendingMark: pendingMark.value }
  return documentNoteDecorationEntries(current, props.kind).map(entry => ({ ...entry,
    custom: entry.dormant || base.value?.[entry.host.decoration] !== undefined,
    inspect: documentNoteDecorationInspectState(current, entry.host, context),
    reset: documentNoteDecorationResetState(current, entry.host, context) }))
})

function cancelMarkInput() { markInput.value?.blur(); markComposing.value = false }
function loadScope() {
  cancelMarkInput()
  const current = source.value ?? props.model
  base.value = getDocumentNoteSettings(current, props.kind, scope.value === 'section' ? props.sectionId : undefined)
  form.position = base.value?.position; form.restart = base.value?.restart; form.numberFormat = base.value?.numberFormat
  form.startNumber = base.value?.startNumber; form.columnCount = base.value && 'columnCount' in base.value ? base.value.columnCount : undefined
  suppressEndnotes.value = documentSectionSuppressesEndnotes(current, props.sectionId)
  draftError.value = ''
}
watch(() => [props.open, props.kind, props.sectionId] as const, ([open]) => {
  cancelMarkInput()
  if (!open) return
  source.value = props.model; customMark.value = ''; useCustomMark.value = false; scope.value = 'document'; loadScope()
}, { immediate: true })
function changeScope(value: unknown) {
  if (value !== 'document' && value !== 'section' || value === scope.value || actionBlocked.value) return
  if (settingsChanged.value) { draftError.value = '当前范围有未应用的设置。请先应用或还原设置草稿，再切换范围。'; return }
  scope.value = value; loadScope()
}
function restoreSettingsDraft() { if (!actionBlocked.value) loadScope() }
function clearMarkDraft() { if (!actionBlocked.value) { customMark.value = ''; useCustomMark.value = false } }
watch(() => props.model, () => { if (stale.value) cancelMarkInput() })
function request(action: 'apply' | 'reset' | 'insert') {
  if (!source.value || actionBlocked.value) return
  try {
    const sectionId = scope.value === 'section' ? props.sectionId : undefined
    const settings = action === 'reset' ? null : documentNoteSettingsFromForm(base.value, form, props.kind, sectionId)
    if (action !== 'insert') {
      if (!settingsState.value.available) return
      emit('submit', source.value, { type: 'settings', kind: props.kind, settings, ...(sectionId === undefined ? {} : { sectionId }),
        ...(props.kind === 'endnote' && sectionId !== undefined ? { suppressEndnotes: suppressEndnotes.value } : {}) })
    } else {
      if (!insertState.value.available) return
      if (settingsChanged.value && !settingsState.value.available) { draftError.value = settingsState.value.reason ?? '此作用域不能修改设置。'; return }
      emit('submit', source.value, { type: 'insert', kind: props.kind, selection: props.selection, options: {
        ...(useCustomMark.value ? { customMark: customMark.value } : {}),
        ...(settingsChanged.value ? { settings, ...(sectionId === undefined ? {} : { sectionId }), ...(props.kind === 'endnote' && sectionId !== undefined ? { suppressEndnotes: suppressEndnotes.value } : {}) } : {}) } })
    }
  } catch (error) { draftError.value = error instanceof Error ? error.message : String(error) }
}
function inspectDecoration(host: DocumentNoteDecorationHost) {
  if (!source.value) return
  const state = documentNoteDecorationInspectState(source.value, host, { permission: props.permission, disabled: actionBlocked.value,
    pendingSettings: settingsChanged.value, pendingMark: pendingMark.value })
  if (!state.available) { draftError.value = state.reason ?? '当前不能打开此内容区。'; return }
  emit('inspect', source.value, { ...host })
}
function resetDecoration(host: DocumentNoteDecorationHost) {
  if (!source.value || host.dormantIndex !== undefined || actionBlocked.value || !settingsState.value.available) return
  const state = documentNoteDecorationResetState(source.value, host, { permission: props.permission })
  if (!state.available) { draftError.value = state.reason ?? '当前不能重置此内容区。'; return }
  const next = { ...base.value }; delete next[host.decoration]; base.value = next
}
</script>

<template>
  <Modal :open="open" :title="title" :width="760" :closable="!markComposing" :keyboard="!markComposing" :mask-closable="false" destroy-on-hidden @cancel="emit('update:open', false)" @after-close="emit('closed')">
    <Alert v-if="error || draftError || stale" type="error" show-icon :message="error || draftError || '文档已变化，请重新打开注释设置以保留最新内容。'" />
    <Alert v-else-if="!settingsState.available" type="info" show-icon :message="settingsState.reason" />
    <Form :model="form" layout="vertical" class="note-settings-form" :disabled="!settingsState.available || stale || disabled">
      <FormItem label="应用范围"><Select :value="scope" :disabled="stale || disabled || markComposing" :options="[{ value: 'document', label: '整个文档' }, { value: 'section', label: '当前节' }]" @change="changeScope" /></FormItem>
      <FormItem label="位置"><Select v-model:value="form.position" allow-clear :options="positionOptions" :placeholder="'继承：' + positionOptions.find(option => option.value === inherited.position)?.label" /></FormItem>
      <FormItem label="编号格式"><Select v-model:value="form.numberFormat" allow-clear :options="numberOptions" :placeholder="'继承：' + (numberOptions.find(option => option.value === inherited.numberFormat)?.label ?? inherited.numberFormat)" /></FormItem>
      <FormItem label="编号方式"><Select v-model:value="form.restart" allow-clear :options="restartOptions" :placeholder="'继承：' + restartOptions.find(option => option.value === inherited.restart)?.label" /></FormItem>
      <FormItem label="起始编号"><InputNumber v-model:value="form.startNumber" :min="1" :max="Number.MAX_SAFE_INTEGER" :precision="0" :placeholder="'继承：' + inherited.startNumber" /></FormItem>
      <FormItem v-if="kind === 'footnote'" label="脚注列数"><Select v-model:value="form.columnCount" allow-clear :options="[0, 1, 2, 3, 4].map(value => ({ value, label: value === 0 ? '跟随正文列数' : value + ' 列' }))" :placeholder="inherited.columnCount === 0 ? '继承：跟随正文列数' : '继承：' + inherited.columnCount + ' 列'" /></FormItem>
      <FormItem v-if="kind === 'endnote' && scope === 'section'" label="节末尾注"><Checkbox v-model:checked="suppressEndnotes">延后到后续允许显示尾注的节</Checkbox></FormItem>
    </Form>
    <p class="note-settings-effective" :style="{ color: token.colorTextSecondary }">当前有效设置：{{ effective.numberFormat }}；起始 {{ effective.startNumber }}。</p>
    <Form layout="vertical" class="note-mark-form" :disabled="!insertState.available || stale || disabled">
      <FormItem label="插入标记"><Checkbox v-model:checked="useCustomMark" :disabled="markComposing">使用自定义标记</Checkbox><Input v-if="useCustomMark" ref="markInput" v-model:value="customMark" :maxlength="4096" aria-label="自定义注释标记" placeholder="可输入符号或留空" @compositionstart="markComposing = true" @compositionend="markComposing = false" /><Button v-if="pendingMark" size="small" :disabled="actionBlocked" @click="clearMarkDraft">清空标记草稿</Button></FormItem>
    </Form>
    <section v-if="scope === 'document'" class="note-decorations">
      <p :style="{ color: token.colorTextSecondary }">分隔内容通过独立检查画布编辑，修改使用文档的命令、撤销和保存。{{ settingsChanged ? '请先应用当前设置，再进入内容检查。' : pendingMark ? '请先插入注释或清空标记草稿，再进入内容检查。' : '' }}</p>
      <div v-for="entry in decorations" :key="entry.key" class="note-decoration-row" :data-document-note-decoration="entry.key">
        <span>{{ entry.label }}{{ entry.nativeId === undefined ? '' : '（原生 ID ' + entry.nativeId + '）' }}</span><span :style="{ color: token.colorTextSecondary }">{{ entry.dormant ? '保留来源内容' : entry.custom ? '自定义内容' : '默认' }}</span>
        <Tooltip :title="entry.inspect.reason"><Button size="small" :disabled="!entry.inspect.available" @click="inspectDecoration(entry.host)">在画布中编辑</Button></Tooltip>
        <Tooltip v-if="!entry.dormant" :title="entry.reset.reason"><Button size="small" :disabled="!settingsState.available || !entry.reset.available || !entry.custom" @click="resetDecoration(entry.host)">恢复默认</Button></Tooltip>
      </div>
    </section>
    <template #footer>
      <Button :disabled="markComposing" @click="emit('update:open', false)">关闭</Button>
      <Button v-if="settingsChanged" :disabled="actionBlocked" @click="restoreSettingsDraft">还原设置草稿</Button>
      <Button :disabled="!settingsState.available || actionBlocked" @click="request('reset')">重置此范围</Button>
      <Button :disabled="!insertState.available || actionBlocked" @click="request('insert')">插入{{ kind === 'footnote' ? '脚注' : '尾注' }}</Button>
      <Button type="primary" :disabled="!settingsState.available || actionBlocked" @click="request('apply')">应用</Button>
    </template>
  </Modal>
</template>

<style scoped>
.note-settings-form { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); column-gap: 20px; padding-top: 12px; }
.note-settings-form :deep(.ant-input-number) { width: 100%; }
.note-settings-effective { margin: 0 0 10px; font-size: 12px; }
.note-mark-form :deep(.ant-checkbox-wrapper) { margin-bottom: 8px; }
.note-decoration-row { display: grid; grid-template-columns: minmax(110px, 1fr) minmax(0, 1fr) auto auto; gap: 8px; align-items: center; margin: 8px 0; }
@media (max-width: 620px) { .note-settings-form { grid-template-columns: 1fr; } .note-decoration-row { grid-template-columns: 1fr 1fr; } }
</style>
