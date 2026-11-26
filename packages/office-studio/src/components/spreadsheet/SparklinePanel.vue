<script setup lang="ts">
import { computed, inject, nextTick, onBeforeUnmount, ref, shallowRef, watch } from 'vue'
import { Alert, Button, Form, FormItem, Input, Modal, Select, Tabs, theme } from 'antdv-next'
import { FOffice } from '../../api/facade.ts'
import type { OfficeStudio } from '../../api/office-studio.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'
import { resolveSparklineSettings, type SparklineMutationRequest, type SparklineSettingsInput } from '../../features/sparkline/sparkline-model.ts'
import { assertSparklinePanelMutation, buildSparklinePanelRequest, canEditSparklinePanelRanges, captureSparklinePanelOwner,
  createSparklinePanelDraft, editSparklineSettingsDraft, readSparklinePanelState, sparklinePanelDraftDirty,
  sparklinePanelOwnerContext, sparklinePanelOwnerCurrent, type SparklinePanelDraft, type SparklinePanelOwner, type SparklinePanelSnapshot } from '../../features/sparkline/sparkline-panel-state.ts'
import { parseSparklineRangeList } from '../../features/sparkline/sparkline-reference.ts'
import SparklineSettingsFields from '../../features/sparkline/SparklineSettingsFields.vue'

const props = defineProps<{ open: boolean; studio: OfficeStudio; workbookId: string; sheetId: string; targetRange: string; targetRanges?: readonly string[] }>()
const emit = defineEmits<{ 'update:open': [value: boolean]; applied: []; closed: [owner: SparklinePanelOwner | undefined] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const snapshot = shallowRef<SparklinePanelSnapshot>(), owner = shallowRef<SparklinePanelOwner>()
const draft = ref<SparklinePanelDraft>(), original = shallowRef<SparklinePanelDraft>()
const scope = ref('*'), tab = ref<'basic' | 'advanced'>('basic'), error = ref(''), status = ref('')
const busy = ref(false), composing = ref(false), generation = ref(0), observation = ref(0)
const { token } = theme.useToken()
const appearance = computed(() => ({ '--sparkline-text': token.value.colorText, '--sparkline-muted': token.value.colorTextSecondary,
  '--sparkline-border': token.value.colorBorderSecondary, '--sparkline-background': token.value.colorBgElevated }))
let disposed = false, applying = false, pendingRefresh = false
let returnFocus: HTMLElement | undefined, closingOwner: SparklinePanelOwner | undefined
let closingFocus: HTMLElement | undefined, closingGeneration = 0
const ranges = computed(() => props.targetRanges?.length ? [...props.targetRanges] : parseSparklineRangeList(props.targetRange))
const existing = computed(() => (snapshot.value?.selectedGroups.length ?? 0) > 0)
const selectedGroups = computed(() => scope.value === '*' ? snapshot.value?.selectedGroups ?? [] : snapshot.value?.selectedGroups.filter(group => group.groupId === scope.value) ?? [])
const settingRanges = computed(() => scope.value === '*' ? snapshot.value?.ranges ?? [] : selectedGroups.value.flatMap(group => group.targetRanges))
const dirty = computed(() => Boolean(draft.value && original.value && sparklinePanelDraftDirty(draft.value, original.value)))
const stale = computed(() => { void observation.value; return Boolean(owner.value && (!sparklinePanelOwnerCurrent(owner.value) || owner.value.studio !== props.studio
  || owner.value.target.workbookId !== props.workbookId || owner.value.target.sheetId !== props.sheetId)) })
const canConfigure = computed(() => {
  void observation.value
  const captured = owner.value
  if (!captured || stale.value) return false
  const targets = existing.value ? selectedGroups.value.flatMap(group => group.targetRanges) : snapshot.value?.ranges ?? []
  try { return canEditSparklinePanelRanges(captured, targets) } catch { return false }
})
const disabled = computed(() => busy.value || composing.value || stale.value || !canConfigure.value)
const textDisabled = computed(() => busy.value || (stale.value || !canConfigure.value) && !composing.value)
const validationError = computed(() => {
  if (!draft.value) return ''
  try { resolveSparklineSettings(draft.value.settings); return '' } catch (failure) { return failure instanceof Error ? failure.message : String(failure) }
})
const scopeOptions = computed(() => [{ label: '选区中的全部组', value: '*' }, ...(snapshot.value?.selectedGroups ?? []).map(group =>
  ({ label: `${group.targetRanges.slice(0, 3).join('、')}${group.members.length > 3 ? '…' : ''}（${group.members.length} 项）`, value: group.groupId }))])
const primaryOptions = computed(() => selectedGroups.value.flatMap(group => group.members.filter(member => scope.value !== '*' || group.selectedAddresses.includes(member.address)))
  .map(member => ({ label: member.address, value: member.address })))
const selectedCount = computed(() => selectedGroups.value.reduce((count, group) => count + (scope.value === '*' ? group.selectedAddresses.length : group.members.length), 0))
const canCombine = computed(() => !dirty.value && selectedCount.value > 1 && (selectedGroups.value.length > 1
  || selectedGroups.value.some(group => scope.value === '*' && group.selectedAddresses.length > 1 && group.members.length > group.selectedAddresses.length)))
const canUncombine = computed(() => !dirty.value && selectedGroups.value.some(group => group.members.length > 1))
const canManage = computed(() => !dirty.value && Boolean(draft.value?.primary))
const tabs = computed(() => [{ key: 'basic', label: t('sheet.sparkline.basic', '基本设置') }, { key: 'advanced', label: t('sheet.sparkline.advanced', '高级设置'), disabled: draft.value?.settings.type === 'pie' }])
const operationOptions = computed(() => [{ label: '只修改设置', value: 'settings' }, { label: '修改当前迷你图的数据与位置', value: 'item' },
  { label: '修改整组迷你图的数据与位置', value: 'group', disabled: selectedGroups.value.length !== 1 }])

function message(failure: unknown): string { return failure instanceof Error ? failure.message : String(failure) }
function sameSession(current: SparklinePanelDraft | undefined, captured: SparklinePanelOwner | undefined, request: number): boolean {
  return !disposed && props.open && draft.value === current && owner.value === captured && generation.value === request
    && !!captured && captured.studio === props.studio && captured.target.workbookId === props.workbookId && captured.target.sheetId === props.sheetId
}
function loadDraft(current: SparklinePanelSnapshot, operation: SparklinePanelDraft['dataOperation'] = draft.value?.dataOperation ?? 'settings', primary = draft.value?.primary): void {
  if (scope.value !== '*' && !current.selectedGroups.some(group => group.groupId === scope.value)) scope.value = '*'
  const next = createSparklinePanelDraft(current, scope.value === '*' ? undefined : scope.value, primary, operation)
  if (operation === 'group' && current.selectedGroups.length !== 1 && scope.value === '*') next.dataOperation = 'settings'
  generation.value += 1; draft.value = next; original.value = cloneModel(next); owner.value = current.owner
  error.value = ''; status.value = ''
  if (next.settings.type === 'pie') tab.value = 'basic'
}
function refresh(force = false): void {
  if (!props.open || disposed) return
  observation.value += 1
  if (applying) { pendingRefresh = true; return }
  try {
    const current = readSparklinePanelState(props.studio, { workbookId: props.workbookId, sheetId: props.sheetId }, ranges.value)
    snapshot.value = current
    if (force || !dirty.value && !composing.value) loadDraft(current)
  } catch (failure) { snapshot.value = undefined; error.value = message(failure) }
}
function reload(): void { if (!busy.value && !composing.value) refresh(true) }

const modalBindings = computed(() => {
  const request = generation.value, studio = props.studio, workbookId = props.workbookId, sheetId = props.sheetId
  const current = () => !disposed && props.open && generation.value === request && props.studio === studio
    && props.workbookId === workbookId && props.sheetId === sheetId
  return { cancel: () => { if (current()) close() }, reload: () => { if (current()) reload() } }
})
const afterClosed = computed(() => {
  void generation.value; void props.open
  const captured = closingOwner, element = closingFocus, request = closingGeneration
  return () => closed(captured, element, request)
})

/** Event callbacks retain the input/draft owner even when a popup or IME publishes after a replacement. */
const bindings = computed(() => {
  const current = draft.value, captured = owner.value, request = generation.value
  const updateText = (key: 'sourceText' | 'targetText', value: unknown) => {
    if (sameSession(current, captured, request) && current && typeof value === 'string' && !busy.value) current[key] = value
  }
  return {
    source: (value: unknown) => updateText('sourceText', value), target: (value: unknown) => updateText('targetText', value),
    settings: (patch: SparklineSettingsInput) => {
      if (!sameSession(current, captured, request) || !current || busy.value || !composing.value && disabled.value) return
      current.settings = editSparklineSettingsDraft(current.settings, patch)
      if (current.settings.type === 'pie') tab.value = 'basic'
    },
    startComposition: () => { if (sameSession(current, captured, request)) composing.value = true },
    endComposition: () => {
      void nextTick(() => {
        if (!sameSession(current, captured, request)) return
        composing.value = false
        if (pendingRefresh) { pendingRefresh = false; refresh() }
      })
    },
    chooseScope: (value: unknown) => {
      if (!sameSession(current, captured, request) || busy.value || composing.value || dirty.value || typeof value !== 'string' || !scopeOptions.value.some(option => option.value === value)) return
      scope.value = value
      if (snapshot.value) loadDraft(snapshot.value, 'settings')
    },
    choosePrimary: (value: unknown) => {
      if (!sameSession(current, captured, request) || busy.value || composing.value || dirty.value || typeof value !== 'string' || !primaryOptions.value.some(option => option.value === value)) return
      const group = snapshot.value?.selectedGroups.find(candidate => candidate.members.some(member => member.address === value))
      if (!snapshot.value || !group) return
      loadDraft(snapshot.value, current?.dataOperation, value)
    },
    chooseOperation: (value: unknown) => {
      if (!sameSession(current, captured, request) || busy.value || composing.value || dirty.value || value !== 'settings' && value !== 'item' && value !== 'group') return
      if (value === 'group' && selectedGroups.value.length !== 1) return
      if (snapshot.value) loadDraft(snapshot.value, value, current?.primary)
    },
    chooseTab: (value: unknown) => {
      if (!sameSession(current, captured, request) || busy.value || composing.value || value !== 'basic' && value !== 'advanced') return
      if (value === 'advanced' && current?.settings.type === 'pie') return
      tab.value = value
    },
    clear: (value: 'item' | 'group') => { if (sameSession(current, captured, request)) clear(value) },
    combine: () => { if (sameSession(current, captured, request)) combine() },
    uncombine: () => { if (sameSession(current, captured, request)) uncombine() },
    apply: () => { if (sameSession(current, captured, request)) apply() },
    close: () => { if (sameSession(current, captured, request)) close() },
    reload: () => { if (sameSession(current, captured, request)) reload() },
  }
})
function activeRanges(): string[] { return [...settingRanges.value] }
function perform(request: SparklineMutationRequest, closeAfter = true): void {
  const captured = owner.value, current = draft.value, actionGeneration = generation.value
  if (!captured || !current || disabled.value) return
  busy.value = true; applying = true; error.value = ''; status.value = ''
  try {
    assertSparklinePanelMutation(captured, request)
    const sheet = FOffice.newAPI(props.studio).getWorkbook(props.workbookId)?.getSheetBySheetId(props.sheetId)
    if (!sheet) throw new Error('迷你图所在工作表已关闭。')
    const accepted = sheet.applySparklineMutation(request, sparklinePanelOwnerContext(captured))
    if (!sameSession(current, captured, actionGeneration)) return
    if (!accepted) { status.value = '当前操作没有产生内容变更。'; return }
    if (!sparklinePanelOwnerCurrent(captured, true)) { error.value = '操作结束后工作台已更改，请重新读取。'; return }
    if (closeAfter) {
      closingOwner = captureSparklinePanelOwner(props.studio, { workbookId: props.workbookId, sheetId: props.sheetId })
      closingFocus = returnFocus; closingGeneration = actionGeneration
      emit('applied')
      if (sameSession(current, captured, actionGeneration) && sparklinePanelOwnerCurrent(closingOwner)) emit('update:open', false)
    } else {
      applying = false; busy.value = false; pendingRefresh = false; refresh(true); emit('applied')
    }
  } catch (failure) { if (sameSession(current, captured, actionGeneration)) error.value = message(failure) }
  finally {
    applying = false
    if (generation.value === actionGeneration) busy.value = false
    if (pendingRefresh && props.open) { pendingRefresh = false; refresh() }
  }
}
function apply(): void {
  if (!draft.value || !original.value || disabled.value || validationError.value) return
  try { perform(buildSparklinePanelRequest(draft.value, original.value, activeRanges(), existing.value)) } catch (failure) { error.value = message(failure) }
}
function clear(scope: 'item' | 'group'): void { if (draft.value?.primary && !dirty.value) perform({ kind: 'clear', ranges: [draft.value.primary], scope }) }
function combine(): void { if (canCombine.value) perform({ kind: 'combine', ranges: activeRanges() }, false) }
function uncombine(): void { if (canUncombine.value) perform({ kind: 'uncombine', ranges: activeRanges() }, false) }
function close(): void {
  if (busy.value || composing.value) return
  closingOwner = owner.value; closingFocus = returnFocus; closingGeneration = generation.value
  emit('update:open', false)
}
function closed(captured: SparklinePanelOwner | undefined, element: HTMLElement | undefined, closeGeneration: number): void {
  if (props.open || disposed || composing.value || closingOwner !== captured || closingGeneration !== closeGeneration) return
  closingOwner = undefined; closingFocus = undefined
  emit('closed', captured)
  void nextTick(() => {
    if (props.open || disposed || composing.value || !captured || !element?.isConnected || closeGeneration !== closingGeneration
      || !sparklinePanelOwnerCurrent(captured)) return
    element.focus({ preventScroll: true })
  })
}

watch([() => props.open, () => props.studio, () => props.workbookId, () => props.sheetId, () => props.targetRange, () => props.targetRanges], ([open]) => {
  generation.value += 1; busy.value = false; composing.value = false; applying = false; pendingRefresh = false
  draft.value = undefined; original.value = undefined; owner.value = undefined; snapshot.value = undefined; scope.value = '*'; tab.value = 'basic'
  error.value = ''; status.value = ''
  if (open) {
    const focused = typeof document === 'undefined' ? undefined : document.activeElement
    returnFocus = typeof HTMLElement !== 'undefined' && focused instanceof HTMLElement ? focused : undefined
    closingOwner = undefined; closingFocus = undefined; closingGeneration = generation.value
    refresh(true)
  }
}, { immediate: true })
watch(() => props.studio, (studio, _previous, onCleanup) => {
  const changed = () => refresh()
  const listeners = [
    studio.events.addEvent('CommandExecuted', event => { if (!event.unitId || event.unitId === props.workbookId) changed() }),
    studio.events.addEvent('Undo', event => { if (event.unitId === props.workbookId) changed() }),
    studio.events.addEvent('Redo', event => { if (event.unitId === props.workbookId) changed() }),
    studio.events.addEvent('WorkbookDisposed', event => { if (event.workbookId === props.workbookId) changed() }),
    studio.events.addEvent('SelectionChanged', changed),
  ]
  const stop = watch(() => studio.editor.workbook.model, changed, { flush: 'post' })
  onCleanup(() => { listeners.forEach(listener => listener.dispose()); stop() })
}, { immediate: true })
onBeforeUnmount(() => { disposed = true; generation.value += 1; closingOwner = undefined; closingFocus = undefined; returnFocus = undefined })
</script>

<template>
  <Modal :open="open" :title="t(existing ? 'sheet.sparkline.title.edit' : 'sheet.sparkline.title.insert', existing ? '编辑迷你图' : '插入迷你图')"
    :width="600" :footer="null" :root-style="appearance" :mask="{ closable: !busy && !composing }" :keyboard="!busy && !composing" :closable="!busy && !composing"
    :focus-trigger-after-close="false" :after-close="afterClosed" destroy-on-hidden @cancel="modalBindings.cancel">
    <Form v-if="draft" :key="generation" :model="draft" layout="vertical" class="sparkline-form" data-testid="sparkline-panel" :style="appearance"
      @submit.prevent="bindings.apply" @compositionstart.capture="bindings.startComposition" @compositionend.capture="bindings.endComposition">
      <Alert v-if="stale" type="warning" title="内容、用户或历史已更改。当前草稿已保留，请重新读取后再应用。" show-icon>
        <template #action><Button size="small" :disabled="busy || composing" @click="bindings.reload">重新读取</Button></template>
      </Alert>
      <Alert v-else-if="!canConfigure" type="info" title="当前用户没有修改目标区域或整组迷你图的权限。" show-icon />
      <template v-if="existing">
        <FormItem name="scope" label="编辑范围"><Select :value="scope" :options="scopeOptions" :disabled="busy || composing || dirty || stale" aria-label="迷你图编辑范围" @update:value="bindings.chooseScope" /></FormItem>
        <FormItem name="primary" label="当前迷你图"><Select :value="draft.primary" :options="primaryOptions" :disabled="busy || composing || dirty || stale" aria-label="当前迷你图单元格" @update:value="bindings.choosePrimary" /></FormItem>
        <FormItem name="dataOperation" label="数据操作"><Select :value="draft.dataOperation" :options="operationOptions" :disabled="disabled || dirty" aria-label="迷你图数据操作" @update:value="bindings.chooseOperation" /></FormItem>
        <Alert v-if="scope === '*' && selectedGroups.length > 1 && draft.dataOperation === 'settings'" type="info" title="只将修改过的设置应用到选区中的全部组；未修改的设置继续分别保留。" show-icon />
      </template>
      <div v-if="!existing || draft.dataOperation !== 'settings'" class="sparkline-data-grid">
        <FormItem name="sourceText" :label="t('sheet.sparkline.source-range', '源数据区域')"><Input :value="draft.sourceText" :disabled="textDisabled" aria-label="迷你图源数据区域" placeholder="B1:E1; '数据表'!B2:E2" @update:value="bindings.source" /></FormItem>
        <FormItem name="targetText" :label="t('sheet.sparkline.target-range', '放置区域')"><Input :value="draft.targetText" :disabled="textDisabled" aria-label="迷你图放置区域" placeholder="F1:F2 或 F1; H3" @update:value="bindings.target" /></FormItem>
        <div class="sparkline-range-help">{{ t('sheet.sparkline.multiple-ranges-help', '多个区域用分号分隔；单项数据修改只接受一个源区域和一个目标格。') }}</div>
      </div>
      <Tabs :active-key="tab" :items="tabs" :animated="false" :destroy-on-hidden="false" aria-label="迷你图设置标签页" @update:active-key="bindings.chooseTab" />
      <SparklineSettingsFields :key="generation" :settings="draft.settings" :tab="tab" :disabled="disabled" :text-disabled="textDisabled" :existing="existing"
        :can-manage="canManage" :can-combine="canCombine" :can-uncombine="canUncombine" @change="bindings.settings" @clear="bindings.clear" @combine="bindings.combine" @uncombine="bindings.uncombine" />
      <Alert v-if="validationError || error" type="error" :title="validationError || error" show-icon />
      <Alert v-if="status" type="info" :title="status" show-icon role="status" />
      <div class="sparkline-actions"><Button :disabled="busy || composing" @click="bindings.close">{{ t('common.cancel', '取消') }}</Button><Button type="primary" :loading="busy" :disabled="disabled || !!validationError || existing && !dirty" @click="bindings.apply">{{ t('common.apply', '应用') }}</Button></div>
    </Form>
    <template v-else><Alert v-if="error" type="error" :title="error" show-icon /><Button :disabled="busy || composing" @click="modalBindings.reload">重新读取迷你图</Button></template>
  </Modal>
</template>

<style scoped>
.sparkline-form { min-width: 0; color: var(--sparkline-text); }
.sparkline-form :deep(.ant-form-item) { margin-bottom: 12px; }
.sparkline-data-grid { display: grid; grid-template-columns: minmax(0, 1fr); gap: 0; }
.sparkline-range-help { margin-bottom: 12px; color: var(--sparkline-muted); font-size: 12px; line-height: 1.6; }
.sparkline-actions { display: flex; justify-content: flex-end; gap: 8px; padding-top: 16px; border-top: 1px solid var(--sparkline-border); margin-top: 16px; }
.sparkline-form :deep(.ant-alert) { margin-bottom: 12px; }
</style>
