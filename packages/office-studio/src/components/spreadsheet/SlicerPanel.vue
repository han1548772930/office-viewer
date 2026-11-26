<script setup lang="ts">
import { Alert, Button, Checkbox, Drawer, Form, FormItem, Input, InputNumber, Modal, Popover, Segmented, Select, Tooltip, theme } from 'antdv-next'
import { computed, nextTick, onBeforeUnmount, ref, shallowRef, watch } from 'vue'
import { ArrowDown, ArrowDownToLine, ArrowUp, ArrowUpToLine, ChevronDown, Link } from 'lucide-vue-next'
import type { OfficeStudio } from '../../api/office-studio'
import type { CellRange, SlicerModel, SlicerType } from '../../domain/model'
import { slicerPalette } from '../../features/slicer/slicer-renderer'
import {
  applyPanelSlicerPresentation, insertPanelSlicer, readPanelSlicerArrangement, readSlicerPanelState, slicerPanelSourceKey,
  type SlicerPanelSnapshot,
} from '../../features/slicer/slicer-panel-state'
import { cloneModel } from '../../domain/serialization.ts'
import { captureSlicerPanelOwner, slicerPanelOwnerContext, slicerPanelOwnerCurrent, type SlicerPanelOwner } from '../../features/slicer/slicer-panel-owner'
import { applyPanelSlicerConnections, readPanelSlicerConnections, slicerConnectionKey, slicerConnectionKeys, slicerConnectionsPanelCurrent,
  type SlicerConnectionsPanelSnapshot } from '../../features/slicer/slicer-connections-panel-state'
import { SLICER_STYLE_NAMES, listSlicerStyleNames } from '../../features/slicer/slicer-style-catalog'
import { applyPanelSlicerSettings, buildPanelSlicerGeometryPatch, buildPanelSlicerSettingsPatch, createSlicerPanelEditDraft,
  readPanelSlicerSettings, slicerSettingsPanelCurrent, slicerSettingsPanelDraftError,
  type SlicerPanelEditDraft, type SlicerSettingsPanelSnapshot } from '../../features/slicer/slicer-settings-panel-state.ts'
import SlicerCacheOptionsDialog from '../../features/slicer/SlicerCacheOptionsDialog.vue'
import { applyPanelTimelineSettings, buildPanelTimelineGeometryPatch, buildPanelTimelineSettingsPatch, createTimelinePanelEditDraft,
  readPanelTimelineSettings, timelineSettingsPanelCurrent, timelineSettingsPanelDraftError, type TimelineSettingsPanelSnapshot } from '../../features/slicer/timeline-settings-panel-state.ts'
import { readPanelTimelineDateFilter, timelinePanelDateFilterSummary, type TimelineRangePanelSnapshot } from '../../features/slicer/timeline-range-panel-state.ts'
import { normalizeTimelineLevel, toTimelinePublicLevel } from '../../features/slicer/timeline-date-filter.ts'
import TimelineDateInput from '../../features/slicer/TimelineDateInput.vue'
import TimelineDateFilterDialog from '../../features/slicer/TimelineDateFilterDialog.vue'

const props = withDefaults(defineProps<{
  open: boolean
  studio: OfficeStudio
  workbookId: string
  sheetId: string
  sourceRange?: CellRange
  initialKind?: SlicerType
  selectedSlicerId?: string
  initialSection?: 'settings' | 'styles' | 'connections'
}>(), { initialKind: 'slicer', initialSection: 'settings' })
const emit = defineEmits<{
  'update:open': [open: boolean]
  changed: [slicerId: string]
  closed: [owner: SlicerPanelOwner | undefined]
  error: [message: string]
}>()

const snapshot = shallowRef<SlicerPanelSnapshot | null>(null)
const fieldId = ref('')
const sourceId = ref('')
const busy = ref(false)
const error = ref('')
const settingsId = ref('')
const draft = ref<SlicerPanelEditDraft | null>(null)
const originalDraft = ref<SlicerPanelEditDraft | null>(null)
const settingsPage = ref('slicer')
const styleGalleryOpen = ref(false)
const settingsOwner = shallowRef<SlicerPanelOwner>()
const settingsSnapshot = shallowRef<SlicerSettingsPanelSnapshot>()
const timelineSettingsSnapshot = shallowRef<TimelineSettingsPanelSnapshot>()
const timelineFilterSnapshot = shallowRef<TimelineRangePanelSnapshot>()
const timelineFilterOpen = ref(false)
const draftGeneration = ref(0)
const cacheOptionsOpen = ref(false)
const composing = ref(false)
const panelOwner = shallowRef<SlicerPanelOwner>()
const connectionsOpen = ref(false)
const connectionsLoading = ref(false)
const connectionsSnapshot = shallowRef<SlicerConnectionsPanelSnapshot>()
const connectionKeys = ref<string[]>([])
const connectionsError = ref('')
const { token } = theme.useToken()
const panelStyle = computed(() => ({
  '--slicer-text': token.value.colorText, '--slicer-muted': token.value.colorTextSecondary,
  '--slicer-border': token.value.colorBorderSecondary, '--slicer-background': token.value.colorBgElevated,
  '--slicer-primary': token.value.colorPrimary, '--slicer-error': token.value.colorError,
}))
let unsubscribe: (() => void) | undefined
let requestGeneration = 0
let connectionGeneration = 0
let busyGeneration = 0
let disposed = false
let compositionDraft: SlicerPanelEditDraft | undefined
let connectionReturnFocus: { owner: SlicerPanelOwner; element: HTMLElement } | undefined
const target = computed(() => ({ workbookId: props.workbookId, sheetId: props.sheetId }))
const setting = computed(() => snapshot.value?.controls.find(({ info }) => info.id === settingsId.value))
const arrangement = computed(() => {
  if (!snapshot.value || !settingsId.value) return { forward: false, backward: false, front: false, back: false }
  return readPanelSlicerArrangement(props.studio, target.value, settingsId.value)
})
const selectedRangeSource = computed(() => {
  const range = props.sourceRange
  if (!range) return undefined
  return snapshot.value?.pivots.find((source) => source.sheetId === props.sheetId
    && range.startRow >= source.range.startRow && range.endRow <= source.range.endRow
    && range.startCol >= source.range.startCol && range.endCol <= source.range.endCol
    && (props.initialKind !== 'timeline' || source.kind !== 'table'))
})
const insertionSource = computed(() => snapshot.value?.pivots.find(source => slicerPanelSourceKey(source) === sourceId.value) ?? selectedRangeSource.value)
const sourceOptions = computed(() => snapshot.value?.pivots.filter(source => props.initialKind !== 'timeline' || source.kind !== 'table')
  .map(source => ({ label: source.label, value: slicerPanelSourceKey(source) })) ?? [])
const fieldOptions = computed(() => insertionSource.value?.fields.flatMap(({ id, name, available, controls, reason }) => {
  if (!available || !controls.includes(props.initialKind)) return []
  return [{ label: name, value: id, title: reason }]
}) ?? [])
const styleOptions = computed(() => {
  void snapshot.value?.revision
  const type = draft.value?.type === 'timeline' ? 'timeline' : 'slicer', workbook = props.studio.getWorkbookModel(props.workbookId)
  const styles = workbook ? listSlicerStyleNames(workbook, type) : [...SLICER_STYLE_NAMES[type]]
  const options = [{ label: '文档默认值', value: '' }, ...styles.map((value, index) =>
    ({ label: SLICER_STYLE_NAMES[type].includes(value) ? `样式 ${index + 1}` : value, value }))]
  const current = draft.value?.style
  if (current && !styles.includes(current)) options.push({ label: `当前样式（${current}）`, value: current })
  return options
})
const selectedStyleLabel = computed(() => {
  const style = draft.value?.style ?? ''
  return styleOptions.value.find(option => option.value === style)?.label ?? '样式'
})
const settingsDirty = computed(() => Boolean(draft.value && originalDraft.value && (draft.value.type === 'slicer'
  ? Object.keys(buildPanelSlicerSettingsPatch(draft.value, originalDraft.value)).length || Object.keys(buildPanelSlicerGeometryPatch(draft.value, originalDraft.value)).length
  : Object.keys(buildPanelTimelineSettingsPatch(draft.value, originalDraft.value)).length || Object.keys(buildPanelTimelineGeometryPatch(draft.value, originalDraft.value)).length)))
const settingsStale = computed(() => Boolean(settingsOwner.value && (snapshot.value?.revision !== settingsOwner.value.revision
  || !slicerPanelOwnerCurrent(settingsOwner.value) || settingsSnapshot.value && !slicerSettingsPanelCurrent(settingsSnapshot.value)
  || timelineSettingsSnapshot.value && !timelineSettingsPanelCurrent(timelineSettingsSnapshot.value))))
const canConfigure = computed(() => setting.value?.canConfigure === true && !settingsStale.value
  && (setting.value.info.type === 'slicer' ? Boolean(settingsSnapshot.value) : Boolean(timelineSettingsSnapshot.value)))
const settingsDisabled = computed(() => busy.value || !canConfigure.value || composing.value)
const textDisabled = computed(() => busy.value || !canConfigure.value && !composing.value)
const settingsValidationError = computed(() => draft.value?.type === 'slicer' && originalDraft.value && settingsSnapshot.value
  ? slicerSettingsPanelDraftError(settingsSnapshot.value, draft.value, originalDraft.value)
  : draft.value?.type === 'timeline' && originalDraft.value && timelineSettingsSnapshot.value
    ? timelineSettingsPanelDraftError(timelineSettingsSnapshot.value, draft.value, originalDraft.value) : undefined)
const timelineFilterSummary = computed(() => timelineFilterSnapshot.value
  ? timelinePanelDateFilterSummary(timelineFilterSnapshot.value.state, timelineFilterSnapshot.value.date1904) : '日期筛选暂不可读取')
const supportsConnections = computed(() => Boolean(setting.value && !setting.value.info.tableId))
const connectionsStale = computed(() => Boolean(connectionsSnapshot.value && (snapshot.value?.revision !== connectionsSnapshot.value.owner.revision
  || !slicerConnectionsPanelCurrent(connectionsSnapshot.value))))
const connectionsDirty = computed(() => {
  const current = connectionsSnapshot.value
  if (!current) return false
  const original = new Set(slicerConnectionKeys(current.connections)), selected = new Set(connectionKeys.value)
  return original.size !== selected.size || [...original].some(key => !selected.has(key))
})
const settingsValid = computed(() => {
  const current = draft.value
  const layout = current?.layout
  if (!current || !layout) return false
  return Boolean(!settingsValidationError.value
    && Number.isFinite(layout.left) && Number.isFinite(layout.top)
    && Number.isFinite(layout.width) && Number.isFinite(layout.height) && layout.width >= 80 && layout.height >= 48
    && (current.type !== 'slicer' || current.columnCount === undefined || Number.isSafeInteger(current.columnCount) && current.columnCount >= 1 && current.columnCount <= 20000))
})
const timelineLevels = [{ label: '年', value: 0 }, { label: '季度', value: 1 }, { label: '月', value: 2 }, { label: '日', value: 3 }]

function refresh(): void {
  if (!props.open || disposed) return
  try {
    snapshot.value = readSlicerPanelState(props.studio, target.value)
    const current = snapshot.value.controls.find(({ info }) => info.id === settingsId.value)?.info
    if (!settingsId.value) return
    if (current?.type === 'timeline') {
      try { timelineFilterSnapshot.value = readPanelTimelineDateFilter(props.studio, target.value, current.id) }
      catch (failure) { timelineFilterSnapshot.value = undefined; error.value = failure instanceof Error ? failure.message : String(failure) }
    }
    if (!current && !composing.value) {
      settingsId.value = ''
      draft.value = null
      originalDraft.value = null
      settingsOwner.value = undefined
      settingsSnapshot.value = undefined
      timelineSettingsSnapshot.value = undefined
      timelineFilterSnapshot.value = undefined
      timelineFilterOpen.value = false
      cacheOptionsOpen.value = false
      closeConnections()
    } else if (current && !settingsDirty.value && !composing.value) {
      loadDraft(current)
    }
  } catch (failure) {
    snapshot.value = null
    error.value = failure instanceof Error ? failure.message : String(failure)
  }
}

function loadDraft(info: SlicerModel): void {
  if (composing.value) return
  const layout = info.layout ? { ...info.layout } : { left: 200, top: 200, width: info.type === 'timeline' ? 420 : 220, height: info.type === 'timeline' ? 140 : 280 }
  settingsSnapshot.value = undefined; timelineSettingsSnapshot.value = undefined
  try {
    if (info.type === 'slicer') settingsSnapshot.value = readPanelSlicerSettings(props.studio, target.value, info.id)
    else timelineSettingsSnapshot.value = readPanelTimelineSettings(props.studio, target.value, info.id)
  } catch (failure) { error.value = failure instanceof Error ? failure.message : String(failure) }
  draft.value = info.type === 'timeline' && timelineSettingsSnapshot.value
    ? createTimelinePanelEditDraft({ ...info, layout }, timelineSettingsSnapshot.value.settings)
    : createSlicerPanelEditDraft({ ...info, layout }, settingsSnapshot.value?.settings)
  draftGeneration.value += 1
  originalDraft.value = cloneModel(draft.value)
  settingsOwner.value = settingsSnapshot.value?.owner ?? timelineSettingsSnapshot.value?.owner ?? captureSlicerPanelOwner(props.studio, target.value, snapshot.value?.revision)
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
watch([() => props.open, () => props.studio, () => props.workbookId, () => props.sheetId, () => props.sourceRange, () => props.initialKind, () => props.initialSection], () => {
  requestGeneration += 1
  connectionGeneration += 1
  busyGeneration += 1
  busy.value = false
  connectionsLoading.value = false
  connectionsSnapshot.value = undefined
  connectionsError.value = ''
  connectionReturnFocus = undefined
  connectionsOpen.value = props.open && props.initialSection === 'connections'
  fieldId.value = ''
  sourceId.value = ''
  settingsId.value = props.selectedSlicerId ?? ''
  draft.value = null
  originalDraft.value = null
  settingsOwner.value = undefined
  settingsSnapshot.value = undefined
  timelineSettingsSnapshot.value = undefined
  timelineFilterSnapshot.value = undefined
  timelineFilterOpen.value = false
  cacheOptionsOpen.value = false
  composing.value = false
  compositionDraft = undefined
  panelOwner.value = undefined
  settingsPage.value = props.initialKind === 'timeline' ? 'timeline' : 'slicer'
  error.value = ''
  refresh()
  if (props.open) {
    try { panelOwner.value = captureSlicerPanelOwner(props.studio, target.value) }
    catch (failure) { error.value = failure instanceof Error ? failure.message : String(failure) }
  }
  sourceId.value = selectedRangeSource.value ? slicerPanelSourceKey(selectedRangeSource.value) : sourceOptions.value[0]?.value ?? ''
  if (connectionsOpen.value && settingsId.value) void loadConnections()
  styleGalleryOpen.value = props.open && props.initialSection === 'styles'
}, { immediate: true })
watch(() => props.selectedSlicerId, (id) => {
  requestGeneration += 1
  connectionGeneration += 1
  busyGeneration += 1
  busy.value = false
  connectionsLoading.value = false
  settingsId.value = id ?? ''
  draft.value = null
  originalDraft.value = null
  settingsOwner.value = undefined
  settingsSnapshot.value = undefined
  timelineSettingsSnapshot.value = undefined
  timelineFilterSnapshot.value = undefined
  timelineFilterOpen.value = false
  cacheOptionsOpen.value = false
  composing.value = false
  compositionDraft = undefined
  connectionsSnapshot.value = undefined
  connectionReturnFocus = undefined
  if (settingsId.value) refresh()
  if (connectionsOpen.value && settingsId.value) void loadConnections()
})
watch(insertionSource, (source) => {
  if (!source?.fields.some(({ id, available, controls }) => id === fieldId.value && available && controls.includes(props.initialKind))) {
    fieldId.value = source?.fields.find(({ available, controls }) => available && controls.includes(props.initialKind))?.id ?? ''
  }
})
watch(setting, (value) => {
  if (!value || !settingsId.value || settingsDirty.value || composing.value) return
  loadDraft(value.info)
})
onBeforeUnmount(() => { disposed = true; requestGeneration += 1; connectionGeneration += 1; compositionDraft = undefined; unsubscribe?.() })

async function perform(action: (owner: SlicerPanelOwner) => Promise<string | void>, expectedRevision?: number): Promise<void> {
  if (busy.value) return
  const generation = requestGeneration, studio = props.studio, workbookId = props.workbookId, sheetId = props.sheetId, userId = props.studio.currentUserId
  const busyToken = ++busyGeneration
  busy.value = true
  error.value = ''
  const samePanel = () => !disposed && props.open && requestGeneration === generation && props.studio === studio
    && props.workbookId === workbookId && props.sheetId === sheetId && studio.currentUserId === userId
  try {
    const owner = captureSlicerPanelOwner(studio, { workbookId, sheetId }, expectedRevision)
    panelOwner.value = owner
    const id = await action(owner)
    if (!samePanel() || !slicerPanelOwnerCurrent(owner, true)) return
    refresh()
    if (id) { emit('changed', id); if (!settingsId.value) close() }
  } catch (failure) {
    if (!samePanel()) return
    error.value = failure instanceof Error ? failure.message : String(failure)
    emit('error', error.value)
  } finally { if (!disposed && busyToken === busyGeneration) busy.value = false }
}

function insert(): void {
  const source = insertionSource.value
  if (!source) return
  const input = { ...(source.kind === 'pivot' ? { pivotTableId: source.id } : source.kind === 'ownedPivotChart' ? { ownedPivotChartId: source.id } : { tableId: source.id }), sourceFieldId: fieldId.value, type: props.initialKind }
  void perform(async (owner) => {
    const id = await insertPanelSlicer(owner.studio, owner.target, input, slicerPanelOwnerContext(owner))
    return id
  })
}

function saveSettings(): void {
  if (!draft.value || !originalDraft.value || !settingsOwner.value || settingsDisabled.value || !settingsValid.value || !settingsDirty.value) return
  const current = draft.value
  const original = originalDraft.value, settings = settingsSnapshot.value, timeline = timelineSettingsSnapshot.value
  const generation = requestGeneration
  void perform(async (owner) => {
    if (current.type === 'slicer' && settings) {
      if (!await applyPanelSlicerSettings(settings, buildPanelSlicerSettingsPatch(current, original), buildPanelSlicerGeometryPatch(current, original),
        () => !disposed && props.open && requestGeneration === generation && settingsId.value === current.id)) throw new Error('切片器显示设置未接受，请检查编辑权限和来源。')
    } else if (current.type === 'timeline' && timeline) {
      if (!await applyPanelTimelineSettings(timeline, buildPanelTimelineSettingsPatch(current, original), buildPanelTimelineGeometryPatch(current, original),
        () => !disposed && props.open && requestGeneration === generation && settingsId.value === current.id)) throw new Error('时间线显示设置未接受，请检查修改权限和来源。')
    } else throw new Error('显示设置尚未读取，请重新读取后再保存。')
    if (generation === requestGeneration && slicerPanelOwnerCurrent(owner, true)) {
      snapshot.value = readSlicerPanelState(owner.studio, owner.target)
      const latest = snapshot.value.controls.find(control => control.info.id === current.id)
      if (latest) loadDraft(latest.info)
    }
    return current.id
  }, settingsOwner.value.revision)
}

function close(): void {
  if (composing.value) return
  const owner = connectionsSnapshot.value?.owner ?? settingsOwner.value ?? panelOwner.value
  requestGeneration += 1; connectionGeneration += 1; busyGeneration += 1; busy.value = false; settingsId.value = ''; connectionsOpen.value = false
  connectionReturnFocus = undefined
  styleGalleryOpen.value = false
  cacheOptionsOpen.value = false
  timelineFilterOpen.value = false
  emit('update:open', false)
  emit('closed', owner)
}
function stylePreview(style: string): Record<string, string> {
  const palette = slicerPalette(style || undefined, draft.value?.type === 'timeline', {
    background: token.value.colorBgContainer, text: token.value.colorText, secondaryFill: token.value.colorFillSecondary, mutedText: token.value.colorTextTertiary,
  }, props.studio.getWorkbookModel(props.workbookId))
  return { '--preview-header': palette.header, '--preview-item': palette.item, '--preview-border': palette.itemBorder, '--preview-text': palette.text }
}
function chooseStyle(style: string): void {
  if (settingsDisabled.value) return
  if (draft.value) draft.value.style = style || undefined
  styleGalleryOpen.value = false
}

function setAnchorType(anchorType: unknown): void {
  if (!draft.value || settingsDisabled.value || (anchorType !== 'both' && anchorType !== 'position' && anchorType !== 'none')) return
  draft.value.drawing = {
    ...(draft.value.drawing ?? { from: { row: 0, column: 0, offsetX: 0, offsetY: 0 }, to: { row: 0, column: 0, offsetX: 0, offsetY: 0 } }),
    anchorType,
  }
}

function setKeepRatio(value: boolean | { target?: { checked?: boolean } }): void {
  if (!draft.value || settingsDisabled.value) return
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
  if (!draft.value || settingsDisabled.value) return
  const slicerId = draft.value.id
  void perform(async (owner) => {
    await applyPanelSlicerPresentation(owner.studio, owner.target, slicerId, { arrange: action }, slicerPanelOwnerContext(owner))
    return slicerId
  })
}

async function loadConnections(): Promise<void> {
  if (!props.open || !settingsId.value || !supportsConnections.value) return
  const generation = ++connectionGeneration, studio = props.studio, workbookId = props.workbookId, sheetId = props.sheetId, slicerId = settingsId.value, userId = props.studio.currentUserId
  connectionsLoading.value = true
  connectionsError.value = ''
  const current = () => !disposed && props.open && connectionsOpen.value && connectionGeneration === generation && props.studio === studio
    && props.workbookId === workbookId && props.sheetId === sheetId && settingsId.value === slicerId && studio.currentUserId === userId
  try {
    const next = await readPanelSlicerConnections(studio, { workbookId, sheetId }, slicerId)
    if (!current()) return
    snapshot.value = readSlicerPanelState(studio, { workbookId, sheetId })
    connectionsSnapshot.value = next
    connectionKeys.value = slicerConnectionKeys(next.connections)
  } catch (failure) {
    if (current()) { connectionsSnapshot.value = undefined; connectionsError.value = failure instanceof Error ? failure.message : String(failure) }
  } finally { if (current()) connectionsLoading.value = false }
}

function openConnections(): void {
  if (composing.value || busy.value || !canConfigure.value) return
  const element = document.activeElement
  if (element instanceof HTMLElement && settingsOwner.value && slicerPanelOwnerCurrent(settingsOwner.value)) connectionReturnFocus = { owner: settingsOwner.value, element }
  connectionsOpen.value = true
  void loadConnections()
}

function setCaptionFromName(value: boolean): void {
  if (!draft.value || settingsDisabled.value) return
  if (!value) draft.value.caption = draft.value.name
  draft.value.captionFromName = value
}
function setViewNumber(key: 'columnCount' | 'rowHeight' | 'startItem', value: unknown): void {
  if (!draft.value || settingsDisabled.value) return
  draft.value[key] = typeof value === 'number' ? value : undefined
}
function resetDisplayDefaults(): void {
  if (!draft.value || settingsDisabled.value) return
  draft.value.captionFromName = true; draft.value.caption = draft.value.name
  if (draft.value.type === 'slicer') {
    draft.value.rowHeight = 24
    for (const key of ['startItem', 'columnCount', 'showCaption', 'lockedPosition', 'style'] as const) delete draft.value[key]
  } else {
    draft.value.level = 'month'; draft.value.selectionLevel = 'month'
    for (const key of ['style', 'scrollPosition', 'showHeader', 'showSelectionLabel', 'showTimeLevel', 'showHorizontalScrollbar'] as const) delete draft.value[key]
  }
  draftGeneration.value += 1
}
function setTimelineLevel(key: 'level' | 'selectionLevel', value: unknown): void {
  if (!draft.value || settingsDisabled.value) return
  try { draft.value[key] = normalizeTimelineLevel(value) }
  catch (failure) { error.value = failure instanceof Error ? failure.message : String(failure) }
}
function setTimelineNavigation(value: number | undefined): void {
  if (!draft.value || draft.value.type !== 'timeline' || textDisabled.value) return
  draft.value.scrollPosition = value
}
function compositionStart(event: CompositionEvent): void {
  if (settingsDisabled.value || !(event.target instanceof HTMLInputElement || event.target instanceof HTMLTextAreaElement)) return
  compositionDraft = draft.value ?? undefined; composing.value = true
}
function compositionEnd(): void {
  if (compositionDraft !== draft.value) return
  composing.value = false; compositionDraft = undefined
  if (!setting.value) refresh()
}
function settingsKey(event: KeyboardEvent): void {
  if ((event.isComposing || event.keyCode === 229 || composing.value) && (event.key === 'Enter' || event.key === 'Escape')) {
    event.preventDefault(); event.stopPropagation()
  }
}
function cacheOptionsChanged(id: string): void { refresh(); emit('changed', id) }
function timelineFilterChanged(id: string): void { refresh(); emit('changed', id) }
function closeConnections(): void {
  if (props.initialSection === 'connections') { close(); return }
  const returnFocus = connectionReturnFocus, generation = requestGeneration
  connectionReturnFocus = undefined
  connectionsOpen.value = false
  connectionGeneration += 1
  connectionsSnapshot.value = undefined
  connectionsLoading.value = false
  if (returnFocus) void nextTick(() => {
    if (disposed || requestGeneration !== generation || connectionsOpen.value || !props.open || props.studio !== returnFocus.owner.studio
      || props.workbookId !== returnFocus.owner.target.workbookId || props.sheetId !== returnFocus.owner.target.sheetId
      || !slicerPanelOwnerCurrent(returnFocus.owner, true) || !returnFocus.element.isConnected) return
    returnFocus.element.focus({ preventScroll: true })
  })
}

async function saveConnections(): Promise<void> {
  const current = connectionsSnapshot.value
  if (!current || busy.value || connectionsStale.value || !connectionsDirty.value || !setting.value?.canConfigure) return
  const generation = connectionGeneration, keys = [...connectionKeys.value], busyToken = ++busyGeneration
  const samePanel = () => !disposed && generation === connectionGeneration && props.open && props.studio === current.owner.studio
    && props.workbookId === current.owner.target.workbookId && props.sheetId === current.owner.target.sheetId
    && settingsId.value === current.slicerId && current.owner.studio.currentUserId === current.owner.userId
  busy.value = true
  connectionsError.value = ''
  try {
    const accepted = await applyPanelSlicerConnections(current, keys)
    if (!samePanel() || !slicerConnectionsPanelCurrent(current, true)) return
    if (!accepted) { connectionsError.value = '连接更新未接受，请检查报表兼容性和编辑权限。'; return }
    refresh()
    emit('changed', current.slicerId)
    closeConnections()
  } catch (failure) {
    if (samePanel()) connectionsError.value = failure instanceof Error ? failure.message : String(failure)
  } finally { if (!disposed && busyToken === busyGeneration) busy.value = false }
}
</script>

<template>
  <Modal :open="open && !settingsId && !connectionsOpen" :focus-trigger-after-close="false" :title="initialKind === 'timeline' ? '插入时间线' : '插入切片器'" :root-style="panelStyle" :width="400" :footer="null" :mask="{ closable: false }" destroy-on-hidden @cancel="close">
    <Form class="slicer-insert" :style="panelStyle" layout="vertical" @submit.prevent="insert">
      <FormItem label="来源"><Select v-model:value="sourceId" :options="sourceOptions" aria-label="切片器来源" placeholder="选择报表或表格" :disabled="busy || !sourceOptions.length" /></FormItem>
      <FormItem label="字段"><Select v-model:value="fieldId" :options="fieldOptions" aria-label="字段" placeholder="选择字段" :disabled="busy || !insertionSource" /></FormItem>
      <p v-if="!insertionSource" class="slicer-muted">请选择数据透视表或表格中的单元格</p>
      <p v-else-if="!fieldOptions.length" class="slicer-muted">当前选区没有可用字段</p>
      <Alert v-if="error" :title="error" type="error" show-icon />
      <div class="slicer-actions"><Button @click="close">取消</Button><Button type="primary" :loading="busy" :disabled="!fieldId || busy" @click="insert">确定</Button></div>
    </Form>
  </Modal>

  <Drawer :open="open && !!settingsId && initialSection !== 'connections'" :root-style="panelStyle" :title="draft?.type === 'timeline' ? '时间线设置' : '切片器设置'" :size="388" placement="right" :mask="false" :keyboard="!composing" destroy-on-hidden @close="close">
    <Form v-if="draft" class="slicer-settings" data-testid="slicer-settings-form" :style="panelStyle" layout="vertical" @submit.prevent="saveSettings"
      @compositionstart="compositionStart" @compositionend="compositionEnd" @keydown="settingsKey">
      <Segmented v-model:value="settingsPage" :options="[{ value: draft.type, label: draft.type === 'timeline' ? '时间线' : '切片器' }, { value: 'size', label: '大小与属性' }]" :disabled="busy || composing" block />
      <Alert v-if="settingsStale" title="内容已更改，当前设置草稿需要重新读取。" type="warning" show-icon>
        <template #action><Button html-type="button" size="small" :disabled="busy || composing || !setting" @click="setting && loadDraft(setting.info)">重新读取</Button></template>
      </Alert>
      <Alert v-else-if="!setting?.canConfigure" title="当前用户没有修改此切片器的权限。" type="info" show-icon />
      <FormItem :label="null">
        <div class="slicer-context-actions">
          <Button v-if="supportsConnections" html-type="button" :disabled="settingsDisabled" @click="openConnections"><Link :size="14" />报表连接</Button>
          <Button v-if="draft.type === 'slicer'" html-type="button" :disabled="busy || composing || !setting?.canConfigure" @click="cacheOptionsOpen = true">排序与无数据项</Button>
          <Button v-else html-type="button" :disabled="busy || composing || !setting?.canConfigure" @click="timelineFilterOpen = true">日期筛选</Button>
        </div>
      </FormItem>
      <p v-if="draft.type === 'timeline'" class="slicer-muted" data-testid="timeline-panel-retained-filter">当前筛选：{{ timelineFilterSummary }}</p>
      <template v-if="settingsPage !== 'size'">
        <section class="slicer-group">
          <header>{{ draft.type === 'timeline' ? '时间线' : '切片器' }}</header>
          <FormItem label="名称"><Input v-model:value="draft.name" :aria-label="draft.type === 'timeline' ? '时间线名称' : '切片器名称'" :disabled="textDisabled" /></FormItem>
            <FormItem label="标题"><Input :value="draft.captionFromName ? draft.name : draft.caption" :maxlength="32767" aria-label="标题"
              :disabled="textDisabled || draft.captionFromName" @update:value="draft.caption = $event" /></FormItem>
            <FormItem :label="null"><Checkbox :checked="draft.captionFromName" :disabled="settingsDisabled" @update:checked="setCaptionFromName">使用名称作为标题</Checkbox></FormItem>
          <FormItem :label="null"><Checkbox v-if="draft.type === 'slicer'" :checked="draft.showCaption ?? true" :disabled="settingsDisabled" @update:checked="draft.showCaption = $event">显示标题</Checkbox><Checkbox v-else :checked="draft.showHeader ?? true" :disabled="settingsDisabled" @update:checked="draft.showHeader = $event">显示标题</Checkbox></FormItem>
          <template v-if="draft.type === 'slicer'">
            <FormItem class="slicer-number" layout="horizontal" label="列数"><InputNumber :value="draft.columnCount ?? 1" :min="1" :max="20000" :precision="0" aria-label="列数" :disabled="settingsDisabled" @update:value="setViewNumber('columnCount', $event)" /></FormItem>
            <FormItem class="slicer-number" layout="horizontal" label="按钮高度"><InputNumber :value="draft.rowHeight" :min="1 / 9525" :max="4294967295 / 9525" :precision="6" aria-label="按钮高度" :disabled="settingsDisabled" @update:value="setViewNumber('rowHeight', $event)" /></FormItem>
            <FormItem class="slicer-number" layout="horizontal" label="首个显示项"><InputNumber :value="draft.startItem === undefined ? 1 : draft.startItem + 1" :min="1" :max="4294967296" :precision="0" aria-label="首个显示项" :disabled="settingsDisabled" @update:value="setViewNumber('startItem', typeof $event === 'number' ? $event - 1 : undefined)" /></FormItem>
            <FormItem :label="null"><Checkbox :checked="draft.lockedPosition ?? false" :disabled="settingsDisabled" @update:checked="draft.lockedPosition = $event">锁定位置</Checkbox></FormItem>
            <FormItem :label="null"><Button html-type="button" size="small" :disabled="settingsDisabled" @click="resetDisplayDefaults">还原显示默认值</Button></FormItem>
          </template>
          <template v-else>
            <FormItem label="时间级别"><Select :value="toTimelinePublicLevel(draft.level ?? 'month')" :options="timelineLevels" aria-label="时间级别" :disabled="settingsDisabled" @update:value="setTimelineLevel('level', $event)" /></FormItem>
            <FormItem label="选择级别"><Select :value="toTimelinePublicLevel(draft.selectionLevel ?? 'month')" :options="timelineLevels" aria-label="选择级别" :disabled="settingsDisabled" @update:value="setTimelineLevel('selectionLevel', $event)" /></FormItem>
            <TimelineDateInput :key="`${draftGeneration}:timeline-navigation`" :value="draft.scrollPosition" :date1904="snapshot?.date1904 ?? false" label="导航日期" aria-label="时间线导航日期" :disabled="textDisabled" :actions-disabled="settingsDisabled" @update:value="setTimelineNavigation" />
            <FormItem :label="null"><Checkbox :checked="draft.showSelectionLabel ?? true" :disabled="settingsDisabled" @update:checked="draft.showSelectionLabel = $event">显示选择标签</Checkbox></FormItem>
            <FormItem :label="null"><Checkbox :checked="draft.showTimeLevel ?? true" :disabled="settingsDisabled" @update:checked="draft.showTimeLevel = $event">显示时间级别</Checkbox></FormItem>
            <FormItem :label="null"><Checkbox :checked="draft.showHorizontalScrollbar ?? true" :disabled="settingsDisabled" @update:checked="draft.showHorizontalScrollbar = $event">显示滚动条</Checkbox></FormItem>
            <FormItem :label="null"><Button html-type="button" size="small" :disabled="settingsDisabled" @click="resetDisplayDefaults">还原显示默认值</Button></FormItem>
          </template>
        </section>
        <section class="slicer-group">
          <header>样式</header>
          <Popover v-model:open="styleGalleryOpen" trigger="click" placement="bottomRight">
            <template #content>
              <div class="slicer-gallery" :style="panelStyle" role="group" aria-label="样式图库">
                <Button v-for="option in styleOptions" :key="option.value || 'default'" html-type="button" class="slicer-style-tile" :class="{ 'is-selected': (draft.style ?? '') === option.value }" :style="stylePreview(option.value)" :title="option.label" :aria-label="option.label" :aria-pressed="(draft.style ?? '') === option.value" :disabled="settingsDisabled" @click="chooseStyle(option.value)">
                  <span v-if="option.value" class="slicer-style-sample"><i /><i /><i /></span><span v-else class="slicer-default-sample">Aa</span>
                </Button>
              </div>
            </template>
            <Button html-type="button" class="slicer-style-trigger" :style="stylePreview(draft.style ?? '')" aria-label="选择样式" :disabled="settingsDisabled">
              <span class="slicer-style-sample"><i /><i /><i /></span><span>{{ selectedStyleLabel }}</span><ChevronDown :size="14" />
            </Button>
          </Popover>
        </section>
      </template>
      <template v-else>
        <section class="slicer-group">
          <header>大小与属性</header>
          <FormItem class="slicer-number" layout="horizontal" label="水平位置"><InputNumber v-model:value="draft.layout!.left" :min="0" :max="1000000" :precision="0" aria-label="水平位置" :disabled="settingsDisabled || draft.lockedPosition" /></FormItem>
          <FormItem class="slicer-number" layout="horizontal" label="垂直位置"><InputNumber v-model:value="draft.layout!.top" :min="0" :max="1000000" :precision="0" aria-label="垂直位置" :disabled="settingsDisabled || draft.lockedPosition" /></FormItem>
          <FormItem class="slicer-number" layout="horizontal" label="宽度"><InputNumber v-model:value="draft.layout!.width" :min="80" :max="100000" :precision="0" aria-label="宽度" :disabled="settingsDisabled || draft.lockedPosition" @change="updateSizeDimension('width', $event)" /></FormItem>
          <FormItem class="slicer-number" layout="horizontal" label="高度"><InputNumber v-model:value="draft.layout!.height" :min="48" :max="100000" :precision="0" aria-label="高度" :disabled="settingsDisabled || draft.lockedPosition" @change="updateSizeDimension('height', $event)" /></FormItem>
          <FormItem label="锚定方式"><Select :value="draft.drawing?.anchorType ?? 'position'" :options="[
            { value: 'both', label: '随单元格移动并调整大小' },
            { value: 'position', label: '随单元格移动，不调整大小' },
            { value: 'none', label: '不随单元格移动或调整大小' },
          ]" aria-label="锚定方式" :disabled="settingsDisabled" @change="setAnchorType" /></FormItem>
          <FormItem :label="null"><Checkbox :checked="draft.drawing?.keepRatio ?? false" :disabled="settingsDisabled" @change="setKeepRatio">保持纵横比</Checkbox></FormItem>
          <FormItem :label="null"><Checkbox :checked="draft.lockedPosition ?? false" :disabled="settingsDisabled" @update:checked="draft.lockedPosition = $event">锁定位置</Checkbox></FormItem>
          <div class="slicer-arrange" role="group" aria-label="对象排列">
            <Tooltip title="置于顶层"><Button html-type="button" aria-label="置于顶层" :disabled="settingsDisabled || !arrangement.front" @click="arrange('front')"><ArrowUpToLine :size="15" /></Button></Tooltip>
            <Tooltip title="上移一层"><Button html-type="button" aria-label="上移一层" :disabled="settingsDisabled || !arrangement.forward" @click="arrange('forward')"><ArrowUp :size="15" /></Button></Tooltip>
            <Tooltip title="下移一层"><Button html-type="button" aria-label="下移一层" :disabled="settingsDisabled || !arrangement.backward" @click="arrange('backward')"><ArrowDown :size="15" /></Button></Tooltip>
            <Tooltip title="置于底层"><Button html-type="button" aria-label="置于底层" :disabled="settingsDisabled || !arrangement.back" @click="arrange('back')"><ArrowDownToLine :size="15" /></Button></Tooltip>
          </div>
        </section>
      </template>
      <Alert v-if="settingsValidationError" :title="settingsValidationError" type="warning" show-icon />
      <Alert v-if="error" :title="error" type="error" show-icon />
      <div class="slicer-actions"><span /><Button :disabled="composing" @click="close">取消</Button><Button type="primary" :loading="busy" :disabled="settingsDisabled || !settingsValid || !settingsDirty" @click="saveSettings">确定</Button></div>
    </Form>
  </Drawer>

  <Modal :open="open && connectionsOpen && !!settingsId" :focus-trigger-after-close="false" title="报表连接" :root-style="panelStyle" :width="520" :footer="null" :mask="{ closable: false }" destroy-on-hidden @cancel="closeConnections">
    <Form class="slicer-connections" :style="panelStyle" layout="vertical" @submit.prevent="saveConnections">
      <p class="slicer-muted">{{ draft?.caption }}：选择使用同一来源和字段的报表。连接更新会应用到共享此切片器缓存的全部视图。</p>
      <Alert v-if="connectionsStale" title="连接列表已过期，请重新读取后再保存。" type="warning" show-icon />
      <Alert v-if="connectionsError" :title="connectionsError" type="error" show-icon />
      <FormItem label="连接的报表">
        <Checkbox.Group v-if="connectionsSnapshot" v-model:value="connectionKeys" class="slicer-connection-list" aria-label="连接的报表" :disabled="busy || connectionsLoading || connectionsStale || !setting?.canConfigure">
          <Checkbox v-for="option in connectionsSnapshot.options" :key="slicerConnectionKey(option)" :value="slicerConnectionKey(option)">
            <span class="slicer-connection-name">{{ option.name }}</span><span class="slicer-connection-detail">{{ option.sheetName }} · {{ option.kind === 'ownedPivotChart' ? '透视图' : '数据透视表' }}</span>
          </Checkbox>
        </Checkbox.Group>
        <p v-if="connectionsSnapshot && !connectionsSnapshot.options.length" class="slicer-muted">当前工作簿没有兼容的报表。</p>
      </FormItem>
      <div class="slicer-actions">
        <Button html-type="button" :loading="connectionsLoading" :disabled="busy || connectionsLoading" @click="loadConnections">重新读取</Button><span />
        <Button html-type="button" @click="closeConnections">取消</Button>
        <Button html-type="button" type="primary" :loading="busy" :disabled="busy || connectionsLoading || connectionsStale || !connectionsDirty || !setting?.canConfigure" @click="saveConnections">确定</Button>
      </div>
    </Form>
  </Modal>

  <SlicerCacheOptionsDialog :open="open && cacheOptionsOpen && !!settingsId && draft?.type === 'slicer'" :studio="studio" :workbook-id="workbookId" :sheet-id="sheetId" :slicer-id="settingsId"
    @update:open="cacheOptionsOpen = $event" @changed="cacheOptionsChanged" />
  <TimelineDateFilterDialog :open="open && timelineFilterOpen && !!settingsId && draft?.type === 'timeline'" :studio="studio" :workbook-id="workbookId" :sheet-id="sheetId" :timeline-id="settingsId"
    @update:open="timelineFilterOpen = $event" @changed="timelineFilterChanged" />
</template>

<style scoped>
.slicer-insert, .slicer-settings, .slicer-connections { display: flex; flex-direction: column; gap: 16px; color: var(--slicer-text); }
.slicer-context-actions { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
.slicer-insert :deep(.ant-form-item), .slicer-settings :deep(.ant-form-item), .slicer-connections :deep(.ant-form-item) { margin-bottom: 0; }
.slicer-settings :deep(.ant-form-item-label) { color: var(--slicer-muted); font-size: 13px; }
.slicer-settings :deep(.slicer-number .ant-form-item-row) { display: grid; grid-template-columns: minmax(0, 1fr) 128px; align-items: center; gap: 8px; }
.slicer-settings :deep(.slicer-number .ant-form-item-label) { text-align: start; }
.slicer-insert :deep(.ant-select), .slicer-settings :deep(.ant-select), .slicer-settings :deep(.ant-input-number) { width: 100%; }
.slicer-group { display: flex; flex-direction: column; gap: 12px; padding: 0 0 16px; border-bottom: 1px solid var(--slicer-border); }
.slicer-group header { color: var(--slicer-text); font-size: 14px; font-weight: 600; }
.slicer-number :deep(.ant-input-number) { width: 128px; }
.slicer-arrange { display: flex; gap: 4px; }
.slicer-arrange .ant-btn { flex: 1; }
.slicer-gallery { width: 320px; max-height: 70vh; overflow: auto; display: grid; grid-template-columns: repeat(3, 72px); justify-content: space-between; gap: 8px; }
.slicer-style-tile { width: 72px; height: 72px; padding: 9px; border: 1px solid var(--slicer-border); background: var(--slicer-background); }
.slicer-style-tile.is-selected { outline: 2px solid var(--slicer-primary); outline-offset: 1px; }
.slicer-style-sample { width: 100%; height: 100%; display: grid; grid-template-rows: 1fr 1fr 1fr; gap: 3px; padding: 4px; background: var(--preview-header); border: 1px solid var(--preview-border); }
.slicer-style-sample i { display: block; background: var(--preview-item); border: 1px solid var(--preview-border); }
.slicer-style-trigger { width: 100%; height: 48px; display: grid; grid-template-columns: 36px 1fr auto; align-items: center; gap: 10px; padding: 5px 9px; text-align: left; color: var(--slicer-text); border: 1px solid var(--slicer-border); background: var(--slicer-background); }
.slicer-style-trigger .slicer-style-sample { width: 36px; height: 36px; }
.slicer-default-sample { width: 100%; height: 100%; display: grid; place-items: center; font-weight: 600; color: var(--preview-text); background: var(--preview-item); }
.slicer-style-tile .slicer-default-sample { font-size: 20px; }
.slicer-actions { display: flex; justify-content: flex-end; gap: 8px; margin-top: auto; }
.slicer-actions > span { flex: 1; }
.slicer-muted, .slicer-error { margin: 0; font-size: 12px; line-height: 1.6; }
.slicer-muted { color: var(--slicer-muted); }
.slicer-error { color: var(--slicer-error); }
.slicer-connection-list { width: 100%; max-height: 340px; overflow: auto; display: flex; flex-direction: column; gap: 12px; }
.slicer-connection-name { display: block; color: var(--slicer-text); }
.slicer-connection-detail { display: block; color: var(--slicer-muted); font-size: 12px; }
</style>
