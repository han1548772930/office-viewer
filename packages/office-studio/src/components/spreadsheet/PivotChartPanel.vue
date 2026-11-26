<script setup lang="ts">
import { computed, h, inject, nextTick, onScopeDispose, ref, shallowRef, watch } from 'vue'
import { Alert, Button, Checkbox, Form, FormItem, Input, Modal, Select, Switch, Tabs, theme } from 'antdv-next'
import { FOffice } from '../../api/facade'
import type { OfficeStudio } from '../../api/office-studio'
import type { CellRange, PivotChartControls, PivotFieldArea, PivotTableOptions } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { createPivotChartDraft, createPivotChartPanelDraft, pivotChartCreateInfo, pivotChartConfigureInput, pivotChartPreviewConfig,
  readPivotChartPanelSource, readPivotChartDraftFieldItems, pivotChartFieldsFromReport, pivotChartPanelInputFromDraft,
  resetPivotChartDraftMapping, canConfigurePivotChartPanel, canViewPivotChartPanelSource,
  PIVOT_CHART_CREATION_MODES, pivotChartCreateOptions, pivotChartCreateOptionsError, setPivotChartDraftControls,
  type PivotChartCreationDraft, type PivotChartDraft, type PivotChartPanelOpenInput, type PivotChartPanelSource } from '../../features/pivot/services/pivot-chart-panel-state'
import { capturePivotChartPanelOwner, pivotChartPanelOwnerContext, pivotChartPanelOwnerCurrent,
  type PivotChartPanelOwner } from '../../features/pivot/services/pivot-chart-panel-owner.ts'
import { addPivotChartDraftField, movePivotChartDraftField, removePivotChartDraftFields, setPivotChartDraftField,
  pivotChartDraftShowDataAsReport, type PivotChartDraftFieldSetting } from '../../features/pivot/services/pivot-chart-field-draft.ts'
import { getPivotShowDataAsBaseItems } from '../../features/pivot/services/pivot-engine.ts'
import { preparePivotShowDataAsDraft } from '../../features/pivot/services/pivot-show-data-as-ui.ts'
import type { PivotChartFieldsConfig } from '../../features/pivot/services/pivot-chart-engine.ts'
import { chartPlacementGeometry } from '../../features/charts/chart-placement'
import type { ChartSettingsSourceLoader, ChartSettingsStatisticsLoader } from '../../features/charts/chart-settings-data'
import type { ChartSettingsDraft } from '../../features/charts/chart-settings-draft.ts'
import { useChartSettingsSource } from '../../features/charts/views/use-chart-settings-source'
import { canEditWorksheetProtection, canViewWorksheetProtection } from '../../features/sheet-protection/range-protection'
import ChartSettingsFields from '../../features/charts/views/ChartSettingsFields.vue'
import PivotChartFieldsForm from '../../features/pivot/views/PivotChartFieldsForm.vue'
import { pivotChartValuesUiOptions, setPivotChartValuesUiDraft } from '../../features/pivot/services/pivot-values-position-ui.ts'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

const props = defineProps<{ open: boolean; studio: OfficeStudio; workbookId: string; sheetId: string; pivotTableId?: string; chartId?: string;
  sourceKind?: 'range' | 'pivotTable'; sourceSheetId?: string; sourceRange?: CellRange; initialFieldId?: string }>()
const emit = defineEmits<{ 'update:open': [value: boolean]; 'busy-change': [value: boolean]; 'owner-change': [owner: PivotChartPanelOwner]; applied: [chartId: string]; closed: [] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const { token } = theme.useToken()
const draft = ref<PivotChartDraft>(createPivotChartDraft())
const panelElement = ref<HTMLElement>()
const owner = shallowRef<PivotChartPanelOwner>(), source = shallowRef<PivotChartPanelSource>()
const error = ref(''), status = ref(''), statusType = ref<'success' | 'info'>('info'), busy = ref(false), composing = ref(false)
const modelRevision = ref(0), sourceKeyRead = ref(''), activeTab = ref('fields')
const showAsValidity = ref<Record<string, boolean>>({})
const filterOpen = ref(false), filterLoading = ref(false), filterId = ref(''), filterAll = ref(true), filterKeys = ref<string[]>([]), filterItems = ref<string[]>([]), filterError = ref('')
let disposed = false, generation = 0, busyGeneration = 0, filterGeneration = 0, compositionGeneration = 0, pendingLoad = false
const workbook = computed(() => { modelRevision.value; return props.studio.getWorkbookModel(props.workbookId) })
const sourceKey = computed(() => JSON.stringify([draft.value.chartId, draft.value.sourceKind, draft.value.pivotTableId, draft.value.sourceSheetId, draft.value.sourceRange]))
const stale = computed(() => { modelRevision.value; return Boolean(owner.value && !pivotChartPanelOwnerCurrent(owner.value)) })
const sourceReady = computed(() => Boolean(source.value && sourceKeyRead.value === sourceKey.value && workbook.value
  && canViewPivotChartPanelSource(workbook.value, source.value.report, hostSheetId.value, chart.value, props.studio.currentUserId, draft.value.sourceKind === 'pivotTable')))
const sheets = computed(() => workbook.value?.sheets.filter(sheet => canViewWorksheetProtection(sheet, props.studio.currentUserId)) ?? [])
const tableTargetSheets = computed(() => sheets.value.filter(sheet => canEditWorksheetProtection(sheet, props.studio.currentUserId)))
const pivotTables = computed(() => workbook.value?.pivotTables?.filter(pivot => [pivot.source.subUnitId, pivot.anchor.subUnitId].every(id =>
  sheets.value.some(sheet => sheet.id === id))) ?? [])
const chart = computed(() => workbook.value?.charts?.find(item => item.id === draft.value.chartId))
const charts = computed(() => workbook.value?.charts?.filter(item => item.sheetId === props.sheetId && (item.pivotTableId || item.ownedPivot)
  && sheets.value.some(sheet => sheet.id === item.sheetId)) ?? [])
const hostSheetId = computed(() => chart.value?.sheetId ?? props.sheetId)
const canConfigure = computed(() => Boolean(workbook.value && canConfigurePivotChartPanel(workbook.value, source.value?.report,
  hostSheetId.value, chart.value, props.studio.currentUserId, draft.value.sourceKind === 'pivotTable')))
const createOptionsError = computed(() => pivotChartCreateOptionsError(draft.value, workbook.value, hostSheetId.value, props.studio.currentUserId))
const navigationDisabled = computed(() => busy.value || composing.value || stale.value || !canConfigure.value || !sourceReady.value)
const inputDisabled = computed(() => busy.value || !composing.value && (!canConfigure.value || !sourceReady.value || stale.value))
const showAsReport = computed(() => sourceReady.value && source.value ? pivotChartDraftShowDataAsReport(source.value.report, draft.value.fieldsConfig) : undefined)
const showAsInvalid = computed(() => draft.value.fieldsConfig.fields.some(field => field.area === 'value' && showAsValidity.value[field.id] === false))
const panelStyle = computed(() => ({ color: token.value.colorText, backgroundColor: token.value.colorBgElevated, '--panel-border': token.value.colorBorderSecondary }))
const sourceRevision = computed(() => JSON.stringify([modelRevision.value, owner.value?.revision, owner.value?.historyGeneration,
  props.studio.currentUserId, sourceKeyRead.value, draft.value.fieldsConfig, draft.value.createMode, draft.value.pivotTableSheetId, draft.value.pivotTableAddress]))
const pivotOptions = computed(() => pivotTables.value.map((pivot, index) => ({ value: pivot.id,
  label: '透视表 ' + (index + 1) + ' · ' + (workbook.value?.sheets.find(sheet => sheet.id === pivot.anchor.subUnitId)?.name ?? pivot.source.sheetName) })))
const chartOptions = computed(() => [{ label: '新建透视图', value: '' }, ...charts.value.map((item, index) => ({
  label: (item.title || '透视图 ' + (index + 1)) + (item.ownedPivot ? ' · 独立报告' : ' · 关联透视表'), value: item.id,
}))])

function requireWorkbook(studio = props.studio, unitId = props.workbookId) {
  const book = FOffice.newAPI(studio).getWorkbook(unitId)
  if (!book) throw new Error('工作簿已关闭，请重新打开透视图面板。')
  return book
}
function currentSession(current: PivotChartDraft, captured: PivotChartPanelOwner, capturedGeneration: number, published = false): boolean {
  return !disposed && props.open && draft.value === current && owner.value === captured && generation === capturedGeneration
    && props.studio === captured.studio && props.workbookId === captured.unitId && props.sheetId === captured.sheetId
    && pivotChartPanelOwnerCurrent(captured, published)
}
function creationSnapshot(current: PivotChartDraft): PivotChartCreationDraft {
  return { chartId: current.chartId, sourceKind: current.sourceKind, sourceSheetId: current.sourceSheetId, sourceRange: current.sourceRange,
    createMode: current.createMode, pivotTableSheetId: current.pivotTableSheetId, pivotTableAddress: current.pivotTableAddress }
}
const loadSource = computed<ChartSettingsSourceLoader>(() => {
  const current = draft.value, captured = owner.value, capturedSource = source.value, capturedGeneration = generation
  const fields = cloneModel(current.fieldsConfig), studio = props.studio, unitId = props.workbookId, host = hostSheetId.value
  const creation = creationSnapshot(current)
  return async (config, options) => {
    if (!captured || !capturedSource || options.signal.aborted || !currentSession(current, captured, capturedGeneration)) throw new Error('透视图数据源读取已过期。')
    const sheet = requireWorkbook(studio, unitId).getSheetBySheetId(host)
    if (!sheet) throw new Error('放置透视图的工作表已被删除。')
    const pairedTable = pivotChartCreateOptions(creation, workbook.value!, host, captured.userId)?.pivotTable
    const data = await sheet.getChartSettingsData(pivotChartPreviewConfig(config, fields, capturedSource, host, pairedTable), options)
    if (options.signal.aborted || !currentSession(current, captured, capturedGeneration)) throw new Error('透视图数据源读取已过期。')
    return data
  }
})
const loadStatistics = computed<ChartSettingsStatisticsLoader>(() => {
  const current = draft.value, captured = owner.value, capturedSource = source.value, capturedGeneration = generation
  const fields = cloneModel(current.fieldsConfig), studio = props.studio, unitId = props.workbookId, host = hostSheetId.value
  const creation = creationSnapshot(current)
  return async (config, options) => {
    if (!captured || !capturedSource || options.signal.aborted || !currentSession(current, captured, capturedGeneration)) throw new Error('透视图统计读取已过期。')
    const sheet = requireWorkbook(studio, unitId).getSheetBySheetId(host)
    if (!sheet) throw new Error('放置透视图的工作表已被删除。')
    const pairedTable = pivotChartCreateOptions(creation, workbook.value!, host, captured.userId)?.pivotTable
    const data = await sheet.getChartSettingsStatistics(pivotChartPreviewConfig(config, fields, capturedSource, host, pairedTable), options)
    if (options.signal.aborted || !currentSession(current, captured, capturedGeneration)) throw new Error('透视图统计读取已过期。')
    return data
  }
})
const { projection: preview, loading: sourceLoading, assertReady, reload } = useChartSettingsSource({
  active: () => props.open && sourceReady.value && !stale.value, draft: () => draft.value.settings,
  hostSheetId: () => hostSheetId.value, date1904: () => workbook.value?.date1904,
  sourceRevision: () => sourceRevision.value, load: () => loadSource.value, loadStatistics: () => loadStatistics.value,
})

function openInput(): PivotChartPanelOpenInput {
  return { chartId: props.chartId, pivotTableId: props.pivotTableId, sourceKind: props.sourceKind,
    sourceSheetId: props.sourceSheetId, sourceRange: props.sourceRange }
}
function invalidate(): void {
  generation++; filterGeneration++; busyGeneration++; busy.value = false
  source.value = undefined; sourceKeyRead.value = ''; owner.value = undefined
  filterOpen.value = false; filterLoading.value = false; filterError.value = ''
  showAsValidity.value = {}
}
function load(input = openInput()): void {
  if (composing.value) { pendingLoad = true; return }
  invalidate(); modelRevision.value++; error.value = ''; status.value = ''
  try {
    const book = workbook.value
    if (!book) throw new Error('工作簿已经关闭。')
    const captured = capturePivotChartPanelOwner(props.studio, props.workbookId, props.sheetId)
    const current = input.chartId ? book.charts?.find(item => item.id === input.chartId) : undefined
    draft.value = createPivotChartPanelDraft(book, props.sheetId, input, chartPlacementGeometry(book, current?.sheetId ?? props.sheetId))
    owner.value = captured
    emit('owner-change', captured)
    readFields()
    locateInitialField()
  } catch (reason) { error.value = message(reason) }
}
function readFields(): void {
  if (busy.value || composing.value || !owner.value) return
  error.value = ''; status.value = ''
  try {
    pivotChartPanelOwnerContext(owner.value)
    const book = workbook.value
    if (!book) throw new Error('工作簿已经关闭。')
    const next = readPivotChartPanelSource(book, draft.value, props.sheetId, props.studio.currentUserId)
    if (sourceKeyRead.value && sourceKeyRead.value !== sourceKey.value) resetPivotChartDraftMapping(draft.value)
    const existing = draft.value.chartId ? requireWorkbook().getSheetBySheetId(hostSheetId.value)?.getPivotChart(draft.value.chartId) : undefined
    draft.value.fieldsConfig = existing ? { ...existing.getFieldsConfig(), options: existing.getReportOptions() } : pivotChartFieldsFromReport(next.report)
    draft.value.initialFieldsConfig = cloneModel(draft.value.fieldsConfig)
    if (existing) next.fields = existing.getSourceFieldsInfo()
    source.value = next; sourceKeyRead.value = sourceKey.value
  } catch (reason) { source.value = undefined; sourceKeyRead.value = ''; error.value = message(reason) }
}
function chooseChart(value: unknown): void {
  if (typeof value !== 'string' || busy.value || composing.value) return
  if (value && !charts.value.some(item => item.id === value)) { error.value = '所选透视图已被删除。'; return }
  load({ ...pivotChartPanelInputFromDraft(draft.value), chartId: value || undefined })
}
function chooseSourceKind(value: unknown): void {
  if ((value !== 'range' && value !== 'pivotTable') || draft.value.chartId || busy.value || composing.value || stale.value) return
  draft.value.sourceKind = value
  if (value === 'pivotTable') draft.value.createMode = 'chartOnly'
  if (value === 'pivotTable' && !draft.value.pivotTableId) draft.value.pivotTableId = pivotTables.value[0]?.id ?? ''
  resetPivotChartDraftMapping(draft.value)
  source.value = undefined; sourceKeyRead.value = ''
  readFields()
}
function choosePivot(value: unknown): void {
  if (typeof value !== 'string' || draft.value.chartId || busy.value || composing.value || stale.value) return
  draft.value.pivotTableId = value
  resetPivotChartDraftMapping(draft.value)
  source.value = undefined; sourceKeyRead.value = ''
  readFields()
}
function chooseCreateMode(value: unknown): void {
  if ((value !== 'chartOnly' && value !== 'withPivotTable') || draft.value.chartId || draft.value.sourceKind !== 'range'
    || busy.value || composing.value || stale.value) return
  draft.value.createMode = value; error.value = ''; status.value = ''
}
function editFields(edit: () => PivotChartFieldsConfig, composedText = false): void {
  if (busy.value || composing.value && !composedText || !composing.value && (stale.value || !canConfigure.value || !sourceReady.value)) return
  try { draft.value.fieldsConfig = edit(); error.value = ''; status.value = '' }
  catch (reason) { error.value = message(reason) }
}
function addField(id: string, area: PivotFieldArea): void {
  const info = source.value?.fields.find(item => item.fieldId === id)
  if (!info) { error.value = '所选来源字段已失效。'; return }
  editFields(() => addPivotChartDraftField(draft.value.fieldsConfig, info, area))
}
function moveField(id: string, area: PivotFieldArea, index: number): void { editFields(() => movePivotChartDraftField(draft.value.fieldsConfig, id, area, index)) }
function removeField(id: string): void { editFields(() => removePivotChartDraftFields(draft.value.fieldsConfig, [id])) }
function setField(id: string, patch: PivotChartDraftFieldSetting): void {
  const textOnly = Object.keys(patch).every(key => key === 'displayName' || key === 'format')
  editFields(() => setPivotChartDraftField(draft.value.fieldsConfig, id, patch), textOnly)
}
function readShowAsBaseItems(id: string): string[] {
  const current = draft.value, captured = owner.value, currentGeneration = generation, report = showAsReport.value
  if (!captured || !report || !sourceReady.value || !currentSession(current, captured, currentGeneration)) throw new Error('基准成员读取已过期，请重新读取透视图。')
  const items = getPivotShowDataAsBaseItems(workbook.value!, report, id)
  if (!sourceReady.value || !currentSession(current, captured, currentGeneration)) throw new Error('透视图基准成员读取已过期。')
  return items
}
function validateShowAsDraft(): void {
  const report = showAsReport.value
  if (!report) throw new Error('请重新读取透视图报告字段。')
  for (const field of report.fields.filter(field => field.area === 'value' && field.showDataAs)) {
    const rule = field.showDataAs!
    preparePivotShowDataAsDraft(report, rule, rule.baseItem ? readShowAsBaseItems(rule.baseFieldId) : undefined)
  }
}
function setOptions(options: Partial<PivotTableOptions>): void {
  editFields(() => {
    const current = draft.value.fieldsConfig, next = { ...cloneModel(current), options: { ...current.options, ...options } }
    if (options.valuePosition === undefined && options.valueIndex === undefined) return next
    const state = pivotChartValuesUiOptions(current)
    return setPivotChartValuesUiDraft(next, options.valuePosition ?? state.valuePosition, options.valueIndex ?? state.valueIndex ?? -1)
  })
}
function setControl(key: keyof PivotChartControls, value: unknown): void {
  if (navigationDisabled.value || typeof value !== 'boolean') return
  try { setPivotChartDraftControls(draft.value, { [key]: value }); error.value = ''; status.value = '' }
  catch (reason) { error.value = message(reason) }
}
const presentationControls = [
  { key: 'showFieldButtons', label: '显示字段按钮' }, { key: 'showExpandCollapseButtons', label: '显示展开折叠按钮' },
] as const
const tabItems = computed(() => [
  { key: 'fields', label: '字段', forceRender: true, content: h(PivotChartFieldsForm, { config: sourceReady.value || composing.value ? draft.value.fieldsConfig : { fields: [] },
    sourceFields: sourceReady.value ? source.value?.fields ?? [] : [], disabled: inputDisabled.value, navigationDisabled: navigationDisabled.value,
    showDataAsReport: showAsReport.value, loadShowDataAsBaseItems: readShowAsBaseItems, showDataAsReadable: sourceReady.value && !stale.value,
    showDataAsContextKey: JSON.stringify([generation, owner.value?.revision, owner.value?.historyGeneration, props.studio.currentUserId]),
    onShowAsValidity: (id: string, valid: boolean) => { showAsValidity.value[id] = valid },
    onAdd: addField, onMove: moveField, onRemove: removeField, onSetting: setField, onFilter: openFilter, onOptions: setOptions }) },
  { key: 'appearance', label: '图表设置', forceRender: true, content: h('div', { class: 'pivot-chart-appearance' }, [h(ChartSettingsFields, { draft: draft.value.settings,
    'onUpdate:draft': (value: ChartSettingsDraft) => { if (!busy.value && (!stale.value || composing.value)) draft.value.settings = value },
    projection: preview.value, date1904: workbook.value?.date1904, hostSheetId: hostSheetId.value, lockedSource: true, disabled: inputDisabled.value }),
    h('div', { class: 'pivot-chart-controls', 'data-testid': 'pivot-chart-presentation-controls' }, presentationControls.map(control =>
      h(FormItem, { label: control.label }, () => h(Switch, { checked: draft.value.controls[control.key], 'aria-label': control.label,
        disabled: navigationDisabled.value, onChange: (value: unknown) => setControl(control.key, value) })))),
  ]) },
])

async function openFilter(id: string): Promise<void> {
  if (navigationDisabled.value || !owner.value || !source.value) return
  const current = draft.value, captured = owner.value, capturedSource = source.value, capturedGeneration = generation, request = ++filterGeneration
  const field = current.fieldsConfig.fields.find(item => item.id === id)
  if (!field || field.area === 'value') return
  filterId.value = id; filterOpen.value = true; filterLoading.value = true; filterError.value = ''
  filterItems.value = []; filterAll.value = field.selectedItems === undefined; filterKeys.value = [...(field.selectedItems ?? [])]
  const sameRequest = (): boolean => !disposed && props.open && filterOpen.value && request === filterGeneration
  const active = (): boolean => sameRequest() && currentSession(current, captured, capturedGeneration)
  try {
    const target = current.chartId ? requireWorkbook().getSheetBySheetId(hostSheetId.value)?.getPivotChart(current.chartId) : undefined
    const items = target ? await target.getFieldItems(id) : readPivotChartDraftFieldItems(workbook.value!, current, capturedSource, hostSheetId.value, id)
    if (!active()) { if (sameRequest()) filterError.value = '标签读取已过期，当前筛选草稿已保留。请关闭后重新读取。'; return }
    filterItems.value = [...new Set([...items, ...(field.selectedItems ?? [])])]
    filterAll.value = field.selectedItems === undefined
    filterKeys.value = [...(field.selectedItems ?? filterItems.value)]
  } catch (reason) { if (sameRequest()) filterError.value = message(reason) }
  finally { if (sameRequest()) filterLoading.value = false }
}
function saveFilter(): void {
  if (!filterOpen.value || navigationDisabled.value || filterLoading.value) return
  editFields(() => setPivotChartDraftField(draft.value.fieldsConfig, filterId.value, { selectedItems: filterAll.value ? undefined : filterKeys.value }))
  closeFilter()
}
function closeFilter(): void { if (!composing.value) { filterGeneration++; filterOpen.value = false; filterLoading.value = false } }
function setFilterKeys(value: unknown): void {
  if (!Array.isArray(value) || value.some(item => typeof item !== 'string')) return
  filterKeys.value = value; filterAll.value = false
}

async function mutate(kind: 'apply' | 'refresh' | 'remove'): Promise<void> {
  const current = draft.value, captured = owner.value, capturedGeneration = generation
  if (!props.open || busy.value || composing.value || stale.value || !canConfigure.value || !sourceReady.value || !captured || createOptionsError.value) return
  const request = ++busyGeneration
  busy.value = true; error.value = ''; status.value = ''
  try {
    const context = pivotChartPanelOwnerContext(captured), book = requireWorkbook(captured.studio, captured.unitId)
    const targetSheet = book.getSheetBySheetId(hostSheetId.value)
    if (!targetSheet) throw new Error('放置透视图的工作表已被删除。')
    let chartId = current.chartId, accepted: unknown
    if (kind === 'apply') {
      validateShowAsDraft()
      assertReady()
      const geometry = chartPlacementGeometry(workbook.value, hostSheetId.value)
      if (chartId) {
        const target = targetSheet.getPivotChart(chartId)
        if (!target) throw new Error('所选透视图已被删除。')
        accepted = await target.configure(pivotChartConfigureInput(current, geometry), context)
      } else {
        const options = pivotChartCreateOptions(current, workbook.value!, hostSheetId.value, captured.userId)
        const created = await targetSheet.insertPivotChart(pivotChartCreateInfo(current, workbook.value!, hostSheetId.value, geometry, captured.userId), context, options)
        chartId = created.getChartId()
      }
    } else {
      const target = chartId ? targetSheet.getPivotChart(chartId) : undefined
      if (!target) throw new Error('请先选择已有透视图。')
      accepted = kind === 'remove' ? await target.remove(context) : await target.refresh(context)
    }
    if (accepted === false) throw new Error(kind === 'remove' ? '透视图未能移除。' : '透视图命令被拒绝，内容没有更新。')
    if (!currentSession(current, captured, capturedGeneration, true)) return
    const changed = captured.studio.getWorkbookRevision(captured.unitId) !== captured.revision
    load({ ...pivotChartPanelInputFromDraft(current), chartId: kind === 'remove' ? undefined : chartId })
    statusType.value = changed ? 'success' : 'info'
    status.value = changed ? kind === 'remove' ? current.sourceKind === 'pivotTable' ? '透视图已移除，关联透视表和来源数据保留。' : '透视图已移除，来源数据保留。' : kind === 'refresh' ? '透视图已刷新。'
      : current.chartId ? '透视图设置已更新。' : current.createMode === 'withPivotTable' ? '数据透视表和透视图已一起创建。' : '已创建透视图。'
      : kind === 'refresh' ? '已重新读取数据源，内容设置没有变化。' : '设置没有变化。'
    if (changed) emit('applied', chartId)
  } catch (reason) { if (currentSession(current, captured, capturedGeneration, true)) error.value = message(reason) }
  finally { if (!disposed && request === busyGeneration) busy.value = false }
}
function startComposition(): void { compositionGeneration++; composing.value = true }
function endComposition(): void {
  const ended = compositionGeneration
  // Capture runs before Ant Input publishes its final composition text.
  void nextTick(() => {
    if (disposed || ended !== compositionGeneration) return
    composing.value = false
    if (pendingLoad) { pendingLoad = false; if (props.open) load() }
    if (!props.open) invalidate()
  })
}
function close(): void { if (!busy.value && !composing.value) { closeFilter(); emit('update:open', false) } }
function locateInitialField(): void {
  const id = props.initialFieldId, current = draft.value, captured = owner.value, request = generation
  if (!id || !captured || !sourceReady.value || busy.value || composing.value) return
  activeTab.value = 'fields'
  void nextTick(() => {
    if (!currentSession(current, captured, request) || !sourceReady.value || busy.value || composing.value || props.initialFieldId !== id) return
    const card = Array.from(panelElement.value?.querySelectorAll<HTMLElement>('[data-pivot-field]') ?? [])
      .find(element => element.getAttribute('data-pivot-field') === id)
    if (!card) { error.value = '所选透视图字段已被移除，请选择当前字段。'; return }
    card.scrollIntoView({ block: 'nearest', inline: 'nearest' })
    card.querySelector<HTMLInputElement>('input[type="text"], input:not([type])')?.focus({ preventScroll: true })
  })
}
function opened(value: boolean): void { if (value) locateInitialField(); else if (!props.open) emit('closed') }
function message(reason: unknown): string { return reason instanceof Error ? reason.message : String(reason) }

watch([() => props.open, () => props.chartId, () => props.pivotTableId, () => props.sourceKind, () => props.sourceSheetId,
  () => props.sourceRange, () => props.workbookId, () => props.sheetId, () => props.studio], ([open]) => {
  if (open) load()
  else if (!composing.value) invalidate()
}, { immediate: true })
watch(busy, value => emit('busy-change', value), { immediate: true, flush: 'sync' })
watch(() => props.initialFieldId, () => { if (props.open) locateInitialField() })
watch([() => props.studio, () => props.workbookId], ([studio, unitId], _previous, onCleanup) => {
  const changed = (id: string): void => { if (!disposed && id === unitId) modelRevision.value++ }
  const listeners = [
    studio.events.addEvent('CommandExecuted', event => { if (event.unitId) changed(event.unitId) }),
    studio.events.addEvent('Undo', event => { if (event.kind === 'workbook') changed(event.unitId) }),
    studio.events.addEvent('Redo', event => { if (event.kind === 'workbook') changed(event.unitId) }),
    studio.events.addEvent('PivotTableChanged', event => changed(event.unitId)),
    studio.events.addEvent('PivotTableRendered', event => changed(event.unitId)),
    studio.events.addEvent('SheetValueChanged', event => changed(event.workbookId)),
    studio.events.addEvent('SheetStructureChanged', event => changed(event.workbookId)),
    studio.events.addEvent('WorkbookDisposed', event => { if (event.workbookId === unitId) { changed(unitId); invalidate(); emit('update:open', false) } }),
  ]
  onCleanup(() => listeners.forEach(listener => listener.dispose()))
}, { immediate: true })
onScopeDispose(() => { disposed = true; invalidate() })
</script>

<template>
  <Modal :open="open" :title="t('sheet.pivot-chart.title', '透视图')" :width="860" :footer="null" :root-style="panelStyle"
    :mask="{ closable: !busy && !composing }" :keyboard="!busy && !composing && !filterOpen" :closable="!busy && !composing"
    :focus-trigger-after-close="false" @cancel="close" @after-open-change="opened">
    <div ref="panelElement">
    <Form layout="vertical" class="pivot-chart-panel" :style="panelStyle" data-testid="pivot-chart-panel" @submit.prevent="mutate('apply')"
      @compositionstart.capture="startComposition" @compositionend.capture="endComposition">
      <Alert v-if="stale" type="warning" title="内容、用户或历史已更改。当前草稿已保留，请重新读取后再应用。" show-icon>
        <template #action><Button size="small" :disabled="busy || composing" @click="load(pivotChartPanelInputFromDraft(draft))">重新读取</Button></template>
      </Alert>
      <FormItem v-if="charts.length" label="选择透视图"><Select :value="draft.chartId" :options="chartOptions" aria-label="选择透视图" :disabled="busy || composing" @change="chooseChart" /></FormItem>
      <FormItem label="来源方式"><Select :value="draft.sourceKind" :options="[{ label: '区域', value: 'range' }, { label: '关联已有透视表', value: 'pivotTable' }]"
        aria-label="透视图来源方式" :disabled="busy || composing || stale || !!draft.chartId" @change="chooseSourceKind" /></FormItem>
      <template v-if="draft.sourceKind === 'range'">
        <FormItem label="来源工作表"><Select v-model:value="draft.sourceSheetId" :options="sheets.map(sheet => ({ label: sheet.name, value: sheet.id }))" aria-label="透视图来源工作表" :disabled="busy || composing || stale || !!draft.chartId" /></FormItem>
        <FormItem label="来源区域"><Input v-model:value="draft.sourceRange" aria-label="透视图来源区域" placeholder="A1:C20" :disabled="busy || !!draft.chartId || stale && !composing" /></FormItem>
      </template>
      <FormItem v-else label="关联透视表"><Select :value="draft.pivotTableId" :options="pivotOptions" aria-label="关联透视表" :disabled="busy || composing || stale || !!draft.chartId" @change="choosePivot" /></FormItem>
      <template v-if="!draft.chartId && draft.sourceKind === 'range'">
        <FormItem label="创建方式"><Select :value="draft.createMode" :options="PIVOT_CHART_CREATION_MODES" aria-label="透视图创建方式"
          :disabled="busy || composing || stale" @change="chooseCreateMode" /></FormItem>
        <template v-if="draft.createMode === 'withPivotTable'">
          <FormItem label="透视表放置工作表"><Select v-model:value="draft.pivotTableSheetId" :options="tableTargetSheets.map(sheet => ({ label: sheet.name, value: sheet.id }))"
            aria-label="配对透视表放置工作表" :disabled="busy || composing || stale" /></FormItem>
          <FormItem label="透视表起始单元格"><Input v-model:value="draft.pivotTableAddress" aria-label="配对透视表起始单元格" placeholder="例如 A1" :disabled="inputDisabled" /></FormItem>
          <Alert v-if="createOptionsError" type="warning" :title="createOptionsError" show-icon />
        </template>
      </template>
      <Button v-if="!sourceReady" :disabled="busy || composing || stale" @click="readFields">读取来源字段</Button>
      <Alert v-if="sourceReady && !canConfigure" type="info" title="当前用户没有修改透视图或关联报告的权限。" show-icon />
      <FormItem label="放置单元格"><Input v-model:value="draft.position" aria-label="透视图放置单元格" placeholder="F2" :disabled="inputDisabled" /></FormItem>
      <Tabs :active-key="activeTab" :items="tabItems" :destroy-on-hidden="false" @change="value => { if (!composing && !busy) activeTab = value }" />
      <Alert v-if="sourceLoading" type="info" title="正在读取透视图字段草稿的数据源…" show-icon role="status" />
      <Button v-if="preview.error" :disabled="busy || composing || sourceLoading || stale" @click="reload">重新读取数据源</Button>
      <Alert v-if="error || preview.error" type="error" :title="error || preview.error" show-icon />
      <Alert v-if="status" :type="statusType" :title="status" show-icon role="status" />
      <footer>
        <Button v-if="draft.chartId" danger :disabled="navigationDisabled" @click="mutate('remove')">移除透视图</Button>
        <Button v-if="draft.chartId" :disabled="navigationDisabled" @click="mutate('refresh')">刷新透视图</Button>
        <span />
        <Button :disabled="busy || composing" @click="close">关闭</Button>
        <Button type="primary" :loading="busy" :disabled="navigationDisabled || sourceLoading || showAsInvalid || !!createOptionsError" @click="mutate('apply')">{{ draft.chartId ? '应用透视图' : draft.createMode === 'withPivotTable' ? '同时创建透视表和透视图' : '创建透视图' }}</Button>
      </footer>
    </Form>
    </div>
  </Modal>
  <Modal :open="open && filterOpen" title="标签筛选" :width="480" :footer="null" :root-style="panelStyle" :mask="{ closable: false }"
    :keyboard="!composing" :focus-trigger-after-close="true" @cancel="closeFilter">
    <Form layout="vertical" class="pivot-chart-filter" :style="panelStyle" @submit.prevent="saveFilter" @compositionstart.capture="startComposition" @compositionend.capture="endComposition">
      <FormItem :label="null"><Checkbox v-model:checked="filterAll" :disabled="filterLoading || navigationDisabled">显示全部标签</Checkbox></FormItem>
      <FormItem label="选择标签"><Select :value="sourceReady ? filterKeys : []" :options="(sourceReady ? filterItems : []).map(item => ({ value: item, label: item || '（空白）' }))" mode="multiple" show-search
        aria-label="透视图筛选标签" :disabled="filterLoading || navigationDisabled" @change="setFilterKeys" /></FormItem>
      <Alert v-if="filterLoading" type="info" title="正在读取真实来源标签…" show-icon />
      <Alert v-if="filterError" type="error" :title="filterError" show-icon />
      <Button :disabled="composing" @click="closeFilter">取消</Button>
      <Button type="primary" :disabled="filterLoading || !!filterError || navigationDisabled" @click="saveFilter">应用标签筛选</Button>
    </Form>
  </Modal>
</template>

<style scoped>
.pivot-chart-panel { display: grid; gap: 14px; max-height: 72vh; min-width: 0; overflow-y: auto; padding-right: 8px; }
.pivot-chart-panel :deep(.ant-form-item), .pivot-chart-filter :deep(.ant-form-item) { min-width: 0; margin-bottom: 0; }
.pivot-chart-panel footer { display: flex; align-items: center; flex-wrap: wrap; gap: 8px; padding-top: 12px; border-top: 1px solid var(--panel-border); }
.pivot-chart-panel footer span { flex: 1; }
.pivot-chart-filter { display: grid; gap: 12px; min-width: 0; }
.pivot-chart-appearance, .pivot-chart-controls { display: grid; gap: 14px; min-width: 0; }
.pivot-chart-controls { padding-top: 12px; border-top: 1px solid var(--panel-border); }
</style>
