<script setup lang="ts">
import { Alert, Button, Checkbox, Dropdown, Form, FormItem, Input, Modal, RadioGroup, Select, Switch, theme } from 'antdv-next'
import { computed, nextTick, onScopeDispose, ref, shallowRef, watch } from 'vue'
import { ChevronDown, Ellipsis, GripVertical } from 'lucide-vue-next'
import { FOffice } from '../../api/facade'
import type { OfficeStudio } from '../../api/office-studio'
import type { CellRange, PivotAggregation, PivotShowDataAsInfo, PivotValuePosition } from '../../domain/model'
import {
  CONFIGURE_PIVOT_COMMAND, createPivotPanelDraft, pivotRangeAddress, preparePivotPanelInput,
  readPivotPanelSourceFields, pivotPanelDraftReport, type PivotPanelDraft, type PivotPanelField,
} from '../../features/pivot/services/pivot-panel-state'
import PivotFieldToolsDialog from '../../features/pivot/views/PivotFieldToolsDialog.vue'
import { useOfficeRegion } from '../../i18n/office-i18n'
import type { PivotFieldToolChange } from '../../features/pivot/services/pivot-field-tools'
import { formatSheetValue, SHEET_NUMBER_FORMATS } from '../../features/sheet-workbench/sheet-number-format'
import { capturePivotChartPanelOwner, pivotChartPanelOwnerContext, pivotChartPanelOwnerCurrent,
  type PivotChartPanelOwner } from '../../features/pivot/services/pivot-chart-panel-owner.ts'
import { assertPivotReportEditPermission, assertPivotReportReadPermission } from '../../features/pivot/services/pivot-report-permissions.ts'
import { getPivotShowDataAsBaseItems } from '../../features/pivot/services/pivot-engine.ts'
import { createPivotShowDataAsDraft, preparePivotShowDataAsDraft } from '../../features/pivot/services/pivot-show-data-as-ui.ts'
import PivotShowDataAsFields from '../../features/pivot/views/PivotShowDataAsFields.vue'
import PivotChartFieldFilterPopup from '../../features/pivot/views/PivotChartFieldFilterPopup.vue'
import { usePivotTableFieldFilter } from '../../features/pivot/views/use-pivot-table-field-filter.ts'
import { pivotTableFieldFilterDraftCompatible, setPivotTablePanelMultipleFilters, synchronizePivotTableFieldFilterDraft,
  type PivotTableFieldFilterAction } from '../../features/pivot/services/pivot-table-field-filter-state.ts'
import type { PivotChartFieldButtonAction } from '../../features/pivot/services/pivot-chart-field-button-state.ts'
import { adjustPivotTableValuesUiFieldEdit, pivotTableValuesUiFieldOrder, pivotValuesUiEntries, readPivotValuesUiState,
  setPivotTableValuesUiDraft } from '../../features/pivot/services/pivot-values-position-ui.ts'
import PivotValuesFieldCard from '../../features/pivot/views/PivotValuesFieldCard.vue'
import PivotValuesPositionFields from '../../features/pivot/views/PivotValuesPositionFields.vue'

const props = defineProps<{ open: boolean; studio?: OfficeStudio; workbookId?: string; sheetId: string; sourceRange: string }>()
interface PivotPanelChange { pivotId: string; sheetId: string; outputRange?: CellRange }
const emit = defineEmits<{
  'update:open': [open: boolean]
  created: [result: PivotPanelChange]
  updated: [result: PivotPanelChange]
}>()
const modelRevision = ref(0), owner = shallowRef<PivotChartPanelOwner>(), composing = ref(false)
const workbook = computed(() => { modelRevision.value; return props.studio?.getWorkbookModel(props.workbookId ?? props.studio.activeWorkbookId) })
const draft = ref<PivotPanelDraft | null>(null), selectedId = ref(''), error = ref(''), status = ref(''), busy = ref(false)
const loadedSource = ref(''), fieldSearch = ref('')
const draggedFieldId = ref('')
const draggedValues = shallowRef<{ draft: PivotPanelDraft; owner: PivotChartPanelOwner; generation: number; signature: string }>()
const settingsFieldId = ref(''), panelElement = ref<HTMLElement>()
const fieldDisplayName = ref(''), fieldAggregation = ref<PivotAggregation>('sum'), fieldFormat = ref('')
const fieldDialogBusy = ref(false)
const settingsBusy = ref(false), settingsMode = ref<'all' | 'showDataAs'>('all'), fieldShowDataAs = ref<PivotShowDataAsInfo | null>(null), fieldShowAsValid = ref(true)
interface FieldSettingsSession { draft: PivotPanelDraft; owner: PivotChartPanelOwner; pivotId: string; fieldId: string; signature: string; generation: number }
const settingsSession = shallowRef<FieldSettingsSession>()
let disposed = false, panelGeneration = 0, settingsGeneration = 0, settingsRequest = 0, fieldActionRequest = 0, compositionGeneration = 0, pendingLoad: string | undefined
const { token } = theme.useToken()
const panelStyle = computed(() => ({ color: token.value.colorText, backgroundColor: token.value.colorBgElevated,
  '--pivot-border': token.value.colorBorder, '--pivot-surface': token.value.colorFillAlter, '--pivot-text': token.value.colorText,
  '--pivot-secondary': token.value.colorTextSecondary, '--pivot-muted': token.value.colorTextTertiary, '--pivot-hover': token.value.colorFillSecondary,
  '--pivot-error': token.value.colorError, '--pivot-error-bg': token.value.colorErrorBg, '--pivot-success': token.value.colorSuccess }))
const panelStale = computed(() => { modelRevision.value; return Boolean(owner.value && !pivotChartPanelOwnerCurrent(owner.value)) })
const settingsDialogBusy = computed(() => fieldDialogBusy.value || settingsBusy.value || filterBusy.value)
const panelNavigationDisabled = computed(() => busy.value || fieldDialogBusy.value || settingsBusy.value || composing.value || filterOpen.value || filterComposing.value)
const toolsFieldId = ref(''), toolsKind = ref<'filter' | 'group'>('filter')
const numberFormatOpen = ref(false), numberFormatDraft = ref('')
const toolsField = computed(() => selectedPivot.value?.fields.find(field => field.id === toolsFieldId.value))
const toolsDateSupported = computed(() => {
  if (!toolsField.value || !settingsReadable.value) return false
  try { return requireApi().book.getPivotTables().find(pivot => pivot.getPivotTableId() === selectedId.value)?.getFilterCapabilities(toolsField.value.id).supportsDateFilter === true }
  catch { return false }
})
const numberFormatOptions = Object.entries(SHEET_NUMBER_FORMATS).map(([value, pattern]) => ({ value: pattern, label: ({ general: '常规', number: '数值', currency: '货币', accounting: '会计', percent: '百分比', date: '日期', time: '时间', datetime: '日期时间', text: '文本' } as Record<string, string>)[value] ?? value }))
const region = useOfficeRegion()
const numberFormatPreview = computed(() => formatSheetValue(1234.56, numberFormatDraft.value, workbook.value?.date1904, region.value) ?? '此自定义格式将按原样保存在值字段中')
const aggregationOptions: Array<{ label: string; value: PivotAggregation }> = [
  { label: '平均值', value: 'average' }, { label: '计数', value: 'count' }, { label: '数值计数', value: 'countNums' },
  { label: '最大值', value: 'max' }, { label: '最小值', value: 'min' }, { label: '乘积', value: 'product' },
  { label: '标准偏差', value: 'stdDev' }, { label: '总体标准偏差', value: 'stdDevp' }, { label: '求和', value: 'sum' },
  { label: '方差', value: 'var' }, { label: '总体方差', value: 'varp' },
]
const sheetOptions = computed(() => (workbook.value?.sheets ?? []).filter(sheet => sheet.visible !== false).map(sheet => ({ label: sheet.name, value: sheet.id })))
const pivotOptions = computed(() => [{ label: '新建数据透视表', value: '' }, ...(workbook.value?.pivotTables ?? []).map((pivot, index) => ({ label: `数据透视表 ${index + 1} · ${workbook.value?.sheets.find(sheet => sheet.id === pivot.anchor.subUnitId)?.name ?? ''}!${pivotRangeAddress({ startRow: pivot.anchor.row, endRow: pivot.anchor.row, startCol: pivot.anchor.col, endCol: pivot.anchor.col })}`, value: pivot.id }))])
const selectedPivot = computed(() => workbook.value?.pivotTables?.find(pivot => pivot.id === selectedId.value))
const sourceKey = computed(() => draft.value ? `${draft.value.sourceSheetId}\n${draft.value.sourceRange.trim().toUpperCase()}` : '')
const fieldsStale = computed(() => sourceKey.value !== loadedSource.value)
const filteredFields = computed(() => (draft.value?.fields ?? []).filter(field => field.fieldName.toLocaleLowerCase().includes(fieldSearch.value.toLocaleLowerCase())))
const visibleFields = computed(() => filteredFields.value)
const valuesState = computed(() => readPivotValuesUiState(draft.value?.options ?? { valuePosition: 'column' }, draft.value ? pivotTableValuesUiFieldOrder(draft.value) : []))
const fieldAreas = computed(() => [
  { area: 'filter' as const, label: '筛选', fields: draft.value?.fields.filter(field => field.area === 'filter') ?? [] },
  { area: 'column' as const, label: '列', fields: draft.value?.fields.filter(field => field.area === 'column') ?? [] },
  { area: 'row' as const, label: '行', fields: draft.value?.fields.filter(field => field.area === 'row') ?? [] },
  { area: 'value' as const, label: '值', fields: draft.value?.fields.filter(field => field.area === 'value') ?? [] },
].map(area => ({ ...area, entries: pivotValuesUiEntries(area.fields, valuesState.value, area.area) })))
const settingsField = computed(() => draft.value?.fields.find(field => field.fieldId === settingsFieldId.value))
const history = computed(() => props.studio?.editor.history.units[workbook.value?.id ?? ''])
const settingsReport = computed(() => {
  if (!workbook.value || !draft.value || fieldsStale.value) return undefined
  try { return pivotPanelDraftReport(workbook.value, draft.value) } catch { return undefined }
})
const settingsReadable = computed(() => {
  if (!workbook.value || !settingsReport.value) return false
  try {
    if (selectedPivot.value) assertPivotReportReadPermission(workbook.value, selectedPivot.value, props.studio?.currentUserId ?? '')
    assertPivotReportReadPermission(workbook.value, settingsReport.value, props.studio?.currentUserId ?? '')
    return true
  } catch { return false }
})
const settingsEditable = computed(() => {
  if (!settingsReadable.value || !workbook.value || !settingsReport.value) return false
  try {
    if (selectedPivot.value) assertPivotReportEditPermission(workbook.value, selectedPivot.value, props.studio?.currentUserId ?? '')
    assertPivotReportEditPermission(workbook.value, settingsReport.value, props.studio?.currentUserId ?? '')
    return true
  } catch { return false }
})
const directFieldSetting = computed(() => {
  const report = settingsReport.value, committed = selectedPivot.value
  if (!report || !committed || fieldsStale.value || JSON.stringify(report.source) !== JSON.stringify(committed.source)
    || report.anchor.subUnitId !== committed.anchor.subUnitId || report.anchor.row !== committed.anchor.row || report.anchor.col !== committed.anchor.col) return false
  const fields = (items: typeof report.fields) => items.map(field => [field.id, field.sourceColumn, field.area, field.position])
  return JSON.stringify(fields(report.fields)) === JSON.stringify(fields(committed.fields))
})
const settingsHostError = computed(() => directFieldSetting.value && selectedPivot.value?.anchor.subUnitId !== props.sheetId
  ? '请切换到透视表所在工作表后设置值字段。' : '')
const settingsDisabled = computed(() => settingsDialogBusy.value || composing.value || panelStale.value || !settingsEditable.value || Boolean(settingsHostError.value))

const filters = usePivotTableFieldFilter({
  get studio() { return props.studio }, get unitId() { return workbook.value?.id }, get sheetId() { return props.sheetId },
  get pivotId() { return selectedId.value }, get panelDraft() { return draft.value }, get owner() { return owner.value },
  get active() { return props.open }, get sourceReady() { return !fieldsStale.value },
  get busy() { return busy.value || fieldDialogBusy.value || settingsBusy.value || composing.value },
  restoreFocus: focusFilterField, applied: fieldFilterApplied,
})
const { field: filterButton, fieldId: filterFieldId, draft: filterDraft, conditions: filterConditions, isOpen: filterOpen,
  loading: filterLoading, busy: filterBusy, composing: filterComposing, stale: filterStale, error: filterError,
  readable: filterReadable, navigationDisabled: filterNavigationDisabled, editDisabled: filterEditDisabled } = filters

function focusFilterField(fieldId: string): void {
  if (!props.open || disposed || composing.value || filterComposing.value || panelStale.value) return
  const card = Array.from(panelElement.value?.querySelectorAll<HTMLElement>('.pivot-field[data-field-id]') ?? [])
    .find(card => card.getAttribute('data-field-id') === fieldId)
  card?.querySelector<HTMLButtonElement>('button:not([disabled])')?.focus({ preventScroll: true })
}
function fieldFilterApplied(fieldId: string, action: PivotTableFieldFilterAction, changed: boolean): void {
  const current = draft.value, report = selectedPivot.value, captured = owner.value
  if (!current || !report || !captured || !pivotChartPanelOwnerCurrent(captured, true)) return
  synchronizePivotTableFieldFilterDraft(current, report, fieldId, action)
  owner.value = capturePivotChartPanelOwner(captured.studio, captured.unitId, captured.sheetId)
  error.value = ''; status.value = changed ? action === 'sortAscending' || action === 'sortDescending' || action === 'clearSort'
    ? '透视字段排序已应用。' : '透视字段筛选已应用。' : '设置没有变化。'
  if (changed) reportChange(report.id)
}
function filterAction(action: PivotChartFieldButtonAction): void { if (action !== 'remove') void filters.mutate(action) }
function setAllowMultipleFilters(value: unknown): void {
  if (!draft.value || typeof value !== 'boolean' || panelNavigationDisabled.value || panelStale.value || !settingsEditable.value) return
  setPivotTablePanelMultipleFilters(draft.value, value)
}
function setValuesPosition(position: PivotValuePosition, index: number): void {
  if (!draft.value || panelNavigationDisabled.value || panelStale.value || !settingsEditable.value) return
  try { setPivotTableValuesUiDraft(draft.value, position, index); error.value = ''; status.value = '' }
  catch (reason) { error.value = reason instanceof Error ? reason.message : String(reason) }
}
function editFieldDraft(fieldId: string, edit: () => void): void {
  const current = draft.value
  if (!current) return
  const previous = pivotTableValuesUiFieldOrder(current)
  edit(); adjustPivotTableValuesUiFieldEdit(current, previous, fieldId)
}
watch(filterError, value => { if (value && !filterOpen.value) error.value = value })

function settingsSignature(current: PivotPanelDraft): string { return JSON.stringify([current.pivotId, current.sourceSheetId, current.sourceRange, current.targetSheetId, current.targetAddress, current.fields]) }
function settingsCurrent(captured: FieldSettingsSession, published = false): boolean {
  return !disposed && props.open && settingsSession.value === captured && draft.value === captured.draft && selectedId.value === captured.pivotId
    && settingsFieldId.value === captured.fieldId && captured.generation === settingsGeneration
    && props.studio === captured.owner.studio && workbook.value?.id === captured.owner.unitId && props.sheetId === captured.owner.sheetId
    && settingsSignature(captured.draft) === captured.signature && pivotChartPanelOwnerCurrent(captured.owner, published)
}
function closeSettings(): void {
  if (settingsBusy.value || composing.value) return
  settingsGeneration++; settingsRequest++; settingsSession.value = undefined; settingsFieldId.value = ''; numberFormatOpen.value = false
}

function loadDraft(pivotId = '') {
  if (composing.value || filterComposing.value) { pendingLoad = pivotId; return }
  filters.close(false)
  draggedValues.value = undefined
  panelGeneration++; settingsGeneration++; settingsRequest++; fieldActionRequest++; settingsBusy.value = false; settingsSession.value = undefined; owner.value = undefined
  const model = workbook.value
  error.value = ''; status.value = ''; selectedId.value = pivotId; fieldSearch.value = ''
  settingsFieldId.value = ''; toolsFieldId.value = ''; numberFormatOpen.value = false; fieldDialogBusy.value = false
  if (!model) { draft.value = null; error.value = '当前工作台尚未连接工作簿命令接口。'; return }
  draft.value = createPivotPanelDraft(model, props.sheetId, props.sourceRange, pivotId)
  loadedSource.value = sourceKey.value
  if (props.studio) {
    try { owner.value = capturePivotChartPanelOwner(props.studio, model.id!, props.sheetId) }
    catch { error.value = '透视表面板已离开原工作簿或工作表，请重新打开。' }
  }
}
watch([() => props.open, () => props.studio, () => props.workbookId, () => props.sheetId], ([open]) => {
  if (open) loadDraft()
  else if (!composing.value && !filterComposing.value) { filters.close(false); panelGeneration++; settingsGeneration++; settingsRequest++; fieldActionRequest++;
    settingsSession.value = undefined; settingsBusy.value = false; settingsFieldId.value = ''; toolsFieldId.value = ''; numberFormatOpen.value = false; owner.value = undefined }
}, { immediate: true })
function loadFields() {
  if (!draft.value || !workbook.value || panelNavigationDisabled.value || settingsFieldId.value || panelStale.value) return
  try {
    const previousOrder = pivotTableValuesUiFieldOrder(draft.value)
    const previous = draft.value.fields
    const next = readPivotPanelSourceFields(workbook.value, draft.value.sourceSheetId, draft.value.sourceRange, selectedPivot.value)
    draft.value.fields = fieldsStale.value ? next : next.flatMap(field => {
      const retained = previous.filter(candidate => candidate.sourceColumn === field.sourceColumn)
      return retained.length ? retained.map(candidate => ({ ...field, ...candidate, fieldName: field.fieldName })) : [field]
    })
    adjustPivotTableValuesUiFieldEdit(draft.value, previousOrder)
    loadedSource.value = sourceKey.value; error.value = ''; status.value = `已读取 ${next.length} 个字段，请选择行、列和汇总值。`
  } catch (reason) { error.value = reason instanceof Error ? reason.message : String(reason) }
}
function removeField(fieldId: string) {
  if (panelNavigationDisabled.value || settingsFieldId.value || panelStale.value) return
  const fields = draft.value?.fields
  if (!fields) return
  const field = fields.find(candidate => candidate.fieldId === fieldId)
  if (field) editFieldDraft(fieldId, () => { field.area = 'unused'; field.showDataAs = null })
}
function toggleField(fieldId: string, checked: boolean) {
  if (panelNavigationDisabled.value || settingsFieldId.value || panelStale.value) return
  const field = draft.value?.fields.find(candidate => candidate.fieldId === fieldId)
  if (!field) return
  if (!checked) { editFieldDraft(fieldId, () => { field.area = 'unused'; field.showDataAs = null }); return }
  const source = workbook.value?.sheets.find(sheet => sheet.id === draft.value?.sourceSheetId)
  const rowNumbers = [...(draft.value?.sourceRange.matchAll(/\$?[A-Z]+\$?(\d+)/gi) ?? [])].map(match => Number(match[1]) - 1)
  const firstRow = rowNumbers[0] ?? 0, lastRow = rowNumbers.at(-1) ?? firstRow
  let foundNumber = false, foundOther = false
  for (let row = firstRow + 1; source && row <= lastRow; row++) {
    const value = source.cells[`${row}:${field.sourceColumn}`]?.value
    if (value == null || value === '') continue
    if (typeof value === 'number' && Number.isFinite(value)) foundNumber = true
    else foundOther = true
  }
  editFieldDraft(fieldId, () => { field.area = foundNumber && !foundOther ? 'value' : 'row' })
}
function fieldMenu(field: PivotPanelField) {
  const committed = selectedPivot.value?.fields.find(candidate => candidate.id === field.fieldId)
  const editableLabel = committed && draft.value && workbook.value && !fieldsStale.value
    && pivotTableFieldFilterDraftCompatible(workbook.value, draft.value, selectedId.value, field.fieldId)
  const items = [
    { key: 'remove', label: '移除字段', danger: true },
    ...(editableLabel ? [{ key: 'filter', label: '筛选字段', disabled: !settingsReadable.value },
      { key: 'value-filter', label: '值筛选', disabled: !settingsEditable.value }, { key: 'group', label: '分组', disabled: !settingsEditable.value }] : []),
    ...(field.area === 'value' ? [{ key: 'settings', label: '字段设置' }, { key: 'show-data-as', label: '值显示方式' }] : []),
    { key: 'up', label: '上移', disabled: draft.value?.fields.filter(candidate => candidate.area === field.area).at(0)?.fieldId === field.fieldId },
    { key: 'down', label: '下移', disabled: draft.value?.fields.filter(candidate => candidate.area === field.area).at(-1)?.fieldId === field.fieldId },
    ...(['filter', 'column', 'row', 'value'] as const).filter(area => area !== field.area).map(area => ({ key: `move-${area}`, label: `移动到${{ filter: '筛选', column: '列', row: '行', value: '值' }[area]}` })),
  ]
  return {
    items,
    onClick: ({ key }: { key: string | number }) => {
      if (panelNavigationDisabled.value || settingsFieldId.value || panelStale.value) return
      if (key === 'remove') removeField(field.fieldId)
      else if (key === 'filter') openFieldFilter(field.fieldId)
      else if (key === 'settings') openFieldSettings(field.fieldId)
      else if (key === 'show-data-as') openFieldSettings(field.fieldId, 'showDataAs')
      else if (key === 'value-filter' || key === 'group') { if (settingsEditable.value) { toolsFieldId.value = field.fieldId; toolsKind.value = key === 'value-filter' ? 'filter' : 'group'; error.value = '' } }
      else if (key === 'up' || key === 'down') moveField(field.fieldId, key === 'up' ? -1 : 1)
      else if (String(key).startsWith('move-')) {
        const area = String(key).slice(5)
        if (area === 'row' || area === 'column' || area === 'filter' || area === 'value') editFieldDraft(field.fieldId, () => { field.area = area; if (area !== 'value') field.showDataAs = null })
      }
    },
  }
}
function moveField(fieldId: string, direction: -1 | 1) {
  if (panelNavigationDisabled.value || settingsFieldId.value || panelStale.value) return
  const fields = draft.value?.fields
  if (!fields) return
  const current = fields.findIndex(field => field.fieldId === fieldId)
  const field = fields[current]
  if (!field) return
  const areaFields = fields.filter(candidate => candidate.area === field.area)
  const sibling = areaFields[areaFields.indexOf(field) + direction]
  if (!sibling) return
  const target = fields.indexOf(sibling)
  editFieldDraft(fieldId, () => { fields[current] = sibling; fields[target] = field })
}
async function applyFieldTool(change: PivotFieldToolChange) {
  const captured = owner.value, current = draft.value, id = selectedId.value, fieldId = toolsFieldId.value, currentGeneration = panelGeneration
  if (!fieldId || !captured || !current || panelNavigationDisabled.value || panelStale.value || !settingsEditable.value) return
  const request = ++fieldActionRequest, signature = JSON.stringify(current)
  const stillCurrent = (published = false): boolean => !disposed && props.open && currentGeneration === panelGeneration && draft.value === current
    && selectedId.value === id && toolsFieldId.value === fieldId && JSON.stringify(current) === signature
    && props.studio === captured.studio && workbook.value?.id === captured.unitId && props.sheetId === captured.sheetId && pivotChartPanelOwnerCurrent(captured, published)
  fieldDialogBusy.value = true; error.value = ''
  try {
    const { book } = requireApi()
    const pivot = book.getPivotTables().find(candidate => candidate.getPivotTableId() === id)
    if (!pivot) throw new Error('数据透视表已被删除。')
    const context = { ...pivotChartPanelOwnerContext(captured), hostSheetId: captured.sheetId }
    if (change.kind === 'valueFilter') await pivot.setValueFilter(fieldId, change.filter, context)
    else if (change.kind === 'numberGroup') await pivot.setNumberGroup(fieldId, change.settings, context)
    else if (change.kind === 'dateGroup') await pivot.setDateGroupType(fieldId, change.type, context)
    else await pivot.setManualGroups(fieldId, change.groups, context)
    if (!stillCurrent(true)) return
    const changed = captured.studio.getWorkbookRevision(captured.unitId) !== captured.revision
    owner.value = capturePivotChartPanelOwner(captured.studio, captured.unitId, captured.sheetId)
    toolsFieldId.value = ''; status.value = changed ? change.kind === 'valueFilter' ? '值筛选已应用。' : '字段分组已应用。' : '设置没有变化。'
    if (changed) reportChange(id)
  } catch (reason) { if (stillCurrent(true)) error.value = reason instanceof Error ? reason.message : String(reason) }
  finally { if (!disposed && request === fieldActionRequest) fieldDialogBusy.value = false }
}
function sourceFieldMenu(field: PivotPanelField) {
  return {
    items: [
      { key: 'row', label: '添加到行' },
      { key: 'column', label: '添加到列' },
      { key: 'value', label: '添加到值' },
      { key: 'filter', label: '添加到筛选' },
    ],
    onClick: ({ key }: { key: string | number }) => {
      if (panelNavigationDisabled.value || settingsFieldId.value || panelStale.value) return
      if (key === 'row' || key === 'column' || key === 'value' || key === 'filter') editFieldDraft(field.fieldId, () => { field.area = key; if (key !== 'value') field.showDataAs = null })
    },
  }
}
function startFieldDrag(fieldId: string, event: DragEvent) {
  if (panelNavigationDisabled.value || settingsFieldId.value || panelStale.value) { event.preventDefault(); return }
  draggedFieldId.value = fieldId
  draggedValues.value = undefined
  if (event.dataTransfer) {
    event.dataTransfer.effectAllowed = 'move'
    event.dataTransfer.setData('text/plain', fieldId)
  }
}
function startValuesDrag(event: DragEvent): void {
  const current = draft.value, captured = owner.value
  if (!current || !captured || panelNavigationDisabled.value || panelStale.value || !settingsEditable.value || !valuesState.value.visible) { event.preventDefault(); return }
  draggedFieldId.value = ''; draggedValues.value = { draft: current, owner: captured, generation: panelGeneration, signature: JSON.stringify(current) }
  if (event.dataTransfer) { event.dataTransfer.effectAllowed = 'move'; event.dataTransfer.setData('application/x-office-pivot-values', 'values') }
}
function dropField(area: PivotPanelField['area'], event: DragEvent, beforeFieldId?: string) {
  event.preventDefault()
  event.stopPropagation()
  if (panelNavigationDisabled.value || settingsFieldId.value || panelStale.value) return
  const valuesDrag = draggedValues.value
  if (valuesDrag) {
    draggedValues.value = undefined
    if ((area !== 'row' && area !== 'column') || valuesDrag.draft !== draft.value || valuesDrag.owner !== owner.value
      || valuesDrag.generation !== panelGeneration || valuesDrag.signature !== JSON.stringify(draft.value) || !pivotChartPanelOwnerCurrent(valuesDrag.owner)) return
    const siblings = draft.value!.fields.filter(field => field.area === area), before = siblings.findIndex(field => field.fieldId === beforeFieldId)
    setValuesPosition(area, before < 0 ? siblings.length : before); return
  }
  const fieldId = event.dataTransfer?.getData('text/plain') || draggedFieldId.value
  const fields = draft.value?.fields
  const currentIndex = fields?.findIndex(candidate => candidate.fieldId === fieldId) ?? -1
  if (!fields || currentIndex < 0) return
  if (beforeFieldId === fieldId) return
  editFieldDraft(fieldId, () => {
    const [field] = fields.splice(currentIndex, 1)
    if (!field) return
    field.area = area
    if (area !== 'value') field.showDataAs = null
    const beforeIndex = beforeFieldId ? fields.findIndex(candidate => candidate.fieldId === beforeFieldId && candidate.area === area) : -1
    if (beforeIndex < 0) fields.push(field)
    else fields.splice(beforeIndex, 0, field)
  })
  draggedFieldId.value = ''
}
function openFieldFilter(fieldId: string) {
  if (panelNavigationDisabled.value || panelStale.value || settingsFieldId.value) return
  error.value = ''; void filters.open(fieldId)
}
function openFieldSettings(fieldId: string, mode: 'all' | 'showDataAs' = 'all') {
  if (panelNavigationDisabled.value || panelStale.value || !settingsEditable.value || !draft.value || !owner.value) return
  const field = draft.value?.fields.find(candidate => candidate.fieldId === fieldId)
  const modelField = selectedPivot.value?.fields.find(candidate => candidate.id === fieldId)
  if (!field || field.area !== 'value') return
  error.value = ''
  settingsFieldId.value = fieldId
  settingsMode.value = mode; settingsGeneration++
  settingsSession.value = { draft: draft.value, owner: owner.value, pivotId: selectedId.value, fieldId, signature: settingsSignature(draft.value), generation: settingsGeneration }
  fieldDisplayName.value = field.displayName ?? modelField?.displayName ?? field.fieldName
  fieldAggregation.value = field.aggregation
  fieldFormat.value = field.format ?? modelField?.format ?? ''
  fieldShowDataAs.value = createPivotShowDataAsDraft(Object.hasOwn(field, 'showDataAs') ? field.showDataAs : modelField?.showDataAs)
  fieldShowAsValid.value = true
}
async function readSettingsBaseItems(fieldId: string): Promise<string[]> {
  const captured = settingsSession.value, report = settingsReport.value
  if (!captured || !report || !settingsReadable.value || !settingsCurrent(captured)) throw new Error('基准成员读取已过期，请重新读取透视表。')
  let items: string[]
  if (directFieldSetting.value && captured.pivotId) {
    const pivot = requireApi().book.getPivotTables().find(candidate => candidate.getPivotTableId() === captured.pivotId)
    if (!pivot) throw new Error('数据透视表已被删除。')
    items = await pivot.getShowDataAsBaseItems(fieldId)
  } else items = getPivotShowDataAsBaseItems(workbook.value!, report, fieldId)
  if (!settingsReadable.value || !settingsCurrent(captured)) throw new Error('透视表基准成员读取已过期。')
  return items
}
async function applyFieldSettings() {
  const captured = settingsSession.value, report = settingsReport.value
  if (!captured || !report || settingsDisabled.value || !fieldShowAsValid.value || !settingsCurrent(captured)) return
  const request = ++settingsRequest, mode = settingsMode.value, direct = directFieldSetting.value
  const displayName = fieldDisplayName.value.trim(), aggregation = fieldAggregation.value, format = fieldFormat.value
  const input = fieldShowDataAs.value ? { ...fieldShowDataAs.value } : null
  settingsBusy.value = true
  error.value = ''
  try {
    if (mode === 'all' && !displayName) throw new Error('字段名称不能为空。')
    const items = input?.baseItem ? await readSettingsBaseItems(input.baseFieldId) : undefined
    const showDataAs = preparePivotShowDataAsDraft(report, input, items)
    if (!settingsCurrent(captured)) throw new Error('值字段设置已过期，草稿已保留。')
    const field = captured.draft.fields.find(candidate => candidate.fieldId === captured.fieldId)
    if (!field) throw new Error('所选字段已被移除。')
    let changed = false
    if (direct) {
      const { book } = requireApi()
      const pivot = book.getPivotTables().find(candidate => candidate.getPivotTableId() === captured.pivotId)
      if (!pivot) throw new Error('数据透视表已被删除。')
      const context = pivotChartPanelOwnerContext(captured.owner)
      if (mode === 'showDataAs') await pivot.setShowDataAs(field.fieldId, showDataAs, context)
      else await pivot.setFieldSetting(field.fieldId, { displayName, subtotalType: aggregation, format, showDataAs }, context)
      if (!settingsCurrent(captured, true)) return
      changed = captured.owner.studio.getWorkbookRevision(captured.owner.unitId) !== captured.owner.revision
      owner.value = capturePivotChartPanelOwner(captured.owner.studio, captured.owner.unitId, captured.owner.sheetId)
    }
    if (mode === 'all') { field.aggregation = aggregation; field.displayName = displayName; field.format = format }
    field.showDataAs = showDataAs
    settingsBusy.value = false; closeSettings()
    status.value = direct ? changed ? mode === 'showDataAs' ? '值显示方式已应用。' : '值字段设置已应用。' : '设置没有变化。'
      : '值字段设置已保存到当前配置，请应用字段配置。'
    if (direct && changed) reportChange(captured.pivotId)
  } catch (reason) { if (settingsCurrent(captured, true)) error.value = reason instanceof Error ? reason.message : String(reason) }
  finally { if (!disposed && request === settingsRequest) settingsBusy.value = false }
}
function requireApi() {
  if (!props.studio || !workbook.value?.id) throw new Error('当前工作台尚未连接工作簿命令接口。')
  const api = FOffice.newAPI(props.studio), book = api.getWorkbook(workbook.value.id)
  if (!book) throw new Error('当前工作簿已关闭，请重新打开面板。')
  return { api, book, model: workbook.value }
}
function reportChange(pivotId: string, created = false) {
  const pivot = workbook.value?.pivotTables?.find(candidate => candidate.id === pivotId)
  if (!pivot) return
  const payload = { pivotId, sheetId: pivot.anchor.subUnitId, outputRange: pivot.outputRange }
  if (created) emit('created', payload)
  else emit('updated', payload)
}
function applyConfiguration() {
  const current = draft.value, captured = owner.value, currentGeneration = panelGeneration
  if (!current || !captured || panelNavigationDisabled.value || settingsFieldId.value || panelStale.value) return
  error.value = ''; status.value = ''; busy.value = true
  try {
    const { api, model } = requireApi()
    if (fieldsStale.value) throw new Error('数据区域已改变，请重新读取字段后配置。')
    const created = !current.pivotId, id = current.pivotId || `pivot-${crypto.randomUUID()}`
    const input = { ...preparePivotPanelInput(model, current, id), ...pivotChartPanelOwnerContext(captured), hostSheetId: captured.sheetId }
    api.executeCommandSync(CONFIGURE_PIVOT_COMMAND, input)
    if (disposed || !props.open || currentGeneration !== panelGeneration || draft.value !== current || !pivotChartPanelOwnerCurrent(captured, true)) return
    const changed = captured.studio.getWorkbookRevision(captured.unitId) !== captured.revision
    loadDraft(id)
    status.value = changed ? created ? '数据透视表已创建，可在画布中查看结果。' : '字段配置已应用。' : '设置没有变化。'
    if (changed) reportChange(id, created)
  } catch (reason) { error.value = reason instanceof Error ? reason.message : String(reason) }
  finally { busy.value = false }
}
async function refreshPivot() {
  if (!selectedId.value || panelNavigationDisabled.value || settingsFieldId.value || panelStale.value) return
  error.value = ''; status.value = ''; busy.value = true
  try {
    const { book } = requireApi(), pivot = book.getPivotTables().find(candidate => candidate.getPivotTableId() === selectedId.value)
    if (!pivot) throw new Error('要刷新的数据透视表已被删除。')
    await pivot.refresh(); status.value = '已根据当前数据源刷新结果。'; reportChange(selectedId.value)
  } catch (reason) { error.value = reason instanceof Error ? reason.message : String(reason) }
  finally { busy.value = false }
}
async function insertPivotChart() {
  if (!selectedId.value || panelNavigationDisabled.value || settingsFieldId.value || panelStale.value) return
  error.value = ''; status.value = ''; busy.value = true
  try {
    const { book } = requireApi()
    const pivot = book.getPivotTables().find(candidate => candidate.getPivotTableId() === selectedId.value)
    if (!pivot) throw new Error('所选数据透视表已被删除。')
    await pivot.insertChart(pivot.newChart('bar').setPosition('F2').setSize(480, 300).build())
    status.value = '透视图已创建，并与此数据透视表关联。'
    reportChange(selectedId.value)
  } catch (reason) { error.value = reason instanceof Error ? reason.message : String(reason) }
  finally { busy.value = false }
}
async function removePivot() {
  if (!selectedId.value || panelNavigationDisabled.value || settingsFieldId.value || panelStale.value) return
  error.value = ''; status.value = ''; busy.value = true
  try {
    const { book } = requireApi(), pivot = book.getPivotTables().find(candidate => candidate.getPivotTableId() === selectedId.value)
    if (!pivot) throw new Error('数据透视表已被删除。')
    await pivot.remove(); loadDraft(); status.value = '数据透视表已删除，可撤销恢复。'
  } catch (reason) { error.value = reason instanceof Error ? reason.message : String(reason) }
  finally { busy.value = false }
}
function navigateHistory(direction: 'undo' | 'redo') {
  if (panelNavigationDisabled.value || settingsFieldId.value) return
  try {
    const { book } = requireApi(), id = selectedId.value
    book[direction]()
    loadDraft(workbook.value?.pivotTables?.some(pivot => pivot.id === id) ? id : '')
    status.value = direction === 'undo' ? '已撤销上一步工作簿操作。' : '已重做工作簿操作。'
  } catch (reason) { error.value = reason instanceof Error ? reason.message : String(reason) }
}
function onPivotSelect(value: string | number) { if (!panelNavigationDisabled.value && !settingsFieldId.value) loadDraft(String(value)) }
function close() { if (!panelNavigationDisabled.value) { filters.close(false); closeSettings(); emit('update:open', false) } }
function startComposition(): void { compositionGeneration++; composing.value = true }
function endComposition(): void {
  const ended = compositionGeneration
  void nextTick(() => {
    if (disposed || ended !== compositionGeneration) return
    composing.value = false
    drainPendingLoad()
    if (!props.open) { closeSettings(); owner.value = undefined }
  })
}
function drainPendingLoad(): void {
  if (disposed || composing.value || filterComposing.value || pendingLoad === undefined) return
  const id = pendingLoad; pendingLoad = undefined
  if (props.open) loadDraft(id)
  else { closeSettings(); owner.value = undefined }
}
watch(filterComposing, value => { if (!value) drainPendingLoad() })
watch([() => props.studio?.currentUserId, () => props.studio && workbook.value?.id
  ? props.studio.editor.history.captureSavepoint(workbook.value.id).generation : -1], () => { modelRevision.value++ })
watch([() => props.studio, () => props.workbookId], ([studio], _previous, onCleanup) => {
  if (!studio) return
  const changed = (id: string): void => { if (!disposed && id === workbook.value?.id) modelRevision.value++ }
  const listeners = [
    studio.events.addEvent('CommandExecuted', event => { if (event.unitId) changed(event.unitId) }),
    studio.events.addEvent('Undo', event => { if (event.kind === 'workbook') changed(event.unitId) }),
    studio.events.addEvent('Redo', event => { if (event.kind === 'workbook') changed(event.unitId) }),
    studio.events.addEvent('PivotTableChanged', event => changed(event.unitId)),
    studio.events.addEvent('SheetValueChanged', event => changed(event.workbookId)),
    studio.events.addEvent('SheetStructureChanged', event => changed(event.workbookId)),
    studio.events.addEvent('WorkbookDisposed', event => {
      if (event.workbookId === owner.value?.unitId) { modelRevision.value++; settingsGeneration++; settingsRequest++; fieldActionRequest++; settingsSession.value = undefined; owner.value = undefined; emit('update:open', false) }
    }),
  ]
  onCleanup(() => listeners.forEach(listener => listener.dispose()))
}, { immediate: true })
onScopeDispose(() => { disposed = true; panelGeneration++; settingsGeneration++; settingsRequest++; fieldActionRequest++; settingsSession.value = undefined })
</script>

<template>
  <Modal :open="open" title="数据透视表" :width="820" :footer="null" :root-style="panelStyle" :mask="{ closable: !panelNavigationDisabled }"
    :keyboard="!panelNavigationDisabled && !settingsFieldId && !numberFormatOpen" :closable="!panelNavigationDisabled" @cancel="close">
    <div ref="panelElement" class="pivot-panel" :style="panelStyle" data-testid="pivot-table-panel">
      <Alert v-if="panelStale" type="warning" title="内容、用户或历史已更改。当前草稿已保留，请重新读取后再应用。" show-icon>
        <template #action><Button size="small" :disabled="panelNavigationDisabled" @click="loadDraft(selectedId)">重新读取</Button></template>
      </Alert>
      <Form layout="vertical" class="pivot-form" :disabled="busy || settingsDialogBusy || !!settingsFieldId || !!toolsFieldId || filterOpen || panelStale && !composing" @submit.prevent
        @compositionstart.capture="startComposition" @compositionend.capture="endComposition">
      <div class="pivot-panel-selector"><FormItem label="选择数据透视表"><Select :value="selectedId" :options="pivotOptions" aria-label="选择数据透视表" @change="onPivotSelect" /></FormItem></div>
      <template v-if="draft">
        <section class="pivot-locations" aria-label="数据透视表数据源和位置">
          <FormItem label="数据源工作表"><Select v-model:value="draft.sourceSheetId" :options="sheetOptions" aria-label="透视表数据源工作表" :disabled="panelNavigationDisabled || !!settingsFieldId || panelStale" /></FormItem>
          <FormItem label="数据区域"><Input v-model:value="draft.sourceRange" aria-label="透视表数据区域" placeholder="例如 A1:C20" /></FormItem>
          <FormItem label="放置工作表"><Select v-model:value="draft.targetSheetId" :options="sheetOptions" aria-label="透视表放置工作表" :disabled="panelNavigationDisabled || !!settingsFieldId || panelStale" /></FormItem>
          <FormItem label="起始单元格"><Input v-model:value="draft.targetAddress" aria-label="透视表起始单元格" placeholder="例如 E1" /></FormItem>
        </section>
        <div class="pivot-fields-header"><span>字段</span><Button size="small" @click="loadFields">读取字段</Button></div>
        <p v-if="fieldsStale" class="pivot-hint">数据区域已改变，请读取字段后重新配置。</p>
        <section class="pivot-area-layout" aria-label="透视字段配置">
          <div class="pivot-source-area" aria-label="未放置字段">
            <span class="pivot-area-title">源字段</span>
            <p class="pivot-source-description">选择字段并将其添加到透视表区域</p>
            <FormItem class="pivot-search-item"><Input v-model:value="fieldSearch" aria-label="查找透视字段" placeholder="查找字段" size="small" /></FormItem>
            <div class="pivot-field-list" @dragover.prevent @drop="dropField('unused', $event)">
              <div v-for="field in visibleFields" :key="field.fieldId" class="pivot-source-field" :class="{ 'is-dragging': draggedFieldId === field.fieldId }" :data-field-id="field.fieldId" draggable="true" @dragstart="startFieldDrag(field.fieldId, $event)" @dragend="draggedFieldId = ''">
                <Checkbox :checked="field.area !== 'unused'" :aria-label="`添加字段 ${field.fieldName}`" @change="toggleField(field.fieldId, field.area === 'unused')" />
                <span class="pivot-field-name" :title="field.fieldName">{{ field.fieldName }}</span>
                  <Dropdown :trigger="['click']" placement="bottomRight" :menu="sourceFieldMenu(field)"><Button type="text" size="small" :aria-label="`字段 ${field.fieldName} 添加到区域`"><ChevronDown :size="16" /></Button></Dropdown>
              </div>
              <p v-if="!visibleFields.length" class="pivot-empty">没有匹配的字段</p>
            </div>
          </div>
          <div class="pivot-area-grid">
            <section v-for="area in fieldAreas" :key="area.area" class="pivot-area" :aria-label="`${area.label}区域`" @dragover.prevent @drop="dropField(area.area, $event)">
              <h3 class="pivot-area-title">{{ area.label }}</h3>
              <div class="pivot-field-list">
              <template v-for="item in area.entries" :key="item.kind === 'field' ? item.field.fieldId : 'synthetic-values'">
                <PivotValuesFieldCard v-if="item.kind === 'values'" :state="item.values" subject-label="透视表"
                  :disabled="panelNavigationDisabled || panelStale || !settingsEditable" @change="setValuesPosition" @drag-start="startValuesDrag"
                  @drag-end="draggedValues = undefined" @dragover.prevent @drop="dropField(area.area, $event, area.fields[valuesState.effectiveIndex]?.fieldId)" />
                <div v-else class="pivot-field" :class="{ 'is-dragging': draggedFieldId === item.field.fieldId }" :data-field-id="item.field.fieldId" draggable="true"
                  @dragover.prevent @drop="dropField(area.area, $event, item.field.fieldId)" @dragstart="startFieldDrag(item.field.fieldId, $event)" @dragend="draggedFieldId = ''">
                <GripVertical class="pivot-field-grip" :size="16" aria-hidden="true" /><span class="pivot-field-name" :title="item.field.displayName ?? item.field.fieldName">{{ item.field.displayName ?? item.field.fieldName }}</span>
                  <Dropdown :trigger="['click']" placement="bottomRight" :menu="fieldMenu(item.field)"><Button type="text" size="small" :aria-label="`字段 ${item.field.fieldName} 操作`"><Ellipsis :size="16" /></Button></Dropdown>
                </div>
              </template>
                <p v-if="!area.entries.length" class="pivot-empty">拖动字段到此处</p>
              </div>
            </section>
          </div>
        </section>
        <section class="pivot-options" aria-label="透视表汇总选项">
          <FormItem class="pivot-option-item"><Switch v-model:checked="draft.options.showRowGrandTotal" size="small" aria-label="显示行总计" />行总计</FormItem>
          <FormItem class="pivot-option-item"><Switch v-model:checked="draft.options.showColGrandTotal" size="small" aria-label="显示列总计" />列总计</FormItem>
          <FormItem class="pivot-option-item"><Switch v-model:checked="draft.options.showRowSubtotals" size="small" aria-label="显示行小计" />行小计</FormItem>
          <FormItem class="pivot-option-item"><Switch v-model:checked="draft.options.showColSubtotals" size="small" aria-label="显示列小计" />列小计</FormItem>
          <FormItem class="pivot-option-item"><Switch v-model:checked="draft.options.repeatRowLabels" size="small" aria-label="重复行标签" />重复行标签</FormItem>
          <FormItem class="pivot-option-item"><Switch :checked="draft.options.allowMultipleFilters === true" size="small" aria-label="允许字段使用多个筛选"
            :disabled="panelNavigationDisabled || panelStale || !settingsEditable" @update:checked="setAllowMultipleFilters" />允许多个筛选</FormItem>
          <PivotValuesPositionFields :state="valuesState" subject-label="透视表" :disabled="panelNavigationDisabled || panelStale || !settingsEditable" @change="setValuesPosition" />
        </section>
        <p v-if="selectedPivot?.outputRange" class="pivot-hint" data-testid="pivot-output-range">结果区域：{{ workbook?.sheets.find(sheet => sheet.id === selectedPivot?.anchor.subUnitId)?.name }}!{{ pivotRangeAddress(selectedPivot.outputRange) }}</p>
      </template>
      <p v-if="error" class="pivot-error" role="alert">{{ error }}</p>
      <p v-if="status" class="pivot-status" role="status">{{ status }}</p>
      <div class="pivot-actions">
        <Button size="small" :disabled="busy || !history?.undo.length" @click="navigateHistory('undo')">撤销透视操作</Button>
        <Button size="small" :disabled="busy || !history?.redo.length" @click="navigateHistory('redo')">重做透视操作</Button>
        <Button v-if="selectedId" size="small" :loading="busy" @click="refreshPivot">刷新数据</Button>
        <Button v-if="selectedId && selectedPivot?.outputRange" size="small" :disabled="busy" @click="insertPivotChart">插入关联透视图</Button>
        <Button v-if="selectedId" size="small" danger :disabled="busy" @click="removePivot">删除数据透视表</Button>
        <span class="pivot-actions-spacer" />
        <Button @click="close">关闭</Button>
        <Button v-if="draft" type="primary" :loading="busy" @click="applyConfiguration">{{ draft.pivotId ? '应用字段配置' : '创建数据透视表' }}</Button>
      </div>
      </Form>
    </div>
  </Modal>
  <Modal :open="filterOpen" :title="null" :width="360" :footer="null" :root-style="panelStyle" :closable="false" :mask="false"
    :keyboard="false" :focus-trigger-after-close="false" destroy-on-hidden @cancel="filters.close()">
    <PivotChartFieldFilterPopup v-if="filterButton" :key="filterFieldId" :field="filterButton" :draft="filterDraft" :conditions="filterConditions"
      subject-label="透视表" test-id="pivot-table-field-filter-popup" :show-sort="filterButton.area === 'row' || filterButton.area === 'column'"
      :show-management="false" :readable="filterReadable" :loading="filterLoading" :busy="filterBusy" :composing="filterComposing"
      :disabled="filterEditDisabled" :navigation-disabled="filterNavigationDisabled" :stale="filterStale" :error="filterError"
      @search="filters.search" @select="filters.select" @select-all="filters.selectAll" @date-change="filters.dateChange" @label-change="filters.labelChange"
      @action="filterAction" @retry="filters.open(filterFieldId)" @close="filters.close()" @composition-start="filters.compositionStart"
      @composition-end="filters.compositionEnd" @keydown="filters.keydown" />
  </Modal>
  <Modal :open="open && Boolean(settingsFieldId)" :title="settingsMode === 'showDataAs' ? '值显示方式' : '字段设置'" :width="460" :footer="null" :root-style="panelStyle"
    :mask="{ closable: false }" :keyboard="!settingsDialogBusy && !composing && !numberFormatOpen" :closable="!settingsDialogBusy && !composing"
    :focus-trigger-after-close="true" destroy-on-hidden @cancel="closeSettings">
    <Form v-if="settingsField" layout="vertical" class="pivot-settings-dialog" :style="panelStyle" data-testid="pivot-field-settings" @submit.prevent="applyFieldSettings"
      @compositionstart.capture="startComposition" @compositionend.capture="endComposition">
      <template v-if="settingsMode === 'all'">
        <FormItem class="pivot-settings-name" label="字段名称"><span v-if="settingsReadable" class="pivot-settings-source">源字段：{{ settingsField.fieldName }}</span><Input v-model:value="fieldDisplayName" size="small" aria-label="值字段显示名称" :disabled="settingsDialogBusy || !composing && (panelStale || !settingsEditable)" /></FormItem>
        <FormItem label="汇总方式"><RadioGroup v-model:value="fieldAggregation" class="pivot-aggregation-list" :options="aggregationOptions" orientation="vertical" :disabled="settingsDisabled" aria-label="汇总方式" /></FormItem>
      </template>
      <PivotShowDataAsFields :value="fieldShowDataAs" :report="settingsReadable ? settingsReport : undefined" :load-base-items="readSettingsBaseItems"
        :readable="settingsReadable && !panelStale" :disabled="settingsDisabled" :context-key="`${settingsGeneration}:${owner?.revision}:${owner?.historyGeneration}:${studio?.currentUserId}`"
        @update:value="fieldShowDataAs = $event" @validity-change="fieldShowAsValid = $event" />
      <Button v-if="settingsMode === 'all'" :disabled="settingsDisabled" @click="numberFormatDraft = fieldFormat || 'General'; numberFormatOpen = true">数字格式</Button>
      <Alert v-if="settingsHostError || !settingsEditable" type="info" :title="settingsHostError || '当前用户没有修改透视表来源或输出范围的权限。'" show-icon />
      <Alert v-if="panelStale" type="warning" title="此设置草稿已过期。请取消并在主面板重新读取。" show-icon />
      <Alert v-if="error" type="error" :title="error" show-icon />
      <div class="pivot-settings-footer"><Button :disabled="settingsDialogBusy || composing" @click="closeSettings">取消</Button><Button type="primary" :loading="settingsBusy" :disabled="settingsDisabled || !fieldShowAsValid || settingsMode === 'all' && !fieldDisplayName.trim()" @click="applyFieldSettings">确定</Button></div>
    </Form>
  </Modal>
  <Modal :open="open && numberFormatOpen" title="数字格式" :width="400" :footer="null" :root-style="panelStyle" :mask="{ closable: false }" :keyboard="!composing && !settingsDialogBusy" :closable="!composing && !settingsDialogBusy" destroy-on-hidden @cancel="numberFormatOpen = false">
    <Form layout="vertical" class="pivot-settings-dialog" :style="panelStyle" @submit.prevent @compositionstart.capture="startComposition" @compositionend.capture="endComposition">
      <FormItem label="格式类型"><Select v-model:value="numberFormatDraft" :options="numberFormatOptions" aria-label="数字格式类型" :disabled="settingsDisabled" /></FormItem>
      <FormItem label="格式代码"><Input v-model:value="numberFormatDraft" aria-label="数字格式代码" :disabled="settingsDialogBusy || !composing && (panelStale || !settingsEditable)" /></FormItem>
      <output class="pivot-format-preview" aria-label="数字格式预览">{{ numberFormatPreview }}</output>
      <div class="pivot-settings-footer"><Button :disabled="composing || settingsDialogBusy" @click="numberFormatOpen = false">取消</Button><Button type="primary" :disabled="settingsDisabled" @click="fieldFormat = numberFormatDraft; numberFormatOpen = false">确定</Button></div>
    </Form>
  </Modal>
  <PivotFieldToolsDialog v-if="toolsField && workbook && selectedPivot" :open="Boolean(toolsFieldId)" :kind="toolsKind" :workbook="workbook" :pivot="selectedPivot" :field="toolsField"
    :busy="fieldDialogBusy" :error="error" :stale="panelStale" :readable="settingsReadable" :editable="settingsEditable" :supports-date-grouping="toolsDateSupported"
    @close="toolsFieldId = ''" @apply="applyFieldTool" @composition-start="startComposition" @composition-end="endComposition" />
</template>

<style scoped>
.pivot-panel { display: flex; flex-direction: column; gap: 12px; min-width: 0; }
.pivot-form { display: flex; flex-direction: column; gap: 12px; min-width: 0; }
.pivot-form :deep(.ant-form-item) { margin-bottom: 0; }
.pivot-form :deep(.ant-form-item-label) { padding-bottom: 5px; }
.pivot-search-item :deep(.ant-form-item-label) { display: none; }
.pivot-panel-selector { max-width: 100%; }
.pivot-locations { display: grid; grid-template-columns: 1fr 1fr; gap: 10px 16px; }
.pivot-fields-header { display: flex; align-items: center; justify-content: space-between; padding-top: 4px; color: var(--pivot-text); font-weight: 600; }
.pivot-area-layout { display: grid; grid-template-columns: minmax(180px, 0.8fr) minmax(0, 1.8fr); gap: 12px; min-height: 200px; }
.pivot-area-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px 16px; }
.pivot-area, .pivot-source-area { box-sizing: border-box; min-height: 200px; padding: 4px; }
.pivot-area-title { display: block; margin: 0 0 8px; color: var(--pivot-text); font-size: 13px; font-weight: 500; line-height: 20px; }
.pivot-source-description { margin: -4px 0 8px; color: var(--pivot-secondary); font-size: 12px; line-height: 16px; }
.pivot-source-area :deep(.ant-input-affix-wrapper) { margin-bottom: 8px; }
.pivot-field-list { box-sizing: border-box; display: flex; flex-direction: column; gap: 4px; height: 161px; overflow-y: auto; padding: 8px; border: 1px solid var(--pivot-border); border-radius: 6px; background: var(--pivot-surface); }
.pivot-source-field { display: flex; align-items: center; gap: 8px; min-height: 32px; padding: 2px 4px; border-radius: 4px; cursor: grab; }
.pivot-field { display: grid; grid-template-columns: 16px minmax(0, 1fr) auto; gap: 5px; align-items: center; min-height: 28px; padding: 2px 4px; border: 1px solid transparent; border-radius: 4px; background: v-bind('token.colorBgContainer'); cursor: grab; }
.pivot-field.is-dragging { opacity: 0.55; }
.pivot-field-grip { color: var(--pivot-muted); }
.pivot-field-name { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-size: 12px; color: var(--pivot-text); }
.pivot-field :deep(.ant-btn) { padding: 0; }
.pivot-empty { margin: auto; color: var(--pivot-muted); font-size: 12px; text-align: center; }
.pivot-options { display: flex; align-items: center; gap: 10px 16px; flex-wrap: wrap; padding: 4px 0; }
.pivot-options :deep(.pivot-option-item) { display: flex; flex: 0 0 auto; align-items: center; gap: 6px; margin: 0; }
.pivot-options :deep(.pivot-option-item .ant-form-item-control-input-content) { display: flex; align-items: center; gap: 6px; }
.pivot-options :deep(.ant-select) { min-width: 76px; }
.pivot-hint, .pivot-status, .pivot-error { margin: 0; font-size: 12px; line-height: 1.6; }
.pivot-hint { color: var(--pivot-secondary); }
.pivot-error { color: var(--pivot-error); background: var(--pivot-error-bg); padding: 8px 10px; border-radius: 4px; }
.pivot-status { color: var(--pivot-success); }
.pivot-actions { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; padding-top: 12px; border-top: 1px solid var(--pivot-border); }
.pivot-actions-spacer { flex: 1; }
.pivot-settings-footer { display: flex; align-items: center; gap: 8px; }
.pivot-settings-dialog { display: flex; flex-direction: column; gap: 16px; }
.pivot-settings-dialog :deep(.ant-form-item) { margin-bottom: 0; }
.pivot-settings-dialog :deep(.ant-form-item-label) { padding-bottom: 6px; }
.pivot-settings-name :deep(.ant-form-item-control-input-content) { display: flex; flex-direction: column; gap: 6px; }
.pivot-settings-source { color: var(--pivot-secondary); font-size: 12px; }
.pivot-aggregation-list { display: flex; flex-direction: column; height: 160px; overflow-y: auto; padding: 4px; border: 1px solid var(--pivot-border); border-radius: 4px; }
.pivot-aggregation-list :deep(.ant-radio-wrapper) { flex: 0 0 28px; align-items: center; margin: 0; padding: 2px 8px; }
.pivot-format-preview { padding: 12px; border: 1px solid var(--pivot-border); border-radius: 4px; text-align: right; }
.pivot-settings-footer { justify-content: flex-end; padding-top: 12px; }
@media (max-width: 620px) { .pivot-locations, .pivot-area-layout { grid-template-columns: 1fr; } .pivot-area-grid { grid-template-columns: 1fr; } .pivot-field { gap: 3px; } .pivot-actions-spacer { display: none; } }
</style>
