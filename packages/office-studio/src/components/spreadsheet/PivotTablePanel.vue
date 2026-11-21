<script setup lang="ts">
import { Button, Checkbox, Dropdown, Input, Modal, Select, Switch } from 'antdv-next'
import { computed, ref, watch } from 'vue'
import { ArrowDownAZ, ArrowUpAZ, Check, ChevronDown, Ellipsis, GripVertical } from 'lucide-vue-next'
import { FOffice } from '../../api/facade'
import type { OfficeStudio } from '../../api/office-studio'
import type { CellRange, PivotAggregation } from '../../domain/model'
import {
  CONFIGURE_PIVOT_COMMAND, createPivotPanelDraft, pivotRangeAddress, preparePivotPanelInput,
  readPivotPanelSourceFields, type PivotPanelDraft, type PivotPanelField,
} from '../../features/pivot/services/pivot-panel-state'
import PivotFieldToolsDialog from '../../features/pivot/views/PivotFieldToolsDialog.vue'
import { readPivotFieldMembers, type PivotFieldToolChange } from '../../features/pivot/services/pivot-field-tools'
import { formatSheetValue, SHEET_NUMBER_FORMATS } from '../../features/sheet-workbench/sheet-number-format'

const props = defineProps<{ open: boolean; studio?: OfficeStudio; workbookId?: string; sheetId: string; sourceRange: string }>()
interface PivotPanelChange { pivotId: string; sheetId: string; outputRange?: CellRange }
const emit = defineEmits<{
  'update:open': [open: boolean]
  created: [result: PivotPanelChange]
  updated: [result: PivotPanelChange]
}>()
const workbook = computed(() => props.studio?.getWorkbookModel(props.workbookId ?? props.studio.activeWorkbookId))
const draft = ref<PivotPanelDraft | null>(null), selectedId = ref(''), error = ref(''), status = ref(''), busy = ref(false)
const loadedSource = ref(''), fieldSearch = ref('')
const draggedFieldId = ref('')
const filterFieldId = ref(''), filterSelection = ref<string[]>([]), settingsFieldId = ref('')
const fieldDisplayName = ref(''), fieldAggregation = ref<PivotAggregation>('sum'), fieldFormat = ref('')
const fieldDialogBusy = ref(false)
const toolsFieldId = ref(''), toolsKind = ref<'filter' | 'group'>('filter')
const numberFormatOpen = ref(false), numberFormatDraft = ref('')
const toolsField = computed(() => selectedPivot.value?.fields.find(field => field.id === toolsFieldId.value))
const numberFormatOptions = Object.entries(SHEET_NUMBER_FORMATS).map(([value, pattern]) => ({ value: pattern, label: ({ general: '常规', number: '数值', currency: '货币', accounting: '会计', percent: '百分比', date: '日期', time: '时间', datetime: '日期时间', text: '文本' } as Record<string, string>)[value] ?? value }))
const numberFormatPreview = computed(() => formatSheetValue(1234.56, numberFormatDraft.value, workbook.value?.date1904) ?? '此自定义格式将按原样保存在值字段中')
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
const fieldAreas = computed(() => [
  { area: 'filter' as const, label: '筛选', fields: draft.value?.fields.filter(field => field.area === 'filter') ?? [] },
  { area: 'column' as const, label: '列', fields: draft.value?.fields.filter(field => field.area === 'column') ?? [] },
  { area: 'row' as const, label: '行', fields: draft.value?.fields.filter(field => field.area === 'row') ?? [] },
  { area: 'value' as const, label: '值', fields: draft.value?.fields.filter(field => field.area === 'value') ?? [] },
])
const filterField = computed(() => draft.value?.fields.find(field => field.fieldId === filterFieldId.value))
const settingsField = computed(() => draft.value?.fields.find(field => field.fieldId === settingsFieldId.value))
const filterItems = computed(() => {
  const field = selectedPivot.value?.fields.find(candidate => candidate.id === filterFieldId.value)
  if (!field || !workbook.value || !selectedPivot.value) return []
  return [...new Set(readPivotFieldMembers(workbook.value, selectedPivot.value, field).map(member => member.value == null ? '' : String(member.value)))]
})
const allFilterItemsSelected = computed(() => filterItems.value.length > 0 && filterItems.value.every(item => filterSelection.value.includes(item)))
const canClearFieldFilter = computed(() => Boolean(selectedPivot.value?.fields.find(field => field.id === filterFieldId.value)?.selectedItems))
const history = computed(() => props.studio?.editor.history.units[workbook.value?.id ?? ''])

function loadDraft(pivotId = '') {
  const model = workbook.value
  error.value = ''; status.value = ''; selectedId.value = pivotId; fieldSearch.value = ''
  filterFieldId.value = ''; settingsFieldId.value = ''; toolsFieldId.value = ''; numberFormatOpen.value = false; fieldDialogBusy.value = false
  if (!model) { draft.value = null; error.value = '当前工作台尚未连接工作簿命令接口。'; return }
  draft.value = createPivotPanelDraft(model, props.sheetId, props.sourceRange, pivotId)
  loadedSource.value = sourceKey.value
}
watch(() => props.open, open => { if (open) loadDraft(); else { filterFieldId.value = ''; settingsFieldId.value = ''; toolsFieldId.value = ''; numberFormatOpen.value = false } }, { immediate: true })
function loadFields() {
  if (!draft.value || !workbook.value) return
  try {
    const previous = draft.value.fields
    const next = readPivotPanelSourceFields(workbook.value, draft.value.sourceSheetId, draft.value.sourceRange)
    draft.value.fields = fieldsStale.value ? next : next.map(field => ({ ...field, ...previous.find(candidate => candidate.fieldId === field.fieldId), fieldName: field.fieldName }))
    loadedSource.value = sourceKey.value; error.value = ''; status.value = `已读取 ${next.length} 个字段，请选择行、列和汇总值。`
  } catch (reason) { error.value = reason instanceof Error ? reason.message : String(reason) }
}
function removeField(fieldId: string) {
  const fields = draft.value?.fields
  if (!fields) return
  const field = fields.find(candidate => candidate.fieldId === fieldId)
  if (field) field.area = 'unused'
}
function toggleField(fieldId: string, checked: boolean) {
  const field = draft.value?.fields.find(candidate => candidate.fieldId === fieldId)
  if (!field) return
  if (!checked) { field.area = 'unused'; return }
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
  field.area = foundNumber && !foundOther ? 'value' : 'row'
}
function fieldMenu(field: PivotPanelField) {
  const committed = selectedPivot.value?.fields.find(candidate => candidate.id === field.fieldId)
  const editableLabel = committed && committed.area !== 'value' && field.area === committed.area && !fieldsStale.value
  const items = [
    { key: 'remove', label: '移除字段', danger: true },
    ...(editableLabel ? [{ key: 'filter', label: '筛选字段' }, { key: 'value-filter', label: '值筛选' }, { key: 'group', label: '分组' }] : []),
    ...(field.area === 'value' ? [{ key: 'settings', label: '字段设置' }] : []),
    { key: 'up', label: '上移', disabled: draft.value?.fields.filter(candidate => candidate.area === field.area).at(0)?.fieldId === field.fieldId },
    { key: 'down', label: '下移', disabled: draft.value?.fields.filter(candidate => candidate.area === field.area).at(-1)?.fieldId === field.fieldId },
    ...(['filter', 'column', 'row', 'value'] as const).filter(area => area !== field.area).map(area => ({ key: `move-${area}`, label: `移动到${{ filter: '筛选', column: '列', row: '行', value: '值' }[area]}` })),
  ]
  return {
    items,
    onClick: ({ key }: { key: string | number }) => {
      if (key === 'remove') removeField(field.fieldId)
      else if (key === 'filter') openFieldFilter(field.fieldId)
      else if (key === 'settings') openFieldSettings(field.fieldId)
      else if (key === 'value-filter' || key === 'group') { toolsFieldId.value = field.fieldId; toolsKind.value = key === 'value-filter' ? 'filter' : 'group'; error.value = '' }
      else if (key === 'up' || key === 'down') moveField(field.fieldId, key === 'up' ? -1 : 1)
      else if (String(key).startsWith('move-')) {
        const area = String(key).slice(5)
        if (area === 'row' || area === 'column' || area === 'filter' || area === 'value') field.area = area
      }
    },
  }
}
function moveField(fieldId: string, direction: -1 | 1) {
  const fields = draft.value?.fields
  if (!fields) return
  const current = fields.findIndex(field => field.fieldId === fieldId)
  const field = fields[current]
  if (!field) return
  const areaFields = fields.filter(candidate => candidate.area === field.area)
  const sibling = areaFields[areaFields.indexOf(field) + direction]
  if (!sibling) return
  const target = fields.indexOf(sibling)
  fields[current] = sibling
  fields[target] = field
}
async function applyFieldTool(change: PivotFieldToolChange) {
  if (!toolsFieldId.value || fieldDialogBusy.value) return
  fieldDialogBusy.value = true; error.value = ''
  try {
    const { book } = requireApi()
    const pivot = book.getPivotTables().find(candidate => candidate.getPivotTableId() === selectedId.value)
    if (!pivot) throw new Error('数据透视表已被删除。')
    if (change.kind === 'valueFilter') await pivot.setValueFilter(toolsFieldId.value, change.filter)
    else if (change.kind === 'numberGroup') await pivot.setNumberGroup(toolsFieldId.value, change.settings)
    else if (change.kind === 'dateGroup') await pivot.setDateGroupType(toolsFieldId.value, change.type)
    else await pivot.setManualGroups(toolsFieldId.value, change.groups)
    toolsFieldId.value = ''; status.value = change.kind === 'valueFilter' ? '值筛选已应用。' : '字段分组已应用。'; reportChange(selectedId.value)
  } catch (reason) { error.value = reason instanceof Error ? reason.message : String(reason) }
  finally { fieldDialogBusy.value = false }
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
      if (key === 'row' || key === 'column' || key === 'value' || key === 'filter') field.area = key
    },
  }
}
function startFieldDrag(fieldId: string, event: DragEvent) {
  draggedFieldId.value = fieldId
  if (event.dataTransfer) {
    event.dataTransfer.effectAllowed = 'move'
    event.dataTransfer.setData('text/plain', fieldId)
  }
}
function dropField(area: PivotPanelField['area'], event: DragEvent, beforeFieldId?: string) {
  event.preventDefault()
  event.stopPropagation()
  const fieldId = event.dataTransfer?.getData('text/plain') || draggedFieldId.value
  const fields = draft.value?.fields
  const currentIndex = fields?.findIndex(candidate => candidate.fieldId === fieldId) ?? -1
  if (!fields || currentIndex < 0) return
  if (beforeFieldId === fieldId) return
  const [field] = fields.splice(currentIndex, 1)
  if (!field) return
  field.area = area
  const beforeIndex = beforeFieldId ? fields.findIndex(candidate => candidate.fieldId === beforeFieldId && candidate.area === area) : -1
  if (beforeIndex < 0) {
    fields.push(field)
  } else {
    fields.splice(beforeIndex, 0, field)
  }
  draggedFieldId.value = ''
}
function openFieldFilter(fieldId: string) {
  const field = draft.value?.fields.find(candidate => candidate.fieldId === fieldId)
  const modelField = selectedPivot.value?.fields.find(candidate => candidate.id === fieldId)
  if (!field || !modelField || modelField.area === 'value' || !selectedId.value) return
  error.value = ''
  filterFieldId.value = fieldId
  const availableItems = new Set(filterItems.value)
  filterSelection.value = modelField.selectedItems?.filter(item => availableItems.has(item)) ?? filterItems.value.slice()
}
function toggleAllFilterItems(checked: boolean) {
  filterSelection.value = checked ? filterItems.value.slice() : []
}
function toggleFilterItem(value: string) {
  filterSelection.value = filterSelection.value.includes(value)
    ? filterSelection.value.filter(candidate => candidate !== value)
    : [...filterSelection.value, value]
}
async function applyLabelFilter(selectedItems: string[] | null) {
  if (!selectedId.value || !filterFieldId.value || fieldDialogBusy.value) return
  fieldDialogBusy.value = true
  error.value = ''
  try {
    const { book } = requireApi()
    const pivot = book.getPivotTables().find(candidate => candidate.getPivotTableId() === selectedId.value)
    if (!pivot) throw new Error('数据透视表已被删除。')
    await pivot.setLabelManualFilter(filterFieldId.value, selectedItems)
    filterFieldId.value = ''
    status.value = '透视字段筛选已应用。'
    reportChange(selectedId.value)
  } catch (reason) { error.value = reason instanceof Error ? reason.message : String(reason) }
  finally { fieldDialogBusy.value = false }
}
async function applyLabelSort(direction: 'ascending' | 'descending') {
  if (!selectedId.value || !filterFieldId.value || fieldDialogBusy.value) return
  fieldDialogBusy.value = true
  error.value = ''
  try {
    const { book } = requireApi()
    const pivot = book.getPivotTables().find(candidate => candidate.getPivotTableId() === selectedId.value)
    if (!pivot) throw new Error('数据透视表已被删除。')
    await pivot.setLabelSort(filterFieldId.value, { type: direction })
    const field = draft.value?.fields.find(candidate => candidate.fieldId === filterFieldId.value)
    if (field) field.sort = direction
    filterFieldId.value = ''
    status.value = '透视字段排序已应用。'
    reportChange(selectedId.value)
  } catch (reason) { error.value = reason instanceof Error ? reason.message : String(reason) }
  finally { fieldDialogBusy.value = false }
}
function openFieldSettings(fieldId: string) {
  const field = draft.value?.fields.find(candidate => candidate.fieldId === fieldId)
  const modelField = selectedPivot.value?.fields.find(candidate => candidate.id === fieldId)
  if (!field || field.area !== 'value') return
  error.value = ''
  settingsFieldId.value = fieldId
  fieldDisplayName.value = field.displayName ?? modelField?.displayName ?? field.fieldName
  fieldAggregation.value = field.aggregation
  fieldFormat.value = field.format ?? modelField?.format ?? ''
}
async function applyFieldSettings() {
  if (!settingsFieldId.value || fieldDialogBusy.value) return
  fieldDialogBusy.value = true
  error.value = ''
  try {
    const field = draft.value?.fields.find(candidate => candidate.fieldId === settingsFieldId.value)
    if (!field) throw new Error('所选字段已被移除。')
    const committed = selectedPivot.value?.fields.find(candidate => candidate.id === field.fieldId)
    if (committed?.area === 'value' && !fieldsStale.value) {
      const { book } = requireApi()
      const pivot = book.getPivotTables().find(candidate => candidate.getPivotTableId() === selectedId.value)
      if (!pivot) throw new Error('数据透视表已被删除。')
      await pivot.setFieldSetting(field.fieldId, { displayName: fieldDisplayName.value.trim(), subtotalType: fieldAggregation.value, format: fieldFormat.value })
      reportChange(selectedId.value)
    }
    field.aggregation = fieldAggregation.value; field.displayName = fieldDisplayName.value.trim(); field.format = fieldFormat.value
    settingsFieldId.value = ''
    status.value = committed?.area === 'value' ? '值字段设置已应用。' : '值字段设置已保存到当前配置。'
  } catch (reason) { error.value = reason instanceof Error ? reason.message : String(reason) }
  finally { fieldDialogBusy.value = false }
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
  if (!draft.value || busy.value) return
  error.value = ''; status.value = ''; busy.value = true
  try {
    const { api, model } = requireApi()
    if (fieldsStale.value) throw new Error('数据区域已改变，请重新读取字段后配置。')
    const created = !draft.value.pivotId, id = draft.value.pivotId || `pivot-${crypto.randomUUID()}`
    const input = preparePivotPanelInput(model, draft.value, id)
    api.executeCommandSync(CONFIGURE_PIVOT_COMMAND, input)
    loadDraft(id)
    status.value = created ? '数据透视表已创建，可在画布中查看结果。' : '字段配置已应用。'
    reportChange(id, created)
  } catch (reason) { error.value = reason instanceof Error ? reason.message : String(reason) }
  finally { busy.value = false }
}
async function refreshPivot() {
  if (!selectedId.value || busy.value) return
  error.value = ''; status.value = ''; busy.value = true
  try {
    const { book } = requireApi(), pivot = book.getPivotTables().find(candidate => candidate.getPivotTableId() === selectedId.value)
    if (!pivot) throw new Error('要刷新的数据透视表已被删除。')
    await pivot.refresh(); status.value = '已根据当前数据源刷新结果。'; reportChange(selectedId.value)
  } catch (reason) { error.value = reason instanceof Error ? reason.message : String(reason) }
  finally { busy.value = false }
}
async function insertPivotChart() {
  if (!selectedId.value || busy.value) return
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
  if (!selectedId.value || busy.value) return
  error.value = ''; status.value = ''; busy.value = true
  try {
    const { book } = requireApi(), pivot = book.getPivotTables().find(candidate => candidate.getPivotTableId() === selectedId.value)
    if (!pivot) throw new Error('数据透视表已被删除。')
    await pivot.remove(); loadDraft(); status.value = '数据透视表已删除，可撤销恢复。'
  } catch (reason) { error.value = reason instanceof Error ? reason.message : String(reason) }
  finally { busy.value = false }
}
function navigateHistory(direction: 'undo' | 'redo') {
  try {
    const { book } = requireApi(), id = selectedId.value
    book[direction]()
    loadDraft(workbook.value?.pivotTables?.some(pivot => pivot.id === id) ? id : '')
    status.value = direction === 'undo' ? '已撤销上一步工作簿操作。' : '已重做工作簿操作。'
  } catch (reason) { error.value = reason instanceof Error ? reason.message : String(reason) }
}
function onPivotSelect(value: string | number) { loadDraft(String(value)) }
function close() { if (!busy.value) { filterFieldId.value = ''; settingsFieldId.value = ''; emit('update:open', false) } }
</script>

<template>
  <Modal :open="open" title="数据透视表" :width="820" :footer="null" :mask-closable="!busy" @cancel="close">
    <div class="pivot-panel" data-testid="pivot-table-panel">
      <div class="pivot-panel-selector"><label>选择数据透视表<Select :value="selectedId" :options="pivotOptions" aria-label="选择数据透视表" @change="onPivotSelect" /></label></div>
      <template v-if="draft">
        <section class="pivot-locations" aria-label="数据透视表数据源和位置">
          <label>数据源工作表<Select v-model:value="draft.sourceSheetId" :options="sheetOptions" aria-label="透视表数据源工作表" /></label>
          <label>数据区域<Input v-model:value="draft.sourceRange" aria-label="透视表数据区域" placeholder="例如 A1:C20" /></label>
          <label>放置工作表<Select v-model:value="draft.targetSheetId" :options="sheetOptions" aria-label="透视表放置工作表" /></label>
          <label>起始单元格<Input v-model:value="draft.targetAddress" aria-label="透视表起始单元格" placeholder="例如 E1" /></label>
        </section>
        <div class="pivot-fields-header"><span>字段</span><Button size="small" @click="loadFields">读取字段</Button></div>
        <p v-if="fieldsStale" class="pivot-hint">数据区域已改变，请读取字段后重新配置。</p>
        <section class="pivot-area-layout" aria-label="透视字段配置">
          <div class="pivot-source-area" aria-label="未放置字段">
            <span class="pivot-area-title">源字段</span>
            <p class="pivot-source-description">选择字段并将其添加到透视表区域</p>
            <Input v-model:value="fieldSearch" aria-label="查找透视字段" placeholder="查找字段" size="small" />
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
              <div v-for="field in area.fields" :key="field.fieldId" class="pivot-field" :class="{ 'is-dragging': draggedFieldId === field.fieldId }" :data-field-id="field.fieldId" draggable="true" @dragover.prevent @drop="dropField(area.area, $event, field.fieldId)" @dragstart="startFieldDrag(field.fieldId, $event)" @dragend="draggedFieldId = ''">
                <GripVertical class="pivot-field-grip" :size="16" aria-hidden="true" /><span class="pivot-field-name" :title="field.displayName ?? field.fieldName">{{ field.displayName ?? field.fieldName }}</span>
                  <Dropdown :trigger="['click']" placement="bottomRight" :menu="fieldMenu(field)"><Button type="text" size="small" :aria-label="`字段 ${field.fieldName} 操作`"><Ellipsis :size="16" /></Button></Dropdown>
                </div>
                <p v-if="!area.fields.length" class="pivot-empty">拖动字段到此处</p>
              </div>
            </section>
          </div>
        </section>
        <section class="pivot-options" aria-label="透视表汇总选项">
          <label><Switch v-model:checked="draft.options.showRowGrandTotal" size="small" aria-label="显示行总计" />行总计</label>
          <label><Switch v-model:checked="draft.options.showColGrandTotal" size="small" aria-label="显示列总计" />列总计</label>
          <label><Switch v-model:checked="draft.options.showRowSubtotals" size="small" aria-label="显示行小计" />行小计</label>
          <label><Switch v-model:checked="draft.options.showColSubtotals" size="small" aria-label="显示列小计" />列小计</label>
          <label><Switch v-model:checked="draft.options.repeatRowLabels" size="small" aria-label="重复行标签" />重复行标签</label>
          <label>值排列<Select v-model:value="draft.options.valuePosition" :options="[{ label: '按列', value: 'column' }, { label: '按行', value: 'row' }]" aria-label="值字段排列" size="small" /></label>
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
    </div>
  </Modal>
  <Modal :open="Boolean(filterFieldId)" :title="null" :width="336" :footer="null" :closable="false" :mask="false" :mask-closable="false" :confirm-loading="fieldDialogBusy" destroy-on-hidden @cancel="filterFieldId = ''">
    <div v-if="filterField" class="pivot-filter-dialog" :aria-label="`筛选字段 ${filterField.fieldName}`">
      <div v-if="filterField.area === 'row' || filterField.area === 'column'" class="pivot-sort-segment" role="group" aria-label="透视字段排序">
        <Button type="text" :disabled="fieldDialogBusy" :class="{ 'is-active': selectedPivot?.fields.find(field => field.id === filterField.fieldId)?.sort === 'ascending' }" :aria-pressed="selectedPivot?.fields.find(field => field.id === filterField.fieldId)?.sort === 'ascending'" @click="applyLabelSort('ascending')"><ArrowUpAZ :size="16" />升序</Button>
        <Button type="text" :disabled="fieldDialogBusy" :class="{ 'is-active': selectedPivot?.fields.find(field => field.id === filterField.fieldId)?.sort === 'descending' }" :aria-pressed="selectedPivot?.fields.find(field => field.id === filterField.fieldId)?.sort === 'descending'" @click="applyLabelSort('descending')"><ArrowDownAZ :size="16" />降序</Button>
      </div>
      <div class="pivot-filter-items">
        <label class="pivot-filter-checkbox pivot-filter-select-all"><Checkbox :checked="allFilterItemsSelected" :indeterminate="filterSelection.length > 0 && !allFilterItemsSelected" @change="toggleAllFilterItems(!allFilterItemsSelected)">全选</Checkbox></label>
        <label v-for="item in filterItems" :key="item" class="pivot-filter-checkbox"><Checkbox :checked="filterSelection.includes(item)" @change="toggleFilterItem(item)">{{ item || '（空白）' }}</Checkbox></label>
      </div>
      <div class="pivot-filter-footer">
        <Button type="link" size="small" :disabled="!canClearFieldFilter || fieldDialogBusy" @click="applyLabelFilter(null)">清除筛选</Button>
        <span class="pivot-filter-footer-spacer" />
        <Button size="small" :disabled="fieldDialogBusy" @click="filterFieldId = ''">取消</Button>
        <Button type="primary" size="small" :disabled="fieldDialogBusy" @click="applyLabelFilter(allFilterItemsSelected ? null : filterSelection)">确定</Button>
      </div>
      <p v-if="error" class="pivot-error" role="alert">{{ error }}</p>
    </div>
  </Modal>
  <Modal :open="Boolean(settingsFieldId)" title="字段设置" :width="400" :footer="null" :mask="false" :mask-closable="false" :confirm-loading="fieldDialogBusy" destroy-on-hidden @cancel="settingsFieldId = ''">
    <div v-if="settingsField" class="pivot-settings-dialog">
      <label class="pivot-settings-name"><span>字段名称</span><span class="pivot-settings-source">源字段：{{ settingsField.fieldName }}</span><Input v-model:value="fieldDisplayName" size="small" /></label>
      <label><span>汇总方式</span><div class="pivot-aggregation-list" role="listbox" aria-label="汇总方式"><Button v-for="option in aggregationOptions" :key="option.value" type="text" role="option" :aria-selected="fieldAggregation === option.value" :disabled="fieldDialogBusy" :class="{ 'is-selected': fieldAggregation === option.value }" @click="fieldAggregation = option.value"><span>{{ option.label }}</span><Check v-if="fieldAggregation === option.value" :size="16" /></Button></div></label>
      <Button :disabled="fieldDialogBusy" @click="numberFormatDraft = fieldFormat || 'General'; numberFormatOpen = true">数字格式</Button>
      <p v-if="error" class="pivot-error" role="alert">{{ error }}</p>
      <div class="pivot-settings-footer"><Button :disabled="fieldDialogBusy" @click="settingsFieldId = ''">取消</Button><Button type="primary" :disabled="!fieldDisplayName.trim() || fieldDialogBusy" @click="applyFieldSettings">确定</Button></div>
    </div>
  </Modal>
  <Modal :open="numberFormatOpen" title="数字格式" :width="400" :footer="null" :mask="false" destroy-on-hidden @cancel="numberFormatOpen = false">
    <div class="pivot-settings-dialog"><label>格式类型<Select v-model:value="numberFormatDraft" :options="numberFormatOptions" aria-label="数字格式类型" /></label><label>格式代码<Input v-model:value="numberFormatDraft" aria-label="数字格式代码" /></label><output class="pivot-format-preview" aria-label="数字格式预览">{{ numberFormatPreview }}</output><div class="pivot-settings-footer"><Button @click="numberFormatOpen = false">取消</Button><Button type="primary" @click="fieldFormat = numberFormatDraft; numberFormatOpen = false">确定</Button></div></div>
  </Modal>
  <PivotFieldToolsDialog v-if="toolsField && workbook && selectedPivot" :open="Boolean(toolsFieldId)" :kind="toolsKind" :workbook="workbook" :pivot="selectedPivot" :field="toolsField" :busy="fieldDialogBusy" :error="error" @close="toolsFieldId = ''" @apply="applyFieldTool" />
</template>

<style scoped>
.pivot-panel { display: flex; flex-direction: column; gap: 12px; min-width: 0; }
.pivot-panel label { display: flex; flex-direction: column; gap: 5px; min-width: 0; color: #5d6979; font-size: 12px; }
.pivot-panel-selector { max-width: 100%; }
.pivot-locations { display: grid; grid-template-columns: 1fr 1fr; gap: 10px 16px; }
.pivot-fields-header { display: flex; align-items: center; justify-content: space-between; padding-top: 4px; color: #354152; font-weight: 600; }
.pivot-area-layout { display: grid; grid-template-columns: minmax(180px, 0.8fr) minmax(0, 1.8fr); gap: 12px; min-height: 200px; }
.pivot-area-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px 16px; }
.pivot-area, .pivot-source-area { box-sizing: border-box; min-height: 200px; padding: 4px; }
.pivot-area-title { display: block; margin: 0 0 8px; color: #354152; font-size: 13px; font-weight: 500; line-height: 20px; }
.pivot-source-description { margin: -4px 0 8px; color: #7b8797; font-size: 12px; line-height: 16px; }
.pivot-source-area :deep(.ant-input-affix-wrapper) { margin-bottom: 8px; }
.pivot-field-list { box-sizing: border-box; display: flex; flex-direction: column; gap: 4px; height: 161px; overflow-y: auto; padding: 8px; border: 1px solid #e7eaf0; border-radius: 6px; background: #f6f7f9; }
.pivot-source-field { display: flex; align-items: center; gap: 8px; min-height: 32px; padding: 2px 4px; border-radius: 4px; cursor: grab; }
.pivot-field { display: grid; grid-template-columns: 16px minmax(0, 1fr) auto; gap: 5px; align-items: center; min-height: 28px; padding: 2px 4px; border: 1px solid transparent; border-radius: 4px; background: #fff; cursor: grab; }
.pivot-field.is-dragging { opacity: 0.55; }
.pivot-field-grip { color: #8490a0; }
.pivot-field-name { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-size: 12px; color: #2c3748; }
.pivot-field :deep(.ant-btn) { padding: 0; }
.pivot-empty { margin: auto; color: #8993a1; font-size: 12px; text-align: center; }
.pivot-options { display: flex; align-items: center; gap: 10px 16px; flex-wrap: wrap; padding: 4px 0; }
.pivot-options label { flex-direction: row; align-items: center; }
.pivot-options :deep(.ant-select) { min-width: 76px; }
.pivot-hint, .pivot-status, .pivot-error { margin: 0; font-size: 12px; line-height: 1.6; }
.pivot-hint { color: #7b8797; }
.pivot-error { color: #b12e35; background: #fff4f3; padding: 8px 10px; border-radius: 4px; }
.pivot-status { color: #267342; }
.pivot-actions { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; padding-top: 12px; border-top: 1px solid #e6ebf1; }
.pivot-actions-spacer { flex: 1; }
.pivot-filter-dialog { display: flex; flex-direction: column; }
.pivot-sort-segment { display: flex; width: 100%; }
.pivot-sort-segment button { box-sizing: border-box; display: flex; flex: 1 1 140px; align-items: center; justify-content: center; gap: 6px; height: 32px; border: 1px solid #e2e6ed; background: transparent; color: #354152; font-size: 14px; cursor: pointer; }
.pivot-sort-segment button:first-child { border-radius: 4px 0 0 4px; }
.pivot-sort-segment button:last-child { margin-left: -1px; border-radius: 0 4px 4px 0; }
.pivot-sort-segment button:hover, .pivot-sort-segment button.is-active { background: #f0f2f5; }
.pivot-filter-items { box-sizing: border-box; display: flex; flex-direction: column; height: 300px; overflow-y: auto; margin: 12px 0 16px; padding: 8px; border: 1px solid #e4e7ec; border-radius: 6px; }
.pivot-filter-checkbox { box-sizing: border-box; display: flex !important; align-items: center; height: 28px; min-height: 28px; margin: 0; padding: 2px 4px; overflow: hidden; color: #303b4a !important; font-size: 14px !important; white-space: nowrap; }
.pivot-filter-checkbox :deep(.ant-checkbox-wrapper) { max-width: 100%; overflow: hidden; text-overflow: ellipsis; }
.pivot-filter-select-all { border-bottom: 1px solid #e9ecf1; margin-bottom: 4px; }
.pivot-filter-footer, .pivot-settings-footer { display: flex; align-items: center; gap: 8px; }
.pivot-filter-footer-spacer { flex: 1; }
.pivot-settings-dialog { display: flex; flex-direction: column; gap: 16px; }
.pivot-settings-dialog label { display: flex; flex-direction: column; gap: 6px; color: #354152; font-size: 14px; }
.pivot-settings-name { display: grid !important; grid-template-columns: 1fr auto; align-items: center; }
.pivot-settings-name :deep(.ant-input-affix-wrapper) { grid-column: 1 / -1; }
.pivot-settings-source { color: #7b8797; font-size: 12px; }
.pivot-aggregation-list { display: flex; flex-direction: column; height: 160px; overflow-y: auto; padding: 4px; border: 1px solid #e4e7ec; border-radius: 4px; }
.pivot-aggregation-list :deep(.ant-btn) { display: flex; flex: 0 0 28px; justify-content: space-between; text-align: left; padding: 2px 8px; }
.pivot-aggregation-list :deep(.is-selected) { color: #1677ff; background: #e6f4ff; }
.pivot-format-preview { padding: 12px; border: 1px solid #e4e7ec; border-radius: 4px; text-align: right; }
.pivot-settings-footer { justify-content: flex-end; padding-top: 12px; }
@media (max-width: 620px) { .pivot-locations, .pivot-area-layout { grid-template-columns: 1fr; } .pivot-area-grid { grid-template-columns: 1fr; } .pivot-field { gap: 3px; } .pivot-actions-spacer { display: none; } }
</style>
