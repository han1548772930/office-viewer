<script setup lang="ts">
import { Button, Input, Modal, Select, Switch } from 'antdv-next'
import { computed, ref, watch } from 'vue'
import { FOffice } from '../../api/facade'
import type { OfficeStudio } from '../../api/office-studio'
import type { CellRange } from '../../domain/model'
import {
  CONFIGURE_PIVOT_COMMAND, createPivotPanelDraft, pivotRangeAddress, preparePivotPanelInput,
  readPivotPanelSourceFields, type PivotPanelDraft,
} from '../../features/pivot/services/pivot-panel-state'

const props = defineProps<{ open: boolean; studio?: OfficeStudio; workbookId?: string; sheetId: string; sourceRange: string }>()
interface PivotPanelChange { pivotId: string; sheetId: string; outputRange?: CellRange }
const emit = defineEmits<{
  'update:open': [open: boolean]
  created: [result: PivotPanelChange]
  updated: [result: PivotPanelChange]
}>()
const workbook = computed(() => props.studio?.getWorkbookModel(props.workbookId ?? props.studio.activeWorkbookId))
const draft = ref<PivotPanelDraft | null>(null), selectedId = ref(''), error = ref(''), status = ref(''), busy = ref(false)
const loadedSource = ref(''), fieldPage = ref(0), fieldSearch = ref('')
const areaOptions = [{ label: '未使用', value: 'unused' }, { label: '筛选', value: 'filter' }, { label: '行', value: 'row' }, { label: '列', value: 'column' }, { label: '值', value: 'value' }]
const aggregationOptions = [{ label: '求和', value: 'sum' }, { label: '计数', value: 'count' }, { label: '平均值', value: 'average' }, { label: '最小值', value: 'min' }, { label: '最大值', value: 'max' }]
const sheetOptions = computed(() => (workbook.value?.sheets ?? []).filter(sheet => sheet.visible !== false).map(sheet => ({ label: sheet.name, value: sheet.id })))
const pivotOptions = computed(() => [{ label: '新建数据透视表', value: '' }, ...(workbook.value?.pivotTables ?? []).map((pivot, index) => ({ label: `数据透视表 ${index + 1} · ${workbook.value?.sheets.find(sheet => sheet.id === pivot.anchor.subUnitId)?.name ?? ''}!${pivotRangeAddress({ startRow: pivot.anchor.row, endRow: pivot.anchor.row, startCol: pivot.anchor.col, endCol: pivot.anchor.col })}`, value: pivot.id }))])
const selectedPivot = computed(() => workbook.value?.pivotTables?.find(pivot => pivot.id === selectedId.value))
const sourceKey = computed(() => draft.value ? `${draft.value.sourceSheetId}\n${draft.value.sourceRange.trim().toUpperCase()}` : '')
const fieldsStale = computed(() => sourceKey.value !== loadedSource.value)
const filteredFields = computed(() => (draft.value?.fields ?? []).filter(field => field.fieldName.toLocaleLowerCase().includes(fieldSearch.value.toLocaleLowerCase())))
const fieldPageCount = computed(() => Math.max(1, Math.ceil(filteredFields.value.length / 12)))
const visibleFields = computed(() => filteredFields.value.slice(fieldPage.value * 12, (fieldPage.value + 1) * 12))
const fieldSummary = computed(() => (['filter', 'row', 'column', 'value'] as const).map(area => ({ area, label: areaOptions.find(option => option.value === area)?.label, names: draft.value?.fields.filter(field => field.area === area).map(field => field.fieldName).join('、') || '未设置' })))
const history = computed(() => props.studio?.editor.history.units[workbook.value?.id ?? ''])

function loadDraft(pivotId = '') {
  const model = workbook.value
  error.value = ''; status.value = ''; selectedId.value = pivotId; fieldPage.value = 0; fieldSearch.value = ''
  if (!model) { draft.value = null; error.value = '当前工作台尚未连接工作簿命令接口。'; return }
  draft.value = createPivotPanelDraft(model, props.sheetId, props.sourceRange, pivotId)
  loadedSource.value = sourceKey.value
}
watch(() => props.open, open => { if (open) loadDraft() }, { immediate: true })
watch(fieldSearch, () => { fieldPage.value = 0 })

function loadFields() {
  if (!draft.value || !workbook.value) return
  try {
    const previous = draft.value.fields
    const next = readPivotPanelSourceFields(workbook.value, draft.value.sourceSheetId, draft.value.sourceRange)
    draft.value.fields = fieldsStale.value ? next : next.map(field => ({ ...field, ...previous.find(candidate => candidate.fieldId === field.fieldId), fieldName: field.fieldName }))
    loadedSource.value = sourceKey.value; fieldPage.value = 0; error.value = ''; status.value = `已读取 ${next.length} 个字段，请选择行、列和汇总值。`
  } catch (reason) { error.value = reason instanceof Error ? reason.message : String(reason) }
}
function moveField(fieldId: string, direction: -1 | 1) {
  const fields = draft.value?.fields
  if (!fields) return
  const index = fields.findIndex(field => field.fieldId === fieldId), other = index + direction
  if (index < 0 || other < 0 || other >= fields.length) return
  const value = fields[index]!
  fields.splice(index, 1); fields.splice(other, 0, value)
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
function close() { if (!busy.value) emit('update:open', false) }
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
        <div class="pivot-fields-header"><span>字段（{{ draft.fields.length }}）</span><Button size="small" @click="loadFields">读取字段</Button></div>
        <p v-if="fieldsStale" class="pivot-hint">数据区域已改变，请读取字段后重新配置。</p>
        <Input v-if="draft.fields.length > 12" v-model:value="fieldSearch" aria-label="查找透视字段" placeholder="查找字段" size="small" />
        <section class="pivot-fields" aria-label="透视字段配置">
          <div v-for="field in visibleFields" :key="field.fieldId" class="pivot-field" :data-field-id="field.fieldId">
            <span class="pivot-field-name" :title="field.fieldName">{{ field.fieldName }}</span>
            <Select v-model:value="field.area" :options="areaOptions" :aria-label="`字段 ${field.fieldName} 区域`" size="small" />
            <Select v-if="field.area === 'value'" v-model:value="field.aggregation" :options="aggregationOptions" :aria-label="`字段 ${field.fieldName} 汇总方式`" size="small" />
            <span v-else class="pivot-field-spacer" />
            <Button type="text" size="small" :aria-label="`上移字段 ${field.fieldName}`" :disabled="draft.fields[0]?.fieldId === field.fieldId" @click="moveField(field.fieldId, -1)">↑</Button>
            <Button type="text" size="small" :aria-label="`下移字段 ${field.fieldName}`" :disabled="draft.fields.at(-1)?.fieldId === field.fieldId" @click="moveField(field.fieldId, 1)">↓</Button>
          </div>
          <p v-if="!draft.fields.length" class="pivot-hint">数据区域首行作为字段标题。填写包含标题和数据的区域后，读取字段。</p>
        </section>
        <div v-if="fieldPageCount > 1" class="pivot-pagination"><Button size="small" :disabled="fieldPage === 0" @click="fieldPage--">上一页字段</Button><span>{{ fieldPage + 1 }} / {{ fieldPageCount }}</span><Button size="small" :disabled="fieldPage >= fieldPageCount - 1" @click="fieldPage++">下一页字段</Button></div>
        <div class="pivot-summary"><span v-for="area in fieldSummary" :key="area.area"><strong>{{ area.label }}：</strong>{{ area.names }}</span></div>
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
</template>

<style scoped>
.pivot-panel { display: flex; flex-direction: column; gap: 12px; min-width: 0; }
.pivot-panel label { display: flex; flex-direction: column; gap: 5px; min-width: 0; color: #5d6979; font-size: 12px; }
.pivot-panel-selector { max-width: 100%; }
.pivot-locations { display: grid; grid-template-columns: 1fr 1fr; gap: 10px 16px; }
.pivot-fields-header { display: flex; align-items: center; justify-content: space-between; padding-top: 4px; color: #354152; font-weight: 600; }
.pivot-fields { display: flex; flex-direction: column; gap: 6px; max-height: 360px; overflow: auto; }
.pivot-field { display: grid; grid-template-columns: minmax(80px, 1fr) 104px 114px 26px 26px; gap: 6px; align-items: center; padding: 4px 6px; border: 1px solid #edf0f4; border-radius: 4px; }
.pivot-field-name { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-size: 12px; color: #2c3748; }
.pivot-field :deep(.ant-btn) { padding: 0; }
.pivot-summary { display: grid; grid-template-columns: 1fr 1fr; gap: 6px 12px; padding: 8px 10px; border-radius: 4px; background: #f6f8fb; font-size: 12px; line-height: 1.6; overflow-wrap: anywhere; }
.pivot-summary strong { color: #47576b; font-weight: 600; }
.pivot-options { display: flex; align-items: center; gap: 10px 16px; flex-wrap: wrap; padding: 4px 0; }
.pivot-options label { flex-direction: row; align-items: center; }
.pivot-options :deep(.ant-select) { min-width: 76px; }
.pivot-hint, .pivot-status, .pivot-error { margin: 0; font-size: 12px; line-height: 1.6; }
.pivot-hint { color: #7b8797; }
.pivot-error { color: #b12e35; background: #fff4f3; padding: 8px 10px; border-radius: 4px; }
.pivot-status { color: #267342; }
.pivot-pagination, .pivot-actions { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
.pivot-pagination { justify-content: flex-end; font-size: 12px; color: #7b8797; }
.pivot-actions { padding-top: 12px; border-top: 1px solid #e6ebf1; }
.pivot-actions-spacer { flex: 1; }
@media (max-width: 620px) { .pivot-locations, .pivot-summary { grid-template-columns: 1fr; } .pivot-field { grid-template-columns: minmax(60px, 1fr) 88px 90px 24px 24px; gap: 3px; } .pivot-actions-spacer { display: none; } }
</style>
