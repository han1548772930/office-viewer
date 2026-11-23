<script setup lang="ts">
import { computed, inject, ref, watch } from 'vue'
import { Alert, Button, Form, FormItem, Input, Modal, Select, theme } from 'antdv-next'
import { FOffice } from '../../api/facade'
import type { OfficeStudio } from '../../api/office-studio'
import { createPivotChartDraft, pivotChartConfig, pivotChartPatch, type PivotChartDraft } from '../../features/pivot/services/pivot-chart-panel-state'
import { readPivotChartSource } from '../../features/charts/chart-pivot-data'
import type { ChartData } from '../../features/charts/chart-engine'
import ChartSettingsFields from '../../features/charts/views/ChartSettingsFields.vue'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

const props = defineProps<{ open: boolean; studio: OfficeStudio; workbookId: string; sheetId: string; pivotTableId?: string; chartId?: string }>()
const emit = defineEmits<{ 'update:open': [value: boolean]; applied: [chartId: string]; closed: [] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const { token } = theme.useToken()
const draft = ref<PivotChartDraft>(createPivotChartDraft())
const error = ref(''), status = ref(''), busy = ref(false)
const workbook = computed(() => props.studio.getWorkbookModel(props.workbookId))
const pivotTables = computed(() => workbook.value?.pivotTables ?? [])
const pivotOptions = computed(() => pivotTables.value.map((pivot, index) => ({
  label: `${t('sheet.pivot-chart.title', '透视表')} ${index + 1} · ${workbook.value?.sheets.find(sheet => sheet.id === pivot.anchor.subUnitId)?.name ?? ''}`,
  value: pivot.id,
})))
const chart = computed(() => workbook.value?.charts?.find(item => item.id === draft.value.chartId))
const linkedCharts = computed(() => workbook.value?.charts?.filter(item => item.pivotTableId === draft.value.pivotTableId) ?? [])
const preview = computed<{ data?: ChartData; error?: string }>(() => {
  if (!props.open || !workbook.value || !draft.value.pivotTableId) return {}
  try {
    const source = readPivotChartSource(workbook.value, draft.value.pivotTableId)
    return { data: { fields: source.fields, records: source.records,
      categories: source.records.map(record => source.categoryFields.map(field => String(record[field] ?? '')).join(' / ')),
      series: source.valueFields.map(field => ({ field, name: source.fields[field]!, values: source.records.map(record => { const value = record[field]; return typeof value === 'number' ? value : null }) })) } }
  } catch (reason) { return { error: reason instanceof Error ? reason.message : String(reason) } }
})

function load(): void {
  error.value = ''; status.value = ''
  const current = props.chartId ? workbook.value?.charts?.find(item => item.id === props.chartId) : undefined
  draft.value = createPivotChartDraft(current, props.pivotTableId ?? pivotTables.value[0]?.id ?? '')
}
watch(() => [props.open, props.chartId, props.pivotTableId, props.workbookId, props.sheetId, props.studio] as const, ([open]) => { if (open) load() }, { immediate: true })

function chooseChart(value: unknown): void {
  if (typeof value !== 'string' || busy.value) return
  const selected = workbook.value?.charts?.find(item => item.id === value)
  if (value && !selected) { error.value = '所选透视图已被删除。'; return }
  draft.value = createPivotChartDraft(selected, draft.value.pivotTableId)
  error.value = ''; status.value = ''
}
function choosePivot(value: unknown): void {
  if (typeof value !== 'string' || busy.value) return
  const current = draft.value
  if (current.pivotTableId === value) return
  current.pivotTableId = value
  current.settings.config.pivotTableId = value
  current.settings.config.valueFields = []
  current.settings.config.categoryField = 0
  current.settings.config.seriesToCategory = {}
  // Field roles belong to the selected pivot output. Keep authored appearance,
  // while the user chooses valid roles for the new source before applying.
  if (current.settings.config.options) current.settings.config.options.mapping = undefined
  error.value = ''; status.value = ''
}
function requireWorkbook(studio: OfficeStudio, unitId: string) {
  const book = FOffice.newAPI(studio).getWorkbook(unitId)
  if (!book) throw new Error(t('sheet.pivot-chart.error.workbook-closed', '工作簿已关闭，请重新打开透视图面板。'))
  return book
}

async function apply(): Promise<void> {
  if (busy.value) return
  const current = draft.value, studio = props.studio, unitId = props.workbookId, hostSheetId = chart.value?.sheetId ?? props.sheetId
  busy.value = true; error.value = ''; status.value = ''
  try {
    const book = requireWorkbook(studio, unitId)
    if (!book.getPivotTables().some(item => item.getPivotTableId() === current.pivotTableId)) throw new Error(t('sheet.pivot-chart.error.pivot-missing', '所选透视表不存在。'))
    const targetSheet = book.getSheetBySheetId(hostSheetId)
    if (!targetSheet) throw new Error('放置图表的工作表已被删除。')
    let chartId = current.chartId
    const updating = Boolean(chartId)
    if (chartId) {
      const existing = targetSheet.getChartById(chartId)
      if (!existing) throw new Error(t('sheet.pivot-chart.error.chart-deleted', '所选透视图已被删除。'))
      await existing.update(pivotChartPatch(current))
    } else {
      const created = await targetSheet.insertPivotChart(pivotChartConfig(current))
      chartId = created.getChartId()
    }
    if (draft.value === current && props.studio === studio && props.workbookId === unitId) {
      const saved = studio.getWorkbookModel(unitId)?.charts?.find(item => item.id === chartId)
      if (saved) draft.value = createPivotChartDraft(saved)
      status.value = t(updating ? 'sheet.pivot-chart.status.updated' : 'sheet.pivot-chart.status.created', updating ? '透视图设置已更新。' : '已创建并关联透视图。')
    }
    emit('applied', chartId)
  } catch (reason) { if (draft.value === current) error.value = reason instanceof Error ? reason.message : String(reason) }
  finally { busy.value = false }
}

async function remove(): Promise<void> {
  if (!draft.value.chartId || busy.value) return
  const current = draft.value, studio = props.studio, unitId = props.workbookId, hostSheetId = chart.value?.sheetId ?? props.sheetId
  busy.value = true; error.value = ''; status.value = ''
  try {
    const target = requireWorkbook(studio, unitId).getSheetBySheetId(hostSheetId)?.getChartById(current.chartId)
    if (!target) throw new Error(t('sheet.pivot-chart.error.chart-deleted', '所选透视图已被删除。'))
    await target.remove()
    if (draft.value === current && props.studio === studio && props.workbookId === unitId) {
      draft.value = createPivotChartDraft(undefined, current.pivotTableId)
      status.value = t('sheet.pivot-chart.status.removed', '透视图已移除，关联透视表仍保留。')
    }
    emit('applied', current.chartId)
  } catch (reason) { if (draft.value === current) error.value = reason instanceof Error ? reason.message : String(reason) }
  finally { busy.value = false }
}
function close(): void { if (!busy.value) emit('update:open', false) }
function opened(value: boolean): void { if (!value) emit('closed') }
</script>

<template>
  <Modal :open="open" :title="t('sheet.pivot-chart.title', '透视图')" :width="760" :footer="null" :mask-closable="!busy" @cancel="close" @after-open-change="opened">
    <Form layout="vertical" class="pivot-chart-panel" data-testid="pivot-chart-panel" @submit.prevent="apply">
      <FormItem :label="t('sheet.pivot-chart.related-pivot', '关联透视表')"><Select :value="draft.pivotTableId" :options="pivotOptions" :aria-label="t('sheet.pivot-chart.related-pivot', '关联透视表')" :disabled="busy" @update:value="choosePivot" /></FormItem>
      <FormItem v-if="linkedCharts.length" :label="t('sheet.pivot-chart.edit-existing', '编辑已有透视图')"><Select :value="draft.chartId" :options="[{ label: t('sheet.pivot-chart.create-new', '新建透视图'), value: '' }, ...linkedCharts.map((item, index) => ({ label: item.title || `${t('sheet.pivot-chart.title', '透视图')} ${index + 1}`, value: item.id }))]" :aria-label="t('sheet.pivot-chart.edit-existing', '选择透视图')" :disabled="busy" @update:value="chooseChart" /></FormItem>
      <FormItem :label="t('sheet.pivot-chart.placement', '放置单元格')"><Input v-model:value="draft.position" :aria-label="t('sheet.pivot-chart.placement', '透视图放置单元格')" placeholder="F2" :disabled="busy" /></FormItem>
      <ChartSettingsFields v-model:draft="draft.settings" :data="preview.data" :host-sheet-id="chart?.sheetId ?? sheetId" locked-source :disabled="busy" />
      <Alert v-if="error || preview.error" type="error" :message="error || preview.error" show-icon />
      <Alert v-if="status" type="success" :message="status" show-icon role="status" />
      <footer :style="{ borderColor: token.colorBorderSecondary }">
        <Button v-if="draft.chartId" danger :disabled="busy" @click="remove">{{ t('sheet.pivot-chart.remove', '移除透视图') }}</Button>
        <span />
        <Button :disabled="busy" @click="close">{{ t('sheet.pivot-chart.close', '关闭') }}</Button>
        <Button type="primary" :loading="busy" :disabled="!pivotTables.length" @click="apply">{{ t(draft.chartId ? 'sheet.pivot-chart.update' : 'sheet.pivot-chart.create', draft.chartId ? '更新透视图' : '创建透视图') }}</Button>
      </footer>
    </Form>
  </Modal>
</template>

<style scoped>
.pivot-chart-panel { display: grid; gap: 14px; max-height: 70vh; min-width: 0; overflow-y: auto; padding-right: 8px; }
.pivot-chart-panel :deep(.ant-form-item) { min-width: 0; margin-bottom: 0; }
.pivot-chart-panel footer { display: flex; align-items: center; gap: 8px; padding-top: 12px; border-top: 1px solid; }
.pivot-chart-panel footer span { flex: 1; }
@media (max-width: 520px) { .pivot-chart-panel footer { flex-wrap: wrap; } }
</style>
