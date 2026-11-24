<script setup lang="ts">
import { computed, inject, onScopeDispose, ref, watch } from 'vue'
import { Alert, Button, Form, FormItem, Input, Modal, Select, theme } from 'antdv-next'
import { FOffice } from '../../api/facade'
import type { OfficeStudio } from '../../api/office-studio'
import { createPivotChartDraft, pivotChartConfig, pivotChartPatch, type PivotChartDraft } from '../../features/pivot/services/pivot-chart-panel-state'
import { chartPlacementGeometry } from '../../features/charts/chart-placement'
import type { ChartSettingsSourceLoader, ChartSettingsStatisticsLoader } from '../../features/charts/chart-settings-data'
import { useChartSettingsSource } from '../../features/charts/views/use-chart-settings-source'
import { canViewWorksheetProtection } from '../../features/sheet-protection/range-protection'
import ChartSettingsFields from '../../features/charts/views/ChartSettingsFields.vue'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

const props = defineProps<{ open: boolean; studio: OfficeStudio; workbookId: string; sheetId: string; pivotTableId?: string; chartId?: string }>()
const emit = defineEmits<{ 'update:open': [value: boolean]; 'busy-change': [value: boolean]; applied: [chartId: string]; closed: [] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const { token } = theme.useToken()
const draft = ref<PivotChartDraft>(createPivotChartDraft())
const error = ref(''), status = ref(''), busy = ref(false)
const modelRevision = ref(0)
let disposed = false
const workbook = computed(() => { modelRevision.value; return props.studio.getWorkbookModel(props.workbookId) })
const pivotTables = computed(() => (workbook.value?.pivotTables ?? []).filter(pivot => [pivot.source.subUnitId, pivot.anchor.subUnitId].every(id => {
  const sheet = workbook.value?.sheets.find(item => item.id === id)
  return sheet && canViewWorksheetProtection(sheet, props.studio.currentUserId)
})))
const pivotOptions = computed(() => pivotTables.value.map((pivot, index) => ({
  label: `${t('sheet.pivot-chart.title', '透视表')} ${index + 1} · ${workbook.value?.sheets.find(sheet => sheet.id === pivot.anchor.subUnitId)?.name ?? ''}`,
  value: pivot.id,
})))
const chart = computed(() => workbook.value?.charts?.find(item => item.id === draft.value.chartId))
const linkedCharts = computed(() => workbook.value?.charts?.filter(item => {
  const sheet = workbook.value?.sheets.find(sheet => sheet.id === item.sheetId)
  return item.pivotTableId === draft.value.pivotTableId && sheet && canViewWorksheetProtection(sheet, props.studio.currentUserId)
}) ?? [])
const sourceHostSheetId = computed(() => chart.value?.sheetId ?? props.sheetId)
const loadSource = computed<ChartSettingsSourceLoader>(() => {
  const studio = props.studio, unitId = props.workbookId, hostSheetId = sourceHostSheetId.value
  return async (config, options) => {
    const sheet = requireWorkbook(studio, unitId).getSheetBySheetId(hostSheetId)
    if (!sheet) throw new Error('放置图表的工作表已被删除。')
    return sheet.getChartSettingsData(config, options)
  }
})
const loadStatistics = computed<ChartSettingsStatisticsLoader>(() => {
  const studio = props.studio, unitId = props.workbookId, hostSheetId = sourceHostSheetId.value
  return async (config, options) => {
    const sheet = requireWorkbook(studio, unitId).getSheetBySheetId(hostSheetId)
    if (!sheet) throw new Error('放置图表的工作表已被删除。')
    return sheet.getChartSettingsStatistics(config, options)
  }
})
const { projection: preview, loading: sourceLoading, assertReady, reload } = useChartSettingsSource({
  active: () => props.open && Boolean(draft.value.pivotTableId), draft: () => draft.value.settings,
  hostSheetId: () => sourceHostSheetId.value, date1904: () => workbook.value?.date1904,
  sourceRevision: () => modelRevision.value, load: () => loadSource.value, loadStatistics: () => loadStatistics.value,
})

function load(): void {
  modelRevision.value += 1
  error.value = ''; status.value = ''
  const current = props.chartId ? workbook.value?.charts?.find(item => item.id === props.chartId) : undefined
  try {
    draft.value = createPivotChartDraft(current, props.pivotTableId ?? pivotTables.value[0]?.id ?? '',
      chartPlacementGeometry(workbook.value, current?.sheetId ?? props.sheetId))
  } catch (reason) { error.value = reason instanceof Error ? reason.message : String(reason) }
}
watch([() => props.open, () => props.chartId, () => props.pivotTableId, () => props.workbookId, () => props.sheetId, () => props.studio], ([open]) => { if (open) load() }, { immediate: true })
watch(busy, value => emit('busy-change', value), { immediate: true, flush: 'sync' })
watch([() => props.studio, () => props.workbookId], ([studio, unitId], _previous, onCleanup) => {
  const changed = (id: string): void => { if (!disposed && id === unitId) modelRevision.value += 1 }
  const listeners = [
    studio.events.addEvent('CommandExecuted', event => { if (event.unitId) changed(event.unitId) }),
    studio.events.addEvent('Undo', event => { if (event.kind === 'workbook') changed(event.unitId) }),
    studio.events.addEvent('Redo', event => { if (event.kind === 'workbook') changed(event.unitId) }),
    studio.events.addEvent('PivotTableChanged', event => changed(event.unitId)),
    studio.events.addEvent('PivotTableRendered', event => changed(event.unitId)),
    studio.events.addEvent('SheetValueChanged', event => changed(event.workbookId)),
    studio.events.addEvent('SheetStructureChanged', event => changed(event.workbookId)),
    studio.events.addEvent('WorkbookDisposed', event => { if (event.workbookId === unitId) { changed(unitId); emit('update:open', false) } }),
  ]
  onCleanup(() => listeners.forEach(listener => listener.dispose()))
}, { immediate: true })
onScopeDispose(() => { disposed = true })

function chooseChart(value: unknown): void {
  if (typeof value !== 'string' || busy.value) return
  const selected = workbook.value?.charts?.find(item => item.id === value)
  if (value && !selected) { error.value = '所选透视图已被删除。'; return }
  try {
    draft.value = createPivotChartDraft(selected, draft.value.pivotTableId, chartPlacementGeometry(workbook.value, selected?.sheetId ?? props.sheetId))
    error.value = ''; status.value = ''
  } catch (reason) { error.value = reason instanceof Error ? reason.message : String(reason) }
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
  if (!props.open || busy.value) return
  const current = draft.value, studio = props.studio, unitId = props.workbookId, sheetId = props.sheetId, hostSheetId = chart.value?.sheetId ?? sheetId
  const active = (): boolean => !disposed && props.open && draft.value === current && props.studio === studio && props.workbookId === unitId && props.sheetId === sheetId
  busy.value = true; error.value = ''; status.value = ''
  try {
    assertReady()
    const book = requireWorkbook(studio, unitId)
    if (!book.getPivotTables().some(item => item.getPivotTableId() === current.pivotTableId)) throw new Error(t('sheet.pivot-chart.error.pivot-missing', '所选透视表不存在。'))
    const targetSheet = book.getSheetBySheetId(hostSheetId)
    if (!targetSheet) throw new Error('放置图表的工作表已被删除。')
    const geometry = chartPlacementGeometry(studio.getWorkbookModel(unitId), hostSheetId)
    let chartId = current.chartId
    const updating = Boolean(chartId)
    if (chartId) {
      const existing = targetSheet.getChartById(chartId)
      if (!existing) throw new Error(t('sheet.pivot-chart.error.chart-deleted', '所选透视图已被删除。'))
      await existing.update(pivotChartPatch(current, geometry))
    } else {
      const created = await targetSheet.insertPivotChart(pivotChartConfig(current, geometry))
      chartId = created.getChartId()
    }
    if (active()) {
      modelRevision.value += 1
      const saved = studio.getWorkbookModel(unitId)?.charts?.find(item => item.id === chartId)
      if (saved) draft.value = createPivotChartDraft(saved, '', chartPlacementGeometry(studio.getWorkbookModel(unitId), saved.sheetId))
      status.value = t(updating ? 'sheet.pivot-chart.status.updated' : 'sheet.pivot-chart.status.created', updating ? '透视图设置已更新。' : '已创建并关联透视图。')
      emit('applied', chartId)
    }
  } catch (reason) { if (active()) error.value = reason instanceof Error ? reason.message : String(reason) }
  finally { busy.value = false }
}

async function remove(): Promise<void> {
  if (!props.open || !draft.value.chartId || busy.value) return
  const current = draft.value, studio = props.studio, unitId = props.workbookId, sheetId = props.sheetId, hostSheetId = chart.value?.sheetId ?? sheetId
  const active = (): boolean => !disposed && props.open && draft.value === current && props.studio === studio && props.workbookId === unitId && props.sheetId === sheetId
  busy.value = true; error.value = ''; status.value = ''
  try {
    const target = requireWorkbook(studio, unitId).getSheetBySheetId(hostSheetId)?.getChartById(current.chartId)
    if (!target) throw new Error(t('sheet.pivot-chart.error.chart-deleted', '所选透视图已被删除。'))
    if (!await target.remove()) throw new Error('透视图未能移除。')
    if (active()) {
      modelRevision.value += 1
      draft.value = createPivotChartDraft(undefined, current.pivotTableId, chartPlacementGeometry(studio.getWorkbookModel(unitId), sheetId))
      status.value = t('sheet.pivot-chart.status.removed', '透视图已移除，关联透视表仍保留。')
      emit('applied', current.chartId)
    }
  } catch (reason) { if (active()) error.value = reason instanceof Error ? reason.message : String(reason) }
  finally { busy.value = false }
}
function close(): void { if (!busy.value) emit('update:open', false) }
function opened(value: boolean): void { if (!value && !props.open) emit('closed') }
</script>

<template>
  <Modal :open="open" :title="t('sheet.pivot-chart.title', '透视图')" :width="760" :footer="null" :mask-closable="!busy" :keyboard="!busy" :closable="!busy" :focus-trigger-after-close="false" @cancel="close" @after-open-change="opened">
    <Form layout="vertical" class="pivot-chart-panel" data-testid="pivot-chart-panel" @submit.prevent="apply">
      <FormItem :label="t('sheet.pivot-chart.related-pivot', '关联透视表')"><Select :value="draft.pivotTableId" :options="pivotOptions" :aria-label="t('sheet.pivot-chart.related-pivot', '关联透视表')" :disabled="busy" @update:value="choosePivot" /></FormItem>
      <FormItem v-if="linkedCharts.length" :label="t('sheet.pivot-chart.edit-existing', '编辑已有透视图')"><Select :value="draft.chartId" :options="[{ label: t('sheet.pivot-chart.create-new', '新建透视图'), value: '' }, ...linkedCharts.map((item, index) => ({ label: item.title || `${t('sheet.pivot-chart.title', '透视图')} ${index + 1}`, value: item.id }))]" :aria-label="t('sheet.pivot-chart.edit-existing', '选择透视图')" :disabled="busy" @update:value="chooseChart" /></FormItem>
      <FormItem :label="t('sheet.pivot-chart.placement', '放置单元格')"><Input v-model:value="draft.position" :aria-label="t('sheet.pivot-chart.placement', '透视图放置单元格')" placeholder="F2" :disabled="busy" /></FormItem>
      <ChartSettingsFields v-model:draft="draft.settings" :projection="preview" :date1904="workbook?.date1904" :host-sheet-id="sourceHostSheetId" locked-source :disabled="busy" />
      <Alert v-if="sourceLoading" type="info" message="正在读取图表数据源…" show-icon role="status" />
      <Button v-if="preview.error" :disabled="busy || sourceLoading" @click="reload">重新读取数据源</Button>
      <Alert v-if="error" type="error" :message="error" show-icon />
      <Alert v-if="status" type="success" :message="status" show-icon role="status" />
      <footer :style="{ borderColor: token.colorBorderSecondary }">
        <Button v-if="draft.chartId" danger :disabled="busy" @click="remove">{{ t('sheet.pivot-chart.remove', '移除透视图') }}</Button>
        <span />
        <Button :disabled="busy" @click="close">{{ t('sheet.pivot-chart.close', '关闭') }}</Button>
        <Button type="primary" :loading="busy" :disabled="!pivotTables.length || sourceLoading" @click="apply">{{ t(draft.chartId ? 'sheet.pivot-chart.update' : 'sheet.pivot-chart.create', draft.chartId ? '更新透视图' : '创建透视图') }}</Button>
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
