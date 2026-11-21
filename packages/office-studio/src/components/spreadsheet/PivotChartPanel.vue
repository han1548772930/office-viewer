<script setup lang="ts">
import { computed, inject, ref, watch } from 'vue'
import { Button, Input, Modal, Select, Switch } from 'antdv-next'
import { FOffice } from '../../api/facade'
import type { OfficeStudio } from '../../api/office-studio'
import type { ChartModel } from '../../domain/model'
import { createPivotChartDraft, pivotChartConfig, pivotChartPatch, type PivotChartDraft } from '../../features/pivot/services/pivot-chart-panel-state'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

const props = defineProps<{ open: boolean; studio: OfficeStudio; workbookId: string; sheetId: string; pivotTableId?: string; chartId?: string }>()
const emit = defineEmits<{ 'update:open': [value: boolean]; applied: [chartId: string] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)

const draft = ref<PivotChartDraft>(createPivotChartDraft())
const error = ref('')
const status = ref('')
const busy = ref(false)
const workbook = computed(() => props.studio.getWorkbookModel(props.workbookId))
const pivotTables = computed(() => workbook.value?.pivotTables ?? [])
const pivotOptions = computed(() => pivotTables.value.map((pivot, index) => ({
  label: `${t('sheet.pivot-chart.title', '透视表')} ${index + 1} · ${workbook.value?.sheets.find(sheet => sheet.id === pivot.anchor.subUnitId)?.name ?? ''}`,
  value: pivot.id,
})))
const chart = computed(() => workbook.value?.charts?.find(item => item.id === (props.chartId || draft.value.chartId)))
const linkedCharts = computed(() => workbook.value?.charts?.filter(item => item.pivotTableId === draft.value.pivotTableId) ?? [])
const typeOptions = computed(() => [
  { label: t('sheet.pivot-chart.type.bar', '簇状柱形图'), value: 'bar' },
  { label: t('sheet.pivot-chart.type.line', '折线图'), value: 'line' },
  { label: t('sheet.pivot-chart.type.area', '面积图'), value: 'area' },
  { label: t('sheet.pivot-chart.type.pie', '饼图'), value: 'pie' },
  { label: t('sheet.pivot-chart.type.doughnut', '圆环图'), value: 'doughnut' },
  { label: t('sheet.pivot-chart.type.stacked-column', '堆积柱形图'), value: 'stacked-column' },
])
const legendOptions = computed(() => [
  { label: t('sheet.pivot-chart.top', '上方'), value: 'top' }, { label: t('sheet.pivot-chart.right', '右侧'), value: 'right' },
  { label: t('sheet.pivot-chart.bottom', '下方'), value: 'bottom' }, { label: t('sheet.pivot-chart.left', '左侧'), value: 'left' },
])

function load() {
  error.value = ''; status.value = ''
  const current = props.chartId ? workbook.value?.charts?.find(item => item.id === props.chartId) : undefined
  draft.value = createPivotChartDraft(current, props.pivotTableId ?? pivotTables.value[0]?.id ?? '')
}

watch(() => [props.open, props.chartId, props.pivotTableId, props.workbookId] as const, ([open]) => { if (open) load() }, { immediate: true })

function chooseChart(chartId: string) {
  const selected = workbook.value?.charts?.find(item => item.id === chartId)
  if (selected) draft.value = createPivotChartDraft(selected)
}

function requireWorkbook() {
  const book = FOffice.newAPI(props.studio).getWorkbook(props.workbookId)
  if (!book) throw new Error(t('sheet.pivot-chart.error.workbook-closed', '工作簿已关闭，请重新打开透视图面板。'))
  return book
}

async function apply() {
  if (busy.value) return
  busy.value = true; error.value = ''; status.value = ''
  try {
    const book = requireWorkbook()
    const patch = pivotChartPatch(draft.value)
    const pivot = book.getPivotTables().find(item => item.getPivotTableId() === draft.value.pivotTableId)
    if (!pivot) throw new Error(t('sheet.pivot-chart.error.pivot-missing', '所选透视表不存在。'))
    let chartId = draft.value.chartId
    if (chartId) {
      const targetSheet = book.getSheetBySheetId(props.sheetId)
      const existing = targetSheet?.getChartById(chartId)
      if (!existing) throw new Error(t('sheet.pivot-chart.error.chart-deleted', '所选透视图已被删除。'))
      await existing.update(patch)
      status.value = t('sheet.pivot-chart.status.updated', '透视图设置已更新。')
    } else {
      const created = await pivot.insertChart(pivotChartConfig(draft.value))
      chartId = created.getChartId()
      draft.value.chartId = chartId
      status.value = t('sheet.pivot-chart.status.created', '已创建并关联透视图。')
    }
    emit('applied', chartId)
  } catch (reason) {
    error.value = reason instanceof Error ? reason.message : String(reason)
  } finally { busy.value = false }
}

async function remove() {
  if (!draft.value.chartId || busy.value) return
  busy.value = true; error.value = ''; status.value = ''
  try {
    const chartId = draft.value.chartId
    const target = FOffice.newAPI(props.studio).getWorkbook(props.workbookId)?.getSheetBySheetId(props.sheetId)?.getChartById(chartId)
    if (!target) throw new Error(t('sheet.pivot-chart.error.chart-deleted', '所选透视图已被删除。'))
    await target.remove()
    draft.value.chartId = ''
    status.value = t('sheet.pivot-chart.status.removed', '透视图已移除，关联透视表仍保留。')
    emit('applied', chartId)
  } catch (reason) { error.value = reason instanceof Error ? reason.message : String(reason) }
  finally { busy.value = false }
}

function close() { if (!busy.value) emit('update:open', false) }
</script>

<template>
  <Modal :open="open" :title="t('sheet.pivot-chart.title', '透视图')" :width="560" :footer="null" :mask-closable="!busy" @cancel="close">
    <div class="pivot-chart-panel" data-testid="pivot-chart-panel">
      <label>{{ t('sheet.pivot-chart.related-pivot', '关联透视表') }}<Select v-model:value="draft.pivotTableId" :options="pivotOptions" :aria-label="t('sheet.pivot-chart.related-pivot', '关联透视表')" :disabled="busy" /></label>
      <label v-if="linkedCharts.length">{{ t('sheet.pivot-chart.edit-existing', '编辑已有透视图') }}<Select :value="draft.chartId" :options="[{ label: t('sheet.pivot-chart.create-new', '新建透视图'), value: '' }, ...linkedCharts.map((item: ChartModel, index: number) => ({ label: item.title || `${t('sheet.pivot-chart.title', '透视图')} ${index + 1}`, value: item.id }))]" :aria-label="t('sheet.pivot-chart.edit-existing', '选择透视图')" @change="chooseChart" /></label>
      <label>{{ t('sheet.pivot-chart.type', '图表类型') }}<Select v-model:value="draft.type" :options="typeOptions" :aria-label="t('sheet.pivot-chart.type', '透视图类型')" :disabled="busy" /></label>
      <label>{{ t('sheet.pivot-chart.chart-title', '标题') }}<Input v-model:value="draft.title" :aria-label="t('sheet.pivot-chart.chart-title', '透视图标题')" :maxlength="120" :disabled="busy" /></label>
      <label>{{ t('sheet.pivot-chart.placement', '放置单元格') }}<Input v-model:value="draft.position" :aria-label="t('sheet.pivot-chart.placement', '透视图放置单元格')" placeholder="F2" :disabled="busy" /></label>
      <div class="pivot-chart-legend">
        <label><Switch v-model:checked="draft.legendVisible" size="small" :aria-label="t('sheet.pivot-chart.show-legend', '显示图例')" :disabled="busy" />{{ t('sheet.pivot-chart.show-legend', '显示图例') }}</label>
        <label v-if="draft.legendVisible">{{ t('sheet.pivot-chart.legend-position', '图例位置') }}<Select v-model:value="draft.legendPosition" :options="legendOptions" :aria-label="t('sheet.pivot-chart.legend-position', '图例位置')" :disabled="busy" /></label>
      </div>
      <p v-if="chart" class="pivot-chart-binding">{{ t('sheet.pivot-chart.binding', '当前图表绑定透视表 {id}').replace('{id}', chart.pivotTableId ?? '') }}</p>
      <p v-if="error" class="pivot-chart-error" role="alert">{{ error }}</p>
      <p v-if="status" class="pivot-chart-status" role="status">{{ status }}</p>
      <footer>
        <Button v-if="draft.chartId" danger :disabled="busy" @click="remove">{{ t('sheet.pivot-chart.remove', '移除透视图') }}</Button>
        <span />
        <Button :disabled="busy" @click="close">{{ t('sheet.pivot-chart.close', '关闭') }}</Button>
        <Button type="primary" :loading="busy" :disabled="!pivotTables.length" @click="apply">{{ t(draft.chartId ? 'sheet.pivot-chart.update' : 'sheet.pivot-chart.create', draft.chartId ? '更新透视图' : '创建透视图') }}</Button>
      </footer>
    </div>
  </Modal>
</template>

<style scoped>
.pivot-chart-panel { display: grid; grid-template-columns: minmax(0, 1fr) minmax(0, 1fr); gap: 14px 18px; }
.pivot-chart-panel > label, .pivot-chart-legend > label { display: grid; min-width: 0; gap: 5px; color: #536174; font-size: 12px; }
.pivot-chart-legend { display: flex; align-items: center; gap: 16px; }
.pivot-chart-legend > label:first-child { display: flex; align-items: center; white-space: nowrap; }
.pivot-chart-binding, .pivot-chart-error, .pivot-chart-status { grid-column: 1 / -1; margin: 0; font-size: 12px; }
.pivot-chart-binding { color: #788596; }
.pivot-chart-error { color: #b42318; }
.pivot-chart-status { color: #267342; }
.pivot-chart-panel footer { grid-column: 1 / -1; display: flex; align-items: center; gap: 8px; padding-top: 12px; border-top: 1px solid #e7ebf0; }
.pivot-chart-panel footer span { flex: 1; }
@media (max-width: 520px) { .pivot-chart-panel { grid-template-columns: 1fr; } .pivot-chart-binding, .pivot-chart-error, .pivot-chart-status, .pivot-chart-panel footer { grid-column: auto; } .pivot-chart-panel footer { flex-wrap: wrap; } }
</style>
