<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { Button, Checkbox, Drawer, Form, FormItem, Input, Select } from 'antdv-next'
import type { ChartLegendPosition, ChartModel, ChartSourceOrientation } from '../../domain/model'
import type { SpreadsheetRenderSheet } from '../../rendering/spreadsheet/layout'

type ChartPatch = Partial<Omit<ChartModel, 'id' | 'sheetId' | 'revision'>>

const props = defineProps<{
  open: boolean
  chart: ChartModel | null
  sheet: SpreadsheetRenderSheet | undefined
}>()

const emit = defineEmits<{
  (event: 'update:open', value: boolean): void
  (event: 'apply', patch: ChartPatch): void
}>()

interface ChartDraft {
  title: string
  source: string
  orientation: ChartSourceOrientation
  categoryField: number
  valueFields: number[]
  legendVisible: boolean
  legendPosition: ChartLegendPosition
  xAxisTitle: string
  yAxisTitle: string
}

const draft = ref<ChartDraft>(emptyDraft())
const error = ref('')

const fieldOptions = computed(() => {
  const range = parseRange(draft.value.source)
  if (!range || !props.sheet) return []
  const count = draft.value.orientation === 'columns'
    ? range.endCol - range.startCol + 1
    : range.endRow - range.startRow + 1
  return Array.from({ length: count }, (_value, index) => {
    const row = range.startRow + (draft.value.orientation === 'rows' ? index : 0)
    const column = range.startCol + (draft.value.orientation === 'columns' ? index : 0)
    const address = `${columnName(column)}${row + 1}`
    const cellValue = props.sheet?.cells[address]?.value
    const label = cellValue == null || cellValue === '' ? address : `${address} · ${cellValue}`
    return { value: index, label }
  })
})

watch([() => props.open, () => props.chart], ([isOpen, chart]) => {
  if (!isOpen || !chart) return
  draft.value = {
    title: chart.title ?? '',
    source: rangeNotation(chart.source.range),
    orientation: chart.source.orientation,
    categoryField: chart.categoryField,
    valueFields: [...chart.valueFields],
    legendVisible: chart.legend.visible,
    legendPosition: chart.legend.position,
    xAxisTitle: chart.xAxis.title ?? '',
    yAxisTitle: chart.yAxis.title ?? '',
  }
  error.value = ''
}, { immediate: true })

function emptyDraft(): ChartDraft {
  return {
    title: '', source: 'A1:A1', orientation: 'columns', categoryField: 0, valueFields: [],
    legendVisible: true, legendPosition: 'bottom', xAxisTitle: '', yAxisTitle: '',
  }
}

function apply(): void {
  const chart = props.chart
  if (!chart) return
  const range = parseRange(draft.value.source)
  if (!range) {
    error.value = '请输入有效的数据范围，例如 A1:C4。'
    return
  }
  const fieldCount = draft.value.orientation === 'columns'
    ? range.endCol - range.startCol + 1
    : range.endRow - range.startRow + 1
  if (draft.value.categoryField < 0 || draft.value.categoryField >= fieldCount) {
    error.value = '分类字段必须位于当前数据范围内。'
    return
  }
  if (draft.value.valueFields.some(field => field < 0 || field >= fieldCount)) {
    error.value = '数值系列字段必须位于当前数据范围内。'
    return
  }
  if (chart.pivotTableId && (draft.value.source !== rangeNotation(chart.source.range) || draft.value.orientation !== chart.source.orientation)) {
    error.value = '透视图的数据范围由关联的数据透视表管理。'
    return
  }

  const patch: ChartPatch = {
    title: draft.value.title.trim() || undefined,
    legend: { visible: draft.value.legendVisible, position: draft.value.legendPosition },
    xAxis: { ...chart.xAxis, title: draft.value.xAxisTitle.trim() || undefined },
    yAxis: { ...chart.yAxis, title: draft.value.yAxisTitle.trim() || undefined },
  }
  if (!chart.pivotTableId) {
    patch.source = { range, orientation: draft.value.orientation }
    patch.categoryField = draft.value.categoryField
    patch.valueFields = [...draft.value.valueFields]
  }
  error.value = ''
  emit('apply', patch)
}

function parseRange(value: string) {
  const parts = value.trim().toUpperCase().split(':')
  if (parts.length > 2 || !parts[0]) return undefined
  const first = parseAddress(parts[0])
  const last = parseAddress(parts[1] ?? parts[0])
  if (!first || !last) return undefined
  return {
    startRow: Math.min(first.row, last.row),
    endRow: Math.max(first.row, last.row),
    startCol: Math.min(first.column, last.column),
    endCol: Math.max(first.column, last.column),
  }
}

function parseAddress(value: string): { row: number; column: number } | undefined {
  const match = /^([A-Z]+)([1-9]\d*)$/.exec(value)
  if (!match) return undefined
  let column = 0
  for (const letter of match[1]!) column = column * 26 + letter.charCodeAt(0) - 64
  return { row: Number(match[2]) - 1, column: column - 1 }
}

function rangeNotation(range: ChartModel['source']['range']): string {
  const start = `${columnName(range.startCol)}${range.startRow + 1}`
  const end = `${columnName(range.endCol)}${range.endRow + 1}`
  return start === end ? start : `${start}:${end}`
}

function columnName(index: number): string {
  let result = ''
  for (let value = index; value >= 0; value = Math.floor(value / 26) - 1) {
    result = String.fromCharCode((value % 26) + 65) + result
  }
  return result
}
</script>

<template>
  <Drawer
    :open="open"
    title="图表设置"
    placement="right"
    :size="440"
    :footer="null"
    @close="emit('update:open', false)"
  >
    <Form layout="vertical" class="chart-editor-form" @submit.prevent="apply">
      <FormItem label="图表标题">
        <Input v-model:value="draft.title" aria-label="图表标题" allow-clear />
      </FormItem>
      <FormItem label="数据区域" :validate-status="error ? 'error' : undefined" :help="error || undefined">
        <Input v-model:value="draft.source" aria-label="数据区域" :disabled="Boolean(chart?.pivotTableId)" />
      </FormItem>
      <FormItem label="系列方向">
        <Select
          v-model:value="draft.orientation"
          aria-label="系列方向"
          :disabled="Boolean(chart?.pivotTableId)"
          :options="[{ label: '按列', value: 'columns' }, { label: '按行', value: 'rows' }]"
        />
      </FormItem>
      <FormItem label="分类字段">
        <Select v-model:value="draft.categoryField" aria-label="分类字段" :disabled="Boolean(chart?.pivotTableId)" :options="fieldOptions" />
      </FormItem>
      <FormItem label="数值系列字段" extra="不选择时使用范围内除分类字段外的数值字段。">
        <Select
          v-model:value="draft.valueFields"
          aria-label="数值系列字段"
          mode="multiple"
          :disabled="Boolean(chart?.pivotTableId)"
          :options="fieldOptions.filter(option => option.value !== draft.categoryField)"
        />
      </FormItem>
      <FormItem label="X 轴标题">
        <Input v-model:value="draft.xAxisTitle" aria-label="X 轴标题" allow-clear />
      </FormItem>
      <FormItem label="Y 轴标题">
        <Input v-model:value="draft.yAxisTitle" aria-label="Y 轴标题" allow-clear />
      </FormItem>
      <FormItem label="图例位置">
        <Select
          v-model:value="draft.legendPosition"
          aria-label="图例位置"
          :options="[
            { label: '上方', value: 'top' }, { label: '右侧', value: 'right' },
            { label: '下方', value: 'bottom' }, { label: '左侧', value: 'left' },
          ]"
        />
      </FormItem>
      <FormItem label="图例">
        <Checkbox v-model:checked="draft.legendVisible">显示图例</Checkbox>
      </FormItem>
      <p v-if="chart?.pivotTableId" class="chart-editor-note">数据区域与系列方向由关联的数据透视表管理。</p>
      <div class="chart-editor-actions">
        <Button @click="emit('update:open', false)">取消</Button>
        <Button type="primary" @click="apply">应用</Button>
      </div>
    </Form>
  </Drawer>
</template>

<style scoped>
.chart-editor-form { display: grid; gap: 2px; }
.chart-editor-note { color: var(--ant-color-text-secondary); }
.chart-editor-actions { display: flex; justify-content: flex-end; gap: 8px; padding-top: 12px; }
</style>
