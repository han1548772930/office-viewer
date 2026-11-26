<script setup lang="ts">
import { Button, Checkbox, FormItem, Input, Select, theme } from 'antdv-next'
import { ArrowDown, ArrowUp, Filter, Trash2 } from 'lucide-vue-next'
import { computed, shallowRef } from 'vue'
import type { PivotFieldArea, PivotFieldModel, PivotTableModel, PivotTableOptions, PivotValuePosition } from '../../../domain/model.ts'
import type { PivotChartFieldsConfig } from '../services/pivot-chart-engine.ts'
import type { PivotSourceFieldInfo } from '../services/pivot-engine.ts'
import { PIVOT_CHART_AGGREGATION_OPTIONS, PIVOT_CHART_FIELD_AREAS, pivotChartFieldsInArea, type PivotChartDraftFieldSetting } from '../services/pivot-chart-field-draft.ts'
import type { PivotShowDataAsBaseItemsLoader } from '../services/pivot-show-data-as-ui.ts'
import PivotShowDataAsFields from './PivotShowDataAsFields.vue'
import { pivotChartValuesUiOptions, pivotValuesUiEntries, readPivotValuesUiState } from '../services/pivot-values-position-ui.ts'
import PivotValuesFieldCard from './PivotValuesFieldCard.vue'
import PivotValuesPositionFields from './PivotValuesPositionFields.vue'

const props = defineProps<{ config: PivotChartFieldsConfig; sourceFields: PivotSourceFieldInfo[]; disabled?: boolean; navigationDisabled?: boolean;
  showDataAsReport?: PivotTableModel; loadShowDataAsBaseItems: PivotShowDataAsBaseItemsLoader; showDataAsContextKey?: string; showDataAsReadable?: boolean }>()
const emit = defineEmits<{
  add: [id: string, area: PivotFieldArea]
  move: [id: string, area: PivotFieldArea, index: number]
  remove: [id: string]
  setting: [id: string, patch: PivotChartDraftFieldSetting]
  'show-as-validity': [id: string, valid: boolean]
  filter: [id: string]
  options: [options: Partial<PivotTableOptions>]
}>()
const { token } = theme.useToken()
const areas = PIVOT_CHART_FIELD_AREAS.map(item => ({ ...item }))
const aggregations = PIVOT_CHART_AGGREGATION_OPTIONS.map(item => ({ ...item }))
const valuesState = computed(() => readPivotValuesUiState(pivotChartValuesUiOptions(props.config), props.config.fields))
const sections = computed(() => areas.map(area => {
  const fields = pivotChartFieldsInArea(props.config, area.value)
  return { ...area, fields, entries: pivotValuesUiEntries(fields, valuesState.value, area.value) }
}))
const draggedValues = shallowRef<{ config: PivotChartFieldsConfig; signature: string; context?: string }>()
const sourceOptions = computed(() => props.sourceFields.map(info => ({ label: info.fieldName, value: info.fieldId })))
const options = computed(() => ({ repeatRowLabels: false, showRowGrandTotal: true, showColGrandTotal: true,
  showRowSubtotals: true, showColSubtotals: true, valuePosition: 'column' as const, ...props.config.options }))
const optionControls: Array<{ key: 'repeatRowLabels' | 'showRowGrandTotal' | 'showColGrandTotal' | 'showRowSubtotals' | 'showColSubtotals'; label: string }> = [
  { key: 'repeatRowLabels', label: '重复行标签' }, { key: 'showRowGrandTotal', label: '显示行总计' },
  { key: 'showColGrandTotal', label: '显示列总计' }, { key: 'showRowSubtotals', label: '显示行小计' }, { key: 'showColSubtotals', label: '显示列小计' },
]
function add(value: unknown, area: PivotFieldArea): void { if (typeof value === 'string') emit('add', value, area) }
function move(value: unknown, field: PivotFieldModel): void {
  const area = areas.find(item => item.value === value)?.value
  if (area) emit('move', field.id, area, pivotChartFieldsInArea(props.config, area).filter(item => item.id !== field.id).length)
}
function aggregate(value: unknown, id: string): void {
  const aggregation = aggregations.find(item => item.value === value)?.value
  if (aggregation) emit('setting', id, { aggregation })
}
function valuesPosition(position: PivotValuePosition, index: number): void {
  if (!props.navigationDisabled) emit('options', { valuePosition: position, valueIndex: index })
}
function startValuesDrag(event: DragEvent): void {
  if (props.navigationDisabled || !valuesState.value.visible) { event.preventDefault(); return }
  draggedValues.value = { config: props.config, signature: JSON.stringify(props.config), context: props.showDataAsContextKey }
  if (event.dataTransfer) { event.dataTransfer.effectAllowed = 'move'; event.dataTransfer.setData('application/x-office-pivot-values', 'values') }
}
function dropValues(area: PivotFieldArea, event: DragEvent, index: number): void {
  event.preventDefault(); event.stopPropagation()
  const captured = draggedValues.value; draggedValues.value = undefined
  if (!captured || props.navigationDisabled || captured.config !== props.config || captured.signature !== JSON.stringify(props.config)
    || captured.context !== props.showDataAsContextKey || area !== 'row' && area !== 'column') return
  valuesPosition(area, index)
}
</script>

<template>
  <div class="pivot-chart-fields" :style="{ '--field-border': token.colorBorderSecondary, '--field-surface': token.colorFillAlter }" data-testid="pivot-chart-fields">
    <div class="field-areas">
      <section v-for="area in sections" :key="area.value" class="field-area" :data-pivot-area="area.value" @dragover.prevent @drop="dropValues(area.value, $event, area.fields.length)">
        <h3>{{ area.label }}</h3>
        <FormItem :label="`添加${area.label}字段`"><Select :value="undefined" :options="sourceOptions" show-search option-filter-prop="label" :aria-label="`添加${area.label}字段`" placeholder="选择来源字段" :disabled="navigationDisabled" @change="add($event, area.value)" /></FormItem>
        <p v-if="!area.entries.length" class="field-empty">此区域尚无字段</p>
        <template v-for="item in area.entries" :key="item.kind === 'field' ? item.field.id : 'synthetic-values'">
        <PivotValuesFieldCard v-if="item.kind === 'values'" :state="item.values" subject-label="透视图" :disabled="navigationDisabled"
          @change="valuesPosition" @drag-start="startValuesDrag" @drag-end="draggedValues = undefined"
          @dragover.prevent @drop="dropValues(area.value, $event, valuesState.effectiveIndex)" />
        <div v-else class="field-card" :data-pivot-field="item.field.id" @dragover.prevent @drop="dropValues(area.value, $event, item.field.position)">
          <template v-for="field in [item.field]" :key="field.id">
          <FormItem :label="`${field.name}的显示名称`"><Input :value="field.displayName ?? field.name" :aria-label="`${field.name}的显示名称`" :disabled="disabled" @update:value="emit('setting', field.id, { displayName: $event })" /></FormItem>
          <FormItem :label="`${field.name}的字段区域`"><Select :value="field.area" :options="areas" :aria-label="`${field.name}的字段区域`" :disabled="navigationDisabled" @change="move($event, field)" /></FormItem>
          <FormItem v-if="field.area === 'value'" :label="`${field.name}的汇总方式`"><Select :value="field.aggregation ?? 'sum'" :options="aggregations" :aria-label="`${field.name}的汇总方式`" :disabled="navigationDisabled" @change="aggregate($event, field.id)" /></FormItem>
          <PivotShowDataAsFields v-if="field.area === 'value'" :value="field.showDataAs" :report="showDataAsReport" :label-prefix="field.name"
            :load-base-items="loadShowDataAsBaseItems" :context-key="showDataAsContextKey" :readable="showDataAsReadable" :disabled="navigationDisabled"
            @update:value="emit('setting', field.id, { showDataAs: $event })" @validity-change="emit('show-as-validity', field.id, $event)" />
          <FormItem :label="`${field.name}的数字格式`"><Input :value="field.format ?? ''" placeholder="General" :aria-label="`${field.name}的数字格式`" :disabled="disabled" @update:value="emit('setting', field.id, { format: $event || undefined })" /></FormItem>
          <div class="field-actions">
            <Button size="small" :aria-label="`上移${field.name}`" :disabled="navigationDisabled || field.position === 0" @click="emit('move', field.id, field.area, field.position - 1)"><ArrowUp :size="14" /></Button>
            <Button size="small" :aria-label="`下移${field.name}`" :disabled="navigationDisabled || field.position === area.fields.length - 1" @click="emit('move', field.id, field.area, field.position + 1)"><ArrowDown :size="14" /></Button>
            <Button v-if="field.area !== 'value'" size="small" :type="field.selectedItems === undefined ? 'default' : 'primary'" :disabled="navigationDisabled" :aria-label="`筛选${field.name}`" @click="emit('filter', field.id)"><Filter :size="14" />标签筛选</Button>
            <Button size="small" danger :disabled="navigationDisabled" :aria-label="`移除${field.name}`" @click="emit('remove', field.id)"><Trash2 :size="14" />移除</Button>
          </div>
          </template>
        </div>
        </template>
      </section>
    </div>
    <div class="report-options">
      <PivotValuesPositionFields :state="valuesState" subject-label="透视图" :disabled="navigationDisabled" @change="valuesPosition" />
      <FormItem v-for="control in optionControls" :key="control.key" :label="null"><Checkbox :checked="options[control.key]" :disabled="navigationDisabled" @update:checked="emit('options', { [control.key]: $event })">{{ control.label }}</Checkbox></FormItem>
    </div>
  </div>
</template>

<style scoped>
.pivot-chart-fields { display: grid; gap: 16px; min-width: 0; }
.field-areas { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px; }
.field-area { min-width: 0; padding: 12px; border: 1px solid var(--field-border); border-radius: 6px; }
.field-area h3 { margin: 0 0 10px; font-size: 14px; }
.field-card { display: grid; gap: 8px; margin-top: 10px; padding: 10px; background: var(--field-surface); border-radius: 4px; }
.field-actions { display: flex; flex-wrap: wrap; gap: 6px; }
.field-actions :deep(.ant-btn) { display: inline-flex; align-items: center; gap: 4px; }
.field-empty { margin: 12px 0 0; opacity: .65; }
.report-options { display: grid; gap: 8px; }
@media (max-width: 620px) { .field-areas { grid-template-columns: minmax(0, 1fr); } }
</style>
