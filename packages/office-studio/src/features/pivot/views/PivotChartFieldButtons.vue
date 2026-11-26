<script setup lang="ts">
import { Alert, Button, Dropdown, Tooltip, theme } from 'antdv-next'
import { ChevronDown, Filter } from 'lucide-vue-next'
import { computed, h, ref, watch, type VNode } from 'vue'
import type { PivotFieldArea, PivotValuePosition } from '../../../domain/model.ts'
import type { PivotChartFieldButton, PivotChartFieldButtonAction, PivotChartFieldButtonSection } from '../services/pivot-chart-field-button-state.ts'
import { usePivotChartFieldButtons, type PivotChartFieldButtonsOptions } from './use-pivot-chart-field-buttons.ts'
import PivotChartFieldFilterPopup from './PivotChartFieldFilterPopup.vue'
import { pivotValuesUiEntries, type PivotValuesUiEntry, type PivotValuesUiState } from '../services/pivot-values-position-ui.ts'
import PivotValuesFieldCard from './PivotValuesFieldCard.vue'

const props = defineProps<PivotChartFieldButtonsOptions & { maxVisibleFields?: number }>()
const emit = defineEmits<{ 'busy-change': [value: boolean]; applied: [fieldId: string, action: PivotChartFieldButtonAction]; 'values-applied': [position: PivotValuePosition, index: number] }>()
const { token } = theme.useToken()
const controls = usePivotChartFieldButtons(props)
const { snapshot, visible, field, draft, conditions, openAnchor, loading, busy, composing, error, readable, stale, navigationDisabled, editDisabled } = controls
const emptyMenu = { items: [] }
const maxVisible = computed(() => Number.isSafeInteger(props.maxVisibleFields) && props.maxVisibleFields! > 0
  ? Math.min(10, props.maxVisibleFields!) : 2)
type UiSection = PivotChartFieldButtonSection & { entries: PivotValuesUiEntry<PivotChartFieldButton>[] }
const sections = computed(() => snapshot.value?.sections.map(section => {
  const entries = pivotValuesUiEntries(section.fields, snapshot.value!.values, section.area)
  return { ...section, entries, visible: entries.slice(0, maxVisible.value), overflow: entries.length > maxVisible.value }
}) ?? [])
const draggingValues = ref(false)
const appearance = computed(() => ({ color: token.value.colorText, '--field-button-border': token.value.colorBorderSecondary,
  '--field-button-surface': token.value.colorBgElevated }))
function overflowAnchor(section: PivotChartFieldButtonSection): string { return `overflow:${section.area}` }
function fieldAnchor(fieldId: string): string { return `field:${fieldId}` }
function valuesOverflowAnchor(section: PivotChartFieldButtonSection): string { return `values:overflow:${section.area}` }
function openValues(value: boolean, anchor = 'values:position'): void { if (value) controls.openValues(anchor); else if (openAnchor.value === anchor) controls.close() }
function openField(value: boolean, fieldId: string): void { if (value) void controls.openField(fieldId, fieldAnchor(fieldId)); else controls.close() }
function openOverflow(value: boolean, section: PivotChartFieldButtonSection): void { if (value) controls.openList(overflowAnchor(section)); else controls.close() }
async function action(value: PivotChartFieldButtonAction): Promise<void> {
  const id = field.value?.fieldId
  if (id && await controls.mutate(value)) emit('applied', id, value)
}
async function valuesChange(position: PivotValuePosition, index: number): Promise<void> {
  if (await controls.mutateValues(position, index)) emit('values-applied', position, index)
}
function startValuesDrag(event: DragEvent): void {
  if (navigationDisabled.value || !snapshot.value?.editable) { event.preventDefault(); return }
  controls.openValues('values:drag'); draggingValues.value = openAnchor.value === 'values:drag'
  if (!draggingValues.value) { event.preventDefault(); return }
  if (event.dataTransfer) { event.dataTransfer.effectAllowed = 'move'; event.dataTransfer.setData('application/x-office-pivot-values', 'values') }
}
function endValuesDrag(): void { if (draggingValues.value) { draggingValues.value = false; if (openAnchor.value === 'values:drag') controls.close() } }
function dropValues(area: PivotFieldArea, event: DragEvent, index: number): void {
  event.preventDefault(); event.stopPropagation()
  if (!draggingValues.value) return
  draggingValues.value = false
  if (openAnchor.value !== 'values:drag') return
  if (area !== 'row' && area !== 'column') { controls.close(); return }
  void valuesChange(area, index)
}
function valuesCard(state: PivotValuesUiState, anchor: string): VNode {
  return h(PivotValuesFieldCard, { key: 'synthetic-values', state, subjectLabel: '透视图', dense: true, testId: 'pivot-chart-values-field-button',
    menuOpen: openAnchor.value === anchor, disabled: navigationDisabled.value || !snapshot.value?.editable,
    onChange: valuesChange, onMenuOpen: (open: boolean) => openValues(open, anchor), onDragStart: startValuesDrag, onDragEnd: endValuesDrag,
    onDragover: (event: DragEvent) => event.preventDefault(), onDrop: (event: DragEvent) => dropValues(state.area, event, state.effectiveIndex) })
}
function filterPopup(): VNode {
  if (!field.value) return h('span')
  return h(PivotChartFieldFilterPopup, { key: field.value.fieldId, field: field.value, draft: draft.value, conditions: conditions.value,
    readable: readable.value, loading: loading.value, navigationDisabled: navigationDisabled.value,
    busy: busy.value, disabled: editDisabled.value, composing: composing.value, error: error.value, stale: stale.value,
    onSearch: controls.search, onSelect: controls.select, onSelectAll: controls.selectAll, onAction: action, onSettings: controls.settings,
    onDateChange: controls.dateChange, onLabelChange: controls.labelChange,
    onRetry: () => { if (field.value) void controls.openField(field.value.fieldId, openAnchor.value) }, onClose: controls.close,
    onCompositionStart: controls.compositionStart, onCompositionEnd: controls.compositionEnd, onKeydown: controls.keydown })
}
function overflowPopup(section: UiSection): VNode {
  if (field.value && openAnchor.value === overflowAnchor(section)) return filterPopup()
  return h('div', { class: 'pivot-chart-field-overflow', 'data-testid': 'pivot-chart-field-overflow', role: 'group',
    'aria-label': `${section.label}全部字段`, style: { color: token.value.colorText, background: token.value.colorBgElevated,
      border: `1px solid ${token.value.colorBorderSecondary}`, borderRadius: '6px', padding: '4px', width: '224px',
      maxWidth: 'calc(100vw - 16px)', maxHeight: 'min(320px, calc(100vh - 24px))', overflowY: 'auto' },
    onPointerdown: (event: PointerEvent) => event.stopPropagation(), onClick: (event: MouseEvent) => event.stopPropagation(),
    onKeydown: (event: KeyboardEvent) => { event.stopPropagation(); controls.keydown(event) },
  }, section.entries.map(item => {
    if (item.kind === 'values') return valuesCard(item.values, valuesOverflowAnchor(section))
    const button = item.field
    return h(Button, { key: button.fieldId, size: 'small', type: button.filtered ? 'primary' : 'text',
    'aria-label': `打开透视图字段${button.label}`, 'aria-pressed': button.filtered, 'data-pivot-field': button.fieldId,
    disabled: navigationDisabled.value, style: { display: 'flex', width: '100%', textAlign: 'start', justifyContent: 'flex-start' },
    onClick: () => controls.openField(button.fieldId, overflowAnchor(section)),
    onDragover: (event: DragEvent) => event.preventDefault(), onDrop: (event: DragEvent) => dropValues(button.area, event, button.field.position),
  }, () => h('span', { style: { overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' } }, button.label))
  }))
}
function tooltip(button: PivotChartFieldButton): string { return `${button.label}${button.filtered ? '（已筛选）' : ''}` }
watch(busy, value => emit('busy-change', value), { immediate: true, flush: 'sync' })
defineExpose({ synchronize: controls.synchronize, close: controls.close })
</script>

<template>
  <div v-if="visible" class="pivot-chart-field-buttons" :style="appearance" data-testid="pivot-chart-field-buttons"
    :data-pivot-chart="chartId" role="group" aria-label="透视图字段按钮" @pointerdown.stop @click.stop @dblclick.stop @contextmenu.stop @keydown.stop>
    <section v-for="section in sections" v-show="section.entries.length" :key="section.area" class="field-button-area"
      :class="`field-button-area-${section.area}`" :data-pivot-button-area="section.area" :aria-label="`${section.label}字段`" @dragover.prevent @drop="dropValues(section.area, $event, section.fields.length)">
      <span class="field-area-label">{{ section.label }}</span>
      <template v-for="item in section.visible" :key="item.kind === 'field' ? item.field.fieldId : 'synthetic-values'">
      <PivotValuesFieldCard v-if="item.kind === 'values'" :state="item.values" subject-label="透视图" dense test-id="pivot-chart-values-field-button"
        :menu-open="openAnchor === 'values:position'" :disabled="navigationDisabled || !snapshot?.editable" @menu-open="openValues"
        @change="valuesChange" @drag-start="startValuesDrag" @drag-end="endValuesDrag" @dragover.prevent @drop="dropValues(section.area, $event, item.values.effectiveIndex)" />
      <template v-else>
      <Dropdown :open="openAnchor === fieldAnchor(item.field.fieldId)" :menu="emptyMenu" :trigger="['click']"
        :popup-render="filterPopup" :auto-focus="false" :destroy-on-hidden="true" :auto-adjust-overflow="true" placement="bottomLeft"
        :disabled="navigationDisabled" @open-change="openField($event, item.field.fieldId)">
        <Tooltip :title="tooltip(item.field)" :open="openAnchor === fieldAnchor(item.field.fieldId) ? false : undefined"><Button size="small" :type="item.field.filtered ? 'primary' : 'default'"
          :disabled="navigationDisabled" :aria-label="`打开透视图字段${item.field.label}`" :aria-pressed="item.field.filtered" :aria-expanded="openAnchor === fieldAnchor(item.field.fieldId)"
          aria-haspopup="dialog" :data-pivot-field="item.field.fieldId" @dragover.prevent @drop="dropValues(section.area, $event, item.field.field.position)"><Filter v-if="item.field.filtered" :size="12" aria-hidden="true" /><span class="field-label">{{ item.field.label }}</span><ChevronDown :size="12" aria-hidden="true" /></Button></Tooltip>
      </Dropdown>
      </template>
      </template>
      <Dropdown v-if="section.overflow" :open="openAnchor === overflowAnchor(section) || openAnchor === valuesOverflowAnchor(section)" :menu="emptyMenu" :trigger="['click']"
        :popup-render="() => overflowPopup(section)" :auto-focus="false" :destroy-on-hidden="true" :auto-adjust-overflow="true" placement="bottomLeft"
        :disabled="navigationDisabled" @open-change="openOverflow($event, section)">
        <Tooltip :title="`${section.label}全部字段`" :open="openAnchor === overflowAnchor(section) ? false : undefined"><Button size="small" :disabled="navigationDisabled"
          :aria-label="`更多${section.label}字段`" :aria-expanded="openAnchor === overflowAnchor(section)" aria-haspopup="dialog">更多<ChevronDown :size="12" aria-hidden="true" /></Button></Tooltip>
      </Dropdown>
    </section>
    <Alert v-if="error && (!openAnchor || openAnchor.startsWith('values:'))" type="error" :title="error" show-icon class="field-button-error" />
  </div>
</template>

<style scoped>
.pivot-chart-field-buttons { position: absolute; inset: 6px; display: grid; grid-template-columns: minmax(0, 1fr) minmax(0, 1fr); grid-template-rows: auto 1fr auto; gap: 4px; pointer-events: none; font-size: 11px; }
.field-button-area { display: flex; align-items: center; flex-wrap: wrap; gap: 3px; min-width: 0; pointer-events: auto; }
.field-button-area-filter { grid-column: 1; grid-row: 1; align-self: start; }
.field-button-area-value { grid-column: 2; grid-row: 1; align-self: start; justify-content: flex-end; }
.field-button-area-row { grid-column: 1; grid-row: 3; align-self: end; }
.field-button-area-column { grid-column: 2; grid-row: 3; align-self: end; justify-content: flex-end; }
.field-area-label { opacity: .75; white-space: nowrap; }
.field-button-area :deep(.ant-btn) { display: inline-flex; align-items: center; gap: 3px; max-width: 100%; height: 22px; padding-inline: 5px; font-size: 11px; }
.field-label { display: block; max-width: 116px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.field-button-error { grid-row: 2; grid-column: 1 / -1; align-self: start; pointer-events: auto; }
</style>
