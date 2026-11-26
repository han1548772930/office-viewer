<script setup lang="ts">
import { Alert, Button, Checkbox, Form, FormItem, Modal, Select, theme } from 'antdv-next'
import { computed, nextTick, onScopeDispose, ref, watch } from 'vue'
import type { OfficeStudio } from '../../api/office-studio.ts'
import { slicerPanelOwnerCurrent, type SlicerPanelOwner } from './slicer-panel-owner.ts'
import { TIMELINE_DATE_CONDITION_OPTIONS, timelinePanelConditionBetween, timelinePanelDateCalendarInput, timelinePanelDateFilterSummary } from './timeline-range-panel-state.ts'
import { useTimelineDateFilter } from './use-timeline-date-filter.ts'
import TimelineDateInput from './TimelineDateInput.vue'

const props = defineProps<{ open: boolean; studio: OfficeStudio; workbookId: string; sheetId: string; timelineId: string }>()
const emit = defineEmits<{ 'update:open': [open: boolean]; changed: [timelineId: string] }>()
const { token } = theme.useToken(), root = ref<HTMLElement>()
const appearance = computed(() => ({ '--timeline-filter-text': token.value.colorText, '--timeline-filter-muted': token.value.colorTextSecondary }))
let focusGeneration = 0, returnFocus: { owner: SlicerPanelOwner; element: HTMLElement } | undefined
const controls = useTimelineDateFilter({
  get studio() { return props.studio }, get workbookId() { return props.workbookId }, get sheetId() { return props.sheetId },
  get timelineId() { return props.timelineId }, get open() { return props.open }, changed: id => emit('changed', id),
  close: owner => {
    const captured = returnFocus, request = ++focusGeneration
    returnFocus = undefined; emit('update:open', false)
    if (captured && owner) void nextTick(() => {
      if (request !== focusGeneration || props.open || props.studio !== owner.studio || props.workbookId !== owner.target.workbookId
        || props.sheetId !== owner.target.sheetId || captured.owner.studio !== owner.studio || captured.owner.userId !== owner.userId
        || captured.owner.target.workbookId !== owner.target.workbookId || captured.owner.target.sheetId !== owner.target.sheetId
        || !slicerPanelOwnerCurrent(owner, true) || !captured.element.isConnected) return
      captured.element.focus({ preventScroll: true })
    })
  },
})
watch([() => props.open, () => props.studio, () => props.workbookId, () => props.sheetId, () => props.timelineId], (current, previous) => {
  if (current[0] || !previous || current.slice(1).some((value, index) => value !== previous[index + 1])) focusGeneration += 1
  returnFocus = undefined
  if (props.open && controls.snapshot.value && document.activeElement instanceof HTMLElement) returnFocus = { owner: controls.snapshot.value.owner, element: document.activeElement }
}, { immediate: true, flush: 'sync' })
onScopeDispose(() => { focusGeneration += 1; returnFocus = undefined })
const disabled = computed(() => controls.busy.value || controls.composing.value || !controls.canConfigure.value)
const inputDisabled = computed(() => controls.busy.value || !controls.canConfigure.value && !controls.composing.value)
const date1904 = computed(() => controls.snapshot.value?.date1904 ?? false)
const modes = [{ value: 'none', label: '全部日期' }, { value: 'range', label: '日期范围（结束值不含）' }, { value: 'condition', label: '日期条件' }]
const summary = computed(() => controls.snapshot.value ? timelinePanelDateFilterSummary(controls.snapshot.value.state, date1904.value) : '')
const boundLabel = (value: number | undefined): string => value === undefined ? '未提供' : timelinePanelDateCalendarInput(value, date1904.value) || String(value)
const popupContainer = (trigger: HTMLElement): HTMLElement => root.value ?? trigger.parentElement ?? trigger
function keydown(event: KeyboardEvent): void {
  if ((event.isComposing || event.keyCode === 229 || controls.composing.value) && (event.key === 'Enter' || event.key === 'Escape')) { event.preventDefault(); event.stopPropagation() }
}
function conditionOperator(value: unknown): void {
  const option = TIMELINE_DATE_CONDITION_OPTIONS.find(option => option.value === value)
  if (option) controls.setCondition({ operator: option.value })
}
</script>

<template>
  <Modal :open="open" title="时间线日期筛选" :width="520" :root-style="appearance" :footer="null" :mask="{ closable: false }"
    :keyboard="!controls.composing.value" :focus-trigger-after-close="false" destroy-on-hidden @cancel="controls.close">
    <div ref="root">
      <Form class="timeline-date-filter" data-testid="timeline-date-filter-form" :style="appearance" layout="vertical" @submit.prevent="controls.apply"
        @compositionstart="controls.compositionStart" @compositionend="controls.compositionEnd" @keydown="keydown">
        <p class="timeline-filter-note">筛选会同步到同一日期字段的时间线和已连接报表，显示与选择粒度分别保留。</p>
        <p v-if="summary" class="timeline-filter-author" data-testid="timeline-retained-date-filter">当前筛选：{{ summary }}</p>
        <Alert v-if="controls.stale.value" title="日期筛选已过期，当前输入已保留；请重新读取后再保存。" type="warning" show-icon />
        <Alert v-else-if="controls.snapshot.value && !controls.canConfigure.value" title="当前用户没有修改此日期筛选的权限。" type="info" show-icon />
        <Alert v-if="controls.boundsLoading.value" title="正在读取来源中的真实日期边界…" type="info" show-icon />
        <p v-else-if="controls.bounds.value" class="timeline-filter-note" data-testid="timeline-inclusive-date-bounds">来源最早日期（含）：{{ boundLabel(controls.bounds.value.minDate) }}；最晚日期（含）：{{ boundLabel(controls.bounds.value.maxDate) }}。{{ date1904 ? '1904' : '1900' }} 日期系统。</p>
        <Alert v-if="controls.boundsError.value" :title="controls.boundsError.value" type="warning" show-icon />
        <Alert v-else-if="!controls.boundsLoading.value && controls.bounds.value && !controls.hasDateBounds.value" title="来源没有可靠的日期边界；当前保留的筛选仍显示在上方。" type="info" show-icon />
        <FormItem label="筛选方式"><Select :value="controls.draft.value.mode" :options="modes" aria-label="时间线筛选方式" :disabled="disabled" :get-popup-container="popupContainer" @update:value="controls.setMode" /></FormItem>
        <template v-if="controls.draft.value.mode === 'range'">
          <TimelineDateInput :key="`${controls.readSession.value}:range-start`" :value="controls.draft.value.range.startDate" :date1904="date1904" label="开始日期（含）" aria-label="时间线开始日期" :disabled="inputDisabled" :actions-disabled="disabled" :popup-container="popupContainer" @update:value="controls.setRange({ startDate: $event })" />
          <TimelineDateInput :key="`${controls.readSession.value}:range-end`" :value="controls.draft.value.range.endDate" :date1904="date1904" label="结束日期（不含）" aria-label="时间线结束日期" :disabled="inputDisabled" :actions-disabled="disabled" :popup-container="popupContainer" @update:value="controls.setRange({ endDate: $event })" />
          <FormItem :label="null"><Checkbox :checked="controls.draft.value.range.useWholeDay ?? false" :disabled="disabled" @update:checked="controls.setRange({ useWholeDay: $event })">按整日比较</Checkbox></FormItem>
          <p class="timeline-filter-note">按整日比较时开始值向当天起点取整、结束值向上取整；结束值仍不包含。选择同一天须把结束值设到次日。</p>
          <Button html-type="button" size="small" :disabled="disabled || controls.boundsLoading.value || !controls.hasDateBounds.value" @click="controls.fillSourceDays">填入来源日期的整日范围</Button>
        </template>
        <template v-else-if="controls.draft.value.mode === 'condition'">
          <FormItem label="日期条件"><Select :value="controls.draft.value.condition.operator" :options="TIMELINE_DATE_CONDITION_OPTIONS" aria-label="时间线日期条件" :disabled="disabled" :get-popup-container="popupContainer" @update:value="conditionOperator" /></FormItem>
          <TimelineDateInput :key="`${controls.readSession.value}:condition-first`" :value="controls.draft.value.condition.first" :date1904="date1904" :label="timelinePanelConditionBetween(controls.draft.value.condition.operator) ? '开始日期（含）' : '比较日期'" aria-label="时间线条件日期" :disabled="inputDisabled" :actions-disabled="disabled" :popup-container="popupContainer" @update:value="controls.setCondition({ first: $event })" />
          <TimelineDateInput v-if="timelinePanelConditionBetween(controls.draft.value.condition.operator)" :key="`${controls.readSession.value}:condition-second`" :value="controls.draft.value.condition.second" :date1904="date1904" label="结束日期（含）" aria-label="时间线条件结束日期" :disabled="inputDisabled" :actions-disabled="disabled" :popup-container="popupContainer" @update:value="controls.setCondition({ second: $event })" />
          <FormItem :label="null"><Checkbox :checked="controls.draft.value.condition.useWholeDay ?? false" :disabled="disabled" @update:checked="controls.setCondition({ useWholeDay: $event })">按整日比较</Checkbox></FormItem>
          <p class="timeline-filter-note">日期条件保留单值或双值比较；介于条件包含两端。</p>
        </template>
        <Alert v-if="controls.validation.value && controls.draft.value.mode !== 'none'" :title="controls.validation.value" type="warning" show-icon />
        <Alert v-if="controls.error.value" :title="controls.error.value" type="error" show-icon />
        <div class="timeline-filter-actions">
          <Button html-type="button" :disabled="controls.busy.value || controls.composing.value" @click="controls.read">重新读取</Button><span />
          <Button html-type="button" :disabled="controls.composing.value" @click="controls.close">取消</Button>
          <Button html-type="button" type="primary" :loading="controls.busy.value" :disabled="!controls.canApply.value" @click="controls.apply">确定</Button>
        </div>
      </Form>
    </div>
  </Modal>
</template>

<style scoped>
.timeline-date-filter { color: var(--timeline-filter-text); }
.timeline-filter-note { color: var(--timeline-filter-muted); font-size: 12px; line-height: 1.6; margin: 0 0 12px; }
.timeline-filter-author { font-size: 12px; line-height: 1.6; }
.timeline-filter-actions { display: flex; align-items: center; flex-wrap: wrap; gap: 8px; margin-top: 16px; }
.timeline-filter-actions > span { flex: 1; }
</style>
