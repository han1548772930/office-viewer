<script setup lang="ts">
import { Button, Checkbox, DatePicker, FormItem, Input, Select } from 'antdv-next'
import dayjs, { type Dayjs } from 'dayjs'
import { computed, ref } from 'vue'
import { PIVOT_CHART_DATE_CONDITIONS, PIVOT_CHART_LABEL_CONDITIONS, pivotChartDateConditionBetween,
  pivotChartFilterDateCalendarInput, pivotChartFilterDateCompatibilityDay, pivotChartFilterDateFromCalendarInput,
  pivotChartLabelConditionBetween, type PivotChartDateConditionDraft, type PivotChartFieldConditionsDraft,
  type PivotChartLabelConditionDraft } from '../services/pivot-chart-field-filter-condition.ts'

const props = defineProps<{ kind: 'date' | 'label'; draft: PivotChartFieldConditionsDraft; disabled: boolean;
  actionsDisabled: boolean; popupContainer: (trigger: HTMLElement) => HTMLElement; subjectLabel?: string }>()
const emit = defineEmits<{ 'date-change': [patch: Partial<PivotChartDateConditionDraft>]; 'label-change': [patch: Partial<PivotChartLabelConditionDraft>] }>()
type Operand = 'first' | 'second'
const date1904 = computed(() => props.draft.capabilities.date1904)
const subject = computed(() => props.subjectLabel ?? '透视图')
// Ant's calendar uses real Gregorian dates and cannot represent Excel's serial day 60.
// Pin this minimal Ant Input adapter while typing; an incomplete value must not unmount its IME input.
const compatibility = ref({ first: pivotChartFilterDateCompatibilityDay(props.draft.date.first, date1904.value),
  second: pivotChartFilterDateCompatibilityDay(props.draft.date.second, date1904.value) })
const compatibilityText = ref({ first: pivotChartFilterDateCalendarInput(props.draft.date.first, date1904.value),
  second: pivotChartFilterDateCalendarInput(props.draft.date.second, date1904.value) })
const operands = computed<Operand[]>(() => pivotChartDateConditionBetween(props.draft.date.operator) ? ['first', 'second'] : ['first'])
const dateFormat = computed(() => props.draft.date.useWholeDay ? 'YYYY-MM-DD' : 'YYYY-MM-DD HH:mm:ss.SSS')
function pickerValue(operand: Operand): Dayjs | undefined {
  const value = pivotChartFilterDateCalendarInput(props.draft.date[operand], date1904.value)
  const parsed = value ? dayjs(value) : undefined
  return parsed?.isValid() ? parsed : undefined
}
function dateOperator(value: unknown): void {
  const operator = PIVOT_CHART_DATE_CONDITIONS.find(option => option.value === value)?.value
  if (operator) emit('date-change', { operator })
}
function labelOperator(value: unknown): void {
  const operator = PIVOT_CHART_LABEL_CONDITIONS.find(option => option.value === value)?.value
  if (operator) emit('label-change', { operator })
}
function chooseDate(operand: Operand, value: Dayjs | Dayjs[] | null): void {
  if (props.disabled || Array.isArray(value)) return
  emit('date-change', { [operand]: value?.isValid() ? pivotChartFilterDateFromCalendarInput(value.format('YYYY-MM-DD HH:mm:ss.SSS'), date1904.value) : undefined })
}
function calendarText(operand: Operand, value: string): void {
  if (props.disabled) return
  compatibilityText.value[operand] = value
  emit('date-change', { [operand]: pivotChartFilterDateFromCalendarInput(value, date1904.value) })
}
function chooseRealDate(operand: Operand): void {
  if (props.actionsDisabled) return
  compatibility.value[operand] = false; emit('date-change', { [operand]: undefined })
}
function wholeDay(value: unknown): void { if (typeof value === 'boolean') emit('date-change', { useWholeDay: value }) }
</script>

<template>
  <div class="field-condition-editor" :data-pivot-filter-condition="kind">
    <template v-if="kind === 'date'">
      <FormItem label="日期条件"><Select :value="draft.date.operator" :options="PIVOT_CHART_DATE_CONDITIONS" :aria-label="`${subject}日期条件`"
        :disabled="actionsDisabled" :get-popup-container="popupContainer" @update:value="dateOperator" /></FormItem>
      <FormItem v-for="operand in operands" :key="operand" :label="operands.length === 1 ? '日期' : operand === 'first' ? '开始日期' : '结束日期'">
        <template v-if="compatibility[operand]">
          <Input :value="compatibilityText[operand]" :aria-label="`${subject}${operand === 'first' ? '筛选日期' : '筛选结束日期'}`" :disabled="disabled"
            placeholder="YYYY-MM-DD HH:mm:ss.SSS" @update:value="calendarText(operand, $event)" />
          <span class="compatibility-note">保留 Excel 1900-02-29 日期。</span>
          <Button type="link" size="small" :disabled="actionsDisabled" @click="chooseRealDate(operand)">选择其他日期</Button>
        </template>
        <DatePicker v-else :value="pickerValue(operand)" :format="dateFormat" :show-time="draft.date.useWholeDay ? false : { format: 'HH:mm:ss' }"
          :aria-label="`${subject}${operand === 'first' ? '筛选日期' : '筛选结束日期'}`" :disabled="disabled" :get-popup-container="popupContainer"
          :need-confirm="false" allow-clear @change="value => chooseDate(operand, value)" />
      </FormItem>
      <FormItem :label="null"><Checkbox :checked="draft.date.useWholeDay" :disabled="actionsDisabled" @update:checked="wholeDay">按整天比较</Checkbox></FormItem>
    </template>
    <template v-else>
      <FormItem label="标签条件"><Select :value="draft.label.operator" :options="PIVOT_CHART_LABEL_CONDITIONS" :aria-label="`${subject}标签条件`"
        :disabled="actionsDisabled" :get-popup-container="popupContainer" @update:value="labelOperator" /></FormItem>
      <FormItem label="比较标签"><Input :value="draft.label.first" :aria-label="`${subject}条件标签`" :disabled="disabled" allow-clear
        @update:value="emit('label-change', { first: $event })" /></FormItem>
      <FormItem v-if="pivotChartLabelConditionBetween(draft.label.operator)" label="结束标签"><Input :value="draft.label.second" :aria-label="`${subject}条件结束标签`"
        :disabled="disabled" allow-clear @update:value="emit('label-change', { second: $event })" /></FormItem>
    </template>
  </div>
</template>

<style scoped>
.field-condition-editor { display: grid; gap: 10px; }
.field-condition-editor :deep(.ant-picker), .field-condition-editor :deep(.ant-select) { width: 100%; }
.compatibility-note { font-size: 12px; opacity: .75; }
</style>
