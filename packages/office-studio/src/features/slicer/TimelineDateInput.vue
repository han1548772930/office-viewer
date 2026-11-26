<script setup lang="ts">
import { Button, DatePicker, FormItem, Input, InputNumber } from 'antdv-next'
import dayjs, { type Dayjs } from 'dayjs'
import { computed, ref } from 'vue'
import { timelinePanelDateCalendarInput, timelinePanelDateCompatibilityDay, timelinePanelDateFromCalendarInput } from './timeline-range-panel-state.ts'

const props = defineProps<{ value?: number; date1904: boolean; label: string; ariaLabel: string; disabled: boolean;
  actionsDisabled: boolean; popupContainer?: (trigger: HTMLElement) => HTMLElement }>()
const emit = defineEmits<{ 'update:value': [value: number | undefined] }>()
const compatibility = ref(timelinePanelDateCompatibilityDay(props.value, props.date1904))
const compatibilityText = ref(timelinePanelDateCalendarInput(props.value, props.date1904))
const numeric = ref(props.value !== undefined && !compatibilityText.value)
const pickerValue = computed<Dayjs | undefined>(() => {
  const text = timelinePanelDateCalendarInput(props.value, props.date1904), value = text ? dayjs(text) : undefined
  return value?.isValid() ? value : undefined
})
function chooseDate(value: Dayjs | Dayjs[] | null): void {
  if (props.disabled || Array.isArray(value)) return
  emit('update:value', value?.isValid() ? timelinePanelDateFromCalendarInput(value.format('YYYY-MM-DD HH:mm:ss.SSS'), props.date1904) : undefined)
}
function calendarText(value: string): void {
  if (props.disabled) return
  compatibilityText.value = value
  emit('update:value', timelinePanelDateFromCalendarInput(value, props.date1904))
}
function chooseCalendar(): void {
  if (props.actionsDisabled) return
  if (compatibility.value) { compatibility.value = false; emit('update:value', undefined) }
  numeric.value = false
}
function serialValue(value: unknown): void { if (!props.disabled) emit('update:value', typeof value === 'number' ? value : undefined) }
</script>

<template>
  <FormItem class="timeline-date-input" :label="label">
    <InputNumber v-if="numeric" :value="value" :aria-label="`${ariaLabel}精确日期值`" :disabled="disabled" @update:value="serialValue" />
    <template v-else-if="compatibility">
      <Input :value="compatibilityText" :aria-label="ariaLabel" :disabled="disabled" placeholder="YYYY-MM-DD HH:mm:ss.SSS" @update:value="calendarText" />
      <span class="timeline-date-note">保留 Excel 1900-02-29 日期。</span>
    </template>
    <DatePicker v-else :value="pickerValue" format="YYYY-MM-DD HH:mm:ss.SSS" :show-time="{ format: 'HH:mm:ss' }"
      :aria-label="ariaLabel" :disabled="disabled" :get-popup-container="popupContainer" :need-confirm="false" allow-clear @change="chooseDate" />
    <div class="timeline-date-actions">
      <Button v-if="!numeric" type="link" html-type="button" size="small" :disabled="actionsDisabled" @click="numeric = true">精确日期值</Button>
      <Button v-if="numeric || compatibility" type="link" html-type="button" size="small" :disabled="actionsDisabled" @click="chooseCalendar">{{ compatibility ? '选择其他日期' : '日历输入' }}</Button>
    </div>
    <span v-if="numeric" class="timeline-date-note">日期值可保留一天内的精确时间，切换输入方式不会改写未修改的值。</span>
  </FormItem>
</template>

<style scoped>
.timeline-date-input :deep(.ant-picker), .timeline-date-input :deep(.ant-input-number) { width: 100%; }
.timeline-date-actions { display: flex; gap: 8px; }
.timeline-date-note { font-size: 12px; opacity: .75; }
</style>
