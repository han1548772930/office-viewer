<script setup lang="ts">
import { computed, inject, ref } from 'vue'
import { Alert, Button, DatePicker, Form, FormItem, Input, InputNumber, TimePicker, theme } from 'antdv-next'
import dayjs, { type Dayjs } from 'dayjs'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'
import { parseSheetNumberEdit, sheetNumberEditText } from '../../features/sheet-workbench/sheet-number-format-edit.ts'
import type { SheetCellInputCommandResult } from '../../features/sheet-workbench/sheet-cell-input-command.ts'
import type { SheetDateInputChange, SheetDateInputState } from '../../features/sheet-validation/validation-date-input.ts'

const props = defineProps<{ state: SheetDateInputState; disabled?: boolean; message?: string; region: string; currentYear: number;
  submit: (change: SheetDateInputChange) => SheetCellInputCommandResult }>()
const emit = defineEmits<{ cancel: []; composition: [value: boolean] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback), { token } = theme.useToken()
const composing = ref(false), busy = ref(false), error = ref(''), dateDirty = ref(false), timeDirty = ref(false), serialDirty = ref(false)
const parts = props.state.parts, today = dayjs(), original = props.state.originalSerial
const exceptional = parts && (parts.day === 0 || parts.year === 1900 && parts.month === 2 && parts.day === 29)
const dateAnchor = parts ? today.year(parts.year).month(parts.month - 1).date(parts.day === 0 ? 1 : exceptional ? 28 : parts.day)
  .hour(0).minute(0).second(0).millisecond(0) : today.startOf('day')
const dateValue = ref<Dayjs | null>(parts && !exceptional ? dateAnchor : null)
const minDate = today.year(props.state.date1904 ? 1904 : 1900).month(0).date(1).startOf('day')
const maxDate = today.year(9999).month(11).date(31).startOf('day')
const timeValue = ref(today.startOf('day').hour(parts?.hours ?? 0).minute(parts?.minutes ?? 0).second(parts?.seconds ?? 0)
  .millisecond(Math.floor((parts?.fractionalSecond ?? 0) * 1000)))
const fraction = ref(parts?.fractionalSecond ?? 0), serialValue = ref<number | null>(original)
const durationText = ref(sheetNumberEditText(original ?? 0, props.state.numberFormat, { date1904: props.state.date1904,
  region: props.region, currentYear: props.currentYear }) ?? String(original ?? 0))
const numeric = props.state.mode !== 'duration' && original !== null && !parts
const clockVisible = props.state.mode === 'time' || props.state.mode === 'datetime'
const exceptionalLabel = parts && exceptional ? `${parts.year}-${String(parts.month).padStart(2, '0')}-${String(parts.day).padStart(2, '0')}` : ''
const blocked = computed(() => props.disabled || busy.value || composing.value)
const message = computed(() => props.message || error.value)
const getPopupContainer = (trigger: HTMLElement) => trigger.closest<HTMLElement>('.sheet-date-picker-form') ?? trigger.parentElement ?? trigger

function changeDate(value: Dayjs | null): void { if (!props.disabled && !composing.value && value?.isValid()) { dateValue.value = value; dateDirty.value = true; error.value = '' } }
function changeTime(value: Dayjs | null): void { if (!props.disabled && !composing.value && value?.isValid()) { timeValue.value = value; timeDirty.value = true; error.value = '' } }
function changeFraction(value: number | string | null): void {
  if (!props.disabled && !composing.value && typeof value === 'number' && Number.isFinite(value)) { fraction.value = value; timeDirty.value = true; error.value = '' }
}
function changeSerial(value: number | string | null): void {
  if (!props.disabled && !composing.value) { serialValue.value = typeof value === 'number' ? value : null; serialDirty.value = true; error.value = '' }
}
function composition(value: boolean): void { composing.value = value; emit('composition', value) }
function draftChange(): SheetDateInputChange {
  if (numeric) {
    if (!serialDirty.value || serialValue.value === null) throw new Error(t('sheet.validation.date.correct-number', '请输入可以在当前日历中表示的日期数值。'))
    return { kind: 'serial', serial: serialValue.value }
  }
  if (props.state.mode === 'duration') {
    if (!serialDirty.value && original !== null) return { kind: 'unchanged' }
    const parsed = parseSheetNumberEdit(durationText.value, original, props.state.numberFormat,
      { date1904: props.state.date1904, region: props.region, currentYear: props.currentYear })
    if (!parsed || typeof parsed.value !== 'number') throw new Error(t('sheet.validation.date.duration-invalid', '请输入有效的持续时间。'))
    return { kind: 'serial', serial: parsed.value }
  }
  if (!dateDirty.value && !timeDirty.value && original !== null) return { kind: 'unchanged' }
  const date = dateValue.value ?? dateAnchor, time = timeValue.value
  const dateParts = { year: date.year(), month: date.month() + 1, day: date.date() }
  const timeParts = { hours: time.hour(), minutes: time.minute(), seconds: time.second(), fractionalSecond: fraction.value }
  if (props.state.mode === 'time') return { kind: 'time', ...timeParts }
  if (props.state.mode === 'datetime' && (dateDirty.value && timeDirty.value || !parts && timeDirty.value)) return { kind: 'date-time', ...dateParts, ...timeParts }
  if (timeDirty.value) return { kind: 'time', ...timeParts }
  return { kind: 'date', ...dateParts }
}
function commit(): void {
  if (blocked.value) return
  busy.value = true; error.value = ''
  try { const result = props.submit(draftChange()); if (!result.accepted) error.value = result.message ?? t('sheet.validation.date.save-failed', '日期或时间未能保存，草稿已保留。') }
  catch (cause) { error.value = cause instanceof Error ? cause.message : t('sheet.validation.date.save-failed', '日期或时间未能保存，草稿已保留。') }
  finally { busy.value = false }
}
function keydown(event: KeyboardEvent): void {
  if (composing.value || event.isComposing || event.keyCode === 229) { event.stopPropagation(); return }
  if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); if (!busy.value) emit('cancel') }
}
</script>

<template>
  <Form class="sheet-date-picker-form" layout="vertical" :style="{ '--date-border': token.colorBorderSecondary }" :disabled="disabled || busy"
    @submit.prevent @keydown.capture="keydown" @compositionstart.capture="composition(true)" @compositionend.capture="composition(false)">
    <Alert v-if="exceptionalLabel" type="info" :title="`${t('sheet.validation.date.compatibility', 'Excel 兼容日期')}：${exceptionalLabel}`" show-icon />
    <Alert v-if="numeric" type="info" :title="t('sheet.validation.date.out-of-calendar', '当前数值无法在日历中表示，请修改日期数值。')" show-icon />
    <FormItem v-if="numeric" :label="t('sheet.validation.date.number', '日期数值')">
      <InputNumber :value="serialValue" :min="0" :controls="false" :aria-label="t('sheet.validation.date.number', '日期数值')" @update:value="changeSerial" />
    </FormItem>
    <FormItem v-else-if="state.mode === 'duration'" :label="t('sheet.validation.date.duration', '持续时间')">
      <Input v-model:value="durationText" :aria-label="t('sheet.validation.date.duration', '持续时间')" @update:value="serialDirty = true; error = ''" />
    </FormItem>
    <template v-else>
      <FormItem v-if="state.mode !== 'time'" :label="t('sheet.validation.date.calendar', '日期')">
        <DatePicker :value="dateValue" :default-picker-value="dateAnchor" :min-date="minDate" :max-date="maxDate" :allow-clear="false" format="YYYY-MM-DD" :get-popup-container="getPopupContainer"
          :aria-label="t('sheet.validation.date.calendar', '日期')" @change="changeDate" />
      </FormItem>
      <FormItem v-if="clockVisible" :label="t('sheet.validation.date.clock', '时间')">
        <TimePicker :value="timeValue" :allow-clear="false" :need-confirm="false" format="HH:mm:ss" :get-popup-container="getPopupContainer"
          :aria-label="t('sheet.validation.date.clock', '时间')" @change="changeTime" />
      </FormItem>
      <FormItem v-if="clockVisible" :label="t('sheet.validation.date.fraction', '小数秒')">
        <InputNumber :value="fraction" :min="0" :max="0.9999999999999999" :step="0.001" :controls="false"
          :aria-label="t('sheet.validation.date.fraction', '小数秒')" @update:value="changeFraction" />
      </FormItem>
    </template>
    <Alert v-if="message" type="error" :title="message" show-icon />
    <div class="sheet-date-picker-actions">
      <Button :disabled="busy || composing" @click="emit('cancel')">{{ t('common.cancel', '取消') }}</Button>
      <Button type="primary" :disabled="blocked || numeric && !serialDirty" @click="commit">{{ t('common.confirm', '确定') }}</Button>
    </div>
  </Form>
</template>

<style scoped>
.sheet-date-picker-form { position: relative; width: 288px; max-width: calc(100vw - 48px); }
.sheet-date-picker-form :deep(.ant-form-item) { margin-block: 12px; }
.sheet-date-picker-form :deep(.ant-picker), .sheet-date-picker-form :deep(.ant-input-number) { width: 100%; }
.sheet-date-picker-actions { display: flex; justify-content: flex-end; gap: 8px; padding-block-start: 12px; margin-block-start: 12px; border-block-start: 1px solid var(--date-border); }
</style>
