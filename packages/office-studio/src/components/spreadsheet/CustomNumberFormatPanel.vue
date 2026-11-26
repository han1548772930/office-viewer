<script setup lang="ts">
import { computed, inject, nextTick, ref, watch } from 'vue'
import { Alert, Button, Card, Form, FormItem, Input, InputNumber, Menu, Select, Tooltip, theme } from 'antdv-next'
import type { MenuInfo } from 'antdv-next'
import { X } from 'lucide-vue-next'
import { previewCustomNumberFormat, validateCustomNumberFormat } from '../../features/sheet-workbench/custom-number-format.ts'
import { renderSheetNumberFormat } from '../../features/sheet-workbench/number-format-render.ts'
import { setNumberFormatDecimalPlaces } from '../../features/sheet-workbench/number-format-pattern.ts'
import {
  createSheetNumberFormatDraft, isSheetNumberFormatCategory, orderNumberFormatPreferences, readSheetNumberFormatPreferences,
  rememberSheetNumberFormat, SHEET_DATE_FORMAT_PREVIEW_VALUE, SHEET_GROUPED_NUMBER_PATTERNS, SHEET_NUMBER_FORMAT_CATEGORIES,
  SHEET_NUMBER_FORMAT_PANEL_CURRENCIES, sheetCurrencyNumberPatterns, sheetCustomNumberPatterns, sheetDateNumberPatterns,
  sheetNumberFormatDraftPattern, writeSheetNumberFormatPreferences,
} from '../../features/sheet-workbench/number-format-panel.ts'
import type { SheetNumberFormatPreferences } from '../../features/sheet-workbench/number-format-panel.ts'
import type { CellScalar } from '../../domain/model.ts'
import type { OfficeLocale } from '../../i18n/regional-settings.ts'
import { OfficeDirectionKey, OfficeLocaleKey, OfficeTranslatorKey, useOfficeRegion } from '../../i18n/office-i18n.ts'

const props = withDefaults(defineProps<{
  open: boolean; currentPattern: string; selectionKey?: string; sampleValue?: CellScalar; date1904?: boolean; disabled?: boolean
}>(), { disabled: false, date1904: false })
const emit = defineEmits<{
  'update:open': [value: boolean]; apply: [pattern: string, complete: (applied: boolean) => void]; preview: [pattern: string | undefined]; 'preference-error': [message: string]
}>()
const region = useOfficeRegion()
const locale = inject(OfficeLocaleKey, ref<OfficeLocale>('zh-CN'))
const direction = inject(OfficeDirectionKey, ref<'ltr' | 'rtl'>('ltr'))
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const { token } = theme.useToken()
const themeStyle = computed(() => ({ '--numfmt-border': token.value.colorBorderSecondary, '--numfmt-secondary': token.value.colorTextSecondary }))
const preferences = ref<SheetNumberFormatPreferences>({ history: [], currencies: [] })
const draft = ref(createSheetNumberFormatDraft('', locale.value, region.value, []))
const pristine = ref(true), composing = ref(false)
const closeButton = ref<{ focus: (options?: FocusOptions) => void }>()
const storageFailed = ref(false)
const applicationFailed = ref(false)
const currencies = computed(() => orderNumberFormatPreferences(
  [...SHEET_NUMBER_FORMAT_PANEL_CURRENCIES, draft.value.currency], preferences.value.currencies).map(value => ({ label: value, value })))
const categoryOptions = computed(() => SHEET_NUMBER_FORMAT_CATEGORIES.map(value => ({
  value, label: t('sheet.number-format.' + value, { general: '常规', accounting: '会计', currency: '货币', date: '日期', grouped: '千分位符', custom: '自定义格式' }[value]),
})))
const pattern = computed(() => pristine.value ? props.currentPattern || 'General' : sheetNumberFormatDraftPattern(draft.value))
const validation = computed(() => {
  try { return { pattern: validateCustomNumberFormat(pattern.value).pattern, error: '' } }
  catch (reason) { return { pattern: undefined, error: reason instanceof Error ? reason.message : t('sheet.number-format.invalid', '格式代码无效。') } }
})
const example = computed(() => validation.value.pattern ? previewCustomNumberFormat(validation.value.pattern,
  typeof props.sampleValue === 'number' ? props.sampleValue : 12345678, region.value, props.date1904) : undefined)
const currencyCategory = computed(() => draft.value.category === 'accounting' || draft.value.category === 'currency')
const negativeOptions = computed(() => {
  const patterns = draft.value.category === 'currency' ? sheetCurrencyNumberPatterns(draft.value.currency) : SHEET_GROUPED_NUMBER_PATTERNS
  return patterns.map((value, index) => {
    const sample = renderSheetNumberFormat(-1234.56, setNumberFormatDecimalPlaces(value, draft.value.decimals), props.date1904, region.value)
    return { key: String(index), label: sample?.text.trim() ?? value, color: sample?.color, disabled: props.disabled }
  })
})
const dateOptions = computed(() => sheetDateNumberPatterns(locale.value, draft.value.category === 'date' ? draft.value.pattern : undefined).map(value => ({
  key: value, label: previewCustomNumberFormat(value, SHEET_DATE_FORMAT_PREVIEW_VALUE, locale.value) ?? value, disabled: props.disabled,
})))
const customOptions = computed(() => sheetCustomNumberPatterns(locale.value, preferences.value.history).map(value => ({ key: value, label: value, disabled: props.disabled })))
const description = computed(() => t('sheet.number-format.description.' + draft.value.category, {
  general: '常规格式不包含任何特定的数字格式。', accounting: '会计数字格式可对齐货币符号和小数点。',
  currency: '货币格式用于表示一般货币数值。', date: '日期格式将日期和时间系列数值显示为日期值。',
  grouped: '千分位符格式用于一般数字的表示。', custom: '根据现有格式生成自定义数字格式。',
}[draft.value.category]))

function publishPreview(): void {
  if (composing.value) return
  applicationFailed.value = false
  if (!props.open || props.disabled) { emit('preview', undefined); return }
  emit('preview', validation.value.pattern)
}
function onCustomPatternEnter(event: KeyboardEvent): void {
  if (event.defaultPrevented || event.isComposing || event.keyCode === 229 || composing.value) return
  publishPreview()
}
function changeCategory(value: unknown): void {
  if (props.disabled || !isSheetNumberFormatCategory(value)) return
  draft.value = createSheetNumberFormatDraft(props.currentPattern, locale.value, region.value, preferences.value.currencies, value)
  pristine.value = false
  publishPreview()
}
function changeDecimals(value: number | string | null): void {
  const decimals = value ?? 0
  if (props.disabled || typeof decimals !== 'number' || !Number.isFinite(decimals)) return
  draft.value.decimals = Math.max(0, Math.min(20, Math.floor(decimals)))
  pristine.value = false
  publishPreview()
}
function changeCurrency(value: unknown): void {
  if (props.disabled || typeof value !== 'string' || !currencies.value.some(option => option.value === value)) return
  draft.value.currency = value
  draft.value.negative = 0
  draft.value.pattern = ''
  pristine.value = false
  publishPreview()
}
function chooseNegative({ key }: MenuInfo): void {
  if (draft.value.category !== 'currency' && draft.value.category !== 'grouped') return
  const index = Number(key)
  if (props.disabled || !Number.isSafeInteger(index) || index < 0 || index >= negativeOptions.value.length) return
  draft.value.negative = index
  draft.value.pattern = ''
  pristine.value = false
  publishPreview()
}
function choosePattern({ key }: MenuInfo): void {
  if (props.disabled) return
  const selected = String(key)
  const options = draft.value.category === 'date' ? dateOptions.value : draft.value.category === 'custom' ? customOptions.value : []
  if (!options.some(option => option.key === selected)) return
  draft.value.pattern = selected
  pristine.value = false
  publishPreview()
}
function changeCustomPattern(value: string): void {
  if (props.disabled || draft.value.category !== 'custom') return
  applicationFailed.value = false
  draft.value.pattern = value
  pristine.value = false
}
function close(): void {
  emit('preview', undefined)
  emit('update:open', false)
}
function onKeydown(event: KeyboardEvent): void {
  if (event.defaultPrevented || event.isComposing || event.keyCode === 229 || composing.value) return
  if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); close() }
}
function apply(): void {
  if (props.disabled || composing.value || !validation.value.pattern) return
  const selected = validation.value.pattern
  let applied = false
  emit('apply', selected, result => { applied = result })
  if (!applied) { applicationFailed.value = true; return }
  preferences.value = rememberSheetNumberFormat(preferences.value, selected, draft.value.category === 'custom')
  try { if (typeof window !== 'undefined') writeSheetNumberFormatPreferences(window.localStorage, preferences.value) }
  catch { emit('preference-error', t('sheet.number-format.preference-error', '格式已应用，但此浏览器未能保存格式偏好。')) }
  close()
}
watch(() => [props.open, props.currentPattern, props.selectionKey] as const, async ([open], previous, onCleanup) => {
  let active = true
  onCleanup(() => { active = false })
  emit('preview', undefined)
  composing.value = false
  if (!open) return
  storageFailed.value = false
  applicationFailed.value = false
  try { preferences.value = readSheetNumberFormatPreferences(typeof window === 'undefined' ? undefined : window.localStorage) }
  catch { preferences.value = { history: [], currencies: [] }; storageFailed.value = true }
  draft.value = createSheetNumberFormatDraft(props.currentPattern, locale.value, region.value, preferences.value.currencies)
  pristine.value = true
  if (previous?.[0]) return
  await nextTick()
  if (active && props.open) closeButton.value?.focus({ preventScroll: true })
}, { immediate: true })
watch(() => props.disabled, disabled => { if (disabled) emit('preview', undefined) })
</script>

<template>
  <Card v-if="open" class="sheet-number-format-panel" size="small" :title="t('sheet.number-format.title', '数字格式')" :style="themeStyle"
    :dir="direction" role="complementary" :aria-label="t('sheet.number-format.title', '数字格式')" @keydown="onKeydown">
    <template #extra>
      <Tooltip :title="t('common.close', '关闭')"><Button ref="closeButton" type="text" size="small" :aria-label="t('common.close', '关闭')" @click="close"><X :size="16" aria-hidden="true" /></Button></Tooltip>
    </template>
    <div class="number-format-panel-body">
      <Form layout="vertical" class="number-format-form" @submit.prevent>
        <FormItem :label="t('sheet.number-format.type', '格式类型')">
          <Select :value="draft.category" :options="categoryOptions" :disabled="disabled" :aria-label="t('sheet.number-format.type', '格式类型')" @change="changeCategory" />
        </FormItem>
        <div v-if="currencyCategory || draft.category === 'grouped'" class="number-format-fields">
          <FormItem :label="t('sheet.number-format.decimals', '小数位数')"><InputNumber :value="draft.decimals" :min="0" :max="20" :step="1" :precision="0" :disabled="disabled" :aria-label="t('sheet.number-format.decimals', '小数位数')" @change="changeDecimals" /></FormItem>
          <FormItem v-if="currencyCategory" :label="t('sheet.number-format.currency-type', '货币类型')"><Select :value="draft.currency" :options="currencies" :disabled="disabled" :aria-label="t('sheet.number-format.currency-type', '货币类型')" @change="changeCurrency" /></FormItem>
        </div>
        <FormItem v-if="draft.category === 'currency' || draft.category === 'grouped'" :label="t('sheet.number-format.negative', '负数类型')">
          <Menu class="number-format-options" :items="negativeOptions" :selected-keys="[String(draft.negative)]" :aria-label="t('sheet.number-format.negative', '负数类型')" @click="chooseNegative">
            <template #labelRender="item"><span class="number-format-choice" :style="{ color: negativeOptions.find(option => option.key === String(item.key))?.color }">{{ item.label }}</span></template>
          </Menu>
        </FormItem>
        <FormItem v-if="draft.category === 'date'" :label="t('sheet.number-format.date-type', '日期类型')">
          <Menu class="number-format-options" :items="dateOptions" :selected-keys="[draft.pattern]" :aria-label="t('sheet.number-format.date-type', '日期类型')" @click="choosePattern" />
        </FormItem>
        <template v-if="draft.category === 'custom'">
          <FormItem :label="t('sheet.number-format.custom', '自定义格式')" :validate-status="validation.error ? 'error' : undefined" :help="validation.error || undefined">
            <Input :value="draft.pattern || (pristine ? currentPattern : '')" :disabled="disabled" :maxlength="255" :aria-label="t('sheet.number-format.pattern-aria', '自定义数字格式代码')"
              @update:value="changeCustomPattern" @blur="publishPreview" @press-enter="onCustomPatternEnter" @compositionstart="composing = true" @compositionend="composing = false" />
          </FormItem>
          <Menu class="number-format-options number-format-custom-options" :items="customOptions" :selected-keys="[draft.pattern]" :aria-label="t('sheet.number-format.custom', '自定义格式')" @click="choosePattern" />
        </template>
        <p class="number-format-description">{{ description }}</p>
        <div class="number-format-preview" aria-live="polite"><span>{{ t('sheet.number-format.preview', '预览') }}</span><strong>{{ example ?? t('sheet.number-format.unavailable', '无法预览') }}</strong></div>
        <Alert v-if="validation.error && draft.category !== 'custom'" type="error" :title="validation.error" show-icon />
        <Alert v-if="applicationFailed" type="error" :title="t('sheet.number-format.apply-error', '数字格式未应用，请检查选区后重试。')" show-icon />
        <Alert v-if="storageFailed" type="warning" :title="t('sheet.number-format.preference-read-error', '此浏览器未能读取格式偏好，可继续设置数字格式。')" show-icon />
      </Form>
    </div>
    <div class="number-format-panel-footer">
      <Button @click="close">{{ t('sheet.number-format.cancel', '取消') }}</Button>
      <Button type="primary" :disabled="disabled || composing || Boolean(validation.error)" @click="apply">{{ t('sheet.number-format.confirm', '确认') }}</Button>
    </div>
  </Card>
</template>

<style scoped>
.sheet-number-format-panel { display: flex; height: 100%; min-width: 0; flex-direction: column; border-radius: 0; border-block: 0; border-inline-end: 0; }
.sheet-number-format-panel :deep(.ant-card-body) { display: flex; min-height: 0; flex: 1; flex-direction: column; padding: 16px; }
.number-format-panel-body { min-height: 0; flex: 1; overflow: auto; }
.number-format-form { display: flex; flex-direction: column; gap: 16px; }
.number-format-form :deep(.ant-form-item) { margin-bottom: 0; }
.number-format-fields { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px; }
.number-format-fields :deep(.ant-input-number), .number-format-fields :deep(.ant-select) { width: 100%; min-width: 0; }
.number-format-options { max-height: 400px; overflow: auto; border: 1px solid var(--numfmt-border); border-radius: 4px; }
.number-format-options :deep(.ant-menu-item) { height: 32px; margin-block: 2px; line-height: 32px; }
.number-format-custom-options :deep(.ant-menu-title-content) { white-space: normal; overflow-wrap: anywhere; }
.number-format-custom-options :deep(.ant-menu-item) { height: auto; min-height: 32px; line-height: 20px; padding-block: 6px; }
.number-format-description { margin: 0; color: var(--numfmt-secondary); font-size: 13px; line-height: 20px; }
.number-format-preview { display: flex; min-height: 36px; align-items: center; justify-content: space-between; gap: 12px; border: 1px solid var(--numfmt-border); border-radius: 4px; padding: 6px 10px; color: var(--numfmt-secondary); font-size: 12px; }
.number-format-preview strong { overflow-wrap: anywhere; color: inherit; font-size: 14px; font-weight: 500; }
.number-format-panel-footer { display: flex; flex: 0 0 auto; justify-content: flex-end; gap: 8px; padding-top: 20px; }
</style>
