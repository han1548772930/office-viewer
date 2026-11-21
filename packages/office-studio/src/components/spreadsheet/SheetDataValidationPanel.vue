<script setup lang="ts">
import { computed, inject, ref, watch } from 'vue'
import { Button, Checkbox, Input, Modal, Select } from 'antdv-next'
import { FOffice, type DataValidationRule } from '../../api/facade'
import type { OfficeStudio } from '../../api/office-studio'
import type { SheetDataValidationErrorStyle, SheetDataValidationOperator, SheetDataValidationType } from '../../domain/model'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

const props = defineProps<{ open: boolean; studio: OfficeStudio; workbookId: string; sheetId: string; sourceRange: string; initialType?: SheetDataValidationType }>()
const emit = defineEmits<{ 'update:open': [value: boolean]; applied: [] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)

const types = computed<Array<{ label: string; value: SheetDataValidationType }>>(() =>
  (['whole', 'decimal', 'date', 'time', 'textLength', 'list', 'listMultiple', 'checkbox', 'custom'] as const)
    .map(value => ({ value, label: t(`sheet.validation.type.${value}`, value) })))
const operators = computed<Array<{ label: string; value: SheetDataValidationOperator }>>(() =>
  (['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual'] as const)
    .map(value => ({ value, label: t(`sheet.validation.operator.${value}`, value) })))
const errorStyles = computed<Array<{ label: string; value: SheetDataValidationErrorStyle }>>(() =>
  (['stop', 'warning', 'info'] as const).map(value => ({ value, label: t(`sheet.validation.errorStyle.${value}`, value) })))
const rule = ref<DataValidationRule | null>(null)
const type = ref<SheetDataValidationType>('whole')
const operator = ref<SheetDataValidationOperator>('between')
const formula1 = ref('1')
const formula2 = ref('10')
const listValues = ref('')
const allowBlank = ref(true)
const showDropDown = ref(true)
const showErrorMessage = ref(true)
const errorStyle = ref<SheetDataValidationErrorStyle>('stop')
const errorTitle = ref(t('sheet.validation.error-default', '数据验证未通过'))
const error = ref('')
const showInputMessage = ref(false)
const promptTitle = ref('')
const prompt = ref('')
const busy = ref(false)
const message = ref('')
const needsFormula = computed(() => type.value !== 'checkbox')
const needsSecondFormula = computed(() => operator.value === 'between' || operator.value === 'notBetween')
const isList = computed(() => type.value === 'list' || type.value === 'listMultiple')
const title = computed(() => t(rule.value ? 'sheet.validation.edit-title' : 'sheet.validation.title', rule.value ? '编辑数据验证' : '数据验证'))

watch(() => [props.open, props.sourceRange, props.sheetId] as const, ([open]) => {
  if (!open) return
  loadRule()
}, { immediate: true })

function loadRule() {
  message.value = ''
  rule.value = FOffice.newAPI(props.studio).getWorkbook(props.workbookId)?.getSheetBySheetId(props.sheetId)?.getDataValidation(props.sourceRange) ?? null
  type.value = rule.value?.type ?? props.initialType ?? 'whole'
  operator.value = rule.value?.operator ?? 'between'
  formula1.value = rule.value?.formula1 ?? '1'
  formula2.value = rule.value?.formula2 ?? '10'
  listValues.value = rule.value?.formula1?.replace(/^"|"$/g, '').replace(/""/g, '"') ?? ''
  allowBlank.value = rule.value?.allowBlank ?? true
  showDropDown.value = rule.value?.showDropDown ?? true
  showErrorMessage.value = rule.value?.showErrorMessage ?? true
  errorStyle.value = rule.value?.errorStyle ?? 'stop'
  errorTitle.value = rule.value?.errorTitle ?? t('sheet.validation.error-default', '数据验证未通过')
  error.value = rule.value?.error ?? ''
  showInputMessage.value = rule.value?.showInputMessage ?? false
  promptTitle.value = rule.value?.promptTitle ?? ''
  prompt.value = rule.value?.prompt ?? ''
}

function listFormula(): string {
  const source = listValues.value.trim()
  if (/^=?\s*(?:(?:'[^']+'|[\w ]+)!\s*)?\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+$/i.test(source)) return source.startsWith('=') ? source : `=${source}`
  const values = listValues.value.split(',').map(value => value.trim()).filter(Boolean)
  if (!values.length) throw new Error(t('sheet.validation.list-empty', '请输入至少一个下拉选项。'))
  if (values.some(value => value.includes(','))) throw new Error(t('sheet.validation.list-comma', '下拉选项暂不接受包含逗号的单项。'))
  return `"${values.join(',').replaceAll('"', '""')}"`
}

function apply(): void {
  message.value = ''
  const sheet = FOffice.newAPI(props.studio).getWorkbook(props.workbookId)?.getSheetBySheetId(props.sheetId)
  if (!sheet) { message.value = t('sheet.validation.sheet-closed', '工作表已关闭，请重新打开数据验证。'); return }
  busy.value = true
  try {
    const nextFormula1 = isList.value ? listFormula() : formula1.value.trim()
    const next = {
      ...(rule.value ? { id: rule.value.id } : {}), type: type.value,
      ...(type.value !== 'checkbox' ? { formula1: nextFormula1 } : {}),
      ...(needsSecondFormula.value ? { formula2: formula2.value.trim() } : {}),
      ...(type.value !== 'checkbox' && type.value !== 'list' && type.value !== 'listMultiple' && type.value !== 'custom' ? { operator: operator.value } : {}),
      allowBlank: allowBlank.value, showDropDown: showDropDown.value, showErrorMessage: showErrorMessage.value,
      errorStyle: errorStyle.value, errorTitle: errorTitle.value, error: error.value,
      showInputMessage: showInputMessage.value, promptTitle: promptTitle.value, prompt: prompt.value,
      renderMode: 'arrow' as const,
    }
    sheet.setDataValidation(props.sourceRange, next)
    rule.value = sheet.getDataValidation(props.sourceRange)
    emit('applied')
    emit('update:open', false)
  } catch (reason) {
    message.value = reason instanceof Error ? reason.message : t('sheet.validation.save-failed', '数据验证规则未能保存。')
  } finally { busy.value = false }
}

function clear(): void {
  const sheet = FOffice.newAPI(props.studio).getWorkbook(props.workbookId)?.getSheetBySheetId(props.sheetId)
  if (!sheet) { message.value = t('sheet.validation.clear-sheet-closed', '工作表已关闭。'); return }
  busy.value = true
  try {
    sheet.setDataValidation(props.sourceRange, null)
    rule.value = null
    emit('applied')
    emit('update:open', false)
  } catch (reason) {
    message.value = reason instanceof Error ? reason.message : t('sheet.validation.clear-failed', '无法清除当前区域的数据验证。')
  } finally { busy.value = false }
}
</script>

<template>
  <Modal :open="open" :title="title" :ok-text="t('common.apply', '应用')" :cancel-text="t('common.cancel', '取消')" :confirm-loading="busy" :width="520" @ok="apply" @cancel="emit('update:open', false)">
    <div class="validation-form">
      <label>{{ t('sheet.validation.range', '应用区域') }}<Input :value="sourceRange" readonly :aria-label="t('sheet.validation.range-aria', '数据验证应用区域')" /></label>
      <label for="sheet-validation-type">{{ t('sheet.validation.type', '验证类型') }}</label>
      <Select id="sheet-validation-type" v-model:value="type" :options="types" :aria-label="t('sheet.validation.type', '数据验证类型')" />
      <label v-if="needsFormula && !isList" for="sheet-validation-operator">{{ t('sheet.validation.condition', '条件') }}</label>
      <Select v-if="needsFormula && !isList" id="sheet-validation-operator" v-model:value="operator" :options="operators" :aria-label="t('sheet.validation.condition', '数据验证条件')" />
      <label v-if="isList">{{ t('sheet.validation.options', '选项') }}<Input v-model:value="listValues" :aria-label="t('sheet.validation.options-aria', '数据验证列表选项')" :placeholder="t('sheet.validation.prompt-range', '例如：高,中,低')" /></label>
      <label v-else-if="needsFormula">{{ t('sheet.validation.rule-value', '规则值') }}<Input v-model:value="formula1" :aria-label="t('sheet.validation.rule-value-aria', '数据验证规则值')" :placeholder="t('sheet.validation.rule-placeholder', '数值、日期、时间、长度或公式')" /></label>
      <label v-if="needsSecondFormula">{{ t('sheet.validation.second-rule', '第二个规则值') }}<Input v-model:value="formula2" :aria-label="t('sheet.validation.second-rule', '数据验证第二个规则值')" /></label>
      <div class="validation-switches">
        <Checkbox v-model:checked="allowBlank">{{ t('sheet.validation.allow-blank', '允许空白') }}</Checkbox>
        <Checkbox v-if="isList" v-model:checked="showDropDown">{{ t('sheet.validation.show-dropdown', '显示下拉选项') }}</Checkbox>
        <Checkbox v-model:checked="showInputMessage">{{ t('sheet.validation.show-input-prompt', '显示输入提示') }}</Checkbox>
        <Checkbox v-model:checked="showErrorMessage">{{ t('sheet.validation.show-error', '输入无效时显示消息') }}</Checkbox>
      </div>
      <label v-if="showInputMessage">{{ t('sheet.validation.prompt-title', '提示标题') }}<Input v-model:value="promptTitle" :aria-label="t('sheet.validation.prompt-title', '数据验证提示标题')" /></label>
      <label v-if="showInputMessage">{{ t('sheet.validation.input-prompt', '输入提示') }}<Input v-model:value="prompt" :aria-label="t('sheet.validation.input-prompt', '数据验证输入提示')" /></label>
      <template v-if="showErrorMessage">
        <label for="sheet-validation-error-style">{{ t('sheet.validation.invalid-action', '无效输入处理') }}</label>
        <Select id="sheet-validation-error-style" v-model:value="errorStyle" :options="errorStyles" :aria-label="t('sheet.validation.invalid-action', '无效输入处理')" />
        <label>{{ t('sheet.validation.error-title', '错误标题') }}<Input v-model:value="errorTitle" :aria-label="t('sheet.validation.error-title', '数据验证错误标题')" /></label>
        <label>{{ t('sheet.validation.error-message', '错误消息') }}<Input v-model:value="error" :aria-label="t('sheet.validation.error-message', '数据验证错误消息')" /></label>
      </template>
      <p v-if="message" class="validation-error" role="alert">{{ message }}</p>
      <Button v-if="rule" type="text" danger :disabled="busy" @click="clear">{{ t('sheet.validation.clear', '清除当前区域验证') }}</Button>
    </div>
  </Modal>
</template>

<style scoped>
.validation-form { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px; }
.validation-form > label { display: grid; min-width: 0; gap: 5px; color: #435066; font-size: 13px; }
.validation-switches { grid-column: 1 / -1; display: flex; flex-wrap: wrap; gap: 12px; }
.validation-error { grid-column: 1 / -1; margin: 0; color: #b42318; }
@media (max-width: 560px) { .validation-form { grid-template-columns: minmax(0, 1fr); } }
</style>
