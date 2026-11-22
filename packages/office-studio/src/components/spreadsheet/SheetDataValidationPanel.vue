<script setup lang="ts">
import { computed, inject, nextTick, onBeforeUnmount, ref, shallowRef, watch } from 'vue'
import { Button, Checkbox, Collapse, Form, FormItem, Input, Select, theme } from 'antdv-next'
import { MousePointer2, Trash2, X } from 'lucide-vue-next'
import { FOffice, type DataValidationRule } from '../../api/facade'
import type { OfficeStudio } from '../../api/office-studio'
import type { CellRange, SheetDataValidationErrorStyle, SheetDataValidationOperator, SheetDataValidationRenderMode, SheetDataValidationType } from '../../domain/model'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'
import { createSheetDataValidationRule } from '../../features/sheet-validation/validation-engine.ts'
import { parseValidationRanges, validationPickedAddress, type ValidationRangeTarget } from '../../features/sheet-validation/validation-panel-state.ts'

const props = defineProps<{ open: boolean; studio: OfficeStudio; workbookId: string; sheetId: string; sourceRange: string; initialType?: SheetDataValidationType; initialRuleId?: string; canvasRanges?: CellRange[]; canAdd?: boolean; selectedRange?: { target: ValidationRangeTarget; value: string } }>()
const emit = defineEmits<{ 'update:open': [value: boolean]; applied: []; add: []; 'request-range-selection': [payload: { target: ValidationRangeTarget; active: boolean }] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)

const types = computed<Array<{ label: string; value: SheetDataValidationType }>>(() =>
  (['whole', 'decimal', 'date', 'time', 'textLength', 'list', 'listMultiple', 'checkbox', 'custom'] as const)
    .map(value => ({ value, label: t(`sheet.validation.type.${value}`, value) })))
const operators = computed<Array<{ label: string; value: SheetDataValidationOperator }>>(() =>
  (['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual'] as const)
    .map(value => ({ value, label: t(`sheet.validation.operator.${value}`, value) })))
const errorStyles = computed<Array<{ label: string; value: SheetDataValidationErrorStyle }>>(() =>
  (['stop', 'warning', 'info'] as const).map(value => ({ value, label: t(`sheet.validation.errorStyle.${value}`, value) })))
const rule = shallowRef<DataValidationRule | null>(null)
const rulesRevision = ref(0)
const selectedRuleId = ref<string>()
const detailOpen = ref(false)
const advancedOptionsOpen = ref(false)
const advancedOptionsKeys = computed({
  get: () => advancedOptionsOpen.value ? ['advanced'] : [],
  set: (keys: string | number | Array<string | number>) => {
    advancedOptionsOpen.value = (Array.isArray(keys) ? keys : [keys]).includes('advanced')
  },
})
const rangeValue = ref('')
const type = ref<SheetDataValidationType>('whole')
const operator = ref<SheetDataValidationOperator | ''>('between')
const formula1 = ref('1')
const formula2 = ref('10')
const listValues = ref('')
const allowBlank = ref(true)
const showDropDown = ref(true)
const ruleRenderMode = ref<SheetDataValidationRenderMode>('arrow')
const showErrorMessage = ref(true)
const errorStyle = ref<SheetDataValidationErrorStyle>('stop')
const errorTitle = ref(t('sheet.validation.error-default', '数据验证未通过'))
const error = ref('')
const showInputMessage = ref(false)
const promptTitle = ref('')
const prompt = ref('')
const listMode = ref<'inline' | 'reference'>('inline')
const checkboxCustomValues = ref(false)
const rangeSelectionTarget = ref<ValidationRangeTarget | null>(null)
const pickedValue = computed(() => rangeSelectionTarget.value && validationPickedAddress(rangeSelectionTarget.value, props.canvasRanges ?? []))
const panel = ref<HTMLElement>()
const { token } = theme.useToken()
const busy = ref(false)
const message = ref('')
const needsFormula = computed(() => type.value === 'list' || type.value === 'listMultiple' || type.value === 'custom' || type.value !== 'checkbox' && operator.value !== '')
const needsSecondFormula = computed(() => operator.value === 'between' || operator.value === 'notBetween')
const isList = computed(() => type.value === 'list' || type.value === 'listMultiple')
const title = computed(() => detailOpen.value
  ? t(rule.value ? 'sheet.validation.edit-title' : 'sheet.validation.title', rule.value ? '编辑数据验证' : '数据验证')
  : t('sheet.validation.title', '数据验证'))
const sheet = computed(() => FOffice.newAPI(props.studio).getWorkbook(props.workbookId)?.getSheetBySheetId(props.sheetId) ?? null)
const rules = computed(() => {
  rulesRevision.value
  return sheet.value?.getDataValidations() ?? []
})

watch(() => [props.open, props.sourceRange, props.sheetId, props.initialType, props.initialRuleId] as const, ([open]) => {
  if (!open) return
  stopRangeSelection()
  message.value = ''
  selectedRuleId.value = undefined
  if (props.initialRuleId) {
    rulesRevision.value++
    const target = rules.value.find(candidate => candidate.id === props.initialRuleId)
    if (target) { editRule(target); return }
  }
  detailOpen.value = Boolean(props.initialType)
  if (props.initialType) {
    loadRule(null, props.initialType)
    return
  }
  loadRule(null)
}, { immediate: true })

watch(() => props.selectedRange, selected => {
  if (!selected || selected.target !== rangeSelectionTarget.value) return
  acceptPickedRange(selected.target, selected.value)
})

function acceptPickedRange(target: ValidationRangeTarget, value: string) {
  if (target === 'range') rangeValue.value = value
  else if (target === 'list') { listMode.value = 'reference'; listValues.value = value }
  else if (target === 'formula1') formula1.value = value
  else formula2.value = value
  finishRangePick()
}

function finishRangePick() {
  const target = rangeSelectionTarget.value
  stopRangeSelection()
  if (target) void nextTick(() => panel.value?.querySelector<HTMLElement>(`[data-pick-target="${target}"]`)?.focus())
}

function confirmRangePick() {
  if (rangeSelectionTarget.value && pickedValue.value) acceptPickedRange(rangeSelectionTarget.value, pickedValue.value)
}

watch(() => props.open, open => {
  if (!open && rangeSelectionTarget.value) {
    emit('request-range-selection', { target: rangeSelectionTarget.value, active: false })
    rangeSelectionTarget.value = null
  }
})
onBeforeUnmount(stopRangeSelection)

function loadRule(current: DataValidationRule | null, initialType?: SheetDataValidationType) {
  rule.value = current
  rangeValue.value = current ? rangeLabel(current) : props.sourceRange
  type.value = current?.type ?? initialType ?? 'whole'
  operator.value = current ? current.operator ?? '' : 'between'
  formula1.value = current?.formula1 ?? (type.value === 'checkbox' ? 'TRUE' : '1')
  formula2.value = current?.formula2 ?? (type.value === 'checkbox' ? 'FALSE' : '10')
  listValues.value = current?.formula1?.replace(/^"|"$/g, '').replace(/""/g, '"') ?? ''
  listMode.value = current?.formula1?.trim().startsWith('=') ? 'reference' : 'inline'
  checkboxCustomValues.value = current?.type === 'checkbox' && (current.formula1 !== undefined || current.formula2 !== undefined)
  allowBlank.value = current?.allowBlank ?? true
  showDropDown.value = current?.showDropDown ?? true
  ruleRenderMode.value = current?.renderMode ?? 'arrow'
  showErrorMessage.value = current?.showErrorMessage ?? true
  errorStyle.value = current?.errorStyle ?? 'stop'
  errorTitle.value = current?.errorTitle ?? t('sheet.validation.error-default', '数据验证未通过')
  error.value = current?.error ?? ''
  showInputMessage.value = current?.showInputMessage ?? false
  promptTitle.value = current?.promptTitle ?? ''
  prompt.value = current?.prompt ?? ''
  advancedOptionsOpen.value = false
}

function editRule(current: DataValidationRule) {
  stopRangeSelection()
  selectedRuleId.value = current.id
  detailOpen.value = true
  message.value = ''
  loadRule(current)
}

function listFormula(): string {
  const source = listValues.value.trim()
  if (listMode.value === 'reference') {
    if (!source) throw new Error(t('sheet.validation.list-reference-empty', '请选择或输入一个有效的单元格区域。'))
    return source.startsWith('=') ? source : `=${source}`
  }
  const values = listValues.value.split(',').map(value => value.trim()).filter(Boolean)
  if (!values.length) throw new Error(t('sheet.validation.list-empty', '请输入至少一个下拉选项。'))
  if (values.some(value => value.includes(','))) throw new Error(t('sheet.validation.list-comma', '下拉选项暂不接受包含逗号的单项。'))
  return `"${values.join(',').replaceAll('"', '""')}"`
}

function apply(): void {
  message.value = ''
  const activeSheet = sheet.value
  if (!activeSheet) { message.value = t('sheet.validation.sheet-closed', '工作表已关闭，请重新打开数据验证。'); return }
  busy.value = true
  try {
    const nextFormula1 = isList.value ? listFormula() : formula1.value.trim()
    const ranges = parseRanges(rangeValue.value)
    const next = {
      ...(rule.value ? { id: rule.value.id, metadata: rule.value.metadata ? structuredClone(rule.value.metadata) : undefined } : {}), type: type.value,
      ...(type.value === 'checkbox' ? checkboxCustomValues.value ? { formula1: formula1.value.trim(), formula2: formula2.value.trim() } : {} : needsFormula.value ? { formula1: nextFormula1 } : {}),
      ...(needsSecondFormula.value ? { formula2: formula2.value.trim() } : {}),
      ...(type.value !== 'checkbox' && type.value !== 'list' && type.value !== 'listMultiple' && type.value !== 'custom' && operator.value ? { operator: operator.value } : {}),
      allowBlank: allowBlank.value, showDropDown: showDropDown.value, showErrorMessage: showErrorMessage.value,
      errorStyle: errorStyle.value, errorTitle: errorTitle.value, error: error.value,
      showInputMessage: showInputMessage.value, promptTitle: promptTitle.value, prompt: prompt.value,
      renderMode: ruleRenderMode.value,
    }
    const updatedRule = createSheetDataValidationRule({ ...next, ...(rule.value ? { id: rule.value.id } : {}) }, ranges)
    props.studio.executeCommandSync('sheet.command.replace-data-validation', {
      unitId: props.workbookId,
      sheetId: props.sheetId,
      rule: updatedRule,
    })
    const savedRule = activeSheet.getDataValidations().find(candidate => candidate.id === updatedRule.id) ?? null
    rulesRevision.value += 1
    emit('applied')
    if (savedRule) {
      selectedRuleId.value = savedRule.id
      loadRule(savedRule)
      detailOpen.value = false
    }
  } catch (reason) {
    message.value = reason instanceof Error ? reason.message : t('sheet.validation.save-failed', '数据验证规则未能保存。')
  } finally { busy.value = false }
}

function clearRule(current: DataValidationRule): void {
  const activeSheet = sheet.value
  if (!activeSheet) { message.value = t('sheet.validation.clear-sheet-closed', '工作表已关闭。'); return }
  busy.value = true
  message.value = ''
  try {
    props.studio.executeCommandSync('sheet.command.remove-data-validation', {
      unitId: props.workbookId,
      sheetId: props.sheetId,
      ruleId: current.id,
    })
    rulesRevision.value += 1
    if (selectedRuleId.value === current.id) {
      selectedRuleId.value = undefined
      detailOpen.value = false
    }
    emit('applied')
  } catch (reason) {
    message.value = reason instanceof Error ? reason.message : t('sheet.validation.clear-failed', '无法清除当前区域的数据验证。')
  } finally { busy.value = false }
}

function clearAll(): void {
  const currentRules = rules.value
  if (!currentRules.length) return
  clearRuleRanges(currentRules)
}

function requestRangeSelection(target: ValidationRangeTarget): void {
  if (rangeSelectionTarget.value === target) {
    stopRangeSelection()
    return
  }
  rangeSelectionTarget.value = target
  emit('request-range-selection', { target, active: true })
}

function stopRangeSelection(): void {
  if (!rangeSelectionTarget.value) return
  emit('request-range-selection', { target: rangeSelectionTarget.value, active: false })
  rangeSelectionTarget.value = null
}

function clearRuleRanges(currentRules: DataValidationRule[]): void {
  busy.value = true
  message.value = ''
  try {
    props.studio.executeCommandSync('sheet.command.clear-data-validation', {
      unitId: props.workbookId,
      sheetId: props.sheetId,
      ranges: currentRules.flatMap(current => current.ranges),
    })
    rulesRevision.value += 1
    selectedRuleId.value = undefined
    detailOpen.value = false
    emit('applied')
  } catch (reason) {
    message.value = reason instanceof Error ? reason.message : t('sheet.validation.clear-failed', '无法清除当前区域的数据验证。')
  } finally { busy.value = false }
}

function rangeLabel(current: DataValidationRule): string {
  return current.ranges.map(range => `${columnLabel(range.startCol)}${range.startRow + 1}${range.startRow === range.endRow && range.startCol === range.endCol ? '' : `:${columnLabel(range.endCol)}${range.endRow + 1}`}`).join(', ')
}

function ruleLabel(current: DataValidationRule): string {
  const typeLabel = types.value.find(option => option.value === current.type)?.label ?? current.type
  if (current.type === 'list' || current.type === 'listMultiple' || current.type === 'checkbox') return typeLabel
  if (current.type === 'custom') return `${typeLabel} ${current.formula1 ?? ''}`.trim()
  if (!current.operator) return `${typeLabel} ${t('sheet.validation.legal', '有效值')}`
  const operatorLabel = operators.value.find(option => option.value === current.operator)?.label ?? current.operator
  const formulas = [current.formula1, needsSecondFormulaFor(current.operator) ? current.formula2 : undefined].filter(Boolean)
  return `${typeLabel} ${operatorLabel} ${formulas.join(' ')}`.trim()
}

function needsSecondFormulaFor(value?: SheetDataValidationOperator): boolean {
  return value === 'between' || value === 'notBetween'
}

function parseRanges(value: string): CellRange[] {
  try {
  return parseValidationRanges(value, sheet.value?.getSheetName())
  } catch {
    throw new TypeError(t('sheet.validation.range-invalid', '请输入有效的单元格区域，例如 A1、A1:B4、A:A 或 1:3。'))
  }
}

function columnLabel(index: number): string {
  let current = index + 1
  let label = ''
  while (current > 0) {
    const remainder = (current - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    current = Math.floor((current - 1) / 26)
  }
  return label
}
</script>

<template>
  <aside v-if="open" ref="panel" class="validation-sidebar" :aria-label="t('sheet.validation.title', '数据验证')" :style="{ '--dv-border': token.colorBorderSecondary, '--dv-text': token.colorText, '--dv-secondary': token.colorTextSecondary, '--dv-surface': token.colorBgContainer, '--dv-hover': token.colorFillTertiary, '--dv-active': token.colorPrimaryBorder, '--dv-error': token.colorError }">
    <header class="validation-header">
      <strong>{{ title }}</strong>
      <div class="validation-header-actions">
        <Button v-if="!detailOpen" type="primary" size="small" :disabled="!canAdd" @click="emit('add')">{{ t('sheet.validation.add', '添加规则') }}</Button>
        <Button type="text" size="small" :aria-label="t('common.close', '关闭')" @click="emit('update:open', false)"><X :size="16" /></Button>
      </div>
    </header>
    <div v-if="!detailOpen" class="validation-body validation-list" aria-label="数据验证规则">
      <p v-if="!rules.length" class="validation-empty">{{ t('sheet.validation.empty', '此工作表尚无数据验证规则。') }}</p>
      <article v-for="item in rules" :key="item.id" class="validation-rule" :class="{ 'is-selected': item.id === selectedRuleId }">
        <Button type="text" class="validation-rule-main" @click="editRule(item)">
          <strong>{{ ruleLabel(item) }}</strong>
          <span>{{ rangeLabel(item) }}</span>
        </Button>
        <Button class="validation-rule-delete" type="text" danger size="small" :disabled="busy" :aria-label="`${t('sheet.validation.remove-rule', '删除规则')} ${rangeLabel(item)}`" @click.stop="clearRule(item)"><Trash2 :size="14" /></Button>
      </article>
      <Button v-if="rules.length" danger type="text" :disabled="busy" @click="clearAll">{{ t('sheet.validation.remove-all', '清除全部规则') }}</Button>
      <p v-if="message" class="validation-error" role="alert">{{ message }}</p>
    </div>
    <Form v-else class="validation-body validation-form" layout="vertical" @submit.prevent="apply">
      <div class="validation-heading">
        <Button type="text" size="small" @click="stopRangeSelection(); detailOpen = false">{{ t('common.back', '返回规则列表') }}</Button>
      </div>
      <FormItem v-if="rangeSelectionTarget && canvasRanges" :label="t('sheet.validation.canvas-selection', '画布选区')">
        <div class="validation-range-input">
          <Input :value="pickedValue || ''" readonly :aria-label="t('sheet.validation.canvas-selection', '画布选区')" />
          <Button size="small" type="primary" :disabled="!pickedValue" @click="confirmRangePick">{{ t('sheet.validation.use-selection', '使用选区') }}</Button>
          <Button size="small" @click="finishRangePick">{{ t('common.cancel', '取消') }}</Button>
        </div>
      </FormItem>
      <FormItem :label="t('sheet.validation.range', '应用区域')">
        <div class="validation-range-input">
          <Input id="sheet-validation-range" v-model:value="rangeValue" :aria-label="t('sheet.validation.range-aria', '数据验证应用区域')" placeholder="A1:B4" />
          <Button data-pick-target="range" :type="rangeSelectionTarget === 'range' ? 'primary' : 'text'" size="small" :aria-label="t('sheet.validation.select-range', '从画布选择区域')" :aria-pressed="rangeSelectionTarget === 'range'" @click="requestRangeSelection('range')"><MousePointer2 :size="15" /></Button>
        </div>
      </FormItem>
      <FormItem :label="t('sheet.validation.type', '验证类型')">
        <Select id="sheet-validation-type" v-model:value="type" :options="types" :aria-label="t('sheet.validation.type', '数据验证类型')" />
      </FormItem>
      <FormItem v-if="type !== 'checkbox' && !isList && type !== 'custom'" :label="t('sheet.validation.condition', '条件')">
        <Select id="sheet-validation-operator" v-model:value="operator" :options="[{ value: '', label: t('sheet.validation.legal', '有效值') }, ...operators]" :aria-label="t('sheet.validation.condition', '数据验证条件')" />
      </FormItem>
      <FormItem v-if="isList" :label="t('sheet.validation.options', '选项')">
        <Select v-model:value="listMode" :options="[{ value: 'inline', label: t('sheet.validation.inline-list', '手动输入') }, { value: 'reference', label: t('sheet.validation.reference-list', '单元格区域') }]" :aria-label="t('sheet.validation.options', '下拉列表来源')" />
        <div class="validation-range-input">
          <Input v-model:value="listValues" :aria-label="t('sheet.validation.options-aria', '数据验证列表选项')" :placeholder="listMode === 'reference' ? t('sheet.validation.prompt-range', '例如：A1:A5') : t('sheet.validation.prompt-items', '例如：高,中,低')" />
          <Button v-if="listMode === 'reference'" data-pick-target="list" :type="rangeSelectionTarget === 'list' ? 'primary' : 'text'" size="small" :aria-label="t('sheet.validation.select-range', '从画布选择区域')" :aria-pressed="rangeSelectionTarget === 'list'" @click="requestRangeSelection('list')"><MousePointer2 :size="15" /></Button>
        </div>
      </FormItem>
      <FormItem v-else-if="needsFormula" :label="t('sheet.validation.rule-value', '规则值')">
        <div class="validation-range-input">
          <Input v-model:value="formula1" :aria-label="t('sheet.validation.rule-value-aria', '数据验证规则值')" :placeholder="t('sheet.validation.rule-placeholder', '数值、日期、时间、长度或公式')" />
          <Button data-pick-target="formula1" :type="rangeSelectionTarget === 'formula1' ? 'primary' : 'text'" size="small" :aria-label="t('sheet.validation.select-range', '从画布选择区域')" :aria-pressed="rangeSelectionTarget === 'formula1'" @click="requestRangeSelection('formula1')"><MousePointer2 :size="15" /></Button>
        </div>
      </FormItem>
      <template v-if="type === 'checkbox'">
        <Checkbox v-model:checked="checkboxCustomValues">{{ t('sheet.validation.checkbox-custom-values', '自定义复选框值') }}</Checkbox>
        <template v-if="checkboxCustomValues">
          <FormItem :label="t('sheet.validation.checked-value', '选中值')"><Input v-model:value="formula1" /></FormItem>
          <FormItem :label="t('sheet.validation.unchecked-value', '未选中值')"><Input v-model:value="formula2" /></FormItem>
        </template>
      </template>
      <FormItem v-if="needsSecondFormula" :label="t('sheet.validation.second-rule', '第二个规则值')">
        <div class="validation-range-input">
          <Input v-model:value="formula2" :aria-label="t('sheet.validation.second-rule', '数据验证第二个规则值')" />
          <Button data-pick-target="formula2" :type="rangeSelectionTarget === 'formula2' ? 'primary' : 'text'" size="small" :aria-label="t('sheet.validation.select-range', '从画布选择区域')" :aria-pressed="rangeSelectionTarget === 'formula2'" @click="requestRangeSelection('formula2')"><MousePointer2 :size="15" /></Button>
        </div>
      </FormItem>
      <Checkbox v-model:checked="allowBlank">{{ t('sheet.validation.allow-blank', '允许空白') }}</Checkbox>
      <Collapse v-model:active-key="advancedOptionsKeys" class="validation-options" :items="[{ key: 'advanced', label: t('sheet.validation.options', '选项') }]" ghost size="small" expand-icon-placement="end" destroy-on-hidden>
        <template #contentRender>
          <div class="validation-options-content">
          <template v-if="isList">
            <Checkbox v-model:checked="showDropDown">{{ t('sheet.validation.show-dropdown', '显示下拉选项') }}</Checkbox>
            <FormItem :label="t('sheet.validation.display-mode', '显示方式')">
              <Select v-model:value="ruleRenderMode" :options="[
                { value: 'custom', label: t('sheet.validation.renderMode.chip', '标签') },
                { value: 'arrow', label: t('sheet.validation.renderMode.arrow', '箭头') },
                { value: 'text', label: t('sheet.validation.renderMode.text', '文本') },
              ]" />
            </FormItem>
          </template>
          <Checkbox v-model:checked="showInputMessage">{{ t('sheet.validation.show-input-prompt', '显示输入提示') }}</Checkbox>
          <FormItem v-if="showInputMessage" :label="t('sheet.validation.prompt-title', '提示标题')"><Input v-model:value="promptTitle" :aria-label="t('sheet.validation.prompt-title', '数据验证提示标题')" /></FormItem>
          <FormItem v-if="showInputMessage" :label="t('sheet.validation.input-prompt', '输入提示')"><Input v-model:value="prompt" :aria-label="t('sheet.validation.input-prompt', '数据验证输入提示')" /></FormItem>
          <FormItem :label="t('sheet.validation.invalid-action', '无效输入处理')"><Select id="sheet-validation-error-style" v-model:value="errorStyle" :options="errorStyles" :aria-label="t('sheet.validation.invalid-action', '无效输入处理')" /></FormItem>
          <Checkbox v-model:checked="showErrorMessage">{{ t('sheet.validation.show-error', '输入无效时显示消息') }}</Checkbox>
          <FormItem v-if="showErrorMessage" :label="t('sheet.validation.error-title', '错误标题')"><Input v-model:value="errorTitle" :aria-label="t('sheet.validation.error-title', '数据验证错误标题')" /></FormItem>
          <FormItem v-if="showErrorMessage" :label="t('sheet.validation.error-message', '错误消息')"><Input v-model:value="error" :aria-label="t('sheet.validation.error-message', '数据验证错误消息')" /></FormItem>
          </div>
        </template>
      </Collapse>
      <p v-if="message" class="validation-error" role="alert">{{ message }}</p>
      <div class="validation-actions">
        <Button v-if="rule" danger :disabled="busy" @click="clearRule(rule)">{{ t('sheet.validation.clear', '清除规则') }}</Button>
        <Button type="primary" :loading="busy" html-type="submit">{{ t('common.apply', '应用') }}</Button>
      </div>
    </Form>
  </aside>
</template>

<style scoped>
.validation-sidebar { display: flex; flex: 0 1 312px; width: 312px; max-width: 100%; min-width: min(280px, 100%); min-height: 0; flex-direction: column; border-inline-start: 1px solid var(--dv-border); color: var(--dv-text); background: var(--dv-surface); }
.validation-header { display: flex; box-sizing: border-box; min-height: 48px; flex: 0 0 auto; align-items: center; justify-content: space-between; gap: 8px; padding: 16px 16px 8px; color: inherit; font-size: 16px; font-weight: 500; }
.validation-header-actions { display: flex; align-items: center; gap: 4px; }
.validation-body { box-sizing: border-box; min-width: 0; min-height: 0; flex: 1; overflow: auto; padding: 0 16px 16px; }
.validation-list { display: flex; flex-direction: column; gap: 8px; }
.validation-empty { margin: 0; color: var(--dv-secondary); }
.validation-rule { position: relative; display: flex; align-items: center; gap: 8px; margin: 0 -8px; border-radius: 6px; padding: 8px 36px 8px 8px; }
.validation-rule:hover, .validation-rule:focus-within { background: var(--dv-hover); }
.validation-rule.is-selected { outline: 1px solid var(--dv-active); outline-offset: -1px; }
.validation-rule-main { display: flex; min-width: 0; flex: 1; flex-direction: column; align-items: flex-start; gap: 3px; padding: 4px; color: inherit; text-align: start; }
.validation-rule-main span { overflow: hidden; max-width: 100%; color: var(--dv-secondary); font-size: 12px; text-overflow: ellipsis; white-space: nowrap; }
.validation-rule-delete { position: absolute; top: 50%; right: 8px; display: none; transform: translateY(-50%); }
.validation-rule:hover .validation-rule-delete, .validation-rule:focus-within .validation-rule-delete { display: inline-flex; }
.validation-form { display: grid; grid-template-columns: minmax(0, 1fr); gap: 10px; padding-top: 8px; padding-bottom: 8px; }
.validation-form .ant-form-item { margin-bottom: 0; }
.validation-range-input { display: flex; min-width: 0; align-items: center; gap: 4px; }
.validation-range-input > :first-child { min-width: 0; flex: 1; }
.validation-heading { display: flex; min-width: 0; align-items: center; justify-content: space-between; gap: 8px; color: var(--dv-secondary); font-size: 12px; }
.validation-heading span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.validation-options { min-width: 0; }
.validation-options :deep(.ant-collapse-header) { padding-inline: 0; }
.validation-options :deep(.ant-collapse-body) { padding-inline: 0; }
.validation-options-content { display: grid; gap: 10px; }
.validation-options-content .ant-form-item { margin-bottom: 0; }
.validation-actions { display: flex; justify-content: flex-end; gap: 8px; margin-top: 8px; }
.validation-error { margin: 0; color: var(--dv-error); }
</style>
