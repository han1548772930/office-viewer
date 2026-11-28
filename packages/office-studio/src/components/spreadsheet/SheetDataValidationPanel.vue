<script setup lang="ts">
import { computed, inject, nextTick, onBeforeUnmount, ref, shallowRef, watch } from 'vue'
import { Alert, Button, Checkbox, Collapse, ColorPicker, Form, FormItem, Input, Progress, Select, Spin, theme } from 'antdv-next'
import { ArrowDown, ArrowUp, MousePointer2, Plus, Trash2, X } from 'lucide-vue-next'
import type { CellRange, SheetDataValidationErrorStyle, SheetDataValidationOperator, SheetDataValidationRenderMode, SheetDataValidationRule as DataValidationRule, SheetDataValidationType } from '../../domain/model'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'
import { serializeValidationListItems } from '../../features/sheet-validation/validation-engine.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { moveValidationListItem, parseValidationRanges, validationListItemsFromCsv, validationPanelSettings, validationPickedAddress,
  validationRuleFromPanel, validationUsesOperator, validationUsesTwoFormulas, type ValidationListItem, type ValidationRangeTarget } from '../../features/sheet-validation/validation-panel-state.ts'
import { sheetOpaqueColorHex } from '../../features/sheet-workbench/sheet-colors.ts'
import type { SheetDataValidationPanelContext } from './use-sheet-data-validation-workbench.ts'

const props = defineProps<{ open: boolean; context: SheetDataValidationPanelContext; sourceRange: string; canvasRanges?: CellRange[]; selectedRange?: { target: ValidationRangeTarget; value: string } }>()
const emit = defineEmits<{ 'update:open': [value: boolean]; applied: []; add: []; 'request-range-selection': [payload: { target: ValidationRangeTarget; active: boolean }] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)

const types = computed<Array<{ label: string; value: SheetDataValidationType }>>(() =>
  (['any', 'whole', 'decimal', 'date', 'time', 'textLength', 'list', 'listMultiple', 'checkbox', 'custom'] as const)
    .map(value => ({ value, label: t(`sheet.validation.type.${value}`, {
      any: '任意值', whole: '整数', decimal: '小数', date: '日期', time: '时间', textLength: '文本长度',
      list: '下拉列表', listMultiple: '下拉多选', checkbox: '复选框', custom: '自定义公式',
    }[value]) })))
const operators = computed<Array<{ label: string; value: SheetDataValidationOperator }>>(() =>
  (['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual'] as const)
    .map(value => ({ value, label: t(`sheet.validation.operator.${value}`, value) })))
const errorStyles = computed<Array<{ label: string; value: SheetDataValidationErrorStyle }>>(() =>
  (['stop', 'warning', 'info'] as const).map(value => ({ value, label: t(`sheet.validation.errorStyle.${value}`, value) })))
const rule = shallowRef<DataValidationRule | null>(null)
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
const listItems = ref<ValidationListItem[]>([])
let syncingList = false, loading = false
let inlineDraft = '', referenceDraft = '='
const allowBlank = ref(true)
const showDropDown = ref(true)
const showDateTime = ref(false)
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
const getPopupContainer = (trigger: HTMLElement) => trigger.closest<HTMLElement>('.validation-sidebar') ?? trigger.parentElement ?? trigger
const busy = ref(false)
const composing = ref(false)
const disabled = computed(() => busy.value || props.context.disabled)
const message = ref('')
const needsFormula = computed(() => type.value === 'list' || type.value === 'listMultiple' || type.value === 'custom' || validationUsesOperator(type.value) && operator.value !== '')
const needsSecondFormula = computed(() => validationUsesTwoFormulas(type.value, operator.value))
const isList = computed(() => type.value === 'list' || type.value === 'listMultiple')
const title = computed(() => detailOpen.value
  ? t(rule.value ? 'sheet.validation.edit-title' : 'sheet.validation.title', rule.value ? '编辑数据验证' : '数据验证')
  : t('sheet.validation.title', '数据验证'))
const rules = computed(() => props.context.rules)
const referenceOptions = computed(() => listMode.value === 'reference' ? props.context.listOptions(listValues.value.trim().startsWith('=') ? listValues.value.trim() : `=${listValues.value.trim()}`) : [])

watch(() => [props.open, props.context.key, props.context.navigationRevision, props.context.initialType, props.context.initialRuleId] as const, ([open]) => {
  if (!open) return
  stopRangeSelection()
  message.value = ''
  selectedRuleId.value = undefined
  if (props.context.initialRuleId) {
    const target = rules.value.find(candidate => candidate.id === props.context.initialRuleId)
    if (target) { editRule(target); return }
  }
  detailOpen.value = Boolean(props.context.initialType)
  if (props.context.initialType) {
    loadRule(null, props.context.initialType)
    return
  }
  loadRule(null)
}, { immediate: true })

watch(() => props.selectedRange, selected => {
  if (!selected || selected.target !== rangeSelectionTarget.value) return
  acceptPickedRange(selected.target, selected.value)
})

function acceptPickedRange(target: ValidationRangeTarget, value: string) {
  if (disabled.value || composing.value) return
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
onBeforeUnmount(() => { stopRangeSelection(); props.context.setComposing(false); props.context.closed() })

function loadRule(current: DataValidationRule | null, initialType?: SheetDataValidationType) {
  loading = true
  rule.value = current ? cloneModel(current) : null
  rangeValue.value = current ? rangeLabel(current) : props.sourceRange
  type.value = current?.type ?? initialType ?? 'whole'
  operator.value = current ? current.operator ?? '' : validationUsesOperator(type.value) && type.value !== 'date' ? 'between' : ''
  formula1.value = current?.formula1 ?? (type.value === 'checkbox' ? 'TRUE' : '1')
  formula2.value = current?.formula2 ?? (type.value === 'checkbox' ? 'FALSE' : type.value === 'list' || type.value === 'listMultiple' ? '' : '10')
  listValues.value = current?.formula1?.replace(/^"|"$/g, '') ?? ''
  listItems.value = validationPanelSettings(current, initialType).listItems
  listMode.value = current?.formula1?.trim().startsWith('=') ? 'reference' : 'inline'
  inlineDraft = listMode.value === 'inline' ? listValues.value : ''
  referenceDraft = listMode.value === 'reference' ? listValues.value : '='
  checkboxCustomValues.value = current?.type === 'checkbox' && (current.formula1 !== undefined || current.formula2 !== undefined)
  allowBlank.value = current?.allowBlank ?? true
  showDropDown.value = current?.showDropDown ?? true
  showDateTime.value = current?.metadata?.showTime === true
  ruleRenderMode.value = current?.renderMode ?? 'arrow'
  showErrorMessage.value = current?.showErrorMessage ?? true
  errorStyle.value = current?.errorStyle ?? 'stop'
  errorTitle.value = current?.errorTitle ?? t('sheet.validation.error-default', '数据验证未通过')
  error.value = current?.error ?? ''
  showInputMessage.value = current?.showInputMessage ?? false
  promptTitle.value = current?.promptTitle ?? ''
  prompt.value = current?.prompt ?? ''
  advancedOptionsOpen.value = false
  loading = false
}

function editRule(current: DataValidationRule) {
  if (disabled.value || composing.value || !props.context.canEditRule(current)) return
  stopRangeSelection()
  selectedRuleId.value = current.id
  detailOpen.value = true
  message.value = ''
  loadRule(current)
}

watch(listValues, value => {
  if (loading) return
  if (listMode.value === 'reference') referenceDraft = value
  else { inlineDraft = value; if (!syncingList) listItems.value = validationListItemsFromCsv(value, listItems.value) }
}, { flush: 'sync' })
watch(listMode, (value, previous) => {
  if (loading || value === previous) return
  if (previous === 'inline') inlineDraft = listValues.value
  else referenceDraft = listValues.value
  syncingList = true
  try { listValues.value = value === 'inline' ? inlineDraft : referenceDraft }
  finally { syncingList = false }
}, { flush: 'sync' })
watch(type, (value, previous) => {
  if (loading || value === previous || (value === 'list' || value === 'listMultiple') && (previous === 'list' || previous === 'listMultiple')) return
  operator.value = validationUsesOperator(value) && value !== 'date' ? 'between' : ''
  formula1.value = value === 'checkbox' ? 'TRUE' : '1'; formula2.value = value === 'checkbox' ? 'FALSE' : value === 'list' || value === 'listMultiple' ? '' : '10'
  checkboxCustomValues.value = false; listValues.value = ''; listItems.value = []; listMode.value = 'inline'
  inlineDraft = ''; referenceDraft = '='
}, { flush: 'sync' })
function updateItems(items: ValidationListItem[]): void {
  if (disabled.value || composing.value) return
  syncingList = true
  try { listItems.value = items; listValues.value = serializeValidationListItems(items.map(item => item.label)).slice(1, -1) }
  finally { syncingList = false }
}
function updateItem(id: string, value: string, color?: string): void {
  updateItems(listItems.value.map(item => item.id === id ? { ...item, label: value, color: color ?? item.color } : item))
}
function addListItem(): void { updateItems([...listItems.value, { id: crypto.randomUUID(), label: '', color: '' }]) }
function composition(value: boolean): void { composing.value = value; props.context.setComposing(value) }
function setReferenceColor(index: number, value: unknown): void {
  if (disabled.value || composing.value) return
  const colors = formula2.value.split(','); colors[index] = typeof value === 'string' ? value : ''; formula2.value = colors.join(',')
}
function keydown(event: KeyboardEvent): void {
  if (composing.value || event.isComposing || event.keyCode === 229) { event.stopPropagation(); return }
  if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); if (!busy.value) props.context.setOpen(false) }
}

async function apply(): Promise<void> {
  if (disabled.value || composing.value || !props.context.isCurrent()) return
  const owner = props.context
  message.value = ''
  busy.value = true
  try {
    const ranges = parseRanges(rangeValue.value)
    const updatedRule = validationRuleFromPanel({ type: type.value, operator: operator.value, formula1: formula1.value.trim(), formula2: formula2.value.trim(),
      allowBlank: allowBlank.value, showDropDown: showDropDown.value, showErrorMessage: showErrorMessage.value,
      errorStyle: errorStyle.value, errorTitle: errorTitle.value, error: error.value,
      showInputMessage: showInputMessage.value, promptTitle: promptTitle.value, prompt: prompt.value,
      renderMode: ruleRenderMode.value, checkboxCustomValues: checkboxCustomValues.value, listMode: listMode.value === 'inline' ? 'custom' : 'reference',
      listItems: listItems.value, listReference: listValues.value.trim().startsWith('=') ? listValues.value.trim() : `=${listValues.value.trim()}` }, ranges,
      rule.value?.metadata || type.value === 'date' ? { metadata: { ...cloneModel(rule.value?.metadata ?? {}), ...(type.value === 'date' ? { showTime: showDateTime.value } : {}) } } : {}, rule.value)
    const result = await owner.submitAsync({ action: 'replace', rule: updatedRule })
    if (props.context !== owner) return
    if (!result.accepted) { message.value = result.error; return }
    if (!owner.isCurrent()) return
    const savedRule = rules.value.find(candidate => candidate.id === updatedRule.id) ?? null
    if (result.changed) emit('applied')
    if (savedRule) {
      selectedRuleId.value = savedRule.id
      loadRule(savedRule)
      detailOpen.value = false
    }
  } catch (reason) {
    if (props.context === owner) message.value = reason instanceof Error ? reason.message : t('sheet.validation.save-failed', '数据验证规则未能保存。')
  } finally { if (props.context === owner) busy.value = false }
}

function clearRule(current: DataValidationRule): void {
  if (disabled.value || composing.value || !props.context.canEditRule(current)) return
  busy.value = true
  message.value = ''
  try {
    const result = props.context.submit({ action: 'remove', ruleId: current.id })
    if (!result.accepted) { message.value = result.error; return }
    if (!props.context.isCurrent()) return
    if (selectedRuleId.value === current.id) {
      selectedRuleId.value = undefined
      detailOpen.value = false
    }
    if (result.changed) emit('applied')
  } catch (reason) {
    message.value = reason instanceof Error ? reason.message : t('sheet.validation.clear-failed', '无法清除当前区域的数据验证。')
  } finally { busy.value = false }
}

function clearAll(): void {
  if (disabled.value || composing.value || !props.context.canRemoveAll) return
  const currentRules = rules.value
  if (!currentRules.length) return
  clearRuleRanges(currentRules)
}

function requestRangeSelection(target: ValidationRangeTarget): void {
  if (disabled.value || composing.value) return
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

function clearRuleRanges(currentRules: readonly DataValidationRule[]): void {
  busy.value = true
  message.value = ''
  try {
    const result = props.context.submit({ action: 'clear', ranges: currentRules.flatMap(current => current.ranges) })
    if (!result.accepted) { message.value = result.error; return }
    if (!props.context.isCurrent()) return
    selectedRuleId.value = undefined
    detailOpen.value = false
    if (result.changed) emit('applied')
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
  return parseValidationRanges(value, props.context.sheetName)
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
  <aside v-if="open" ref="panel" class="validation-sidebar" :aria-label="t('sheet.validation.title', '数据验证')" :data-ribbon-editor-owner="context.inputOwner"
    :style="{ '--dv-border': token.colorBorderSecondary, '--dv-text': token.colorText, '--dv-secondary': token.colorTextSecondary, '--dv-surface': token.colorBgContainer, '--dv-hover': token.colorFillTertiary, '--dv-active': token.colorPrimaryBorder, '--dv-error': token.colorError }"
    @keydown.capture="keydown" @compositionstart.capture="composition(true)" @compositionend.capture="composition(false)">
    <header class="validation-header">
      <strong>{{ title }}</strong>
      <div class="validation-header-actions">
        <Button v-if="!detailOpen" type="primary" size="small" :disabled="disabled || composing || !context.canAdd" @click="emit('add')">{{ t('sheet.validation.add', '添加规则') }}</Button>
        <Button type="text" size="small" :disabled="busy || composing" :aria-label="t('common.close', '关闭')" @click="context.setOpen(false)"><X :size="16" /></Button>
      </div>
    </header>
    <div v-if="context.pending" class="validation-progress" role="status" aria-live="polite">
      <Spin size="small" />
      <span>{{ t(`sheet.validation.bulk.${context.progress?.phase ?? 'processing'}`, '正在处理单元格') }}<template v-if="context.progress"> {{ context.progress.processed }}<template v-if="context.progress.total !== null"> / {{ context.progress.total }}</template></template></span>
      <Progress v-if="context.progress?.total" :percent="Math.min(100, context.progress.processed / context.progress.total * 100)" :show-info="false" />
      <Button size="small" @click="context.cancel()">{{ t('common.cancel', '取消') }}</Button>
    </div>
    <div v-if="!detailOpen" class="validation-body validation-list" aria-label="数据验证规则">
      <p v-if="!rules.length" class="validation-empty">{{ t('sheet.validation.empty', '此工作表尚无数据验证规则。') }}</p>
      <article v-for="item in rules" :key="item.id" class="validation-rule" :class="{ 'is-selected': item.id === selectedRuleId }">
        <Button type="text" class="validation-rule-main" :disabled="disabled || composing || !context.canEditRule(item)" @click="editRule(item)">
          <strong>{{ ruleLabel(item) }}</strong>
          <span>{{ rangeLabel(item) }}</span>
        </Button>
        <Button class="validation-rule-delete" type="text" danger size="small" :disabled="disabled || composing || !context.canEditRule(item)" :aria-label="`${t('sheet.validation.remove-rule', '删除规则')} ${rangeLabel(item)}`" @click.stop="clearRule(item)"><Trash2 :size="14" /></Button>
      </article>
      <Button v-if="rules.length" danger type="text" :disabled="disabled || composing || !context.canRemoveAll" @click="clearAll">{{ t('sheet.validation.remove-all', '清除全部规则') }}</Button>
      <Alert v-if="message" type="error" :title="message" show-icon />
    </div>
    <Form v-else class="validation-body validation-form" layout="vertical" :disabled="disabled" @submit.prevent="apply">
      <div class="validation-heading">
        <Button type="text" size="small" :disabled="busy || composing" @click="stopRangeSelection(); detailOpen = false">{{ t('common.back', '返回规则列表') }}</Button>
      </div>
      <FormItem v-if="rangeSelectionTarget && canvasRanges" :label="t('sheet.validation.canvas-selection', '画布选区')">
        <div class="validation-range-input">
          <Input :value="pickedValue || ''" readonly :aria-label="t('sheet.validation.canvas-selection', '画布选区')" />
          <Button size="small" type="primary" :disabled="composing || !pickedValue" @click="confirmRangePick">{{ t('sheet.validation.use-selection', '使用选区') }}</Button>
          <Button size="small" :disabled="composing" @click="finishRangePick">{{ t('common.cancel', '取消') }}</Button>
        </div>
      </FormItem>
      <FormItem :label="t('sheet.validation.range', '应用区域')">
        <div class="validation-range-input">
          <Input id="sheet-validation-range" v-model:value="rangeValue" :aria-label="t('sheet.validation.range-aria', '数据验证应用区域')" placeholder="A1:B4" />
          <Button data-pick-target="range" :type="rangeSelectionTarget === 'range' ? 'primary' : 'text'" size="small" :aria-label="t('sheet.validation.select-range', '从画布选择区域')" :aria-pressed="rangeSelectionTarget === 'range'" @click="requestRangeSelection('range')"><MousePointer2 :size="15" /></Button>
        </div>
      </FormItem>
      <FormItem :label="t('sheet.validation.type', '验证类型')">
        <Select id="sheet-validation-type" v-model:value="type" :options="types" :get-popup-container="getPopupContainer" :aria-label="t('sheet.validation.type', '数据验证类型')" />
      </FormItem>
      <FormItem v-if="type === 'date'">
        <Checkbox v-model:checked="showDateTime">{{ t('sheet.validation.date.show-time', '显示时间') }}</Checkbox>
      </FormItem>
      <FormItem v-if="type !== 'checkbox' && !isList && type !== 'custom'" :label="t('sheet.validation.condition', '条件')">
        <Select id="sheet-validation-operator" v-model:value="operator" :options="[{ value: '', label: t('sheet.validation.legal', '有效值') }, ...operators]" :get-popup-container="getPopupContainer" :aria-label="t('sheet.validation.condition', '数据验证条件')" />
      </FormItem>
      <FormItem v-if="isList" :label="t('sheet.validation.options', '选项')">
        <Select v-model:value="listMode" :options="[{ value: 'inline', label: t('sheet.validation.inline-list', '手动输入') }, { value: 'reference', label: t('sheet.validation.reference-list', '单元格区域') }]" :get-popup-container="getPopupContainer" :aria-label="t('sheet.validation.options', '下拉列表来源')" />
        <div class="validation-range-input">
          <Input.TextArea v-if="listMode === 'inline'" v-model:value="listValues" :auto-size="{ minRows: 1, maxRows: 4 }" :aria-label="t('sheet.validation.options-aria', '数据验证列表选项')" :placeholder="t('sheet.validation.prompt-items', '例如：高,中,低')" />
          <Input v-else v-model:value="listValues" :aria-label="t('sheet.validation.options-aria', '数据验证列表选项')" :placeholder="t('sheet.validation.prompt-range', '例如：A1:A5')" />
          <Button v-if="listMode === 'reference'" data-pick-target="list" :type="rangeSelectionTarget === 'list' ? 'primary' : 'text'" size="small" :aria-label="t('sheet.validation.select-range', '从画布选择区域')" :aria-pressed="rangeSelectionTarget === 'list'" @click="requestRangeSelection('list')"><MousePointer2 :size="15" /></Button>
        </div>
      </FormItem>
      <div v-if="isList && listMode === 'inline'" class="validation-list-editor">
        <FormItem v-for="(item, index) in listItems" :key="item.id" :label="`${t('sheet.validation.list.item', '列表项')} ${index + 1}`">
          <div class="validation-list-item">
            <ColorPicker :value="sheetOpaqueColorHex(item.color)" value-format="hex" disabled-alpha allow-clear size="small" :get-popup-container="getPopupContainer"
              @update:value="value => updateItem(item.id, item.label, typeof value === 'string' ? value : '')">
              <Button type="text" size="small" :disabled="disabled || composing" :aria-label="`${t('sheet.validation.list.color', '列表项颜色')} ${index + 1}`"><span class="validation-color-swatch" :style="{ background: sheetOpaqueColorHex(item.color) ?? token.colorFillSecondary }" aria-hidden="true" /></Button>
            </ColorPicker>
            <Input.TextArea :value="item.label" :auto-size="{ minRows: 1, maxRows: 3 }" :aria-label="`${t('sheet.validation.list.item', '列表项')} ${index + 1}`" @update:value="value => updateItem(item.id, value)" />
            <Button type="text" size="small" :disabled="composing || index === 0" :aria-label="`${t('sheet.validation.list.move-up', '上移列表项')} ${index + 1}`" @click="updateItems(moveValidationListItem(listItems, item.id, -1))"><ArrowUp :size="13" /></Button>
            <Button type="text" size="small" :disabled="composing || index === listItems.length - 1" :aria-label="`${t('sheet.validation.list.move-down', '下移列表项')} ${index + 1}`" @click="updateItems(moveValidationListItem(listItems, item.id, 1))"><ArrowDown :size="13" /></Button>
            <Button type="text" danger size="small" :disabled="composing" :aria-label="`${t('sheet.validation.list.remove', '删除列表项')} ${index + 1}`" @click="updateItems(listItems.filter(value => value.id !== item.id))"><Trash2 :size="13" /></Button>
          </div>
        </FormItem>
        <Button type="dashed" size="small" :disabled="composing" @click="addListItem"><Plus :size="13" />{{ t('sheet.validation.list.add', '添加列表项') }}</Button>
      </div>
      <div v-else-if="isList && listMode === 'reference'" class="validation-list-editor">
        <FormItem v-for="(item, index) in referenceOptions" :key="item.value" :label="`${t('sheet.validation.list.item', '列表项')} ${index + 1}`">
          <div class="validation-list-item">
            <ColorPicker :value="sheetOpaqueColorHex(formula2.split(',')[index] ?? '')" value-format="hex" disabled-alpha allow-clear size="small" :get-popup-container="getPopupContainer"
              @update:value="value => setReferenceColor(index, value)">
              <Button type="text" size="small" :disabled="disabled || composing" :aria-label="`${t('sheet.validation.list.color', '列表项颜色')} ${index + 1}`"><span class="validation-color-swatch" :style="{ background: sheetOpaqueColorHex(formula2.split(',')[index] ?? '') ?? token.colorFillSecondary }" aria-hidden="true" /></Button>
            </ColorPicker>
            <Input :value="item.label" readonly :aria-label="`${t('sheet.validation.list.item', '列表项')} ${index + 1}`" />
          </div>
        </FormItem>
      </div>
      <FormItem v-if="!isList && needsFormula" :label="t('sheet.validation.rule-value', '规则值')">
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
              ]" :get-popup-container="getPopupContainer" />
            </FormItem>
          </template>
          <Checkbox v-model:checked="showInputMessage">{{ t('sheet.validation.show-input-prompt', '显示输入提示') }}</Checkbox>
          <FormItem v-if="showInputMessage" :label="t('sheet.validation.prompt-title', '提示标题')"><Input v-model:value="promptTitle" :aria-label="t('sheet.validation.prompt-title', '数据验证提示标题')" /></FormItem>
          <FormItem v-if="showInputMessage" :label="t('sheet.validation.input-prompt', '输入提示')"><Input v-model:value="prompt" :aria-label="t('sheet.validation.input-prompt', '数据验证输入提示')" /></FormItem>
          <FormItem :label="t('sheet.validation.invalid-action', '无效输入处理')"><Select id="sheet-validation-error-style" v-model:value="errorStyle" :options="errorStyles" :get-popup-container="getPopupContainer" :aria-label="t('sheet.validation.invalid-action', '无效输入处理')" /></FormItem>
          <Checkbox v-model:checked="showErrorMessage">{{ t('sheet.validation.show-error', '输入无效时显示消息') }}</Checkbox>
          <FormItem v-if="showErrorMessage" :label="t('sheet.validation.error-title', '错误标题')"><Input v-model:value="errorTitle" :aria-label="t('sheet.validation.error-title', '数据验证错误标题')" /></FormItem>
          <FormItem v-if="showErrorMessage" :label="t('sheet.validation.error-message', '错误消息')"><Input v-model:value="error" :aria-label="t('sheet.validation.error-message', '数据验证错误消息')" /></FormItem>
          </div>
        </template>
      </Collapse>
      <Alert v-if="message" type="error" :title="message" show-icon />
      <div class="validation-actions">
        <Button v-if="rule" danger :disabled="disabled || composing || !context.canEditRule(rule)" @click="clearRule(rule)">{{ t('sheet.validation.clear', '清除规则') }}</Button>
        <Button type="primary" :loading="busy" :disabled="disabled || composing" html-type="submit">{{ t('common.apply', '应用') }}</Button>
      </div>
    </Form>
  </aside>
</template>

<style scoped>
.validation-sidebar { position: relative; display: flex; flex: 0 1 312px; width: 312px; max-width: 100%; min-width: min(280px, 100%); min-height: 0; flex-direction: column; border-inline-start: 1px solid var(--dv-border); color: var(--dv-text); background: var(--dv-surface); }
.validation-header { display: flex; box-sizing: border-box; min-height: 48px; flex: 0 0 auto; align-items: center; justify-content: space-between; gap: 8px; padding: 16px 16px 8px; color: inherit; font-size: 16px; font-weight: 500; }
.validation-header-actions { display: flex; align-items: center; gap: 4px; }
.validation-progress { display: flex; flex-wrap: wrap; align-items: center; gap: 8px; padding: 8px 16px; color: var(--dv-secondary); }
.validation-progress :deep(.ant-progress) { flex: 1 1 120px; min-width: 80px; }
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
.validation-list-editor { display: grid; gap: 8px; min-width: 0; }
.validation-list-item { display: flex; align-items: center; gap: 3px; min-width: 0; }
.validation-list-item :deep(.ant-input), .validation-list-item :deep(.ant-input-textarea) { flex: 1; min-width: 0; }
.validation-list-item :deep(.ant-btn) { flex-shrink: 0; padding-inline: 2px; }
.validation-color-swatch { display: inline-block; width: 16px; height: 16px; border: 1px solid var(--dv-border); border-radius: 3px; }
.validation-actions { display: flex; justify-content: flex-end; gap: 8px; margin-top: 8px; }
.validation-error { margin: 0; color: var(--dv-error); }
</style>
