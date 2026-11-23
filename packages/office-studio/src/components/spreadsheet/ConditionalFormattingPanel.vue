<script setup lang="ts">
import { computed, inject, ref, watch } from 'vue'
import { Alert, Button, Checkbox, ColorPicker, Drawer, Form, FormItem, Input, InputNumber, Select, Tooltip } from 'antdv-next'
import { ArrowDown, ArrowLeft, ArrowUp, Plus, Trash2 } from 'lucide-vue-next'
import type { CellRange, SheetConditionalFormatFormulaAnchor, SheetConditionalFormatRule, SheetConditionalFormatType, SheetConditionalFormatValueConfig } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { conditionalFormatFormulaAnchor, reanchorConditionalFormatRule } from '../../features/conditional-formatting/conditional-format-formulas.ts'
import { conditionalFormatAverageEditorOperator, conditionalFormatAverageRuleOperator, conditionalFormatEditorFormula, conditionalFormatEditorText, mergeConditionalFormatRuleDraft, moveConditionalFormatRuleDraft,
  prependConditionalFormatRuleDraft, removeConditionalFormatRuleDraft } from '../../features/conditional-formatting/conditional-format-rule-editor.ts'
import { assertSheetConditionalFormatRuleSnapshot, type SheetConditionalFormatSubmitResult } from '../../features/conditional-formatting/conditional-format-command.ts'
import { conditionalIconPresets } from '../../features/conditional-formatting/conditional-icon-presets.ts'
import { conditionalIconSequence, createConditionalIconConfig, validateConditionalIconConfig, validateConditionalThreshold } from '../../features/conditional-formatting/icon-set-config.ts'
import ConditionalIcon from './ConditionalIcon.vue'
import ConditionalThresholdInput from './ConditionalThresholdInput.vue'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'
import { useRibbonTheme } from '../../platform/ribbon/ribbon-theme.ts'

const props = defineProps<{ open: boolean; rules: SheetConditionalFormatRule[]; defaultRange: string; submit: (rules: readonly SheetConditionalFormatRule[]) => SheetConditionalFormatSubmitResult }>()
const emit = defineEmits<{ 'update:open': [value: boolean]; closed: [] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const themeStyle = useRibbonTheme()
const draftId = ref<string | null>(null)
const editing = ref(false)
const ruleType = ref<SheetConditionalFormatType>('highlightCell')
const highlightSubtype = ref<'text' | 'number' | 'timePeriod' | 'duplicateValues' | 'uniqueValues'>('number')
const iconSet = ref('3Arrows')
const rank = ref(10)
const rankPercent = ref(false)
const showValue = ref(true)
const reverseIcons = ref(false)
type IconConfig = NonNullable<NonNullable<SheetConditionalFormatRule['config']>['icons']>[number]
const iconThresholds = ref<IconConfig[]>(createConditionalIconConfig('3Arrows'))
const minimum = ref<SheetConditionalFormatValueConfig>({ type: 'min' })
const midpoint = ref<SheetConditionalFormatValueConfig>({ type: 'percentile', value: 50 })
const maximum = ref<SheetConditionalFormatValueConfig>({ type: 'max' })
const useMidpoint = ref(true)
const isGradient = ref(false)
const negativeColor = ref('#f8696b')
const draftRanges = ref('A1')
const draftOrigin = ref<SheetConditionalFormatFormulaAnchor>()
const operator = ref<string>('greaterThan')
const formula1 = ref('0')
const formula2 = ref('1')
const fill = ref('#fff2cc')
const fontColor = ref('#9c6500')
const useFill = ref(true)
const useFontColor = ref(true)
const minimumColor = ref('#f8696b')
const maximumColor = ref('#63be7b')
const bold = ref(false)
const italic = ref(false)
const underline = ref(false)
const strike = ref(false)
const stopIfTrue = ref(false)
const error = ref('')
const localRules = ref<SheetConditionalFormatRule[]>([])
const previewIcons = computed(() => conditionalIconSequence({ icons: iconThresholds.value, reverse: reverseIcons.value }))
const needsSecondValue = computed(() => operator.value === 'between' || operator.value === 'notBetween')
const operatorOptions = computed(() => [
  ['equal', t('conditional.operator.equal', '等于')],
  ['notEqual', t('conditional.operator.notEqual', '不等于')],
  ['greaterThan', t('conditional.operator.greaterThan', '大于')],
  ['greaterThanOrEqual', t('conditional.operator.greaterThanOrEqual', '大于或等于')],
  ['lessThan', t('conditional.operator.lessThan', '小于')],
  ['lessThanOrEqual', t('conditional.operator.lessThanOrEqual', '小于或等于')],
  ['between', t('conditional.operator.between', '介于')],
  ['notBetween', t('conditional.operator.notBetween', '不介于')],
].map(([value, label]) => ({ value, label })))
const typeOptions = [
  ['highlightCell', '突出显示单元格规则'], ['top10', '排名和平均值规则'], ['dataBar', '数据条规则'],
  ['colorScale', '色阶规则'], ['expression', '公式规则'], ['iconSet', '图标集规则'],
].map(([value, label]) => ({ value, label }))
const iconSetOptions = Object.keys(conditionalIconPresets).map(value => ({ value, label: iconSetLabel(value) }))
const individualIconOptions = [{ value: 'EMPTY_ICON_TYPE:', label: '无图标' }, ...Object.entries(conditionalIconPresets).flatMap(([type, icons]) => icons.map((icon, index) => ({ value: `${type}:${index}`, label: `${iconSetLabel(type)} ${index + 1} ${icon.symbol}` })))]
const visualRule = computed(() => ['dataBar', 'colorScale', 'iconSet'].includes(ruleType.value))
const styleRule = computed(() => ruleType.value === 'highlightCell' || ruleType.value === 'top10' || ruleType.value === 'expression' || ruleType.value === 'office')
const typeNeedsValue = computed(() => ruleType.value === 'expression' || ruleType.value === 'highlightCell' && (highlightSubtype.value === 'number' || highlightSubtype.value === 'text' && !['containsBlanks', 'notContainsBlanks', 'containsErrors', 'notContainsErrors'].includes(operator.value)))
const formulaLabel = computed(() => ruleType.value === 'expression' ? '公式' : highlightSubtype.value === 'number' ? t('conditional.value-formula', '数值或公式') : highlightSubtype.value === 'timePeriod' ? '日期期间' : '匹配内容')
const formulaBasis = computed(() => {
  const hasFormula = ruleType.value === 'expression' || ruleType.value === 'highlightCell' && highlightSubtype.value === 'number'
    && [formula1.value, ...(needsSecondValue.value ? [formula2.value] : [])].some(value => !Number.isFinite(Number(value.trim().replace(/^=/, ''))))
    || visualRule.value && [minimum.value, maximum.value, ...(useMidpoint.value ? [midpoint.value] : []), ...iconThresholds.value.map(icon => icon.value)].some(value => value.type === 'formula')
  if (!hasFormula) return undefined
  try {
    const origin = draftOrigin.value ?? conditionalFormatFormulaAnchor({ ranges: parseRanges(draftRanges.value) })
    return origin ? `公式以 ${formatCell(origin.row, origin.col)} 为基准，其他单元格按相对引用计算。` : undefined
  } catch { return undefined }
})

watch([ruleType, highlightSubtype], ([type]) => {
  if (type === 'highlightCell' && highlightSubtype.value === 'timePeriod') operator.value = 'today'
  else if (type === 'top10') operator.value = 'isNotBottom'
  else if (type === 'highlightCell' && highlightSubtype.value === 'text') operator.value = 'containsText'
  else if (type === 'highlightCell' && highlightSubtype.value === 'number') operator.value = 'greaterThan'
}, { flush: 'sync' })

watch(() => props.open, open => {
  if (!open) return
  localRules.value = props.rules.map(cloneRule).sort((left, right) => left.priority - right.priority)
  error.value = ''
  if (localRules.value.length) {
    selectRule(localRules.value[0]!)
    editing.value = false
  } else {
    newRule()
    editing.value = false
  }
}, { immediate: true })

function newRule() {
  editing.value = true
  draftId.value = null
  draftOrigin.value = undefined
  ruleType.value = 'highlightCell'
  highlightSubtype.value = 'number'
  draftRanges.value = props.defaultRange || 'A1'
  operator.value = 'greaterThan'
  formula1.value = '0'
  formula2.value = '1'
  fill.value = '#fff2cc'
  fontColor.value = '#9c6500'
  useFill.value = true
  useFontColor.value = true
  minimumColor.value = '#f8696b'
  maximumColor.value = '#63be7b'
  bold.value = false
  italic.value = false
  underline.value = false
  strike.value = false
  stopIfTrue.value = false
  rank.value = 10
  rankPercent.value = false
  showValue.value = true
  reverseIcons.value = false
  iconSet.value = '3Arrows'
  iconThresholds.value = createConditionalIconConfig(iconSet.value)
  minimum.value = { type: 'min' }
  midpoint.value = { type: 'percentile', value: 50 }
  maximum.value = { type: 'max' }
  useMidpoint.value = true
  isGradient.value = false
  negativeColor.value = '#f8696b'
  error.value = ''
}

function selectRule(rule: SheetConditionalFormatRule) {
  editing.value = true
  draftId.value = rule.id
  draftOrigin.value = conditionalFormatFormulaAnchor(rule)
  ruleType.value = rule.type === 'office' ? 'office' : rule.type === 'expression' || rule.subType === 'formula' ? 'expression'
    : rule.type === 'average' || rule.subType === 'average' || rule.type === 'top10' || rule.subType === 'rank' ? 'top10'
      : ['cellIs', 'highlightCell', 'containsText', 'notContainsText', 'beginsWith', 'endsWith', 'duplicateValues', 'uniqueValues', 'timePeriod'].includes(rule.type) ? 'highlightCell' : rule.type
  highlightSubtype.value = rule.subType === 'text' || ['containsText', 'notContainsText', 'beginsWith', 'endsWith'].includes(rule.type) ? 'text'
    : rule.subType === 'timePeriod' || rule.type === 'timePeriod' ? 'timePeriod'
      : rule.subType === 'duplicateValues' || rule.type === 'duplicateValues' ? 'duplicateValues'
        : rule.subType === 'uniqueValues' || rule.type === 'uniqueValues' ? 'uniqueValues' : 'number'
  draftRanges.value = rule.ranges.map(formatRange).join(';')
  operator.value = ['notContainsText', 'beginsWith', 'endsWith'].includes(rule.type) ? rule.type : rule.operator || (highlightSubtype.value === 'text' ? 'containsText' : 'greaterThan')
  formula1.value = highlightSubtype.value === 'text' ? String(rule.value ?? rule.formula1)
    : rule.formula1 || (ruleType.value === 'expression' && typeof rule.value === 'string' ? rule.value : '')
  formula2.value = rule.formula2 ?? '1'
  fill.value = rule.style.fill ?? rule.config?.positiveColor ?? (rule.config?.thresholds?.length === 3 ? rule.config.thresholds[1]?.color : undefined) ?? '#ffffff'
  fontColor.value = rule.style.fontColor ?? '#1f2937'
  useFill.value = rule.style.fill !== undefined
  useFontColor.value = rule.style.fontColor !== undefined
  minimumColor.value = rule.config?.thresholds?.[0]?.color ?? '#f8696b'
  maximumColor.value = rule.config?.thresholds?.at(-1)?.color ?? '#63be7b'
  bold.value = rule.style.bold ?? false
  italic.value = rule.style.italic ?? false
  underline.value = rule.style.underline ?? false
  strike.value = rule.style.strike ?? false
  stopIfTrue.value = rule.stopIfTrue
  rank.value = Number(rule.value) || 10
  rankPercent.value = rule.isPercent ?? false
  if (rule.type === 'top10' || rule.subType === 'rank') operator.value = rule.isBottom ? 'isBottom' : 'isNotBottom'
  if (rule.type === 'average' || rule.subType === 'average') operator.value = conditionalFormatAverageEditorOperator(rule.operator) ?? 'greaterThanAverage'
  showValue.value = rule.config?.isShowValue ?? true
  iconSet.value = rule.config?.iconSet ?? rule.config?.icons?.[0]?.iconType ?? '3Arrows'
  minimum.value = { ...(rule.type === 'colorScale' ? rule.config?.thresholds?.[0]?.value : rule.config?.min) ?? { type: 'min' } }
  maximum.value = { ...(rule.type === 'colorScale' ? rule.config?.thresholds?.at(-1)?.value : rule.config?.max) ?? { type: 'max' } }
  midpoint.value = { ...(rule.config?.thresholds?.length === 3 ? rule.config.thresholds[1]?.value : undefined) ?? { type: 'percentile', value: 50 } }
  useMidpoint.value = rule.config?.thresholds?.length === 3
  isGradient.value = rule.config?.isGradient ?? false
  negativeColor.value = rule.config?.nativeColor ?? '#f8696b'
  reverseIcons.value = rule.config?.reverse ?? false
  iconThresholds.value = rule.config?.icons?.map(icon => ({ ...icon, value: { ...icon.value } })) ?? createConditionalIconConfig(iconSet.value)
  error.value = ''
}

function saveDraft() {
  try {
    const ranges = parseRanges(draftRanges.value)
    const existingIndex = localRules.value.findIndex(rule => rule.id === draftId.value)
    const previous = existingIndex >= 0 ? localRules.value[existingIndex] : undefined
    const previousStyle = previous?.style ?? {}
    const style = {
      fill: useFill.value ? fill.value : undefined,
      fontColor: useFontColor.value ? fontColor.value : undefined,
      bold: bold.value || previousStyle.bold !== undefined ? bold.value : undefined,
      italic: italic.value || previousStyle.italic !== undefined ? italic.value : undefined,
      underline: underline.value || previousStyle.underline !== undefined ? underline.value : undefined,
      strike: strike.value || previousStyle.strike !== undefined ? strike.value : undefined,
    }
    const secondFormula = needsSecondValue.value ? formula2.value : undefined
    const normalizedType = ruleType.value === 'highlightCell'
      ? highlightSubtype.value === 'number' ? 'cellIs' : highlightSubtype.value === 'text' ? operator.value === 'notContainsText' ? 'notContainsText' : operator.value === 'beginsWith' ? 'beginsWith' : operator.value === 'endsWith' ? 'endsWith' : 'containsText' : highlightSubtype.value
      : ruleType.value === 'top10' && conditionalFormatAverageRuleOperator(operator.value) !== undefined ? 'average' : ruleType.value === 'expression' ? 'expression' : ruleType.value
    const normalizedSubType = ruleType.value === 'highlightCell' ? highlightSubtype.value : ruleType.value === 'top10' ? normalizedType === 'average' ? 'average' : 'rank' : ruleType.value === 'expression' ? 'formula' : undefined
    if (normalizedType === 'expression' && !formula1.value.trim().replace(/^=/, '')) throw new TypeError('请输入条件公式。')
    if (normalizedType === 'top10' && (!Number.isInteger(rank.value) || rank.value < 1 || rank.value > (rankPercent.value ? 100 : 1000))) throw new RangeError('排名项数无效。')
    if (ruleType.value === 'iconSet') validateConditionalIconConfig({ icons: iconThresholds.value })
    else if (visualRule.value) {
      const thresholds = [minimum.value, maximum.value, ...(ruleType.value === 'colorScale' && useMidpoint.value ? [midpoint.value] : [])]
      for (const threshold of thresholds) validateConditionalThreshold(threshold, ruleType.value === 'dataBar')
    }
    const id = draftId.value ?? crypto.randomUUID()
    const text = conditionalFormatEditorText(formula1.value, previous)
    const nextRule = reanchorConditionalFormatRule(mergeConditionalFormatRuleDraft(previous, {
      id, type: normalizedType, ranges, operator: ruleType.value === 'office' ? previous?.operator ?? '' : ruleType.value === 'top10'
        ? conditionalFormatAverageRuleOperator(operator.value) ?? 'greaterThan' : operator.value,
      ...(draftOrigin.value ? { formulaAnchor: { ...draftOrigin.value } } : {}),
      subType: normalizedSubType,
      formula1: ruleType.value === 'office' ? previous?.formula1 ?? '' : normalizedSubType === 'text' ? text.formula1
        : normalizedType === 'expression' && previous?.formula1 === '' && typeof previous.value === 'string' ? '' : conditionalFormatEditorFormula(formula1.value, previous?.formula1),
      formula2: ruleType.value === 'office' ? previous?.formula2 : secondFormula === undefined ? undefined : conditionalFormatEditorFormula(secondFormula, previous?.formula2),
      value: ruleType.value === 'office' ? previous?.value : normalizedType === 'top10' ? rank.value : normalizedSubType === 'text' ? text.value
        : normalizedType === 'expression' && typeof previous?.value === 'string' ? conditionalFormatEditorFormula(formula1.value, previous.formula1 || previous.value) : previous?.value,
      isBottom: normalizedType === 'top10' ? operator.value === 'isBottom' : previous?.isBottom,
      isPercent: normalizedType === 'top10' ? rankPercent.value : previous?.isPercent,
      config: ruleType.value === 'dataBar' ? { min: { ...minimum.value }, max: { ...maximum.value }, isGradient: isGradient.value, positiveColor: fill.value, nativeColor: negativeColor.value, isShowValue: showValue.value }
        : ruleType.value === 'colorScale' ? { thresholds: [{ value: { ...minimum.value }, color: minimumColor.value }, ...(useMidpoint.value ? [{ value: { ...midpoint.value }, color: fill.value }] : []), { value: { ...maximum.value }, color: maximumColor.value }] }
          : ruleType.value === 'iconSet' ? { iconSet: iconSet.value, isShowValue: showValue.value, reverse: reverseIcons.value, icons: iconThresholds.value.map(icon => ({ ...icon, value: { ...icon.value } })) }
            : ruleType.value === 'office' ? previous?.config : undefined,
      style, priority: existingIndex < 0 ? localRules.value[0]?.priority ?? 1 : localRules.value[existingIndex]!.priority,
      stopIfTrue: stopIfTrue.value,
    }), ranges)
    assertSheetConditionalFormatRuleSnapshot(nextRule, previous)
    if (existingIndex < 0) localRules.value = prependConditionalFormatRuleDraft(localRules.value, nextRule)
    else localRules.value = localRules.value.map(rule => rule.id === id ? nextRule : rule)
    selectRule(localRules.value.find(rule => rule.id === id)!)
    editing.value = false
    error.value = ''
  } catch (reason) {
    error.value = reason instanceof Error ? reason.message : t('conditional.error.invalid', '条件格式规则无效。')
  }
}

function removeRule(id: string) {
  const wasEditing = editing.value
  localRules.value = removeConditionalFormatRuleDraft(localRules.value, id)
  if (draftId.value === id) {
    if (localRules.value.length) {
      selectRule(localRules.value[0]!)
      editing.value = wasEditing
    } else if (wasEditing) newRule()
    else {
      draftId.value = null
      newRule()
      editing.value = false
    }
  }
}

function moveRule(id: string, offset: -1 | 1) {
  localRules.value = moveConditionalFormatRuleDraft(localRules.value, id, offset)
}

function apply() {
  error.value = ''
  if (editing.value) saveDraft()
  if (error.value) return
  try {
    const result = props.submit(localRules.value.map(cloneRule))
    if (!result.accepted) { error.value = result.error; return }
    emit('update:open', false)
  } catch (reason) {
    error.value = reason instanceof Error ? reason.message : t('conditional.error.apply', '无法应用条件格式，草稿已保留。')
  }
}

function afterDrawerOpenChange(open: boolean) {
  if (!open) emit('closed')
}

function returnToRules() {
  editing.value = false
  error.value = ''
}

function changeIconSet(value: unknown) {
  iconSet.value = String(value)
  iconThresholds.value = createConditionalIconConfig(iconSet.value)
}
function changeIndividualIcon(icon: IconConfig, value: unknown) {
  const [type, id] = String(value).split(':')
  if (!type || id === undefined) return
  icon.iconType = type
  icon.iconId = id
}

function parseRanges(value: string): CellRange[] {
  const ranges = value.split(/[;,]/).map(token => token.trim()).filter(Boolean).map(token => {
    const bounds = token.split(':')
    if (bounds.length > 2) throw new TypeError(t('conditional.error.range', '范围格式无效。'))
    const start = parseCell(bounds[0]!)
    const end = parseCell(bounds[1] ?? bounds[0]!)
    if (!start || !end || end.row < start.row || end.column < start.column) throw new TypeError(t('conditional.error.range', '范围格式无效。'))
    return { startRow: start.row, startCol: start.column, endRow: end.row, endCol: end.column }
  })
  if (!ranges.length) throw new TypeError(t('conditional.error.range', '范围格式无效。'))
  return ranges
}

function parseCell(address: string): { row: number; column: number } | undefined {
  const match = /^\$?([A-Z]+)\$?(\d+)$/i.exec(address.trim())
  if (!match) return undefined
  let column = 0
  for (const letter of match[1]!.toUpperCase()) column = column * 26 + letter.charCodeAt(0) - 64
  const row = Number(match[2])
  if (row < 1 || row > 1_048_576 || column > 16_384 || !Number.isSafeInteger(row)) return undefined
  return { row: row - 1, column: column - 1 }
}

function formatRange(range: CellRange): string {
  const start = formatCell(range.startRow, range.startCol)
  const end = formatCell(range.endRow, range.endCol)
  return start === end ? start : `${start}:${end}`
}

function formatCell(row: number, column: number): string {
  let letters = ''
  for (let value = column + 1; value; value = Math.floor((value - 1) / 26)) letters = String.fromCharCode((value - 1) % 26 + 65) + letters
  return `${letters}${row + 1}`
}

function iconSetLabel(value: string): string {
  return ({ '3Arrows': '三向箭头', '3ArrowsGray': '三向灰箭头', '4Arrows': '四向箭头', '4ArrowsGray': '四向灰箭头', '5Arrows': '五向箭头', '5ArrowsGray': '五向灰箭头', '3Triangles': '三角形', '3TrafficLights1': '交通灯', '3Signs': '标志', '3TrafficLights2': '交通灯 2', '4RedToBlack': '红至黑', '4TrafficLights': '四向交通灯', '3Symbols': '符号', '3Symbols2': '符号 2', '3Flags': '旗帜', '4Rating': '四级评级', '5Rating': '五级评级', '5Quarters': '五级进度', '_5Felling': '五级表情', '5Boxes': '五级方格', '3Stars': '星级', EMPTY_ICON_TYPE: '无图标' } as Record<string, string>)[value] ?? value
}

function cloneRule(rule: SheetConditionalFormatRule): SheetConditionalFormatRule {
  return cloneModel(rule)
}
</script>

<template>
  <Drawer :open="open" :title="t('conditional.title', '条件格式')" placement="right" :size="384" :mask="false" :closable="true" :focusable="{ focusTriggerAfterClose: false }" :after-open-change="afterDrawerOpenChange" :styles="{ header: { padding: '16px 16px 8px' }, body: { padding: '16px' }, footer: { padding: '16px', display: 'flex', justifyContent: 'flex-end', gap: '8px' } }" @close="emit('update:open', false)">
    <div class="conditional-panel" :style="themeStyle">
      <aside v-if="!editing" class="conditional-rule-list" :aria-label="t('conditional.rules', '规则')">
        <Button size="small" type="primary" @click="newRule"><Plus :size="14" />{{ t('conditional.add', '新建规则') }}</Button>
        <p v-if="!localRules.length" class="conditional-empty">{{ t('conditional.empty', '此工作表还没有条件格式规则。') }}</p>
        <div v-for="(rule, index) in localRules" :key="rule.id" class="conditional-rule-row">
          <Button class="conditional-rule-select" type="text" :aria-pressed="draftId === rule.id" @click="selectRule(rule)">
            {{ index + 1 }}. {{ t(`conditional.operator.${rule.operator}`, rule.operator) }} {{ rule.formula1 }}<template v-if="rule.formula2"> – {{ rule.formula2 }}</template>
          </Button>
          <Tooltip :title="t('conditional.move-up', '上移规则')"><Button type="text" size="small" :disabled="index === 0" :aria-label="t('conditional.move-up', '上移规则')" @click="moveRule(rule.id, -1)"><ArrowUp :size="14" /></Button></Tooltip>
          <Tooltip :title="t('conditional.move-down', '下移规则')"><Button type="text" size="small" :disabled="index === localRules.length - 1" :aria-label="t('conditional.move-down', '下移规则')" @click="moveRule(rule.id, 1)"><ArrowDown :size="14" /></Button></Tooltip>
          <Tooltip :title="t('conditional.delete', '删除规则')"><Button type="text" size="small" danger :aria-label="t('conditional.delete', '删除规则')" @click="removeRule(rule.id)"><Trash2 :size="14" /></Button></Tooltip>
        </div>
      </aside>
      <Form v-else layout="vertical" class="conditional-rule-editor" @submit.prevent>
        <Button type="text" class="conditional-back" @click="returnToRules"><ArrowLeft :size="14" />{{ t('conditional.rules', '规则') }}</Button>
        <FormItem :label="t('conditional.range', '应用范围')" :extra="formulaBasis"><Input v-model:value="draftRanges" :aria-label="t('conditional.range', '应用范围')" /></FormItem>
        <FormItem v-if="ruleType !== 'office'" label="规则类型"><Select v-model:value="ruleType" :options="typeOptions" /></FormItem>
        <Alert v-else type="info" title="原有条件规则" description="保留原有条件，可调整应用范围、格式和顺序。该条件的计算尚未验收。" show-icon />
        <FormItem v-if="ruleType === 'highlightCell'" label="条件类型"><Select v-model:value="highlightSubtype" :options="[{value:'text',label:'文本'},{value:'number',label:'数值'},{value:'timePeriod',label:'日期期间'},{value:'duplicateValues',label:'重复值'},{value:'uniqueValues',label:'唯一值'}]" /></FormItem>
        <FormItem v-if="ruleType === 'highlightCell' && highlightSubtype === 'number'" :label="t('conditional.operator', '条件')"><Select v-model:value="operator" :options="operatorOptions" /></FormItem>
        <FormItem v-if="ruleType === 'highlightCell' && highlightSubtype === 'text'" label="文本条件"><Select v-model:value="operator" :options="[{value:'containsText',label:'包含'},{value:'notContainsText',label:'不包含'},{value:'beginsWith',label:'开头为'},{value:'endsWith',label:'结尾为'},{value:'equal',label:'等于'},{value:'notEqual',label:'不等于'},{value:'containsBlanks',label:'包含空白'},{value:'notContainsBlanks',label:'不包含空白'},{value:'containsErrors',label:'包含错误'},{value:'notContainsErrors',label:'不包含错误'}]" /></FormItem>
        <FormItem v-if="ruleType === 'top10'" label="排名类型"><Select v-model:value="operator" :options="[{value:'isNotBottom',label:'前几项'},{value:'isBottom',label:'后几项'},{value:'greaterThanAverage',label:'高于平均值'},{value:'greaterThanOrEqualAverage',label:'高于或等于平均值'},{value:'lessThanAverage',label:'低于平均值'},{value:'lessThanOrEqualAverage',label:'低于或等于平均值'},{value:'equalAverage',label:'等于平均值'},{value:'notEqualAverage',label:'不等于平均值'}]" /></FormItem>
        <div v-if="typeNeedsValue || ruleType === 'highlightCell' && highlightSubtype === 'timePeriod'" class="conditional-values">
          <FormItem :label="formulaLabel"><Select v-if="ruleType === 'highlightCell' && highlightSubtype === 'timePeriod'" v-model:value="operator" :aria-label="formulaLabel" :options="[{value:'today',label:'今天'},{value:'yesterday',label:'昨天'},{value:'tomorrow',label:'明天'},{value:'last7Days',label:'最近 7 天'},{value:'thisWeek',label:'本周'},{value:'lastWeek',label:'上周'},{value:'nextWeek',label:'下周'},{value:'thisMonth',label:'本月'},{value:'lastMonth',label:'上月'},{value:'nextMonth',label:'下月'}]" /><Input v-else v-model:value="formula1" :aria-label="formulaLabel" /></FormItem>
          <FormItem v-if="ruleType === 'highlightCell' && highlightSubtype === 'number' && needsSecondValue" :label="t('conditional.second-value-formula', '第二个数值或公式')"><Input v-model:value="formula2" :aria-label="t('conditional.second-value-formula', '第二个数值或公式')" /></FormItem>
        </div>
        <div v-if="ruleType === 'top10' && (operator === 'isNotBottom' || operator === 'isBottom')" class="conditional-values"><FormItem label="项数"><InputNumber v-model:value="rank" :min="1" :max="1000" /></FormItem><FormItem :label="null"><Checkbox v-model:checked="rankPercent">百分比</Checkbox></FormItem></div>
        <div v-if="ruleType === 'dataBar' || ruleType === 'colorScale'" class="conditional-rule-editor">
          <ConditionalThresholdInput v-model:value="minimum" boundary="min" :automatic="ruleType === 'dataBar'" label="最小值" />
          <template v-if="ruleType === 'colorScale'">
            <Checkbox v-model:checked="useMidpoint">使用中间值</Checkbox>
            <ConditionalThresholdInput v-if="useMidpoint" v-model:value="midpoint" label="中间值" />
          </template>
          <ConditionalThresholdInput v-model:value="maximum" boundary="max" :automatic="ruleType === 'dataBar'" label="最大值" />
          <div v-if="ruleType === 'dataBar'" class="conditional-colors"><FormItem label="正值颜色"><ColorPicker v-model:value="fill" value-format="hex" /></FormItem><FormItem label="负值颜色"><ColorPicker v-model:value="negativeColor" value-format="hex" /></FormItem></div>
          <template v-if="ruleType === 'dataBar'"><FormItem :label="null"><Checkbox v-model:checked="isGradient">渐变填充</Checkbox></FormItem><FormItem :label="null"><Checkbox v-model:checked="showValue">显示数值</Checkbox></FormItem></template>
          <div v-else class="conditional-colors"><FormItem label="最小值颜色"><ColorPicker v-model:value="minimumColor" value-format="hex" /></FormItem><FormItem v-if="useMidpoint" label="中间值颜色"><ColorPicker v-model:value="fill" value-format="hex" /></FormItem><FormItem label="最大值颜色"><ColorPicker v-model:value="maximumColor" value-format="hex" /></FormItem></div>
        </div>
        <div v-if="ruleType === 'iconSet'" class="conditional-rule-editor">
          <FormItem label="图标集"><Select :value="iconSet" :options="iconSetOptions" @update:value="changeIconSet" /></FormItem>
          <div class="conditional-icon-preview" :aria-label="iconSetLabel(iconSet)"><ConditionalIcon v-for="(icon, index) in previewIcons" :key="index" :icon-type="icon.iconType ?? 'EMPTY_ICON_TYPE'" :icon-id="icon.iconId ?? ''" /></div>
          <div v-for="(threshold, index) in iconThresholds" :key="index" class="conditional-icon-threshold">
            <FormItem :label="`图标 ${index + 1}`"><Select :value="`${threshold.iconType}:${threshold.iconId}`" :options="individualIconOptions" @update:value="changeIndividualIcon(threshold, $event)" /></FormItem>
            <FormItem v-if="index < iconThresholds.length - 1" label="比较"><Select v-model:value="threshold.operator" :options="[{value:'greaterThanOrEqual',label:'大于或等于'},{value:'greaterThan',label:'大于'}]" /></FormItem>
            <ConditionalThresholdInput v-if="index < iconThresholds.length - 1" v-model:value="threshold.value" :label="`阈值 ${index + 1}`" />
            <span v-else class="conditional-hint">其余值使用此图标。</span>
          </div>
          <FormItem :label="null"><Checkbox v-model:checked="showValue">显示数值</Checkbox></FormItem><FormItem :label="null"><Checkbox v-model:checked="reverseIcons">反转图标顺序</Checkbox></FormItem>
        </div>
        <div v-if="styleRule" class="conditional-colors">
          <FormItem :label="null"><Checkbox v-model:checked="useFill">{{ t('conditional.fill', '填充') }}</Checkbox><ColorPicker v-model:value="fill" :disabled="!useFill" value-format="hex" :aria-label="t('conditional.fill', '填充')" /></FormItem>
          <FormItem :label="null"><Checkbox v-model:checked="useFontColor">{{ t('conditional.font-color', '文字颜色') }}</Checkbox><ColorPicker v-model:value="fontColor" :disabled="!useFontColor" value-format="hex" :aria-label="t('conditional.font-color', '文字颜色')" /></FormItem>
          <FormItem :label="null"><Checkbox v-model:checked="bold">{{ t('conditional.bold', '加粗') }}</Checkbox></FormItem>
          <FormItem :label="null"><Checkbox v-model:checked="italic">斜体</Checkbox></FormItem><FormItem :label="null"><Checkbox v-model:checked="underline">下划线</Checkbox></FormItem><FormItem :label="null"><Checkbox v-model:checked="strike">删除线</Checkbox></FormItem>
        </div>
        <FormItem v-if="!visualRule" :label="null"><Checkbox v-model:checked="stopIfTrue">{{ t('conditional.stop-if-true', '匹配后停止') }}</Checkbox></FormItem>
        <Button type="primary" @click="saveDraft">{{ draftId ? t('conditional.update', '更新规则') : t('conditional.save-rule', '保存规则') }}</Button>
      </Form>
      <Alert v-if="error" type="error" :title="error" show-icon role="alert" />
    </div>
    <template #footer>
      <Button @click="emit('update:open', false)">{{ t('common.cancel', '取消') }}</Button>
      <Button type="primary" @click="apply">{{ t('common.apply', '应用') }}</Button>
    </template>
  </Drawer>
</template>

<style scoped>
.conditional-panel { display: grid; min-height: 0; }
.conditional-rule-list, .conditional-rule-editor { display: grid; align-content: start; gap: 10px; min-width: 0; }
.conditional-empty { margin: 0; padding: 12px 0; color: var(--ribbon-text-secondary); font-size: 12px; }
.conditional-rule-row { display: grid; grid-template-columns: minmax(0, 1fr) auto auto auto; gap: 4px; align-items: center; }
.conditional-rule-select.ant-btn { min-width: 0; height: auto; justify-content: flex-start; overflow: hidden; padding: 6px 8px; border: 1px solid var(--ribbon-border); background: var(--ribbon-background); color: var(--ribbon-text); text-align: left; text-overflow: ellipsis; white-space: nowrap; }
.conditional-rule-select.ant-btn:hover { background: var(--ribbon-hover-background); }
.conditional-rule-select.ant-btn[aria-pressed="true"] { border-color: var(--ribbon-primary); background: var(--ribbon-active-background); color: var(--ribbon-primary); }
.conditional-rule-select.ant-btn:focus-visible { outline: 2px solid var(--ribbon-primary); outline-offset: 2px; }
.conditional-rule-editor :deep(.ant-form-item) { min-width: 0; margin-bottom: 0; }
.conditional-back { justify-self: start; padding-inline: 0; }
.conditional-values, .conditional-colors { display: flex; flex-wrap: wrap; align-items: end; gap: 10px; }
.conditional-values > :deep(.ant-form-item) { flex: 1; min-width: 0; }
.conditional-colors > :deep(.ant-form-item) { flex: 1; min-width: 0; }
.conditional-colors :deep(.ant-color-picker-trigger) { width: 100%; }
.conditional-icon-preview { display: flex; align-items: center; gap: 4px; min-height: 28px; }
.conditional-icon-swatch { font-size: 11px; color: var(--ribbon-text-secondary); }
.conditional-icon-threshold { display: grid; gap: 8px; padding-block: 10px; border-top: 1px solid var(--ribbon-border); }
.conditional-hint { color: var(--ribbon-text-secondary); font-size: 12px; }
@media (max-width: 620px) { .conditional-panel { grid-template-columns: minmax(0, 1fr); } }
</style>
