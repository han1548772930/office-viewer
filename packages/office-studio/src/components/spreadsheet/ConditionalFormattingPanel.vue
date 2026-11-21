<script setup lang="ts">
import { computed, inject, ref, watch } from 'vue'
import { Button, Checkbox, ColorPicker, Drawer, Input, InputNumber, Select, Tooltip } from 'antdv-next'
import { ArrowDown, ArrowLeft, ArrowUp, Plus, Trash2 } from 'lucide-vue-next'
import type { CellRange, SheetConditionalFormatOperator, SheetConditionalFormatRule, SheetConditionalFormatType, SheetConditionalFormatValueConfig } from '../../domain/model.ts'
import { validateConditionalFormatDraft } from '../../features/conditional-formatting/conditional-formatting.ts'
import { conditionalIconPresets } from '../../features/conditional-formatting/conditional-icon-presets.ts'
import { conditionalIconSequence, createConditionalIconConfig, validateConditionalIconConfig, validateConditionalThreshold } from '../../features/conditional-formatting/icon-set-config.ts'
import ConditionalIcon from './ConditionalIcon.vue'
import ConditionalThresholdInput from './ConditionalThresholdInput.vue'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

const props = defineProps<{ open: boolean; rules: SheetConditionalFormatRule[]; defaultRange: string }>()
const emit = defineEmits<{ 'update:open': [value: boolean]; apply: [rules: SheetConditionalFormatRule[]] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
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
const styleRule = computed(() => ruleType.value === 'highlightCell' || ruleType.value === 'top10' || ruleType.value === 'expression')
const typeNeedsValue = computed(() => ruleType.value === 'expression' || ruleType.value === 'highlightCell' && (highlightSubtype.value === 'number' || highlightSubtype.value === 'text' && !['containsBlanks', 'notContainsBlanks', 'containsErrors', 'notContainsErrors'].includes(operator.value)))

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
  ruleType.value = ['cellIs', 'highlightCell', 'containsText', 'notContainsText', 'beginsWith', 'endsWith', 'duplicateValues', 'uniqueValues', 'timePeriod'].includes(rule.type) ? 'highlightCell' : rule.type === 'average' ? 'top10' : rule.type === 'expression' ? 'expression' : rule.type
  highlightSubtype.value = rule.subType === 'text' || ['containsText', 'notContainsText', 'beginsWith', 'endsWith'].includes(rule.type) ? 'text'
    : rule.subType === 'timePeriod' || rule.type === 'timePeriod' ? 'timePeriod'
      : rule.subType === 'duplicateValues' || rule.type === 'duplicateValues' ? 'duplicateValues'
        : rule.subType === 'uniqueValues' || rule.type === 'uniqueValues' ? 'uniqueValues' : 'number'
  draftRanges.value = rule.ranges.map(formatRange).join(';')
  operator.value = ['notContainsText', 'beginsWith', 'endsWith'].includes(rule.type) ? rule.type : rule.operator || (highlightSubtype.value === 'text' ? 'containsText' : 'greaterThan')
  formula1.value = rule.formula1
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
  if (rule.type === 'top10') operator.value = rule.isBottom ? 'isBottom' : 'isNotBottom'
  if (rule.type === 'average') operator.value = rule.operator === 'lessThan' || rule.operator === 'lessThanOrEqual' ? 'lessThanAverage' : rule.operator === 'equal' ? 'equalAverage' : 'greaterThanAverage'
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
    const previousStyle = existingIndex >= 0 ? localRules.value[existingIndex]!.style : {}
    const style = {
      fill: useFill.value ? fill.value : undefined,
      fontColor: useFontColor.value ? fontColor.value : undefined,
      bold: bold.value || previousStyle.bold !== undefined ? bold.value : undefined,
      italic: italic.value || previousStyle.italic !== undefined ? italic.value : undefined,
      underline: underline.value || previousStyle.underline !== undefined ? underline.value : undefined,
      strike: strike.value || previousStyle.strike !== undefined ? strike.value : undefined,
    }
    const draft = { ranges, operator: operator.value as SheetConditionalFormatOperator, formula1: formula1.value, formula2: needsSecondValue.value ? formula2.value : undefined, style, stopIfTrue: stopIfTrue.value }
    const normalizedType = ruleType.value === 'highlightCell'
      ? highlightSubtype.value === 'number' ? 'cellIs' : highlightSubtype.value === 'text' ? operator.value === 'notContainsText' ? 'notContainsText' : operator.value === 'beginsWith' ? 'beginsWith' : operator.value === 'endsWith' ? 'endsWith' : 'containsText' : highlightSubtype.value
      : ruleType.value === 'top10' && ['greaterThanAverage', 'lessThanAverage', 'equalAverage'].includes(operator.value) ? 'average' : ruleType.value === 'expression' ? 'expression' : ruleType.value
    const normalizedSubType = ruleType.value === 'highlightCell' ? highlightSubtype.value : ruleType.value === 'top10' ? normalizedType === 'average' ? 'average' : 'rank' : ruleType.value === 'expression' ? 'formula' : undefined
    if (normalizedType === 'cellIs') validateConditionalFormatDraft(draft)
    if (normalizedType === 'expression' && !formula1.value.trim().replace(/^=/, '')) throw new TypeError('请输入条件公式。')
    if (normalizedType === 'top10' && (!Number.isInteger(rank.value) || rank.value < 1 || rank.value > (rankPercent.value ? 100 : 1000))) throw new RangeError('排名项数无效。')
    if (ruleType.value === 'iconSet') validateConditionalIconConfig({ icons: iconThresholds.value })
    else if (visualRule.value) {
      const thresholds = [minimum.value, maximum.value, ...(ruleType.value === 'colorScale' && useMidpoint.value ? [midpoint.value] : [])]
      for (const threshold of thresholds) validateConditionalThreshold(threshold)
    }
    const id = draftId.value ?? crypto.randomUUID()
    const nextRule: SheetConditionalFormatRule = {
      id, type: normalizedType, ranges, operator: ruleType.value === 'top10' ? operator.value === 'lessThanAverage' ? 'lessThan' : operator.value === 'equalAverage' ? 'equal' : 'greaterThan' : operator.value,
      subType: normalizedSubType,
      formula1: normalizedSubType === 'text' ? formula1.value : formula1.value.trim().replace(/^=/, ''), formula2: draft.formula2?.trim().replace(/^=/, ''),
      value: ['top10'].includes(ruleType.value) ? rank.value : ruleType.value === 'highlightCell' && highlightSubtype.value === 'text' ? formula1.value : undefined,
      isBottom: ruleType.value === 'top10' ? operator.value === 'isBottom' : undefined,
      isPercent: ruleType.value === 'top10' ? rankPercent.value : undefined,
      config: ruleType.value === 'dataBar' ? { min: { ...minimum.value }, max: { ...maximum.value }, isGradient: isGradient.value, positiveColor: fill.value, nativeColor: negativeColor.value, isShowValue: showValue.value }
        : ruleType.value === 'colorScale' ? { thresholds: [{ value: { ...minimum.value }, color: minimumColor.value }, ...(useMidpoint.value ? [{ value: { ...midpoint.value }, color: fill.value }] : []), { value: { ...maximum.value }, color: maximumColor.value }] }
          : ruleType.value === 'iconSet' ? { iconSet: iconSet.value, isShowValue: showValue.value, reverse: reverseIcons.value, icons: iconThresholds.value.map(icon => ({ ...icon, value: { ...icon.value } })) }
            : undefined,
      ...(existingIndex >= 0 && localRules.value[existingIndex]?.metadata ? { metadata: { ...localRules.value[existingIndex]!.metadata } } : {}),
      style, priority: existingIndex < 0 ? localRules.value.length + 1 : localRules.value[existingIndex]!.priority,
      stopIfTrue: stopIfTrue.value,
    }
    if (existingIndex < 0) localRules.value = reprioritize([nextRule, ...localRules.value])
    else localRules.value = localRules.value.map(rule => rule.id === id ? nextRule : rule)
    selectRule(nextRule)
    editing.value = false
    error.value = ''
  } catch (reason) {
    error.value = reason instanceof Error ? reason.message : t('conditional.error.invalid', '条件格式规则无效。')
  }
}

function removeRule(id: string) {
  const wasEditing = editing.value
  localRules.value = reprioritize(localRules.value.filter(rule => rule.id !== id))
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
  const index = localRules.value.findIndex(rule => rule.id === id)
  const target = index + offset
  if (index < 0 || target < 0 || target >= localRules.value.length) return
  const rules = [...localRules.value]
  const [item] = rules.splice(index, 1)
  rules.splice(target, 0, item!)
  localRules.value = reprioritize(rules)
}

function apply() {
  if (editing.value) saveDraft()
  if (error.value) return
  emit('apply', localRules.value.map(cloneRule))
  emit('update:open', false)
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

function reprioritize(rules: SheetConditionalFormatRule[]): SheetConditionalFormatRule[] {
  return rules.map((rule, index) => ({ ...rule, priority: index + 1 }))
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
  return { ...rule, ranges: rule.ranges.map(range => ({ ...range })), style: { ...rule.style }, config: rule.config ? { ...rule.config, min: rule.config.min ? { ...rule.config.min } : undefined, max: rule.config.max ? { ...rule.config.max } : undefined, thresholds: rule.config.thresholds?.map(item => ({ ...item, value: { ...item.value } })), icons: rule.config.icons?.map(item => ({ ...item, value: { ...item.value } })) } : undefined, metadata: rule.metadata ? { ...rule.metadata } : undefined }
}
</script>

<template>
  <Drawer :open="open" :title="t('conditional.title', '条件格式')" placement="right" :width="384" :mask="false" :closable="true" :styles="{ header: { padding: '16px 16px 8px' }, body: { padding: '16px' }, footer: { padding: '16px' } }" :footer-style="{ display: 'flex', justifyContent: 'flex-end', gap: '8px' }" @close="emit('update:open', false)">
    <div class="conditional-panel">
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
      <section v-else class="conditional-rule-editor">
        <Button type="text" class="conditional-back" @click="returnToRules"><ArrowLeft :size="14" />{{ t('conditional.rules', '规则') }}</Button>
        <label>{{ t('conditional.range', '应用范围') }}<Input v-model:value="draftRanges" :aria-label="t('conditional.range', '应用范围')" /></label>
        <label>规则类型<Select v-model:value="ruleType" :options="typeOptions" /></label>
        <label v-if="ruleType === 'highlightCell'">条件类型<Select v-model:value="highlightSubtype" :options="[{value:'text',label:'文本'},{value:'number',label:'数值'},{value:'timePeriod',label:'日期期间'},{value:'duplicateValues',label:'重复值'},{value:'uniqueValues',label:'唯一值'}]" /></label>
        <label v-if="ruleType === 'highlightCell' && highlightSubtype === 'number'">{{ t('conditional.operator', '条件') }}<Select v-model:value="operator" :options="operatorOptions" /></label>
        <label v-if="ruleType === 'highlightCell' && highlightSubtype === 'text'">文本条件<Select v-model:value="operator" :options="[{value:'containsText',label:'包含'},{value:'notContainsText',label:'不包含'},{value:'beginsWith',label:'开头为'},{value:'endsWith',label:'结尾为'},{value:'equal',label:'等于'},{value:'notEqual',label:'不等于'},{value:'containsBlanks',label:'包含空白'},{value:'notContainsBlanks',label:'不包含空白'},{value:'containsErrors',label:'包含错误'},{value:'notContainsErrors',label:'不包含错误'}]" /></label>
        <label v-if="ruleType === 'top10'">排名类型<Select v-model:value="operator" :options="[{value:'isNotBottom',label:'前几项'},{value:'isBottom',label:'后几项'},{value:'greaterThanAverage',label:'高于平均值'},{value:'lessThanAverage',label:'低于平均值'},{value:'equalAverage',label:'等于平均值'}]" /></label>
        <div v-if="typeNeedsValue || ruleType === 'highlightCell' && highlightSubtype === 'timePeriod'" class="conditional-values">
          <label>{{ ruleType === 'expression' ? '公式' : highlightSubtype === 'number' ? t('conditional.value', '数值') : highlightSubtype === 'timePeriod' ? '日期期间' : '匹配内容' }}<Select v-if="ruleType === 'highlightCell' && highlightSubtype === 'timePeriod'" v-model:value="operator" :options="[{value:'today',label:'今天'},{value:'yesterday',label:'昨天'},{value:'tomorrow',label:'明天'},{value:'last7Days',label:'最近 7 天'},{value:'thisWeek',label:'本周'},{value:'lastWeek',label:'上周'},{value:'nextWeek',label:'下周'},{value:'thisMonth',label:'本月'},{value:'lastMonth',label:'上月'},{value:'nextMonth',label:'下月'}]" /><Input v-else v-model:value="formula1" :inputmode="ruleType === 'highlightCell' && highlightSubtype === 'number' ? 'decimal' : undefined" /></label>
          <label v-if="ruleType === 'highlightCell' && highlightSubtype === 'number' && needsSecondValue">{{ t('conditional.second-value', '第二个数值') }}<Input v-model:value="formula2" inputmode="decimal" /></label>
        </div>
        <div v-if="ruleType === 'top10' && (operator === 'isNotBottom' || operator === 'isBottom')" class="conditional-values"><label>项数<InputNumber v-model:value="rank" :min="1" :max="1000" /></label><Checkbox v-model:checked="rankPercent">百分比</Checkbox></div>
        <div v-if="ruleType === 'dataBar' || ruleType === 'colorScale'" class="conditional-rule-editor">
          <ConditionalThresholdInput v-model:value="minimum" boundary="min" label="最小值" />
          <template v-if="ruleType === 'colorScale'">
            <Checkbox v-model:checked="useMidpoint">使用中间值</Checkbox>
            <ConditionalThresholdInput v-if="useMidpoint" v-model:value="midpoint" label="中间值" />
          </template>
          <ConditionalThresholdInput v-model:value="maximum" boundary="max" label="最大值" />
          <div v-if="ruleType === 'dataBar'" class="conditional-colors"><label>正值颜色<ColorPicker v-model:value="fill" value-format="hex" /></label><label>负值颜色<ColorPicker v-model:value="negativeColor" value-format="hex" /></label></div>
          <template v-if="ruleType === 'dataBar'"><Checkbox v-model:checked="isGradient">渐变填充</Checkbox><Checkbox v-model:checked="showValue">显示数值</Checkbox></template>
          <div v-else class="conditional-colors"><label>最小值颜色<ColorPicker v-model:value="minimumColor" value-format="hex" /></label><label v-if="useMidpoint">中间值颜色<ColorPicker v-model:value="fill" value-format="hex" /></label><label>最大值颜色<ColorPicker v-model:value="maximumColor" value-format="hex" /></label></div>
        </div>
        <div v-if="ruleType === 'iconSet'" class="conditional-rule-editor">
          <label>图标集<Select :value="iconSet" :options="iconSetOptions" @update:value="changeIconSet" /></label>
          <div class="conditional-icon-preview" :aria-label="iconSetLabel(iconSet)"><ConditionalIcon v-for="(icon, index) in previewIcons" :key="index" :icon-type="icon.iconType ?? 'EMPTY_ICON_TYPE'" :icon-id="icon.iconId ?? ''" /></div>
          <div v-for="(threshold, index) in iconThresholds" :key="index" class="conditional-icon-threshold">
            <label>图标 {{ index + 1 }}<Select :value="`${threshold.iconType}:${threshold.iconId}`" :options="individualIconOptions" @update:value="changeIndividualIcon(threshold, $event)" /></label>
            <label v-if="index < iconThresholds.length - 1">比较<Select v-model:value="threshold.operator" :options="[{value:'greaterThanOrEqual',label:'大于或等于'},{value:'greaterThan',label:'大于'}]" /></label>
            <ConditionalThresholdInput v-if="index < iconThresholds.length - 1" v-model:value="threshold.value" :label="`阈值 ${index + 1}`" />
            <span v-else class="conditional-hint">其余值使用此图标。</span>
          </div>
          <Checkbox v-model:checked="showValue">显示数值</Checkbox><Checkbox v-model:checked="reverseIcons">反转图标顺序</Checkbox>
        </div>
        <div v-if="styleRule" class="conditional-colors">
          <label><Checkbox v-model:checked="useFill">{{ t('conditional.fill', '填充') }}</Checkbox><ColorPicker v-model:value="fill" :disabled="!useFill" value-format="hex" :aria-label="t('conditional.fill', '填充')" /></label>
          <label><Checkbox v-model:checked="useFontColor">{{ t('conditional.font-color', '文字颜色') }}</Checkbox><ColorPicker v-model:value="fontColor" :disabled="!useFontColor" value-format="hex" :aria-label="t('conditional.font-color', '文字颜色')" /></label>
          <Checkbox v-model:checked="bold">{{ t('conditional.bold', '加粗') }}</Checkbox>
          <Checkbox v-model:checked="italic">斜体</Checkbox><Checkbox v-model:checked="underline">下划线</Checkbox><Checkbox v-model:checked="strike">删除线</Checkbox>
        </div>
        <Checkbox v-if="!visualRule" v-model:checked="stopIfTrue">{{ t('conditional.stop-if-true', '匹配后停止') }}</Checkbox>
        <p v-if="error" class="conditional-error" role="alert">{{ error }}</p>
        <Button type="primary" @click="saveDraft">{{ draftId ? t('conditional.update', '更新规则') : t('conditional.save-rule', '保存规则') }}</Button>
      </section>
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
.conditional-empty { margin: 0; padding: 12px 0; color: #667085; font-size: 12px; }
.conditional-rule-row { display: grid; grid-template-columns: minmax(0, 1fr) auto auto auto; gap: 4px; align-items: center; }
.conditional-rule-select.ant-btn { min-width: 0; height: auto; justify-content: flex-start; overflow: hidden; padding: 6px 8px; border: 1px solid #d0d5dd; background: #fff; color: #344054; text-align: left; text-overflow: ellipsis; white-space: nowrap; }
.conditional-rule-select.ant-btn[aria-pressed="true"] { border-color: #246bfd; background: #eff4ff; color: #244c91; }
.conditional-rule-editor label { display: grid; gap: 5px; color: #475467; font-size: 12px; }
.conditional-back { justify-self: start; padding-inline: 0; }
.conditional-values, .conditional-colors { display: flex; flex-wrap: wrap; align-items: end; gap: 10px; }
.conditional-values > label { flex: 1; min-width: 0; }
.conditional-colors > label { flex: 1; min-width: 0; }
.conditional-colors :deep(.ant-color-picker-trigger) { width: 100%; }
.conditional-icon-preview { display: flex; align-items: center; gap: 4px; min-height: 28px; }
.conditional-icon-swatch { font-size: 11px; color: #667085; }
.conditional-icon-threshold { display: grid; gap: 8px; padding-block: 10px; border-top: 1px solid #e5e7eb; }
.conditional-hint { color: #667085; font-size: 12px; }
.conditional-error { margin: 0; color: #b42318; font-size: 12px; }
@media (max-width: 620px) { .conditional-panel { grid-template-columns: minmax(0, 1fr); } }
</style>
