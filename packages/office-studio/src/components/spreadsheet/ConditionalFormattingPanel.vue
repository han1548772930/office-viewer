<script setup lang="ts">
import { computed, inject, ref, watch } from 'vue'
import { Button, Input, Modal, Select } from 'antdv-next'
import type { CellRange, SheetConditionalFormatOperator, SheetConditionalFormatRule } from '../../domain/model.ts'
import { validateConditionalFormatDraft } from '../../features/conditional-formatting/conditional-formatting.ts'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

const props = defineProps<{ open: boolean; rules: SheetConditionalFormatRule[]; defaultRange: string }>()
const emit = defineEmits<{ 'update:open': [value: boolean]; apply: [rules: SheetConditionalFormatRule[]] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const draftId = ref<string | null>(null)
const draftRanges = ref('A1')
const operator = ref<SheetConditionalFormatOperator>('greaterThan')
const formula1 = ref('0')
const formula2 = ref('1')
const fill = ref('#fff2cc')
const fontColor = ref('#9c6500')
const bold = ref(false)
const stopIfTrue = ref(false)
const error = ref('')
const localRules = ref<SheetConditionalFormatRule[]>([])
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

watch(() => props.open, open => {
  if (!open) return
  localRules.value = props.rules.map(cloneRule)
  error.value = ''
  if (localRules.value.length) selectRule(localRules.value[0]!)
  else newRule()
})

function newRule() {
  draftId.value = null
  draftRanges.value = props.defaultRange || 'A1'
  operator.value = 'greaterThan'
  formula1.value = '0'
  formula2.value = '1'
  fill.value = '#fff2cc'
  fontColor.value = '#9c6500'
  bold.value = false
  stopIfTrue.value = false
  error.value = ''
}

function selectRule(rule: SheetConditionalFormatRule) {
  draftId.value = rule.id
  draftRanges.value = rule.ranges.map(formatRange).join(';')
  operator.value = rule.operator
  formula1.value = rule.formula1
  formula2.value = rule.formula2 ?? '1'
  fill.value = rule.style.fill ?? '#ffffff'
  fontColor.value = rule.style.fontColor ?? '#1f2937'
  bold.value = rule.style.bold ?? false
  stopIfTrue.value = rule.stopIfTrue
  error.value = ''
}

function saveDraft() {
  try {
    const ranges = parseRanges(draftRanges.value)
    const style = {
      fill: fill.value.toLowerCase() === '#ffffff' ? undefined : fill.value,
      fontColor: fontColor.value.toLowerCase() === '#1f2937' ? undefined : fontColor.value,
      bold: bold.value || undefined,
    }
    const draft = { ranges, operator: operator.value, formula1: formula1.value, formula2: needsSecondValue.value ? formula2.value : undefined, style, stopIfTrue: stopIfTrue.value }
    validateConditionalFormatDraft(draft)
    const existingIndex = localRules.value.findIndex(rule => rule.id === draftId.value)
    const id = draftId.value ?? crypto.randomUUID()
    const nextRule: SheetConditionalFormatRule = {
      id, type: 'cellIs', ranges, operator: operator.value,
      formula1: formula1.value.trim().replace(/^=/, ''), formula2: draft.formula2?.trim().replace(/^=/, ''),
      style, priority: existingIndex < 0 ? localRules.value.length + 1 : localRules.value[existingIndex]!.priority,
      stopIfTrue: stopIfTrue.value,
    }
    if (existingIndex < 0) localRules.value = [...localRules.value, nextRule]
    else localRules.value = localRules.value.map(rule => rule.id === id ? nextRule : rule)
    selectRule(nextRule)
    error.value = ''
  } catch (reason) {
    error.value = reason instanceof Error ? reason.message : t('conditional.error.invalid', '条件格式规则无效。')
  }
}

function removeRule(id: string) {
  localRules.value = reprioritize(localRules.value.filter(rule => rule.id !== id))
  if (draftId.value === id) localRules.value.length ? selectRule(localRules.value[0]!) : newRule()
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
  if (error.value) return
  if (draftId.value === null || !localRules.value.some(rule => rule.id === draftId.value)) saveDraft()
  if (error.value) return
  emit('apply', localRules.value.map(cloneRule))
  emit('update:open', false)
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
  if (row < 1 || !Number.isSafeInteger(row)) return undefined
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

function cloneRule(rule: SheetConditionalFormatRule): SheetConditionalFormatRule {
  return { ...rule, ranges: rule.ranges.map(range => ({ ...range })), style: { ...rule.style } }
}
</script>

<template>
  <Modal :open="open" :title="t('conditional.title', '条件格式')" :width="760" :ok-text="t('common.apply', '应用')" :cancel-text="t('common.cancel', '取消')" @ok="apply" @cancel="emit('update:open', false)">
    <div class="conditional-panel">
      <aside class="conditional-rule-list" :aria-label="t('conditional.rules', '规则')">
        <Button size="small" @click="newRule">{{ t('conditional.add', '新建规则') }}</Button>
        <div v-for="(rule, index) in localRules" :key="rule.id" class="conditional-rule-row">
          <button class="conditional-rule-select" :aria-pressed="draftId === rule.id" @click="selectRule(rule)">
            {{ index + 1 }}. {{ t(`conditional.operator.${rule.operator}`, rule.operator) }} {{ rule.formula1 }}<template v-if="rule.formula2"> – {{ rule.formula2 }}</template>
          </button>
          <Button size="small" :disabled="index === 0" :aria-label="t('conditional.move-up', '上移规则')" @click="moveRule(rule.id, -1)">↑</Button>
          <Button size="small" :disabled="index === localRules.length - 1" :aria-label="t('conditional.move-down', '下移规则')" @click="moveRule(rule.id, 1)">↓</Button>
          <Button size="small" :aria-label="t('conditional.delete', '删除规则')" @click="removeRule(rule.id)">×</Button>
        </div>
      </aside>
      <section class="conditional-rule-editor">
        <label>{{ t('conditional.range', '应用范围') }}<Input v-model:value="draftRanges" :aria-label="t('conditional.range', '应用范围')" /></label>
        <label>{{ t('conditional.operator', '条件') }}<Select v-model:value="operator" :options="operatorOptions" /></label>
        <div class="conditional-values">
          <label>{{ t('conditional.value', '数值') }}<Input v-model:value="formula1" inputmode="decimal" /></label>
          <label v-if="needsSecondValue">{{ t('conditional.second-value', '第二个数值') }}<Input v-model:value="formula2" inputmode="decimal" /></label>
        </div>
        <div class="conditional-colors">
          <label>{{ t('conditional.fill', '填充') }}<input v-model="fill" type="color" :aria-label="t('conditional.fill', '填充')" /></label>
          <label>{{ t('conditional.font-color', '文字颜色') }}<input v-model="fontColor" type="color" :aria-label="t('conditional.font-color', '文字颜色')" /></label>
          <label class="conditional-check"><input v-model="bold" type="checkbox" />{{ t('conditional.bold', '加粗') }}</label>
        </div>
        <label class="conditional-check"><input v-model="stopIfTrue" type="checkbox" />{{ t('conditional.stop-if-true', '匹配后停止') }}</label>
        <p v-if="error" class="conditional-error" role="alert">{{ error }}</p>
        <Button type="primary" @click="saveDraft">{{ draftId ? t('conditional.update', '更新规则') : t('conditional.save-rule', '保存规则') }}</Button>
      </section>
    </div>
  </Modal>
</template>

<style scoped>
.conditional-panel { display: grid; grid-template-columns: minmax(220px, 0.85fr) minmax(280px, 1.15fr); gap: 18px; min-height: 300px; }
.conditional-rule-list, .conditional-rule-editor { display: grid; align-content: start; gap: 10px; min-width: 0; }
.conditional-rule-row { display: grid; grid-template-columns: minmax(0, 1fr) auto auto auto; gap: 4px; align-items: center; }
.conditional-rule-select { min-width: 0; overflow: hidden; padding: 6px 8px; border: 1px solid #d0d5dd; border-radius: 4px; background: #fff; color: #344054; text-align: left; text-overflow: ellipsis; white-space: nowrap; }
.conditional-rule-select[aria-pressed="true"] { border-color: #246bfd; background: #eff4ff; }
.conditional-rule-editor label { display: grid; gap: 5px; color: #475467; font-size: 12px; }
.conditional-values, .conditional-colors { display: flex; align-items: end; gap: 10px; }
.conditional-values > label { flex: 1; min-width: 0; }
.conditional-colors > label { flex: 1; min-width: 0; }
.conditional-colors input[type="color"] { width: 100%; height: 34px; padding: 3px; border: 1px solid #d0d5dd; border-radius: 4px; background: #fff; }
.conditional-rule-editor .conditional-check { display: flex; align-items: center; gap: 8px; }
.conditional-error { margin: 0; color: #b42318; font-size: 12px; }
@media (max-width: 620px) { .conditional-panel { grid-template-columns: minmax(0, 1fr); } }
</style>
