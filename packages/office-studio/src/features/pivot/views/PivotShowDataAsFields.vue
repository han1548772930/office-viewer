<script setup lang="ts">
import { Alert, Button, FormItem, Select } from 'antdv-next'
import { computed, onScopeDispose, ref, watch } from 'vue'
import type { PivotShowDataAsInfo, PivotTableModel } from '../../../domain/model.ts'
import { getPivotShowDataAsRule } from '../services/pivot-show-data-as.ts'
import { changePivotShowDataAsBaseField, changePivotShowDataAsBaseItemType, changePivotShowDataAsType,
  createPivotShowDataAsDraft, PIVOT_SHOW_DATA_AS_OPTIONS, pivotShowDataAsBaseFieldOptions, pivotShowDataAsDraftError,
  pivotShowDataAsMemberRevision, type PivotShowDataAsBaseItemsLoader } from '../services/pivot-show-data-as-ui.ts'

const props = defineProps<{ value?: PivotShowDataAsInfo | null; report?: PivotTableModel; labelPrefix?: string;
  loadBaseItems: PivotShowDataAsBaseItemsLoader; contextKey?: string; readable?: boolean; disabled?: boolean }>()
const emit = defineEmits<{ 'update:value': [value: PivotShowDataAsInfo | null]; 'validity-change': [valid: boolean] }>()
const value = computed(() => createPivotShowDataAsDraft(props.value))
const rule = computed(() => getPivotShowDataAsRule(value.value.type))
const readable = computed(() => props.readable !== false && Boolean(props.report))
const fields = computed(() => readable.value && props.report ? pivotShowDataAsBaseFieldOptions(props.report, value.value.type) : [])
const memberItems = ref<string[]>([]), loading = ref(false), readError = ref(''), loaded = ref(false), retry = ref(0)
let generation = 0, disposed = false
const hasMembers = computed(() => rule.value.baseItemMode !== 'none')
const explicitItem = computed(() => value.value.baseItemType !== 'previous' && value.value.baseItemType !== 'next')
const selectedFieldValid = computed(() => fields.value.some(field => field.value === value.value.baseFieldId))
const memberOptions = computed(() => readable.value ? memberItems.value.map(item => ({ value: item, label: item || '（空白）', disabled: item === '' })) : [])
const relativeOptions = [{ value: 'item', label: '指定成员' }, { value: 'previous', label: '上一项' }, { value: 'next', label: '下一项' }]
const validationError = computed(() => !readable.value ? '请重新读取可访问的报告字段。' : readError.value
  || pivotShowDataAsDraftError(props.report, value.value, hasMembers.value && explicitItem.value && loaded.value ? memberItems.value : undefined))
const valid = computed(() => readable.value && !validationError.value && (!hasMembers.value || !explicitItem.value || loaded.value && !loading.value))
const prefix = computed(() => props.labelPrefix ? `${props.labelPrefix}的` : '')
const memberKey = computed(() => JSON.stringify([readable.value, props.contextKey, pivotShowDataAsMemberRevision(props.report),
  value.value.type, value.value.baseFieldId, hasMembers.value, explicitItem.value, retry.value]))

watch(memberKey, async () => {
  const request = ++generation
  memberItems.value = []; readError.value = ''; loaded.value = false; loading.value = false
  if (!readable.value || !hasMembers.value || !explicitItem.value || !selectedFieldValid.value) return
  loading.value = true
  try {
    const items = await props.loadBaseItems(value.value.baseFieldId)
    if (disposed || request !== generation || !readable.value) return
    if (!Array.isArray(items) || items.some(item => typeof item !== 'string')) throw new TypeError('基准成员读取结果无效。')
    memberItems.value = [...new Set(items)]; loaded.value = true
  } catch (reason) { if (!disposed && request === generation) readError.value = reason instanceof Error ? reason.message : String(reason) }
  finally { if (!disposed && request === generation) loading.value = false }
}, { immediate: true, flush: 'sync' })
watch(valid, current => emit('validity-change', current), { immediate: true, flush: 'sync' })
onScopeDispose(() => { disposed = true; generation++ })

function changeType(input: unknown): void {
  const option = PIVOT_SHOW_DATA_AS_OPTIONS.find(option => option.value === input)
  if (!props.disabled && props.report && readable.value && option) emit('update:value', changePivotShowDataAsType(props.report, value.value, option.value))
}
function changeField(input: unknown): void {
  if (!props.disabled && props.report && readable.value && typeof input === 'string') emit('update:value', changePivotShowDataAsBaseField(props.report, value.value, input))
}
function changeItemType(input: unknown): void {
  if (!props.disabled && readable.value && (input === 'item' || input === 'previous' || input === 'next')) emit('update:value', changePivotShowDataAsBaseItemType(value.value, input))
}
function changeItem(input: unknown): void {
  if (!props.disabled && readable.value && explicitItem.value && typeof input === 'string' && input && memberItems.value.includes(input)) {
    emit('update:value', { type: value.value.type, baseFieldId: value.value.baseFieldId, baseItem: input })
  }
}
</script>

<template>
  <div class="pivot-show-data-as" data-testid="pivot-show-data-as">
    <FormItem :label="`${prefix}值显示方式`"><Select :value="rule.compatibilityOnly ? undefined : value.type" :options="PIVOT_SHOW_DATA_AS_OPTIONS" :virtual="false"
      :aria-label="`${prefix}值显示方式`" :placeholder="rule.compatibilityOnly ? '保留导入的兼容方式' : '选择值显示方式'"
      :disabled="disabled || !readable" @change="changeType" /></FormItem>
    <Alert v-if="rule.compatibilityOnly" type="info" :title="`已保留导入的“${rule.label}”。可选择固定版本方式或恢复无计算。`" show-icon />
    <FormItem v-if="rule.requiresParameters" :label="`${prefix}基准字段`"><Select :value="readable && selectedFieldValid ? value.baseFieldId : undefined" :options="fields"
      :aria-label="`${prefix}基准字段`" placeholder="选择报告中的行或列字段" :disabled="disabled || !readable" @change="changeField" /></FormItem>
    <Alert v-if="rule.requiresParameters && !fields.length && readable" type="info" :title="rule.parentFieldOnly ? '此方式需要具有子层级的行或列字段。' : '此方式需要一个已配置的行或列字段。'" show-icon />
    <FormItem v-if="rule.baseItemMode === 'item-or-relative'" :label="`${prefix}基准项类型`"><Select :value="value.baseItemType ?? 'item'" :options="relativeOptions"
      :aria-label="`${prefix}基准项类型`" :disabled="disabled || !readable || !selectedFieldValid" @change="changeItemType" /></FormItem>
    <FormItem v-if="hasMembers && explicitItem" :label="`${prefix}基准项`"><Select :value="readable && value.baseItem ? value.baseItem : undefined" :options="memberOptions"
      :aria-label="`${prefix}基准项`" placeholder="选择真实来源成员" show-search option-filter-prop="label" :loading="loading"
      :disabled="disabled || !readable || !selectedFieldValid || loading" @change="changeItem" /></FormItem>
    <Alert v-if="memberItems.includes('') && hasMembers && explicitItem && readable" type="info" title="来源包含空白成员；当前成员协议不能将空白设为显式基准项。" show-icon />
    <Alert v-if="validationError" type="warning" :title="validationError" show-icon />
    <Button v-if="readError" size="small" :disabled="disabled || !readable || loading" @click="retry++">重新读取基准成员</Button>
    <Button size="small" :aria-label="`${prefix}恢复无计算`" :disabled="disabled || !readable" @click="emit('update:value', null)">恢复无计算</Button>
  </div>
</template>

<style scoped>
.pivot-show-data-as { display: grid; gap: 8px; min-width: 0; }
.pivot-show-data-as :deep(.ant-form-item) { min-width: 0; margin-bottom: 0; }
.pivot-show-data-as :deep(.ant-select) { width: 100%; }
.pivot-show-data-as > :deep(.ant-btn) { justify-self: start; }
</style>
