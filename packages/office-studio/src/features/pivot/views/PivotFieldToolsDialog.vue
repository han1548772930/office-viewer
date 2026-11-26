<script setup lang="ts">
import { Alert, Button, Form, FormItem, Input, InputNumber, Modal, Select, theme } from 'antdv-next'
import { computed, nextTick, onScopeDispose, ref, watch } from 'vue'
import { Trash2 } from 'lucide-vue-next'
import type { PivotDateGroupType, PivotFieldModel, PivotFilterOperator, PivotManualGroup, PivotTableModel, WorkbookModel } from '../../../domain/model'
import { defaultPivotNumberGroup, PIVOT_VALUE_FILTER_OPTIONS, pivotValueFilterInput, readPivotFieldMembers, type PivotFieldToolChange } from '../services/pivot-field-tools'

const props = defineProps<{ open: boolean; kind: 'filter' | 'group'; workbook: WorkbookModel; pivot: PivotTableModel; field: PivotFieldModel; busy: boolean;
  error?: string; stale?: boolean; readable?: boolean; editable?: boolean; supportsDateGrouping?: boolean }>()
const emit = defineEmits<{ close: []; apply: [change: PivotFieldToolChange]; 'composition-start': []; 'composition-end': [] }>()
const composing = ref(false), { token } = theme.useToken()
let disposed = false, compositionGeneration = 0
const actionsDisabled = computed(() => props.busy || composing.value || props.stale || props.readable === false || props.editable === false)
const inputDisabled = computed(() => props.busy || !composing.value && Boolean(props.stale || props.readable === false || props.editable === false))
const appearance = computed(() => ({ color: token.value.colorText, backgroundColor: token.value.colorBgElevated, '--pivot-tool-error': token.value.colorError }))
const groupMode = ref<'number' | 'date' | 'manual'>('number')
const groupStart = ref<number | null>(0), groupEnd = ref<number | null>(1), groupInterval = ref<number | null>(1)
const dateGroup = ref<PivotDateGroupType>('month')
const manualGroups = ref<PivotManualGroup[]>([]), selectedMembers = ref<string[]>([]), groupName = ref('')
const operator = ref<PivotFilterOperator>('valueGreaterThan'), valueFieldId = ref('')
const expected = ref<number | null>(0), expected2 = ref<number | null>(0), localError = ref('')
const members = computed(() => props.readable === false ? [] : readPivotFieldMembers(props.workbook, props.pivot, props.field))
const numericDefaults = computed(() => defaultPivotNumberGroup(members.value))
const valueFields = computed(() => props.readable === false ? [] : props.pivot.fields.filter(field => field.area === 'value').map(field => ({ label: field.displayName ?? field.name, value: field.id })))
const memberOptions = computed(() => members.value.map(member => ({
  label: members.value.filter(candidate => candidate.label === member.label).length > 1 ? `${member.label}（${typeof member.value === 'number' ? '数值' : typeof member.value === 'boolean' ? '逻辑值' : member.value === null ? '空白' : '文本'}）` : member.label,
  value: member.key,
  disabled: manualGroups.value.some(group => group.members.some(value => Object.is(value, member.value))),
})))
const between = computed(() => operator.value === 'valueBetween' || operator.value === 'valueNotBetween')
const hasGrouping = computed(() => Boolean(props.field.numberGroup || props.field.dateGroupType || props.field.manualGroups?.length))

watch(() => [props.open, props.field.id, props.pivot.id, props.kind] as const, () => {
  if (!props.open || composing.value) return
  localError.value = ''
  const defaults = props.field.numberGroup ?? numericDefaults.value
  groupStart.value = defaults?.start ?? 0
  groupEnd.value = defaults?.end ?? 1
  groupInterval.value = defaults?.interval ?? 1
  groupMode.value = props.field.manualGroups?.length ? 'manual' : props.field.dateGroupType ? 'date' : defaults ? 'number' : 'manual'
  dateGroup.value = props.field.dateGroupType ?? 'month'
  manualGroups.value = (props.field.manualGroups ?? []).map(group => ({ name: group.name, members: [...group.members] }))
  groupName.value = `组 ${manualGroups.value.length + 1}`
  selectedMembers.value = []
  operator.value = props.field.valueFilter?.operator ?? 'valueGreaterThan'
  valueFieldId.value = props.field.valueFilter?.valueFieldId ?? valueFields.value[0]?.value ?? ''
  expected.value = props.field.valueFilter?.expected ?? 0
  expected2.value = props.field.valueFilter?.expected2 ?? 0
}, { immediate: true })

function addManualGroup() {
  if (actionsDisabled.value) return
  localError.value = ''
  const name = groupName.value.trim()
  const values = members.value.filter(member => selectedMembers.value.includes(member.key)).map(member => member.value)
  if (!name) { localError.value = '请输入组名。'; return }
  if (values.length < 2) { localError.value = '请选择至少两个不同的源项目。'; return }
  if (manualGroups.value.some(group => group.name === name)) { localError.value = '组名已存在，请使用不同的名称。'; return }
  manualGroups.value.push({ name, members: values })
  selectedMembers.value = []
  groupName.value = `组 ${manualGroups.value.length + 1}`
}
function clearGrouping() {
  if (actionsDisabled.value) return
  if (props.field.numberGroup) emit('apply', { kind: 'numberGroup', settings: null })
  else if (props.field.dateGroupType) emit('apply', { kind: 'dateGroup', type: null })
  else emit('apply', { kind: 'manualGroup', groups: [] })
}
function confirm() {
  if (actionsDisabled.value) return
  localError.value = ''
  try {
    if (props.kind === 'filter') {
      emit('apply', { kind: 'valueFilter', filter: pivotValueFilterInput(operator.value, valueFieldId.value, expected.value, expected2.value) })
      return
    }
    if (groupMode.value === 'date') {
      if (props.supportsDateGrouping === false) { localError.value = '当前来源不支持日期分组。'; return }
      emit('apply', { kind: 'dateGroup', type: dateGroup.value })
    }
    else if (groupMode.value === 'manual') {
      if (selectedMembers.value.length) { localError.value = '请先将已选项目添加到组。'; return }
      if (manualGroups.value.some(group => !group.name.trim())) { localError.value = '组名不能为空。'; return }
      emit('apply', { kind: 'manualGroup', groups: manualGroups.value })
    } else {
      const start = groupStart.value, end = groupEnd.value, interval = groupInterval.value
      if (start === null || end === null || interval === null || ![start, end, interval].every(Number.isFinite) || end <= start || interval <= 0) { localError.value = '请输入有效的起始值、结束值和正数间隔。'; return }
      emit('apply', { kind: 'numberGroup', settings: { start, end, interval } })
    }
  } catch (reason) { localError.value = reason instanceof Error ? reason.message : String(reason) }
}
function startComposition(): void { compositionGeneration++; composing.value = true; emit('composition-start') }
function endComposition(): void {
  const ended = compositionGeneration
  void nextTick(() => { if (disposed || ended !== compositionGeneration) return; composing.value = false; emit('composition-end') })
}
function close(): void { if (!props.busy && !composing.value) emit('close') }
onScopeDispose(() => { disposed = true; compositionGeneration++ })
</script>

<template>
  <Modal :open="open" :title="`${kind === 'filter' ? '值筛选' : '字段分组'} · ${field.displayName ?? field.name}`" :width="440" :footer="null" :mask="false"
    :root-style="appearance" :keyboard="!busy && !composing" :closable="!busy && !composing" destroy-on-hidden @cancel="close">
    <Form layout="vertical" class="pivot-field-tools" :style="appearance" data-testid="pivot-field-tools" @submit.prevent
      @compositionstart.capture="startComposition" @compositionend.capture="endComposition">
      <Alert v-if="stale || readable === false || editable === false" type="warning" title="来源、用户或历史已更改，或当前用户无法修改。请关闭并重新读取。" show-icon />
      <template v-if="readable !== false || composing">
      <template v-if="kind === 'filter'">
        <FormItem label="值字段"><Select v-model:value="valueFieldId" :options="valueFields" :disabled="actionsDisabled" aria-label="筛选值字段" /></FormItem>
        <FormItem label="筛选条件"><Select v-model:value="operator" :options="PIVOT_VALUE_FILTER_OPTIONS" :disabled="actionsDisabled" aria-label="值筛选条件" /></FormItem>
        <div class="pivot-tool-numbers"><FormItem :label="between ? '开始值' : '数值'"><InputNumber v-model:value="expected" :disabled="inputDisabled" aria-label="筛选数值" /></FormItem><FormItem v-if="between" label="结束值"><InputNumber v-model:value="expected2" :disabled="inputDisabled" aria-label="第二个筛选数值" /></FormItem></div>
      </template>
      <template v-else>
        <FormItem label="分组方式"><Select v-model:value="groupMode" :disabled="actionsDisabled" :options="[{ label: '数值分组', value: 'number', disabled: !numericDefaults }, { label: '日期分组', value: 'date', disabled: supportsDateGrouping === false }, { label: '组合项目', value: 'manual' }]" aria-label="分组方式" /></FormItem>
        <div v-if="groupMode === 'number'" class="pivot-tool-numbers"><FormItem label="起始值"><InputNumber v-model:value="groupStart" :disabled="inputDisabled" aria-label="分组起始值" /></FormItem><FormItem label="结束值"><InputNumber v-model:value="groupEnd" :disabled="inputDisabled" aria-label="分组结束值" /></FormItem><FormItem label="间隔"><InputNumber v-model:value="groupInterval" :disabled="inputDisabled" aria-label="分组间隔" /></FormItem></div>
        <FormItem v-else-if="groupMode === 'date'" label="分组依据"><Select v-model:value="dateGroup" :disabled="actionsDisabled" :options="[{ label: '年', value: 'year' }, { label: '年-季度', value: 'quarter' }, { label: '年-月', value: 'month' }, { label: '年-月-日', value: 'day' }]" aria-label="日期分组依据" /></FormItem>
        <template v-else>
          <div v-for="(group, index) in manualGroups" :key="index" class="pivot-manual-group"><FormItem><Input v-model:value="group.name" :disabled="inputDisabled" :aria-label="`分组 ${index + 1} 名称`" /></FormItem><span v-if="readable !== false">{{ group.members.map(value => value == null || value === '' ? '（空白）' : String(value)).join('、') }}</span><Button type="text" :disabled="actionsDisabled" :aria-label="`移除组 ${group.name}`" @click="manualGroups.splice(index, 1)"><Trash2 :size="16" /></Button></div>
          <FormItem label="组名"><Input v-model:value="groupName" :disabled="inputDisabled" aria-label="新分组名称" /></FormItem>
          <FormItem label="源项目"><Select v-model:value="selectedMembers" mode="multiple" :options="memberOptions" :disabled="actionsDisabled" aria-label="分组源项目" /></FormItem>
          <Button :disabled="actionsDisabled || selectedMembers.length < 2" @click="addManualGroup">添加组</Button>
        </template>
      </template>
      </template>
      <p v-if="localError || error" class="pivot-tool-error" role="alert">{{ localError || error }}</p>
      <div class="pivot-tool-actions"><Button v-if="kind === 'filter'" type="link" :disabled="actionsDisabled || !field.valueFilter" @click="emit('apply', { kind: 'valueFilter', filter: null })">清除筛选</Button><Button v-else type="link" :disabled="actionsDisabled || !hasGrouping" @click="clearGrouping">取消分组</Button><span /><Button :disabled="busy || composing" @click="close">取消</Button><Button type="primary" :loading="busy" :disabled="actionsDisabled" @click="confirm">确定</Button></div>
    </Form>
  </Modal>
</template>

<style scoped>
.pivot-field-tools { display: flex; flex-direction: column; gap: 16px; }
.pivot-field-tools :deep(.ant-form-item) { margin-bottom: 0; }
.pivot-field-tools :deep(.ant-form-item-label) { padding-bottom: 6px; }
.pivot-tool-numbers { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px; }
.pivot-tool-numbers :deep(.ant-input-number) { width: 100%; }
.pivot-manual-group { display: grid; grid-template-columns: minmax(0, 1fr) minmax(0, 1fr) auto; align-items: center; gap: 8px; font-size: 12px; }
.pivot-manual-group :deep(.ant-form-item) { min-width: 0; margin-bottom: 0; }
.pivot-manual-group span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.pivot-tool-actions { display: flex; align-items: center; gap: 8px; padding-top: 12px; }
.pivot-tool-actions > span { flex: 1; }
.pivot-tool-error { margin: 0; color: var(--pivot-tool-error); font-size: 12px; }
</style>
