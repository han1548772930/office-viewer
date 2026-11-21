<script setup lang="ts">
import { Button, Input, InputNumber, Modal, Select } from 'antdv-next'
import { computed, ref, watch } from 'vue'
import { Trash2 } from 'lucide-vue-next'
import type { PivotDateGroupType, PivotFieldModel, PivotFilterOperator, PivotManualGroup, PivotTableModel, WorkbookModel } from '../../../domain/model'
import { defaultPivotNumberGroup, PIVOT_VALUE_FILTER_OPTIONS, pivotValueFilterInput, readPivotFieldMembers, type PivotFieldToolChange } from '../services/pivot-field-tools'

const props = defineProps<{ open: boolean; kind: 'filter' | 'group'; workbook: WorkbookModel; pivot: PivotTableModel; field: PivotFieldModel; busy: boolean; error?: string }>()
const emit = defineEmits<{ close: []; apply: [change: PivotFieldToolChange] }>()
const groupMode = ref<'number' | 'date' | 'manual'>('number')
const groupStart = ref<number | null>(0), groupEnd = ref<number | null>(1), groupInterval = ref<number | null>(1)
const dateGroup = ref<PivotDateGroupType>('month')
const manualGroups = ref<PivotManualGroup[]>([]), selectedMembers = ref<string[]>([]), groupName = ref('')
const operator = ref<PivotFilterOperator>('valueGreaterThan'), valueFieldId = ref('')
const expected = ref<number | null>(0), expected2 = ref<number | null>(0), localError = ref('')
const members = computed(() => readPivotFieldMembers(props.workbook, props.pivot, props.field))
const numericDefaults = computed(() => defaultPivotNumberGroup(members.value))
const valueFields = computed(() => props.pivot.fields.filter(field => field.area === 'value').map(field => ({ label: field.displayName ?? field.name, value: field.id })))
const memberOptions = computed(() => members.value.map(member => ({
  label: members.value.filter(candidate => candidate.label === member.label).length > 1 ? `${member.label}（${typeof member.value === 'number' ? '数值' : typeof member.value === 'boolean' ? '逻辑值' : member.value === null ? '空白' : '文本'}）` : member.label,
  value: member.key,
  disabled: manualGroups.value.some(group => group.members.some(value => Object.is(value, member.value))),
})))
const between = computed(() => operator.value === 'valueBetween' || operator.value === 'valueNotBetween')
const hasGrouping = computed(() => Boolean(props.field.numberGroup || props.field.dateGroupType || props.field.manualGroups?.length))

watch(() => [props.open, props.field.id, props.pivot.id, props.kind] as const, () => {
  if (!props.open) return
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
  if (props.field.numberGroup) emit('apply', { kind: 'numberGroup', settings: null })
  else if (props.field.dateGroupType) emit('apply', { kind: 'dateGroup', type: null })
  else emit('apply', { kind: 'manualGroup', groups: [] })
}
function confirm() {
  localError.value = ''
  try {
    if (props.kind === 'filter') {
      emit('apply', { kind: 'valueFilter', filter: pivotValueFilterInput(operator.value, valueFieldId.value, expected.value, expected2.value) })
      return
    }
    if (groupMode.value === 'date') emit('apply', { kind: 'dateGroup', type: dateGroup.value })
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
</script>

<template>
  <Modal :open="open" :title="`${kind === 'filter' ? '值筛选' : '字段分组'} · ${field.displayName ?? field.name}`" :width="440" :footer="null" :mask="false" :keyboard="!busy" destroy-on-hidden @cancel="!busy && emit('close')">
    <div class="pivot-field-tools" data-testid="pivot-field-tools">
      <template v-if="kind === 'filter'">
        <label>值字段<Select v-model:value="valueFieldId" :options="valueFields" :disabled="busy" aria-label="筛选值字段" /></label>
        <label>筛选条件<Select v-model:value="operator" :options="PIVOT_VALUE_FILTER_OPTIONS" :disabled="busy" aria-label="值筛选条件" /></label>
        <div class="pivot-tool-numbers"><label>{{ between ? '开始值' : '数值' }}<InputNumber v-model:value="expected" :disabled="busy" aria-label="筛选数值" /></label><label v-if="between">结束值<InputNumber v-model:value="expected2" :disabled="busy" aria-label="第二个筛选数值" /></label></div>
      </template>
      <template v-else>
        <label>分组方式<Select v-model:value="groupMode" :disabled="busy" :options="[{ label: '数值分组', value: 'number', disabled: !numericDefaults }, { label: '日期分组', value: 'date' }, { label: '组合项目', value: 'manual' }]" aria-label="分组方式" /></label>
        <div v-if="groupMode === 'number'" class="pivot-tool-numbers"><label>起始值<InputNumber v-model:value="groupStart" :disabled="busy" aria-label="分组起始值" /></label><label>结束值<InputNumber v-model:value="groupEnd" :disabled="busy" aria-label="分组结束值" /></label><label>间隔<InputNumber v-model:value="groupInterval" :disabled="busy" aria-label="分组间隔" /></label></div>
        <label v-else-if="groupMode === 'date'">分组依据<Select v-model:value="dateGroup" :disabled="busy" :options="[{ label: '年', value: 'year' }, { label: '年-季度', value: 'quarter' }, { label: '年-月', value: 'month' }, { label: '年-月-日', value: 'day' }]" aria-label="日期分组依据" /></label>
        <template v-else>
          <div v-for="(group, index) in manualGroups" :key="index" class="pivot-manual-group"><Input v-model:value="group.name" :disabled="busy" :aria-label="`分组 ${index + 1} 名称`" /><span>{{ group.members.map(value => value == null || value === '' ? '（空白）' : String(value)).join('、') }}</span><Button type="text" :disabled="busy" :aria-label="`移除组 ${group.name}`" @click="manualGroups.splice(index, 1)"><Trash2 :size="16" /></Button></div>
          <label>组名<Input v-model:value="groupName" :disabled="busy" aria-label="新分组名称" /></label>
          <label>源项目<Select v-model:value="selectedMembers" mode="multiple" :options="memberOptions" :disabled="busy" aria-label="分组源项目" /></label>
          <Button :disabled="busy || selectedMembers.length < 2" @click="addManualGroup">添加组</Button>
        </template>
      </template>
      <p v-if="localError || error" class="pivot-tool-error" role="alert">{{ localError || error }}</p>
      <div class="pivot-tool-actions"><Button v-if="kind === 'filter'" type="link" :disabled="busy || !field.valueFilter" @click="emit('apply', { kind: 'valueFilter', filter: null })">清除筛选</Button><Button v-else type="link" :disabled="busy || !hasGrouping" @click="clearGrouping">取消分组</Button><span /><Button :disabled="busy" @click="emit('close')">取消</Button><Button type="primary" :loading="busy" @click="confirm">确定</Button></div>
    </div>
  </Modal>
</template>

<style scoped>
.pivot-field-tools { display: flex; flex-direction: column; gap: 16px; }
.pivot-field-tools label { display: flex; flex-direction: column; gap: 6px; min-width: 0; color: #354152; font-size: 13px; }
.pivot-tool-numbers { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px; }
.pivot-tool-numbers :deep(.ant-input-number) { width: 100%; }
.pivot-manual-group { display: grid; grid-template-columns: minmax(0, 1fr) minmax(0, 1fr) auto; align-items: center; gap: 8px; font-size: 12px; }
.pivot-manual-group span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.pivot-tool-actions { display: flex; align-items: center; gap: 8px; padding-top: 12px; }
.pivot-tool-actions > span { flex: 1; }
.pivot-tool-error { margin: 0; color: #b12e35; font-size: 12px; }
</style>
