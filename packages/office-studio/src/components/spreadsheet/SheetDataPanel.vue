<script setup lang="ts">
import { Button, Checkbox, Form, FormItem, Input, Modal, Segmented, Select, Switch } from 'antdv-next'
import { computed, onBeforeUnmount, ref, useId, watch } from 'vue'
import { FOffice } from '../../api/facade'
import type { OfficeStudio } from '../../api/office-studio'
import type { CellRange, SheetFilterConditionOperator, SheetFilterCriteria } from '../../domain/model'
import {
  CONFIGURE_SHEET_FILTER_COMMAND, dataPanelAddress, dataPanelValueKey, parseDataPanelRange, parseDataPanelScalar,
  prepareDataPanelSort, readDataPanelFields, readDataPanelValues,
  type DataPanelField, type DataPanelValue, type PanelSortKey,
} from '../../features/sheet-sort-filter/sheet-data-panel-state'

const props = withDefaults(defineProps<{
  open: boolean; studio: OfficeStudio; workbookId: string; sheetId: string; sourceRange: string; initialMode?: 'filter' | 'sort'
}>(), { initialMode: 'filter' })
const emit = defineEmits<{ 'update:open': [open: boolean]; changed: [] }>()
const formId = useId()
const mode = ref<'filter' | 'sort'>('filter')
const address = ref(''), error = ref(''), status = ref(''), busy = ref(false)
const range = ref<CellRange>(), fields = ref<DataPanelField[]>([]), values = ref<DataPanelValue[]>([])
const revision = ref(0), columns = ref<Record<number, SheetFilterCriteria>>({}), column = ref(0)
const kind = ref<'values' | 'condition'>('values'), selectedKeys = ref<string[]>([]), search = ref(''), page = ref(0)
const operator = ref<SheetFilterConditionOperator>('equals'), valueType = ref<'text' | 'number' | 'boolean'>('text')
const expected = ref(''), second = ref(''), caseSensitive = ref(false)
const hasHeader = ref(true), locale = ref('zh-CN'), keys = ref<PanelSortKey[]>([]), hiddenCount = ref(0)
const modelReadVersion = ref(0)
let generation = 0, controller: AbortController | undefined
const sheet = computed(() => {
  modelReadVersion.value
  return props.studio.getWorkbookModel(props.workbookId)?.sheets.find(item => item.id === props.sheetId)
})
const fieldOptions = computed(() => fields.value.map(field => ({ label: field.label, value: field.column })))
const matched = computed(() => values.value.filter(item => item.label.toLocaleLowerCase().includes(search.value.toLocaleLowerCase())))
const pageCount = computed(() => Math.max(1, Math.ceil(matched.value.length / 20)))
const visibleValues = computed(() => matched.value.slice(page.value * 20, (page.value + 1) * 20))
const needsValue = computed(() => !['isBlank', 'isNotBlank', 'isError'].includes(operator.value))
const needsSecond = computed(() => operator.value === 'between' || operator.value === 'notBetween')
const operators: Array<{ label: string; value: SheetFilterConditionOperator }> = [
  { label: '等于', value: 'equals' }, { label: '不等于', value: 'doesNotEqual' },
  { label: '大于', value: 'greaterThan' }, { label: '大于或等于', value: 'greaterThanOrEqual' },
  { label: '小于', value: 'lessThan' }, { label: '小于或等于', value: 'lessThanOrEqual' },
  { label: '介于', value: 'between' }, { label: '不介于', value: 'notBetween' },
  { label: '包含', value: 'contains' }, { label: '不包含', value: 'doesNotContain' },
  { label: '开头是', value: 'beginsWith' }, { label: '结尾是', value: 'endsWith' },
  { label: '为空', value: 'isBlank' }, { label: '非空', value: 'isNotBlank' }, { label: '为错误值', value: 'isError' },
]

function requireRange(): CellRange {
  if (!sheet.value) throw new Error('目标工作表已关闭。')
  const parsed = parseDataPanelRange(sheet.value, address.value)
  if (!range.value || dataPanelAddress(parsed) !== dataPanelAddress(range.value)) throw new Error('数据区域已改变，请先读取字段。')
  return parsed
}
function loadColumn(): void {
  if (!sheet.value || !range.value) return
  values.value = readDataPanelValues(sheet.value, range.value, column.value)
  const criterion = columns.value[column.value]
  kind.value = criterion?.kind ?? 'values'
  selectedKeys.value = criterion?.kind === 'values' ? criterion.selectedItems.map(dataPanelValueKey) : values.value.map(item => item.key)
  operator.value = criterion?.kind === 'condition' ? criterion.operator : 'equals'
  const raw = criterion?.kind === 'condition' ? criterion.value : ''
  valueType.value = typeof raw === 'number' ? 'number' : typeof raw === 'boolean' ? 'boolean' : 'text'
  expected.value = raw == null ? '' : String(raw)
  second.value = criterion?.kind === 'condition' && criterion.secondValue != null ? String(criterion.secondValue) : ''
  caseSensitive.value = criterion?.kind === 'condition' && criterion.caseSensitive === true
  search.value = ''; page.value = 0
}
function loadFields(): void {
  try {
    modelReadVersion.value++
    if (!sheet.value) throw new Error('目标工作表已关闭。')
    const parsed = parseDataPanelRange(sheet.value, address.value)
    range.value = parsed
    fields.value = readDataPanelFields(sheet.value, parsed, hasHeader.value)
    const filter = sheet.value.autoFilter
    columns.value = filter && dataPanelAddress(filter.range) === dataPanelAddress(parsed) ? JSON.parse(JSON.stringify(filter.columns)) : {}
    if (!fields.value.some(field => field.column === column.value)) column.value = fields.value[0]!.column
    keys.value = keys.value.filter(key => fields.value.some(field => field.column === key.column))
    if (!keys.value.length) keys.value = [{ column: fields.value[0]!.column, direction: 'ascending' }]
    revision.value = props.studio.getWorkbookRevision(props.workbookId)
    hiddenCount.value = filter?.filteredRows.length ?? 0
    error.value = ''; status.value = `已读取 ${fields.value.length} 个字段；筛选条件在列之间按“并且”组合。`
    loadColumn()
  } catch (failure) { fields.value = []; values.value = []; range.value = undefined; error.value = failure instanceof Error ? failure.message : String(failure) }
}
watch(column, loadColumn)
watch(search, () => { page.value = 0 })
watch(hasHeader, () => { if (sheet.value && range.value) fields.value = readDataPanelFields(sheet.value, range.value, hasHeader.value) })
watch([() => props.open, () => props.workbookId, () => props.sheetId, () => props.initialMode], () => {
  generation++; controller?.abort(); busy.value = false
  if (!props.open) return
  mode.value = props.initialMode; hasHeader.value = true; error.value = ''; status.value = ''; keys.value = []
  address.value = mode.value === 'filter' && sheet.value?.autoFilter ? dataPanelAddress(sheet.value.autoFilter.range) : props.sourceRange
  loadFields()
}, { immediate: true })
onBeforeUnmount(() => { generation++; controller?.abort() })

function configure(columnsInput: Record<number, SheetFilterCriteria>, remove = false): void {
  const selectedRange = requireRange()
  props.studio.executeCommandSync(CONFIGURE_SHEET_FILTER_COMMAND, { unitId: props.workbookId, sheetId: props.sheetId, range: remove ? null : selectedRange, columns: columnsInput, expectedRevision: revision.value })
  loadFields(); emit('changed')
}
function runFilter(action: 'apply' | 'clear-column' | 'clear-all' | 'remove'): void {
  try {
    let next = { ...columns.value }
    if (action === 'apply') {
      let criterion: SheetFilterCriteria
      if (kind.value === 'values') criterion = { kind: 'values', selectedItems: values.value.filter(item => selectedKeys.value.includes(item.key)).map(item => item.value) }
      else criterion = { kind: 'condition', operator: operator.value, ...(needsValue.value ? { value: parseDataPanelScalar(expected.value, valueType.value) } : {}), ...(needsSecond.value ? { secondValue: parseDataPanelScalar(second.value, valueType.value) } : {}), caseSensitive: caseSensitive.value }
      next[column.value] = criterion
    } else if (action === 'clear-column') delete next[column.value]
    else next = {}
    configure(next, action === 'remove')
    status.value = action === 'remove' ? '已移除筛选区域。' : `筛选已应用，隐藏 ${hiddenCount.value} 行；手动隐藏的行保持不变。`
  } catch (failure) { error.value = failure instanceof Error ? failure.message : String(failure) }
}
function toggle(key: string): void { selectedKeys.value = selectedKeys.value.includes(key) ? selectedKeys.value.filter(value => value !== key) : [...selectedKeys.value, key] }
function addSortKey(): void {
  const next = fields.value.find(field => !keys.value.some(key => key.column === field.column))
  if (next) keys.value.push({ column: next.column, direction: 'ascending' })
}
async function applySort(): Promise<void> {
  if (busy.value) return
  const version = generation
  busy.value = true; error.value = ''; controller = new AbortController()
  try {
    const selectedRange = requireRange()
    if (revision.value !== props.studio.getWorkbookRevision(props.workbookId)) throw new Error('数据已更改，请重新读取数据区域后再排序。')
    const criteria = prepareDataPanelSort(selectedRange, keys.value)
    const worksheet = FOffice.newAPI(props.studio).getWorkbook(props.workbookId)?.getSheetBySheetId(props.sheetId)
    if (!worksheet) throw new Error('目标工作表已关闭。')
    await worksheet.sortRange(address.value, criteria, { hasHeader: hasHeader.value, locale: locale.value, signal: controller.signal, onProgress: progress => { if (version === generation) status.value = `正在排序：${progress.processed.toLocaleString()} / ${progress.total.toLocaleString()}` } })
    if (version !== generation) return
    loadFields(); status.value = '已完成排序，可通过文档撤销恢复。'; emit('changed')
  } catch (failure) { if (version === generation) error.value = failure instanceof Error ? failure.message : String(failure) }
  finally { if (version === generation) busy.value = false }
}
</script>

<template>
  <Modal :open="open" title="排序与筛选" :footer="null" :width="700" @cancel="emit('update:open', false)">
    <Form layout="vertical" class="sheet-data-panel" :aria-busy="busy" @submit.prevent>
      <Segmented v-model:value="mode" :options="[{ label: '筛选设置', value: 'filter' }, { label: '排序设置', value: 'sort' }]" :disabled="busy" aria-label="筛选或排序设置" />
      <FormItem label="数据区域"><Input v-model:value="address" aria-label="排序筛选数据区域" :disabled="busy" placeholder="A1:C20" @press-enter="loadFields" /></FormItem>
      <Button :disabled="busy" @click="loadFields">读取字段</Button>
      <p v-if="error" class="data-error" role="alert">{{ error }}</p><p v-if="status" role="status">{{ status }}</p>
      <template v-if="fields.length && mode === 'filter'">
        <FormItem label="筛选字段"><Select v-model:value="column" aria-label="筛选字段" :options="fieldOptions" :disabled="busy" /></FormItem>
        <FormItem label="筛选方式"><Select v-model:value="kind" aria-label="筛选方式" :options="[{ label: '按值选择', value: 'values' }, { label: '按条件筛选', value: 'condition' }]" /></FormItem>
        <template v-if="kind === 'values'">
          <FormItem label="搜索筛选值"><Input v-model:value="search" aria-label="搜索筛选值" placeholder="搜索字段值" /></FormItem>
          <div class="data-actions"><Button size="small" @click="selectedKeys = values.map(item => item.key)">全选</Button><Button size="small" @click="selectedKeys = []">全不选</Button><span>{{ selectedKeys.length }} / {{ values.length }} 个值</span></div>
          <div class="data-values"><Checkbox v-for="item in visibleValues" :key="item.key" :checked="selectedKeys.includes(item.key)" @change="toggle(item.key)">{{ item.label }}</Checkbox></div>
          <div class="data-actions"><Button size="small" :disabled="page === 0" @click="page--">上一页</Button><span>{{ page + 1 }} / {{ pageCount }}</span><Button size="small" :disabled="page + 1 >= pageCount" @click="page++">下一页</Button></div>
        </template>
        <template v-else>
          <FormItem label="条件"><Select v-model:value="operator" aria-label="筛选条件" :options="operators" /></FormItem>
          <FormItem v-if="needsValue" label="比较值类型"><Select v-model:value="valueType" aria-label="比较值类型" :options="[{ label: '文本', value: 'text' }, { label: '数值或日期序号', value: 'number' }, { label: '逻辑值', value: 'boolean' }]" /></FormItem>
          <FormItem v-if="needsValue" label="比较值"><Input v-model:value="expected" aria-label="筛选比较值" /></FormItem><FormItem v-if="needsSecond" label="第二个比较值"><Input v-model:value="second" aria-label="筛选第二个比较值" /></FormItem>
          <FormItem :label="null"><Checkbox v-model:checked="caseSensitive">区分大小写</Checkbox></FormItem>
        </template>
        <div class="data-actions"><Button type="primary" @click="runFilter('apply')">应用筛选</Button><Button @click="runFilter('clear-column')">清除此列条件</Button><Button @click="runFilter('clear-all')">清除全部条件</Button><Button danger @click="runFilter('remove')">移除筛选区域</Button></div>
      </template>
      <template v-if="fields.length && mode === 'sort'">
        <FormItem label="数据包含标题行"><Switch v-model:checked="hasHeader" aria-label="排序包含标题行" :disabled="busy" /></FormItem>
        <FormItem label="比较语言"><Select v-model:value="locale" aria-label="排序比较语言" :disabled="busy" :options="[{ label: '中文', value: 'zh-CN' }, { label: '英语', value: 'en-US' }]" /></FormItem>
        <div v-for="(key, index) in keys" :key="index" class="data-sort-key">
          <FormItem class="data-sort-item" :label="`排序字段 ${index + 1}`"><Select :id="`${formId}-field-${index}`" v-model:value="key.column" :aria-label="`排序字段 ${index + 1}`" :disabled="busy" :options="fieldOptions" /></FormItem>
          <FormItem class="data-sort-item" :label="`排序方向 ${index + 1}`"><Select :id="`${formId}-direction-${index}`" v-model:value="key.direction" :aria-label="`排序方向 ${index + 1}`" :disabled="busy" :options="[{ label: '升序', value: 'ascending' }, { label: '降序', value: 'descending' }]" /></FormItem>
          <Button :aria-label="`移除排序条件 ${index + 1}`" :disabled="busy || keys.length === 1" @click="keys.splice(index, 1)">移除</Button>
        </div>
        <div class="data-actions"><Button :disabled="busy || keys.length >= fields.length" @click="addSortKey">添加排序条件</Button><Button type="primary" :loading="busy" @click="applySort">应用排序</Button><Button v-if="busy" @click="controller?.abort()">取消排序</Button></div>
      </template>
    </Form>
  </Modal>
</template>

<style scoped>
.sheet-data-panel { display: grid; gap: 12px; min-width: 0; padding-top: 8px; }
.sheet-data-panel :deep(.ant-form-item) { min-width: 0; margin-bottom: 0; }
.sheet-data-panel > :deep(.ant-form-item) { display: grid; grid-template-columns: 115px minmax(0, 1fr); align-items: center; column-gap: 10px; }
.sheet-data-panel > :deep(.ant-form-item .ant-form-item-row) { display: contents; }
.sheet-data-panel > :deep(.ant-form-item .ant-form-item-label) { padding: 0; text-align: start; }
.sheet-data-panel > :deep(.ant-form-item .ant-form-item-control) { min-width: 0; }
.data-actions { display: flex; flex-wrap: wrap; align-items: center; gap: 8px; }
.data-values { display: grid; gap: 6px; max-height: 220px; padding: 10px; overflow: auto; border: 1px solid #e5e7eb; border-radius: 4px; }
.data-sort-key { display: grid; grid-template-columns: minmax(0, 1fr) 110px auto; gap: 8px; }
.data-sort-item :deep(.ant-form-item-label) { position: absolute; width: 1px; height: 1px; overflow: hidden; white-space: nowrap; clip-path: inset(50%); }
.data-error { margin: 0; color: #b42318; }
.sheet-data-panel p { margin: 0; font-size: 12px; line-height: 1.6; }
@media (max-width: 600px) { .sheet-data-panel > :deep(.ant-form-item) { grid-template-columns: 1fr; } .data-sort-key { grid-template-columns: 1fr 100px; } }
</style>
