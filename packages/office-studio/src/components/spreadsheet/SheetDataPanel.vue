<script setup lang="ts">
import { Button, Checkbox, Form, FormItem, Input, Modal, Radio, RadioGroup, Segmented, Select, Switch, theme } from 'antdv-next'
import { computed, onBeforeUnmount, ref, useId, watch } from 'vue'
import { FOffice } from '../../api/facade'
import type { OfficeStudio } from '../../api/office-studio'
import { bindWorkbookCommandGuard } from '../../api/workbook-command-guards.ts'
import type { CellRange, SheetFilterConditionOperator, SheetFilterCriteria } from '../../domain/model'
import {
  CONFIGURE_SHEET_FILTER_COMMAND, dataPanelAddress, dataPanelValueKey, parseDataPanelRange, parseDataPanelScalar,
  parseDataPanelSourceRange, prepareDataPanelSort, prepareDataPanelSortRange, readDataPanelFields, readDataPanelValues,
  readDataPanelColors, type DataPanelColor, type DataPanelField, type DataPanelValue, type PanelSortKey, type PanelSortRangeChoice, type PanelSortRangeSelection,
} from '../../features/sheet-sort-filter/sheet-data-panel-state'
import { prepareSmartSheetFilterRange } from '../../features/sheet-sort-filter/sheet-data-range.ts'
import { canonicalSheetFilterColor } from '../../features/sheet-sort-filter/sheet-filter-colors.ts'

const props = withDefaults(defineProps<{
  open: boolean; studio: OfficeStudio; workbookId: string; sheetId: string; sourceRange: string; initialMode?: 'filter' | 'sort'
  sourcePrimary?: { row: number; column: number }; initialColumn?: number; editable?: boolean
  requestSortRange?: (input: PanelSortRangeSelection) => boolean; assertCurrent?: (revision?: number) => void
}>(), { initialMode: 'filter', editable: true })
const emit = defineEmits<{ 'update:open': [open: boolean]; changed: [] }>()
const formId = useId()
const { token } = theme.useToken()
const mode = ref<'filter' | 'sort'>('filter')
const address = ref(''), error = ref(''), status = ref(''), busy = ref(false)
const range = ref<CellRange>(), fields = ref<DataPanelField[]>([]), values = ref<DataPanelValue[]>([])
const revision = ref(0), columns = ref<Record<number, SheetFilterCriteria>>({}), column = ref(0)
const kind = ref<'values' | 'condition' | 'color'>('values'), selectedKeys = ref<string[]>([]), search = ref(''), page = ref(0)
const fillColors = ref<DataPanelColor[]>([]), textColors = ref<DataPanelColor[]>([])
const selectedFillColors = ref<Array<string | null>>([]), selectedTextColors = ref<string[]>([])
const operator = ref<SheetFilterConditionOperator>('equals'), valueType = ref<'text' | 'number' | 'boolean'>('text')
const expected = ref(''), second = ref(''), caseSensitive = ref(false)
const hasHeader = ref(true), locale = ref('zh-CN'), keys = ref<PanelSortKey[]>([]), hiddenCount = ref(0)
const modelReadVersion = ref(0)
const choosingRange = ref(false), sortRangeReady = ref(false), sortRangeChoice = ref<PanelSortRangeChoice>('keep')
const sortSource = ref<CellRange>(), sortPrimary = ref<{ row: number; column: number }>(), composing = ref(false)
const formDisabled = computed(() => !props.editable || busy.value)
const controlsDisabled = computed(() => formDisabled.value || composing.value)
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
  assertPanelCurrent(revision.value)
  if (!sheet.value) throw new Error('目标工作表已关闭。')
  const parsed = parseDataPanelRange(sheet.value, address.value)
  if (!range.value || dataPanelAddress(parsed) !== dataPanelAddress(range.value)) throw new Error('数据区域已改变，请先读取字段。')
  return parsed
}
function assertPanelCurrent(expectedRevision?: number): void {
  if (!props.open || !props.editable || composing.value) throw new Error('当前编辑器无法修改排序或筛选，请结束输入后重新打开。')
  props.assertCurrent?.(expectedRevision)
}
function close(): void {
  if (composing.value) return
  generation++; controller?.abort(); controller = undefined; busy.value = false
  emit('update:open', false)
}
function beginSortRange(): void {
  try {
    assertPanelCurrent()
    if (!sheet.value) throw new Error('目标工作表已关闭。')
    sortSource.value = parseDataPanelSourceRange(sheet.value, props.sourceRange)
    sortPrimary.value = props.sourcePrimary ? { ...props.sourcePrimary } : { row: sortSource.value.startRow, column: sortSource.value.startCol }
    revision.value = props.studio.getWorkbookRevision(props.workbookId)
    sortRangeChoice.value = 'keep'; choosingRange.value = true; error.value = ''; status.value = ''
  } catch (failure) { error.value = failure instanceof Error ? failure.message : String(failure) }
}
function confirmSortRange(): void {
  if (controlsDisabled.value || !sheet.value || !sortSource.value || !sortPrimary.value) return
  try {
    assertPanelCurrent(revision.value)
    if (props.studio.getWorkbookRevision(props.workbookId) !== revision.value) throw new Error('数据已更改，请关闭后重新选择排序区域。')
    const selected = prepareDataPanelSortRange(sheet.value, sortSource.value, sortRangeChoice.value, sortPrimary.value)
    if (props.requestSortRange && !props.requestSortRange({ unitId: props.workbookId, sheetId: props.sheetId,
      expectedRevision: revision.value, sourceRange: { ...sortSource.value }, range: selected, primary: { ...sortPrimary.value } })) {
      throw new Error('排序选区已改变，请关闭后重新选择。')
    }
    address.value = dataPanelAddress(selected); sortRangeReady.value = true; choosingRange.value = false
    keys.value = [{ column: sortPrimary.value.column, direction: 'ascending' }]
    loadFields()
  } catch (failure) { error.value = failure instanceof Error ? failure.message : String(failure) }
}
function loadColumn(): void {
  if (!sheet.value || !range.value) return
  const workbook = props.studio.getWorkbookModel(props.workbookId)
  if (!workbook) return
  values.value = readDataPanelValues(sheet.value, range.value, column.value, workbook)
  fillColors.value = readDataPanelColors(workbook, sheet.value, range.value, column.value, 'fill')
  textColors.value = readDataPanelColors(workbook, sheet.value, range.value, column.value, 'text')
  const criterion = columns.value[column.value]
  kind.value = criterion?.kind ?? 'values'
  selectedKeys.value = criterion?.kind === 'values' ? criterion.selectedItems.map(dataPanelValueKey) : values.value.map(item => item.key)
  operator.value = criterion?.kind === 'condition' ? criterion.operator : 'equals'
  const raw = criterion?.kind === 'condition' ? criterion.value : ''
  valueType.value = typeof raw === 'number' ? 'number' : typeof raw === 'boolean' ? 'boolean' : 'text'
  expected.value = raw == null ? '' : String(raw)
  second.value = criterion?.kind === 'condition' && criterion.secondValue != null ? String(criterion.secondValue) : ''
  caseSensitive.value = criterion?.kind === 'condition' && criterion.caseSensitive === true
  selectedFillColors.value = criterion?.kind === 'color' ? (criterion.cellFillColors ?? []).map(color => color === null ? null : canonicalSheetFilterColor(color)) : []
  selectedTextColors.value = criterion?.kind === 'color' && criterion.cellFillColors === undefined ? (criterion.cellTextColors ?? []).map(canonicalSheetFilterColor) : []
  search.value = ''; page.value = 0
}
function loadFields(): void {
  if (composing.value) return
  try {
    assertPanelCurrent()
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
watch(mode, value => { if (value === 'sort' && props.open && !sortRangeReady.value) beginSortRange() })
watch([() => props.open, () => props.studio, () => props.workbookId, () => props.sheetId, () => props.initialMode, () => props.initialColumn], () => {
  generation++; controller?.abort(); busy.value = false
  if (!props.open) return
  composing.value = false; sortRangeReady.value = false; choosingRange.value = false
  mode.value = props.initialMode; hasHeader.value = true; error.value = ''; status.value = ''; keys.value = []
  if (props.initialColumn !== undefined) column.value = props.initialColumn
  if (mode.value === 'sort') { beginSortRange(); return }
  try {
    assertPanelCurrent(); modelReadVersion.value++
    if (!sheet.value) throw new Error('目标工作表已关闭。')
    address.value = sheet.value.autoFilter ? dataPanelAddress(sheet.value.autoFilter.range)
      : dataPanelAddress(prepareSmartSheetFilterRange(sheet.value, parseDataPanelSourceRange(sheet.value, props.sourceRange)))
    loadFields()
  } catch (failure) { error.value = failure instanceof Error ? failure.message : String(failure) }
}, { immediate: true, flush: 'sync' })
watch(() => props.editable, editable => { if (!editable) { generation++; controller?.abort(); controller = undefined; busy.value = false } }, { flush: 'sync' })
onBeforeUnmount(() => { generation++; controller?.abort() })

function configure(columnsInput: Record<number, SheetFilterCriteria>, remove = false): boolean {
  const selectedRange = requireRange()
  const expectedRevision = revision.value
  const studio = props.studio, unitId = props.workbookId, sheetId = props.sheetId, userId = studio.currentUserId
  const model = studio.getWorkbookModel(unitId), historyGeneration = studio.editor.history.captureSavepoint(unitId).generation
  const assertCurrent = () => {
    assertPanelCurrent(expectedRevision)
    if (props.studio !== studio || props.workbookId !== unitId || props.sheetId !== sheetId || studio.currentUserId !== userId
      || studio.getWorkbookModel(unitId) !== model || studio.editor.history.captureSavepoint(unitId).generation !== historyGeneration) {
      throw new Error('筛选的工作簿、用户或内容版本已更改，请重新执行。')
    }
  }
  assertCurrent()
  const accepted = props.studio.executeCommandSync(CONFIGURE_SHEET_FILTER_COMMAND, bindWorkbookCommandGuard({ unitId: props.workbookId,
    sheetId: props.sheetId, range: remove ? null : selectedRange, columns: columnsInput, expectedRevision }, assertCurrent))
  if (!accepted) return false
  loadFields(); emit('changed')
  return true
}
function runFilter(action: 'apply' | 'clear-column' | 'clear-all' | 'remove'): void {
  if (controlsDisabled.value) return
  try {
    let next = { ...columns.value }
    if (action === 'apply') {
      let criterion: SheetFilterCriteria | undefined
      if (kind.value === 'values') criterion = { kind: 'values', selectedItems: matched.value.filter(item => selectedKeys.value.includes(item.key)).map(item => item.value) }
      else if (kind.value === 'condition') criterion = { kind: 'condition', operator: operator.value, ...(needsValue.value ? { value: parseDataPanelScalar(expected.value, valueType.value) } : {}), ...(needsSecond.value ? { secondValue: parseDataPanelScalar(second.value, valueType.value) } : {}), caseSensitive: caseSensitive.value }
      else if (selectedFillColors.value.length) criterion = { kind: 'color', cellFillColors: [...selectedFillColors.value] }
      else if (selectedTextColors.value.length) criterion = { kind: 'color', cellTextColors: [...selectedTextColors.value] }
      if (criterion) next[column.value] = criterion
      else delete next[column.value]
    } else if (action === 'clear-column') delete next[column.value]
    else next = {}
    if (!configure(next, action === 'remove')) return
    status.value = action === 'remove' ? '已移除筛选区域。' : `筛选已应用，隐藏 ${hiddenCount.value} 行；手动隐藏的行保持不变。`
  } catch (failure) { error.value = failure instanceof Error ? failure.message : String(failure) }
}
function toggle(key: string): void {
  if (!controlsDisabled.value) selectedKeys.value = selectedKeys.value.includes(key) ? selectedKeys.value.filter(value => value !== key) : [...selectedKeys.value, key]
}
function toggleColor(item: DataPanelColor, target: 'fill' | 'text'): void {
  if (controlsDisabled.value) return
  if (target === 'fill') {
    selectedFillColors.value = selectedFillColors.value.includes(item.value) ? selectedFillColors.value.filter(value => value !== item.value) : [...selectedFillColors.value, item.value]
    selectedTextColors.value = []
  } else if (item.value !== null) {
    selectedTextColors.value = selectedTextColors.value.includes(item.value) ? selectedTextColors.value.filter(value => value !== item.value) : [...selectedTextColors.value, item.value]
    selectedFillColors.value = []
  }
}
function addSortKey(): void {
  if (controlsDisabled.value) return
  const next = fields.value.find(field => !keys.value.some(key => key.column === field.column))
  if (next) keys.value.push({ column: next.column, direction: 'ascending' })
}
async function applySort(): Promise<void> {
  if (controlsDisabled.value || choosingRange.value || !sortRangeReady.value) return
  const version = generation
  const task = new AbortController(); controller = task; busy.value = true; error.value = ''
  try {
    const selectedRange = requireRange()
    if (revision.value !== props.studio.getWorkbookRevision(props.workbookId)) throw new Error('数据已更改，请重新读取数据区域后再排序。')
    const criteria = prepareDataPanelSort(selectedRange, keys.value)
    const worksheet = FOffice.newAPI(props.studio).getWorkbook(props.workbookId)?.getSheetBySheetId(props.sheetId)
    if (!worksheet) throw new Error('目标工作表已关闭。')
    const expectedRevision = revision.value
    await worksheet.sortRange(address.value, criteria, { hasHeader: hasHeader.value, locale: locale.value, signal: task.signal,
      assertCurrent: () => {
        task.signal.throwIfAborted()
        if (version !== generation) throw new Error('排序请求已被替换，旧结果没有应用。')
        assertPanelCurrent(expectedRevision)
      }, onProgress: progress => { if (version === generation && props.editable) status.value = `正在排序：${progress.processed.toLocaleString()} / ${progress.total.toLocaleString()}` } })
    if (version !== generation) return
    loadFields(); status.value = '已完成排序，可通过文档撤销恢复。'; emit('changed')
  } catch (failure) { if (version === generation) {
    if (task.signal.aborted) status.value = '已取消排序。'
    else error.value = failure instanceof Error ? failure.message : String(failure)
  } }
  finally { if (version === generation) { busy.value = false; controller = undefined } }
}
</script>

<template>
  <Modal :open="open" :title="choosingRange ? '排序区域' : '排序与筛选'" :footer="null" :width="choosingRange ? 400 : 700" :keyboard="!composing" @cancel="close">
    <Form layout="vertical" class="sheet-data-panel" :style="{ '--data-border': token.colorBorderSecondary, '--data-error': token.colorError }"
      :disabled="formDisabled" :aria-busy="busy" @submit.prevent @compositionstart="composing = true" @compositionend="composing = false">
      <p v-if="error" class="data-error" role="alert">{{ error }}</p><p v-if="status" role="status">{{ status }}</p>
      <template v-if="choosingRange">
        <p>请选择保留当前选区，或扩展到相邻的连续数据区域。</p>
        <FormItem label="排序区域选择" name="sortRangeChoice">
          <RadioGroup v-model:value="sortRangeChoice" aria-label="排序区域选择" :disabled="controlsDisabled">
            <Radio value="keep">保持当前选区</Radio><Radio value="extend">扩展选定区域</Radio>
          </RadioGroup>
        </FormItem>
        <div class="data-actions"><Button :disabled="composing" @click="close">取消</Button><Button type="primary" :disabled="controlsDisabled" @click="confirmSortRange">确认区域</Button></div>
      </template>
      <template v-else>
      <Segmented v-model:value="mode" :options="[{ label: '筛选设置', value: 'filter' }, { label: '排序设置', value: 'sort' }]" :disabled="controlsDisabled" aria-label="筛选或排序设置" />
      <FormItem label="数据区域"><Input v-model:value="address" aria-label="排序筛选数据区域" :disabled="formDisabled" placeholder="A1:C20" @press-enter="loadFields" /></FormItem>
      <Button :disabled="controlsDisabled" @click="loadFields">读取字段</Button>
      <template v-if="fields.length && mode === 'filter'">
        <FormItem label="筛选字段"><Select v-model:value="column" aria-label="筛选字段" :options="fieldOptions" :disabled="controlsDisabled" /></FormItem>
        <FormItem label="筛选方式"><Select v-model:value="kind" aria-label="筛选方式" :disabled="controlsDisabled" :options="[{ label: '按值选择', value: 'values' }, { label: '按条件筛选', value: 'condition' }, { label: '按颜色筛选', value: 'color' }]" /></FormItem>
        <template v-if="kind === 'values'">
          <FormItem label="搜索筛选值"><Input v-model:value="search" aria-label="搜索筛选值" placeholder="搜索字段值" /></FormItem>
          <div class="data-actions"><Button size="small" :disabled="controlsDisabled" @click="selectedKeys = matched.map(item => item.key)">全选</Button><Button size="small" :disabled="controlsDisabled" @click="selectedKeys = []">全不选</Button><span>{{ matched.filter(item => selectedKeys.includes(item.key)).length }} / {{ matched.length }} 个值</span></div>
          <div class="data-values"><Checkbox v-for="item in visibleValues" :key="item.key" :checked="selectedKeys.includes(item.key)" :disabled="controlsDisabled" @change="toggle(item.key)">{{ item.label }}</Checkbox></div>
          <div class="data-actions"><Button size="small" :disabled="controlsDisabled || page === 0" @click="page--">上一页</Button><span>{{ page + 1 }} / {{ pageCount }}</span><Button size="small" :disabled="controlsDisabled || page + 1 >= pageCount" @click="page++">下一页</Button></div>
        </template>
        <template v-else-if="kind === 'condition'">
          <FormItem label="条件"><Select v-model:value="operator" aria-label="筛选条件" :disabled="controlsDisabled" :options="operators" /></FormItem>
          <FormItem v-if="needsValue" label="比较值类型"><Select v-model:value="valueType" aria-label="比较值类型" :disabled="controlsDisabled" :options="[{ label: '文本', value: 'text' }, { label: '数值或日期序号', value: 'number' }, { label: '逻辑值', value: 'boolean' }]" /></FormItem>
          <FormItem v-if="needsValue" label="比较值"><Input v-model:value="expected" aria-label="筛选比较值" /></FormItem><FormItem v-if="needsSecond" label="第二个比较值"><Input v-model:value="second" aria-label="筛选第二个比较值" /></FormItem>
          <FormItem :label="null"><Checkbox v-model:checked="caseSensitive" :disabled="controlsDisabled">区分大小写</Checkbox></FormItem>
        </template>
        <template v-else>
          <FormItem label="填充颜色">
            <div class="data-values"><Checkbox v-for="item in fillColors" :key="item.key" :aria-label="`填充颜色 ${item.label}`" :checked="selectedFillColors.includes(item.value)" :disabled="controlsDisabled" @change="toggleColor(item, 'fill')">
              <span class="data-color-swatch" :style="{ background: item.value ?? token.colorBgContainer }" aria-hidden="true" />{{ item.label }}
            </Checkbox><span v-if="!fillColors.length">没有可选颜色</span></div>
          </FormItem>
          <FormItem label="文字颜色">
            <div class="data-values"><Checkbox v-for="item in textColors" :key="item.key" :aria-label="`文字颜色 ${item.label}`" :checked="item.value !== null && selectedTextColors.includes(item.value)" :disabled="controlsDisabled" @change="toggleColor(item, 'text')">
              <span class="data-color-swatch" :style="{ background: item.value ?? token.colorText }" aria-hidden="true" />{{ item.label }}
            </Checkbox><span v-if="!textColors.length">没有可选颜色</span></div>
          </FormItem>
        </template>
        <div class="data-actions"><Button type="primary" :disabled="controlsDisabled" @click="runFilter('apply')">应用筛选</Button><Button :disabled="controlsDisabled" @click="runFilter('clear-column')">清除此列条件</Button><Button :disabled="controlsDisabled" @click="runFilter('clear-all')">清除全部条件</Button><Button danger :disabled="controlsDisabled" @click="runFilter('remove')">移除筛选区域</Button></div>
      </template>
      <template v-if="fields.length && mode === 'sort'">
        <FormItem label="数据包含标题行"><Switch v-model:checked="hasHeader" aria-label="排序包含标题行" :disabled="controlsDisabled" /></FormItem>
        <FormItem label="比较语言"><Select v-model:value="locale" aria-label="排序比较语言" :disabled="controlsDisabled" :options="[{ label: '中文', value: 'zh-CN' }, { label: '英语', value: 'en-US' }]" /></FormItem>
        <div v-for="(key, index) in keys" :key="index" class="data-sort-key">
          <FormItem class="data-sort-item" :label="`排序字段 ${index + 1}`"><Select :id="`${formId}-field-${index}`" v-model:value="key.column" :aria-label="`排序字段 ${index + 1}`" :disabled="formDisabled" :options="fieldOptions" /></FormItem>
          <FormItem class="data-sort-item" :label="`排序方向 ${index + 1}`"><Select :id="`${formId}-direction-${index}`" v-model:value="key.direction" :aria-label="`排序方向 ${index + 1}`" :disabled="formDisabled" :options="[{ label: '升序', value: 'ascending' }, { label: '降序', value: 'descending' }]" /></FormItem>
          <Button :aria-label="`移除排序条件 ${index + 1}`" :disabled="controlsDisabled || keys.length === 1" @click="keys.splice(index, 1)">移除</Button>
        </div>
        <div class="data-actions"><Button :disabled="controlsDisabled || keys.length >= fields.length" @click="addSortKey">添加排序条件</Button><Button type="primary" :disabled="controlsDisabled" :loading="busy" @click="applySort">应用排序</Button><Button v-if="busy" :disabled="composing" @click="controller?.abort()">取消排序</Button></div>
      </template>
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
.data-values { display: grid; gap: 6px; max-height: 220px; padding: 10px; overflow: auto; border: 1px solid var(--data-border); border-radius: 4px; }
.data-color-swatch { display: inline-block; box-sizing: border-box; width: 14px; height: 14px; margin-inline-end: 6px; vertical-align: -2px; border: 1px solid var(--data-border); }
.data-sort-key { display: grid; grid-template-columns: minmax(0, 1fr) 110px auto; gap: 8px; }
.data-sort-item :deep(.ant-form-item-label) { position: absolute; width: 1px; height: 1px; overflow: hidden; white-space: nowrap; clip-path: inset(50%); }
.data-error { margin: 0; color: var(--data-error); }
.sheet-data-panel p { margin: 0; font-size: 12px; line-height: 1.6; }
@media (max-width: 600px) { .sheet-data-panel > :deep(.ant-form-item) { grid-template-columns: 1fr; } .data-sort-key { grid-template-columns: 1fr 100px; } }
</style>
