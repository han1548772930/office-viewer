<script setup lang="ts">
import { Alert, Button, Checkbox, Form, FormItem, Input, Pagination, theme } from 'antdv-next'
import { ArrowDownAZ, ArrowUpAZ, RotateCcw, Search, Settings2, Trash2 } from 'lucide-vue-next'
import { computed, nextTick, onScopeDispose, ref, watch } from 'vue'
import { pivotChartFilterMemberLabel, pivotChartFilterVisibleMembers,
  type PivotChartFieldButton, type PivotChartFieldButtonAction, type PivotChartFieldFilterDraft } from '../services/pivot-chart-field-button-state.ts'
import { pivotChartDateConditionFilter, pivotChartLabelConditionFilter, type PivotChartDateConditionDraft,
  type PivotChartFieldConditionsDraft, type PivotChartLabelConditionDraft } from '../services/pivot-chart-field-filter-condition.ts'
import PivotChartFieldConditionEditor from './PivotChartFieldConditionEditor.vue'

const props = defineProps<{ field: PivotChartFieldButton; draft?: PivotChartFieldFilterDraft; readable: boolean;
  conditions?: PivotChartFieldConditionsDraft; loading?: boolean; busy?: boolean; disabled?: boolean; navigationDisabled?: boolean;
  composing?: boolean; error?: string; stale?: boolean; subjectLabel?: string; testId?: string; showSort?: boolean;
  showClearSort?: boolean; showManagement?: boolean }>()
const emit = defineEmits<{
  search: [value: string]
  select: [key: string, checked: boolean]
  'select-all': [checked: boolean]
  'date-change': [patch: Partial<PivotChartDateConditionDraft>]
  'label-change': [patch: Partial<PivotChartLabelConditionDraft>]
  action: [action: PivotChartFieldButtonAction]
  settings: []
  retry: []
  close: []
  'composition-start': []
  'composition-end': []
  keydown: [event: KeyboardEvent]
}>()
const { token } = theme.useToken()
const root = ref<HTMLElement>(), page = ref(1), pageSize = 80
const mode = ref<'members' | 'date' | 'label'>('members')
const subject = computed(() => props.subjectLabel ?? '透视图')
let disposed = false, focusGeneration = 0
const appearance = computed(() => ({ color: token.value.colorText, background: token.value.colorBgElevated,
  borderColor: token.value.colorBorderSecondary, '--member-border': token.value.colorBorderSecondary }))
const manual = computed(() => props.field.area !== 'value')
const members = computed(() => props.readable && props.draft ? pivotChartFilterVisibleMembers(props.draft) : [])
const pageCount = computed(() => Math.max(1, Math.ceil(members.value.length / pageSize)))
const currentPage = computed(() => Math.min(page.value, pageCount.value))
const pageMembers = computed(() => members.value.slice((currentPage.value - 1) * pageSize, currentPage.value * pageSize))
const selected = computed(() => new Set(props.readable ? props.draft?.selectedKeys ?? [] : []))
const partial = computed(() => new Set(props.readable ? props.draft?.partialKeys ?? [] : []))
const all = computed(() => Boolean(props.draft?.items.length && !props.draft.partialKeys.length && props.draft.selectedKeys.length === props.draft.items.length))
const browseDisabled = computed(() => Boolean(props.navigationDisabled || props.busy || props.loading || props.composing || !props.readable))
const navigationDisabled = computed(() => Boolean(props.disabled || browseDisabled.value))
const inputDisabled = computed(() => Boolean(props.busy || props.loading || !props.composing && (props.navigationDisabled || !props.readable)))
const conditionInputDisabled = computed(() => inputDisabled.value || !props.composing && Boolean(props.disabled))
const applyAction = computed<PivotChartFieldButtonAction>(() => mode.value === 'date' ? 'applyDateFilter' : mode.value === 'label' ? 'applyLabelCondition' : 'applyFilter')
const applyLabel = computed(() => `应用${subject.value}${mode.value === 'date' ? '日期筛选' : mode.value === 'label' ? '标签条件' : '字段筛选'}`)
const clearAction = computed<PivotChartFieldButtonAction>(() => mode.value === 'date' ? 'clearDateFilter' : mode.value === 'label' ? 'clearLabelCondition' : 'clearFilter')
const clearLabel = computed(() => `清除${subject.value}${mode.value === 'date' ? '日期条件' : mode.value === 'label' ? '标签条件' : '字段筛选'}`)
const hasSavedFilter = computed(() => mode.value === 'date' ? Boolean(props.field.field.dateFilter) : mode.value === 'label' ? Boolean(props.field.field.labelFilter) : props.field.filtered)
const validation = computed(() => {
  if (mode.value === 'members') return { valid: selected.value.size + partial.value.size > 0, error: '' }
  if (!props.conditions || !props.readable) return { valid: false, error: '' }
  try {
    if (mode.value === 'date') pivotChartDateConditionFilter(props.conditions)
    else pivotChartLabelConditionFilter(props.conditions)
    return { valid: true, error: '' }
  } catch (reason) { return { valid: false, error: reason instanceof Error ? reason.message : String(reason) } }
})
const popupContainer = (trigger: HTMLElement): HTMLElement => root.value ?? trigger.parentElement ?? trigger
function changePage(value: number): void { if (!browseDisabled.value) page.value = value }
function setMode(value: 'members' | 'date' | 'label'): void { if (!browseDisabled.value) { mode.value = value; focusFirst() } }
function select(key: string, value: unknown): void { if (typeof value === 'boolean') emit('select', key, value) }
function selectAll(value: unknown): void { if (typeof value === 'boolean') emit('select-all', value) }
function focusFirst(): void {
  const request = ++focusGeneration, id = props.field.fieldId
  void nextTick(() => {
    if (disposed || request !== focusGeneration || id !== props.field.fieldId || props.loading || props.busy || props.composing || !props.readable) return
    const target = root.value?.querySelector<HTMLInputElement>('input[type="text"]:not([disabled]), input:not([type]):not([disabled])')
      ?? root.value?.querySelector<HTMLButtonElement>('button:not([disabled])')
    target?.focus({ preventScroll: true })
  })
}
watch(() => props.draft?.search, () => { page.value = 1 })
watch(() => props.field.fieldId, () => { mode.value = 'members'; page.value = 1 })
watch([() => props.field.fieldId, () => props.loading, () => props.readable], focusFirst, { immediate: true })
onScopeDispose(() => { disposed = true; focusGeneration++ })
</script>

<template>
  <div ref="root" class="pivot-chart-field-filter" :style="appearance" :data-testid="testId ?? 'pivot-chart-field-filter-popup'"
    :data-pivot-chart-field="subject === '透视图' ? field.fieldId : undefined" :data-pivot-table-field="subject === '透视表' ? field.fieldId : undefined"
    role="dialog" :aria-label="`${field.label}的${subject}字段操作`"
    @pointerdown.stop @click.stop @dblclick.stop @contextmenu.stop @keyup.stop @keydown.stop="emit('keydown', $event)"
    @compositionstart.capture="emit('composition-start')" @compositionend.capture="emit('composition-end')">
    <Form layout="vertical" @submit.prevent="emit('action', applyAction)">
      <div class="field-heading"><span>{{ field.label }}</span><Button size="small" :disabled="busy || composing" @click="emit('close')">关闭</Button></div>
      <Alert v-if="stale || !readable" type="warning" title="字段内容、用户或历史已更改，请重新打开字段操作。" show-icon />
      <Alert v-if="loading" type="info" :title="`正在读取${subject}真实筛选成员…`" show-icon role="status" />
      <template v-if="manual">
        <div class="field-conditions">
          <Button v-if="conditions?.capabilities.supportsDateFilter" size="small" :type="mode === 'date' ? 'primary' : 'default'" :aria-pressed="mode === 'date'"
            :disabled="browseDisabled" :aria-label="`打开${subject}日期筛选`" @click="setMode('date')">日期筛选</Button>
          <Button size="small" :type="mode === 'label' ? 'primary' : 'default'" :aria-pressed="mode === 'label'" :disabled="browseDisabled"
            :aria-label="`打开${subject}标签条件`" @click="setMode('label')">标签筛选</Button>
          <Button v-if="mode !== 'members'" size="small" :disabled="browseDisabled" @click="setMode('members')">成员筛选</Button>
        </div>
        <PivotChartFieldConditionEditor v-if="mode !== 'members' && conditions" :key="field.fieldId" :kind="mode" :draft="conditions"
          :disabled="conditionInputDisabled" :actions-disabled="navigationDisabled" :popup-container="popupContainer" :subject-label="subject"
          @date-change="emit('date-change', $event)" @label-change="emit('label-change', $event)" />
        <template v-if="mode === 'members'">
        <div v-if="showSort !== false" class="field-sorts">
          <Button size="small" :type="field.field.sort === 'ascending' ? 'primary' : 'default'" :disabled="navigationDisabled"
            :aria-pressed="field.field.sort === 'ascending'" :aria-label="`${subject}字段升序`" @click="emit('action', 'sortAscending')"><ArrowDownAZ :size="14" />升序</Button>
          <Button size="small" :type="field.field.sort === 'descending' ? 'primary' : 'default'" :disabled="navigationDisabled"
            :aria-pressed="field.field.sort === 'descending'" :aria-label="`${subject}字段降序`" @click="emit('action', 'sortDescending')"><ArrowUpAZ :size="14" />降序</Button>
          <Button v-if="showClearSort !== false" size="small" :disabled="navigationDisabled || !field.field.sort" :aria-label="`清除${subject}字段排序`" @click="emit('action', 'clearSort')"><RotateCcw :size="14" />清除排序</Button>
        </div>
        <FormItem label="搜索标签"><Input :value="draft?.search ?? ''" :aria-label="`搜索${subject}筛选标签`" :disabled="inputDisabled" allow-clear
          @update:value="emit('search', $event)"><template #prefix><Search :size="14" aria-hidden="true" /></template></Input></FormItem>
        <FormItem :label="null"><Checkbox :checked="all" :indeterminate="!all && selected.size + partial.size > 0" :disabled="navigationDisabled || !draft?.items.length"
          @update:checked="selectAll">全选所有项目</Checkbox></FormItem>
        <div class="field-members" :aria-busy="loading">
          <FormItem v-for="key in pageMembers" :key="key" :label="null" class="member-row"><Checkbox :checked="selected.has(key)" :indeterminate="partial.has(key)" :disabled="navigationDisabled"
            :aria-label="`选择${subject}标签${pivotChartFilterMemberLabel(key)}`" @update:checked="select(key, $event)">{{ pivotChartFilterMemberLabel(key) }}</Checkbox></FormItem>
          <span v-if="readable && !loading && draft && !members.length" class="member-empty">{{ draft.items.length ? '没有匹配标签' : '当前来源没有可选标签' }}</span>
        </div>
        <Pagination v-if="members.length > pageSize" size="small" :current="currentPage" :page-size="pageSize" :total="members.length"
          :show-size-changer="false" :disabled="browseDisabled" @change="changePage" />
        <span v-if="readable && draft" class="member-count" role="status">已选择 {{ draft.selectedKeys.length }} / {{ draft.items.length }} 项<span v-if="draft.partialKeys.length">，部分选中 {{ draft.partialKeys.length }} 项</span></span>
        </template>
        <Alert v-if="validation.error && !loading" type="warning" :title="validation.error" show-icon />
      </template>
      <Alert v-if="error" type="error" :title="error" show-icon><template #action><Button v-if="manual && readable" size="small" :disabled="busy || composing || loading" @click="emit('retry')">重新读取成员</Button></template></Alert>
      <div v-if="showManagement !== false" class="field-management">
        <Button size="small" :disabled="browseDisabled" :aria-label="`${subject}字段设置`" @click="emit('settings')"><Settings2 :size="14" />字段设置</Button>
        <Button size="small" danger :disabled="navigationDisabled" :aria-label="`移除${subject}字段`" @click="emit('action', 'remove')"><Trash2 :size="14" />移除字段</Button>
      </div>
      <footer>
        <Button v-if="manual" size="small" type="link" :disabled="navigationDisabled || !hasSavedFilter" :aria-label="clearLabel" @click="emit('action', clearAction)">{{ mode === 'members' ? '清除筛选' : '清除条件' }}</Button>
        <span />
        <Button size="small" :disabled="busy || composing" @click="emit('close')">取消</Button>
        <Button v-if="manual" size="small" type="primary" :loading="busy" :disabled="navigationDisabled || !validation.valid"
          :aria-label="applyLabel" @click="emit('action', applyAction)">确认</Button>
      </footer>
    </Form>
  </div>
</template>

<style scoped>
.pivot-chart-field-filter { position: relative; box-sizing: border-box; width: 336px; max-width: calc(100vw - 16px); max-height: calc(100vh - 24px); padding: 12px; border: 1px solid; border-radius: 6px; }
.pivot-chart-field-filter :deep(.ant-form) { display: grid; gap: 10px; max-height: calc(100vh - 50px); overflow: auto; }
.pivot-chart-field-filter :deep(.ant-form-item) { margin-bottom: 0; }
.pivot-chart-field-filter :deep(.ant-btn) { display: inline-flex; align-items: center; justify-content: center; gap: 4px; }
.field-heading { display: flex; align-items: center; justify-content: space-between; gap: 8px; font-size: 13px; font-weight: 600; }
.field-heading > span { overflow: hidden; text-overflow: ellipsis; }
.field-sorts, .field-management, .field-conditions { display: flex; flex-wrap: wrap; gap: 6px; }
.field-members { max-height: 240px; overflow-y: auto; padding: 4px 6px; border: 1px solid var(--member-border); border-radius: 4px; }
.member-row { padding: 3px 0; }
.member-row :deep(.ant-checkbox-wrapper) { max-width: 100%; overflow-wrap: anywhere; }
.member-empty, .member-count { font-size: 12px; opacity: .75; }
footer { display: flex; align-items: center; gap: 6px; }
footer > span { flex: 1; }
</style>
