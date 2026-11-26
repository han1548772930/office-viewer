<script setup lang="ts">
import { Alert, Button, Form, FormItem, Modal, Select, theme } from 'antdv-next'
import { computed, ref, watch } from 'vue'
import type { OfficeStudio } from '../../api/office-studio.ts'
import type { OutlineAxis } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import type { SheetSelectionRange } from '../sheet-workbench/sheet-selection.ts'
import { getSheetOutlinePromptChoices } from './sheet-outline-command.ts'
import { prepareSheetOutlinePanelChange } from './sheet-outline-panel-owner.ts'
import { sheetOutlineErrorMessage } from './sheet-outline-types.ts'
import { useSheetOutlineDialog, useSheetOutlineDialogFocus } from './use-sheet-outline-dialog.ts'

const props = defineProps<{ open: boolean; studio: OfficeStudio; workbookId: string; sheetId: string;
  ranges: readonly SheetSelectionRange[]; operation: 'group' | 'remove' }>()
const emit = defineEmits<{ 'update:open': [open: boolean]; changed: [] }>()
const { token } = theme.useToken()
const appearance = computed(() => ({ '--outline-dialog-text': token.value.colorText, '--outline-dialog-muted': token.value.colorTextSecondary }))
const draftKey = computed(() => `${props.operation}:${props.ranges.map(range =>
  `${range.startRow},${range.endRow},${range.startCol},${range.endCol},${range.rangeType ?? 0}`).join(';')}`)
const controls = useSheetOutlineDialog({
  get studio() { return props.studio }, get workbookId() { return props.workbookId }, get sheetId() { return props.sheetId },
  get open() { return props.open }, get draftKey() { return draftKey.value },
  changed: () => emit('changed'), close: (owner, committed) => focus.close(owner, committed) })
const focus = useSheetOutlineDialogFocus({
  get studio() { return props.studio }, get workbookId() { return props.workbookId }, get sheetId() { return props.sheetId },
  get open() { return props.open }, get draftKey() { return draftKey.value },
  owner: () => controls.owner.value, emitClose: () => emit('update:open', false) })
const axis = ref<OutlineAxis>('row'), groupId = ref(''), selectedRanges = ref<SheetSelectionRange[]>([])
const title = computed(() => props.operation === 'group' ? '组合行列' : '取消行列分组')
watch(controls.owner, () => {
  selectedRanges.value = cloneModel([...props.ranges]); groupId.value = crypto.randomUUID(); axis.value = 'row'
}, { immediate: true, flush: 'sync' })
const choices = computed(() => {
  const owner = controls.owner.value, sheet = controls.sheet.value
  if (!owner || !sheet) return { values: [], error: '' }
  try { return { values: getSheetOutlinePromptChoices({ unitId: props.workbookId, sheetId: props.sheetId,
    sheet, userId: owner.userId, selections: selectedRanges.value }, props.operation), error: '' } }
  catch (failure) { return { values: [], error: sheetOutlineErrorMessage(failure) } }
})
watch(choices, current => {
  if (!current.values.some(choice => choice.axis === axis.value && !choice.disabled)) {
    const first = current.values.find(choice => !choice.disabled)
    if (first) axis.value = first.axis
  }
}, { immediate: true })
const selected = computed(() => choices.value.values.find(choice => choice.axis === axis.value))
const axisOptions = computed(() => choices.value.values.map(choice => ({ value: choice.axis, label: choice.label, disabled: choice.disabled })))
const action = computed(() => selected.value ? { ...selected.value.action,
  ...(props.operation === 'group' ? { groupId: groupId.value } : {}) } : undefined)
const preview = computed(() => {
  if (!controls.owner.value || !action.value || selected.value?.disabled || controls.stale.value) return { valid: false, error: '' }
  try { return { valid: Boolean(prepareSheetOutlinePanelChange(controls.owner.value, action.value)), error: '' } }
  catch (failure) { return { valid: false, error: sheetOutlineErrorMessage(failure) } }
})
const disabled = computed(() => controls.busy.value || controls.stale.value || !preview.value.valid)
function choose(value: unknown): void { if (!controls.busy.value && !controls.stale.value && (value === 'row' || value === 'column')) axis.value = value }
function cancel(): void { controls.close() }
function apply(): void { if (action.value && !disabled.value) controls.apply(action.value) }
</script>

<template>
  <Modal :open="open" :title="title" :root-style="appearance" :width="420" :footer="null" :mask="{ closable: false }"
    :focus-trigger-after-close="false" destroy-on-hidden @cancel="cancel">
    <Form class="outline-dialog" :style="appearance" :model="{ axis }" layout="vertical" @finish="apply">
      <p class="outline-dialog-note">选择组合方向。行列大纲按当前选区的起止位置建立，最多支持八层嵌套。</p>
      <Alert v-if="controls.stale.value" title="选区对应的内容或用户已更改，请重新读取后再操作。" type="warning" show-icon />
      <FormItem label="分组方向" name="axis">
        <Select :value="axis" :options="axisOptions" aria-label="大纲分组方向" :disabled="controls.busy.value || controls.stale.value"
          @change="choose" />
      </FormItem>
      <Alert v-if="choices.error || preview.error || selected?.reason" :title="choices.error || preview.error || selected?.reason" type="warning" show-icon />
      <Alert v-else-if="!choices.values.length && !controls.stale.value && controls.owner.value" title="请选择一个连续的单一区域。" type="info" show-icon />
      <Alert v-if="controls.error.value" :title="controls.error.value" type="error" show-icon />
      <div class="outline-dialog-actions">
        <Button html-type="button" :disabled="controls.busy.value" @click="controls.read">重新读取</Button><span />
        <Button html-type="button" @click="cancel">取消</Button>
        <Button type="primary" html-type="submit" :disabled="disabled" :loading="controls.busy.value">确定</Button>
      </div>
    </Form>
  </Modal>
</template>

<style scoped>
.outline-dialog { color: var(--outline-dialog-text); }
.outline-dialog-note { color: var(--outline-dialog-muted); font-size: 12px; line-height: 1.6; margin: 0 0 12px; }
.outline-dialog-actions { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
.outline-dialog-actions > span { flex: 1; }
</style>
