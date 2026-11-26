<script setup lang="ts">
import { Alert, Button, Form, FormItem, Modal, Select, theme } from 'antdv-next'
import { computed, ref, watch } from 'vue'
import type { OfficeStudio } from '../../api/office-studio.ts'
import { getSheetOutlineUpdate } from './sheet-outline.ts'
import { sheetOutlineErrorMessage } from './sheet-outline-types.ts'
import { useSheetOutlineDialog, useSheetOutlineDialogFocus } from './use-sheet-outline-dialog.ts'

const props = defineProps<{ open: boolean; studio: OfficeStudio; workbookId: string; sheetId: string }>()
const emit = defineEmits<{ 'update:open': [open: boolean]; changed: [] }>()
const { token } = theme.useToken()
const appearance = computed(() => ({ '--outline-dialog-text': token.value.colorText, '--outline-dialog-muted': token.value.colorTextSecondary }))
const controls = useSheetOutlineDialog({
  get studio() { return props.studio }, get workbookId() { return props.workbookId }, get sheetId() { return props.sheetId }, get open() { return props.open },
  changed: () => emit('changed'), close: (owner, committed) => focus.close(owner, committed) })
const focus = useSheetOutlineDialogFocus({
  get studio() { return props.studio }, get workbookId() { return props.workbookId }, get sheetId() { return props.sheetId }, get open() { return props.open },
  owner: () => controls.owner.value, emitClose: () => emit('update:open', false) })
const draft = ref({ summaryAbove: false, summaryLeft: false })
watch(controls.owner, () => {
  draft.value = { summaryAbove: controls.sheet.value?.outline?.summaryAbove ?? false,
    summaryLeft: controls.sheet.value?.outline?.summaryLeft ?? false }
}, { immediate: true, flush: 'sync' })
const action = computed(() => ({ kind: 'summary-settings' as const, ...draft.value }))
const preview = computed(() => {
  const sheet = controls.sheet.value, owner = controls.owner.value
  if (!sheet || !owner || controls.stale.value) return { changed: false, error: '' }
  try { return { changed: Boolean(getSheetOutlineUpdate(sheet, action.value, owner.userId)), error: '' } }
  catch (failure) { return { changed: false, error: sheetOutlineErrorMessage(failure) } }
})
const disabled = computed(() => controls.busy.value || controls.stale.value || !controls.owner.value)
const rowOptions = [{ value: 'below', label: '明细下方' }, { value: 'above', label: '明细上方' }]
const columnOptions = [{ value: 'right', label: '明细右侧' }, { value: 'left', label: '明细左侧' }]
function setRow(value: unknown): void { if (!disabled.value && (value === 'above' || value === 'below')) draft.value.summaryAbove = value === 'above' }
function setColumn(value: unknown): void { if (!disabled.value && (value === 'left' || value === 'right')) draft.value.summaryLeft = value === 'left' }
function cancel(): void { controls.close() }
function apply(): void { if (!disabled.value && preview.value.changed) controls.apply(action.value) }
</script>

<template>
  <Modal :open="open" title="大纲汇总设置" :root-style="appearance" :width="420" :footer="null" :mask="{ closable: false }"
    :focus-trigger-after-close="false" destroy-on-hidden @cancel="cancel">
    <Form class="outline-dialog" :style="appearance" :model="draft" layout="vertical" @finish="apply">
      <p class="outline-dialog-note">汇总位置决定分组折叠按钮所在的行列边界；明细数据和已有分组保留。</p>
      <Alert v-if="controls.stale.value" title="大纲设置已过期，请重新读取后再保存。" type="warning" show-icon />
      <FormItem label="行汇总位置" name="summaryAbove">
        <Select :value="draft.summaryAbove ? 'above' : 'below'" :options="rowOptions" aria-label="大纲行汇总位置" :disabled="disabled" @change="setRow" />
      </FormItem>
      <FormItem label="列汇总位置" name="summaryLeft">
        <Select :value="draft.summaryLeft ? 'left' : 'right'" :options="columnOptions" aria-label="大纲列汇总位置" :disabled="disabled" @change="setColumn" />
      </FormItem>
      <Alert v-if="preview.error" :title="preview.error" type="warning" show-icon />
      <Alert v-if="controls.error.value" :title="controls.error.value" type="error" show-icon />
      <div class="outline-dialog-actions">
        <Button html-type="button" :disabled="controls.busy.value" @click="controls.read">重新读取</Button><span />
        <Button html-type="button" @click="cancel">取消</Button>
        <Button type="primary" html-type="submit" :disabled="disabled || !preview.changed" :loading="controls.busy.value">确定</Button>
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
