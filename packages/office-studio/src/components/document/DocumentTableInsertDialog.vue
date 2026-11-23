<script setup lang="ts">
import { Form, FormItem, Input, InputNumber, Modal, Switch } from 'antdv-next'
import { computed, inject, reactive, watch } from 'vue'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'
import type { DocumentTableInsertOptions } from '../../features/document/commands/table-command-types'
import { documentTableValuesFromTsv } from '../../features/document/services/document-table-data'

const props = defineProps<{ open: boolean }>()
const emit = defineEmits<{
  'update:open': [open: boolean]
  insert: [size: { rows: number; columns: number; options: DocumentTableInsertOptions }]
}>()

const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const form = reactive({ rows: 3, columns: 5, specifyWidth: false, width: 500, rowHeight: 30, headerRowCount: 0, titleRow: false, data: '' })
const values = computed(() => documentTableValuesFromTsv(form.data))
const canInsert = computed(() => Number.isSafeInteger(form.rows) && form.rows >= 1 && form.rows <= 10000 && Number.isSafeInteger(form.columns) && form.columns >= 1 && form.columns <= 512
  && form.rows * form.columns <= 250000 && Number.isSafeInteger(form.headerRowCount) && form.headerRowCount >= 0 && form.headerRowCount <= form.rows
  && Number.isFinite(form.rowHeight) && form.rowHeight > 0 && (!form.specifyWidth || Number.isFinite(form.width) && form.width > 0)
  && (!values.value || values.value.length <= form.rows && values.value.every(row => row.length <= form.columns)))

watch(() => props.open, open => {
  if (open) {
    Object.assign(form, { rows: 3, columns: 5, specifyWidth: false, width: 500, rowHeight: 30, headerRowCount: 0, titleRow: false, data: '' })
  }
})

function insert() {
  if (!canInsert.value) return
  emit('insert', { rows: form.rows, columns: form.columns, options: {
    ...(form.specifyWidth ? { width: form.width } : {}), ...(form.rowHeight !== 30 ? { rowHeights: Array.from({ length: form.rows }, () => form.rowHeight) } : {}),
    titleRow: form.titleRow, headerRowCount: Math.max(form.headerRowCount, form.titleRow ? 1 : 0), ...(values.value ? { values: values.value } : {}),
  } })
  emit('update:open', false)
}
</script>

<template>
  <Modal
    :open="open"
    :title="t('doc.table-insert.title', '插入表格')"
    :ok-text="t('common.insert', '插入')"
    :cancel-text="t('common.cancel', '取消')"
    :ok-button-props="{ disabled: !canInsert }"
    :width="440"
    destroy-on-hidden
    @ok="insert"
    @cancel="emit('update:open', false)"
  >
    <Form :model="form" class="table-insert-form" layout="vertical" @finish="insert">
      <FormItem name="rows" :label="t('doc.table-insert.rows', '行数')">
        <InputNumber v-model:value="form.rows" :min="1" :max="10000" :precision="0" />
      </FormItem>
      <FormItem name="columns" :label="t('doc.table-insert.columns', '列数')">
        <InputNumber v-model:value="form.columns" :min="1" :max="512" :precision="0" />
      </FormItem>
      <FormItem name="specifyWidth" label="指定表格宽度"><Switch v-model:checked="form.specifyWidth" /></FormItem>
      <FormItem name="width" label="表格宽度（px）"><InputNumber v-model:value="form.width" :min="1" :disabled="!form.specifyWidth" /></FormItem>
      <FormItem name="rowHeight" label="最小行高（px）"><InputNumber v-model:value="form.rowHeight" :min="1" /></FormItem>
      <FormItem name="headerRowCount" label="重复表头行数"><InputNumber v-model:value="form.headerRowCount" :min="0" :max="form.rows" :precision="0" /></FormItem>
      <FormItem name="titleRow" label="合并首行为标题"><Switch v-model:checked="form.titleRow" /></FormItem>
      <FormItem name="data" label="初始数据（制表符分列，换行分行）" class="table-insert-data"><Input.TextArea v-model:value="form.data" :rows="4" /></FormItem>
    </Form>
    <p v-if="!canInsert" role="status">请检查尺寸、初始数据和表头行数；单次最多插入 250000 个单元格。</p>
  </Modal>
</template>

<style scoped>
.table-insert-form { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 16px; padding-top: 8px; }
.table-insert-form :deep(.ant-form-item) { min-width: 0; margin-bottom: 0; }
.table-insert-form :deep(.ant-input-number) { width: 100%; }
.table-insert-data { grid-column: 1 / -1; }
@media (max-width: 460px) { .table-insert-form { grid-template-columns: minmax(0, 1fr); } }
</style>
