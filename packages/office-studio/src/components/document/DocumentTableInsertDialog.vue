<script setup lang="ts">
import { InputNumber, Modal } from 'antdv-next'
import { computed, inject, ref, watch } from 'vue'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

const props = defineProps<{ open: boolean }>()
const emit = defineEmits<{
  'update:open': [open: boolean]
  insert: [size: { rows: number; columns: number }]
}>()

const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const rows = ref<number | null>(3)
const columns = ref<number | null>(5)
const canInsert = computed(() => Number.isSafeInteger(rows.value) && rows.value! >= 1 && rows.value! <= 20
  && Number.isSafeInteger(columns.value) && columns.value! >= 1 && columns.value! <= 20)

watch(() => props.open, open => {
  if (open) {
    rows.value = 3
    columns.value = 5
  }
})

function insert() {
  if (!canInsert.value) return
  emit('insert', { rows: rows.value!, columns: columns.value! })
  emit('update:open', false)
}
</script>

<template>
  <Modal
    :open="open"
    :title="t('doc.table-insert.title', 'Insert Table')"
    :ok-text="t('common.insert', 'Insert')"
    :cancel-text="t('common.cancel', 'Cancel')"
    :ok-button-props="{ disabled: !canInsert }"
    :width="440"
    destroy-on-hidden
    @ok="insert"
    @cancel="emit('update:open', false)"
  >
    <div class="table-insert-form">
      <label>
        <span>{{ t('doc.table-insert.rows', 'Rows') }}</span>
        <InputNumber v-model:value="rows" :min="1" :max="20" :precision="0" />
      </label>
      <label>
        <span>{{ t('doc.table-insert.columns', 'Columns') }}</span>
        <InputNumber v-model:value="columns" :min="1" :max="20" :precision="0" />
      </label>
    </div>
  </Modal>
</template>

<style scoped>
.table-insert-form { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 16px; padding-top: 8px; }
.table-insert-form label { display: grid; gap: 7px; min-width: 0; }
.table-insert-form :deep(.ant-input-number) { width: 100%; }
@media (max-width: 460px) { .table-insert-form { grid-template-columns: minmax(0, 1fr); } }
</style>
