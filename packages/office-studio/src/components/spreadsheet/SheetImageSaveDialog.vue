<script setup lang="ts">
import { computed, reactive, ref, watch } from 'vue'
import { Alert, Checkbox, Form, FormItem, Modal, Select } from 'antdv-next'
import { sheetImageSaveNames, type SheetImageSaveConfig, type SheetImageSaveRequest } from '../../features/sheet-images/sheet-image-save.ts'
const props = defineProps<{ request: SheetImageSaveRequest | null; busy: boolean; error: string }>()
const emit = defineEmits<{ confirm: [config: SheetImageSaveConfig]; cancel: []; closed: [] }>()
const draft = reactive({ address: true, column: false, columnIndex: 0 }), draftError = ref('')
watch(() => props.request, request => { draft.address = true; draft.column = false; draft.columnIndex = request?.columns[0]?.index ?? 0; draftError.value = '' }, { flush: 'sync' })
const config = computed<SheetImageSaveConfig>(() => ({ fileNameParts: [...(draft.address ? ['cellAddress' as const] : []), ...(draft.column ? ['columnValue' as const] : [])], columnIndex: draft.column ? draft.columnIndex : undefined }))
const names = computed(() => { try { return props.request ? sheetImageSaveNames(props.request, config.value).slice(0, 5) : [] } catch { return [] } })
function confirm(): void { if (!config.value.fileNameParts.length) { draftError.value = '请至少选择一种文件名组成方式。'; return } emit('confirm', config.value) }
</script>
<template>
  <Modal :open="Boolean(request)" :title="`批量保存图片${request ? `（${request.target}）` : ''}`" :width="440" :confirm-loading="busy"
    :focus-trigger-after-close="false" :mask="{ closable: false }" destroy-on-hidden ok-text="保存" cancel-text="取消"
    @ok="confirm" @cancel="emit('cancel')" @after-close="emit('closed')">
    <Form layout="vertical" :model="draft" :preserve="false" :disabled="busy" @submit.prevent="confirm">
      <p>图片数量：{{ request?.entries.length ?? 0 }}</p>
      <FormItem label="文件名组成">
        <Checkbox v-model:checked="draft.address">单元格地址</Checkbox>
        <Checkbox v-model:checked="draft.column" :disabled="!request?.columns.length || busy">指定列的内容</Checkbox>
      </FormItem>
      <FormItem v-if="draft.column" label="文件名列"><Select v-model:value="draft.columnIndex" :options="request?.columns.map(column => ({ label: column.label, value: column.index }))" aria-label="图片文件名列" /></FormItem>
      <FormItem label="文件名预览"><div v-for="name in names" :key="name">{{ name }}</div></FormItem>
      <Alert v-if="draftError || error" type="error" :title="draftError || error" show-icon />
    </Form>
  </Modal>
</template>
