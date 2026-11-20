<script setup lang="ts">
import { InputNumber, Modal } from 'antdv-next'
import { ref, watch } from 'vue'
import type { DocumentPageSetup } from '../../domain/model'
import { documentPageSettingsError } from './document-page-settings'

const props = defineProps<{ open: boolean; page: DocumentPageSetup }>()
const emit = defineEmits<{ 'update:open': [open: boolean]; apply: [page: DocumentPageSetup] }>()
const draft = ref({ ...props.page })
const error = ref('')
watch(() => props.open, (open) => { if (open) { draft.value = { ...props.page }; error.value = '' } })
function apply() {
  const message = documentPageSettingsError(draft.value)
  if (message) { error.value = message; return }
  emit('apply', { ...draft.value })
  emit('update:open', false)
}
</script>

<template>
  <Modal :open="open" title="页面设置" ok-text="应用" cancel-text="取消" :width="460" destroy-on-hidden @ok="apply" @cancel="emit('update:open', false)">
    <div class="page-settings-form">
      <p>尺寸和页边距的单位为像素（96 DPI）。</p>
      <fieldset><legend>纸张尺寸</legend><label>宽度<InputNumber v-model:value="draft.width" aria-label="纸张宽度" :min="200" :precision="1" /></label><label>高度<InputNumber v-model:value="draft.height" aria-label="纸张高度" :min="240" :precision="1" /></label></fieldset>
      <fieldset><legend>页边距</legend><label>上<InputNumber v-model:value="draft.marginTop" aria-label="上边距" :min="0" :precision="1" /></label><label>下<InputNumber v-model:value="draft.marginBottom" aria-label="下边距" :min="0" :precision="1" /></label><label>左<InputNumber v-model:value="draft.marginLeft" aria-label="左边距" :min="0" :precision="1" /></label><label>右<InputNumber v-model:value="draft.marginRight" aria-label="右边距" :min="0" :precision="1" /></label></fieldset>
      <p v-if="error" class="page-settings-error" role="alert">{{ error }}</p>
    </div>
  </Modal>
</template>

<style scoped>
.page-settings-form { padding-top: 8px; color: #4f5b69; }
.page-settings-form p { margin: 0 0 16px; color: #7b8592; font-size: 12px; }
fieldset { display: grid; grid-template-columns: 1fr 1fr; gap: 14px 18px; min-width: 0; margin: 0 0 20px; padding: 12px 0 0; border: 0; border-top: 1px solid #e5e8ed; }
legend { padding-right: 12px; font-size: 13px; }
label { display: flex; align-items: center; justify-content: space-between; gap: 12px; }
label :deep(.ant-input-number) { width: 114px; }
.page-settings-form .page-settings-error { margin-bottom: 0; color: #c14444; }
</style>
