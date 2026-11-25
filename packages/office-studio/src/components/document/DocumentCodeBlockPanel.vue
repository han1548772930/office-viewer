<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { useI18n } from 'vue-i18n'
import { Form, FormItem, Input, InputNumber, Modal, Select, Switch, theme } from 'antdv-next'
import type { DocumentCodeBlockData } from '../../features/document/code-blocks/code-blocks.ts'
import { documentStructuredCodeLanguageOptions } from '../../features/document/code-blocks/document-code-languages'

const props = defineProps<{ open: boolean; initial?: DocumentCodeBlockData }>()
const { t } = useI18n({ useScope: 'global' })
const emit = defineEmits<{ 'update:open': [value: boolean]; apply: [value: DocumentCodeBlockData]; closed: [] }>()

const source = ref('')
const language = ref('typescript')
const lineNumbers = ref(false)
const tabSize = ref(4)
const wrap = ref(true)
const { token } = theme.useToken()
const languages = computed(() => documentStructuredCodeLanguageOptions(language.value))

watch(() => [props.open, props.initial] as const, ([open, initial]) => {
  if (!open) return
  source.value = initial?.source ?? ''
  language.value = initial?.language ?? 'typescript'
  lineNumbers.value = initial?.lineNumbers ?? false
  tabSize.value = initial?.tabSize ?? 4
  wrap.value = initial?.wrap ?? true
}, { immediate: true })

function apply() {
  emit('apply', { source: source.value, language: language.value, lineNumbers: lineNumbers.value, tabSize: tabSize.value, wrap: wrap.value })
}
</script>

<template>
  <Modal :open="open" :title="t('document.code-block.title')" :ok-text="t('common.apply')" :cancel-text="t('common.cancel')" :width="640" destroy-on-hidden @ok="apply" @cancel="emit('update:open', false)" @after-close="emit('closed')">
    <Form layout="vertical" class="code-block-panel" :style="{ '--code-panel-color': token.colorTextSecondary }">
      <FormItem :label="t('document.code-block.language')"><Select v-model:value="language" :options="languages" :aria-label="t('document.code-block.language-label')" show-search option-filter-prop="label" /></FormItem>
      <FormItem class="code-block-panel__toggle" :label="t('document.code-block.line-numbers')"><Switch v-model:checked="lineNumbers" :aria-label="t('document.code-block.show-line-numbers')" /></FormItem>
      <FormItem label="缩进空格数"><InputNumber v-model:value="tabSize" :min="1" :max="16" :precision="0" aria-label="代码块缩进空格数" /></FormItem>
      <FormItem class="code-block-panel__toggle" label="自动换行"><Switch v-model:checked="wrap" aria-label="代码块自动换行" /></FormItem>
      <FormItem :label="t('document.code-block.source')"><Input.TextArea v-model:value="source" :aria-label="t('document.code-block.source')" :auto-size="{ minRows: 10, maxRows: 24 }" spellcheck="false" /></FormItem>
    </Form>
  </Modal>
</template>

<style scoped>
.code-block-panel { display: flex; flex-direction: column; gap: 14px; }
.code-block-panel :deep(.ant-form-item) { margin-bottom: 0; color: var(--code-panel-color); font-size: 12px; }
.code-block-panel :deep(.code-block-panel__toggle .ant-form-item-row) { display: grid; grid-template-columns: 1fr auto; align-items: center; }
.code-block-panel :deep(.code-block-panel__toggle .ant-form-item-label) { padding-bottom: 0; }
.code-block-panel :deep(textarea) { font: 13px/1.5 Consolas, "Cascadia Code", monospace; }
</style>
