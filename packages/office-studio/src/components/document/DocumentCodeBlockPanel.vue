<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { useI18n } from 'vue-i18n'
import { Input, Modal, Select, Switch } from 'antdv-next'
import type { DocumentCodeBlockData, DocumentCodeLanguage } from '../../features/document/code-blocks/code-blocks.ts'

const props = defineProps<{ open: boolean; initial?: DocumentCodeBlockData }>()
const { t } = useI18n({ useScope: 'global' })
const emit = defineEmits<{ 'update:open': [value: boolean]; apply: [value: DocumentCodeBlockData] }>()

const source = ref('')
const language = ref<DocumentCodeLanguage>('typescript')
const lineNumbers = ref(false)
const languages = computed<Array<{ label: string; value: DocumentCodeLanguage }>>(() => [
  { label: t('document.code-block.plain'), value: 'plain' }, { label: 'JavaScript', value: 'javascript' },
  { label: 'TypeScript', value: 'typescript' }, { label: 'JSON', value: 'json' },
  { label: 'CSS', value: 'css' }, { label: 'HTML', value: 'html' },
  { label: 'Python', value: 'python' }, { label: 'SQL', value: 'sql' }, { label: 'Rust', value: 'rust' },
])

watch(() => [props.open, props.initial] as const, ([open, initial]) => {
  if (!open) return
  source.value = initial?.source ?? ''
  language.value = initial?.language ?? 'typescript'
  lineNumbers.value = initial?.lineNumbers ?? false
}, { immediate: true })

function apply() {
  emit('apply', { source: source.value, language: language.value, lineNumbers: lineNumbers.value })
  emit('update:open', false)
}
</script>

<template>
  <Modal :open="open" :title="t('document.code-block.title')" :ok-text="t('common.apply')" :cancel-text="t('common.cancel')" :width="640" destroy-on-hidden @ok="apply" @cancel="emit('update:open', false)">
    <div class="code-block-panel">
      <label>{{ t('document.code-block.language') }}<Select v-model:value="language" :options="languages" :aria-label="t('document.code-block.language-label')" show-search option-filter-prop="label" /></label>
      <label class="code-block-panel__toggle"><span>{{ t('document.code-block.line-numbers') }}</span><Switch v-model:checked="lineNumbers" :aria-label="t('document.code-block.show-line-numbers')" /></label>
      <label>{{ t('document.code-block.source') }}<Input.TextArea v-model:value="source" :aria-label="t('document.code-block.source')" :auto-size="{ minRows: 10, maxRows: 24 }" spellcheck="false" /></label>
    </div>
  </Modal>
</template>

<style scoped>
.code-block-panel { display: flex; flex-direction: column; gap: 14px; }
.code-block-panel label { display: grid; gap: 6px; color: #465365; font-size: 12px; }
.code-block-panel__toggle { grid-template-columns: 1fr auto; align-items: center; }
.code-block-panel :deep(textarea) { font: 13px/1.5 Consolas, "Cascadia Code", monospace; }
</style>
