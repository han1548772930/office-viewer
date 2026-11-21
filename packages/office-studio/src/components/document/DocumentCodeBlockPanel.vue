<script setup lang="ts">
import { ref, watch } from 'vue'
import { Input, Modal, Select, Switch } from 'antdv-next'
import type { DocumentCodeBlockData, DocumentCodeLanguage } from '../../features/document/code-blocks/code-blocks.ts'

const props = defineProps<{ open: boolean; initial?: DocumentCodeBlockData }>()
const emit = defineEmits<{ 'update:open': [value: boolean]; apply: [value: DocumentCodeBlockData] }>()

const source = ref('')
const language = ref<DocumentCodeLanguage>('typescript')
const lineNumbers = ref(false)
const languages: Array<{ label: string; value: DocumentCodeLanguage }> = [
  { label: 'Plain text', value: 'plain' }, { label: 'JavaScript', value: 'javascript' },
  { label: 'TypeScript', value: 'typescript' }, { label: 'JSON', value: 'json' },
  { label: 'CSS', value: 'css' }, { label: 'HTML', value: 'html' },
  { label: 'Python', value: 'python' }, { label: 'SQL', value: 'sql' }, { label: 'Rust', value: 'rust' },
]

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
  <Modal :open="open" title="Code block" ok-text="Apply" cancel-text="Cancel" :width="640" destroy-on-hidden @ok="apply" @cancel="emit('update:open', false)">
    <div class="code-block-panel">
      <label>Language<Select v-model:value="language" :options="languages" aria-label="Code language" show-search option-filter-prop="label" /></label>
      <label class="code-block-panel__toggle"><span>Line numbers</span><Switch v-model:checked="lineNumbers" aria-label="Show line numbers" /></label>
      <label>Source<Input.TextArea v-model:value="source" aria-label="Code source" :auto-size="{ minRows: 10, maxRows: 24 }" spellcheck="false" /></label>
    </div>
  </Modal>
</template>

<style scoped>
.code-block-panel { display: flex; flex-direction: column; gap: 14px; }
.code-block-panel label { display: grid; gap: 6px; color: #465365; font-size: 12px; }
.code-block-panel__toggle { grid-template-columns: 1fr auto; align-items: center; }
.code-block-panel :deep(textarea) { font: 13px/1.5 Consolas, "Cascadia Code", monospace; }
</style>
