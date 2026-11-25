<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { Button, Form, FormItem, InputNumber, Select, Switch } from 'antdv-next'
import type { DocumentCodeConfig } from '../../domain/model'
import { documentStructuredCodeLanguageOptions } from '../../features/document/code-blocks/document-code-languages'
import { documentCodeFontFamilyOptions } from '../../features/document/code-blocks/document-code-fonts'

const props = defineProps<{ kind: 'font' | 'language' | 'settings'; config: DocumentCodeConfig; fontFamily: string; disabled?: boolean }>()
const emit = defineEmits<{ apply: [value: { patch?: Partial<DocumentCodeConfig>; fontFamily?: string }]; cancel: [] }>()
const draft = ref({ language: props.config.language, showLineNumbers: props.config.showLineNumbers, tabSize: props.config.tabSize,
  wrap: props.config.wrap, fontFamily: props.fontFamily })
watch(() => [props.kind, props.config, props.fontFamily] as const, () => {
  draft.value = { ...props.config, fontFamily: props.fontFamily }
}, { immediate: true })
const languages = computed(() => documentStructuredCodeLanguageOptions(draft.value.language))
const fonts = computed(() => documentCodeFontFamilyOptions(draft.value.fontFamily))
const available = computed(() => !props.disabled && (props.kind === 'font'
  ? fonts.value.some(font => !font.disabled && font.value === draft.value.fontFamily)
  : !!draft.value.language.trim() && Number.isInteger(draft.value.tabSize) && draft.value.tabSize >= 1 && draft.value.tabSize <= 16))
function apply() {
  if (!available.value) return
  if (props.kind === 'font') emit('apply', { fontFamily: draft.value.fontFamily })
  else if (props.kind === 'language') emit('apply', { patch: { language: draft.value.language } })
  else emit('apply', { patch: { language: draft.value.language, showLineNumbers: draft.value.showLineNumbers, tabSize: draft.value.tabSize, wrap: draft.value.wrap } })
}
</script>

<template>
  <Form :model="draft" layout="vertical" :disabled="disabled" class="document-code-controls" @finish="apply">
    <FormItem v-if="kind === 'font'" label="等宽字体" name="fontFamily">
      <Select v-model:value="draft.fontFamily" :options="fonts" show-search option-filter-prop="label" aria-label="代码块等宽字体">
        <template #option="option"><span :style="{ fontFamily: `${option.value}, monospace` }">{{ option.label }}</span></template>
      </Select>
    </FormItem>
    <FormItem v-else label="代码语言" name="language"><Select v-model:value="draft.language" :options="languages" show-search option-filter-prop="label" aria-label="代码块语言" /></FormItem>
    <template v-if="kind === 'settings'">
      <FormItem label="自动换行" name="wrap"><Switch v-model:checked="draft.wrap" aria-label="代码块自动换行" /></FormItem>
      <FormItem label="显示行号" name="showLineNumbers"><Switch v-model:checked="draft.showLineNumbers" aria-label="代码块显示行号" /></FormItem>
      <FormItem label="缩进空格数" name="tabSize"><InputNumber v-model:value="draft.tabSize" :min="1" :max="16" :precision="0" aria-label="代码块缩进空格数" /></FormItem>
    </template>
    <FormItem class="document-code-controls__actions"><Button @click="emit('cancel')">取消</Button><Button type="primary" html-type="submit" :disabled="!available">应用</Button></FormItem>
  </Form>
</template>

<style scoped>
.document-code-controls { min-width: 220px; }
.document-code-controls :deep(.ant-form-item) { margin-bottom: 12px; }
.document-code-controls__actions { margin-bottom: 0 !important; }
.document-code-controls__actions :deep(.ant-form-item-control-input-content) { display: flex; justify-content: flex-end; gap: 8px; }
</style>
