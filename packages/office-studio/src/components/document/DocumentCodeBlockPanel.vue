<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { useI18n } from 'vue-i18n'
import { Form, FormItem, Input, InputNumber, Modal, Select, Switch, theme } from 'antdv-next'
import type { DocumentCodeBlockData } from '../../features/document/code-blocks/code-blocks.ts'
import { documentStructuredCodeLanguageOptions } from '../../features/document/code-blocks/document-code-languages'
import { documentCodeFontFamilyOptions } from '../../features/document/code-blocks/document-code-fonts'

const props = defineProps<{ open: boolean; initial?: DocumentCodeBlockData; initialFocus?: 'font' | 'language' | 'settings' }>()
const { t } = useI18n({ useScope: 'global' })
const emit = defineEmits<{ 'update:open': [value: boolean]; apply: [value: DocumentCodeBlockData]; 'composition-change': [value: boolean]; closed: [] }>()

const source = ref('')
const composing = ref(false)
const language = ref('typescript')
const lineNumbers = ref(false)
const tabSize = ref(4)
const wrap = ref(true)
const fontFamily = ref('Consolas')
const fontSelect = ref<{ focus: () => void }>()
const languageSelect = ref<{ focus: () => void }>()
const indentationInput = ref<{ focus: () => void }>()
const { token } = theme.useToken()
const languages = computed(() => documentStructuredCodeLanguageOptions(language.value))
const fonts = computed(() => documentCodeFontFamilyOptions(fontFamily.value))

watch(() => [props.open, props.initial] as const, ([open, initial]) => {
  composing.value = false
  if (!open) return
  source.value = initial?.source ?? ''
  language.value = initial?.language ?? 'typescript'
  lineNumbers.value = initial?.lineNumbers ?? false
  tabSize.value = initial?.tabSize ?? 4
  wrap.value = initial?.wrap ?? true
  fontFamily.value = initial?.fontFamily ?? 'Consolas'
}, { immediate: true })
watch(composing, value => emit('composition-change', value), { flush: 'sync' })

function apply() {
  if (!props.open || composing.value) return
  emit('apply', { source: source.value, language: language.value, lineNumbers: lineNumbers.value, tabSize: tabSize.value, wrap: wrap.value,
    ...(fontFamily.value === props.initial?.fontFamily ? {} : { fontFamily: fontFamily.value }) })
}
function afterOpenChange(open: boolean) {
  if (!open || !props.open || composing.value) return
  if (props.initialFocus === 'font') fontSelect.value?.focus()
  else if (props.initialFocus === 'language') languageSelect.value?.focus()
  else if (props.initialFocus === 'settings') indentationInput.value?.focus()
}
function guardKey(event: KeyboardEvent) {
  if (!composing.value && !event.isComposing && event.keyCode !== 229) return
  event.stopPropagation()
  if (event.key === 'Escape') event.preventDefault()
}
</script>

<template>
  <Modal :open="open" :title="t('document.code-block.title')" :ok-text="t('common.apply')" :cancel-text="t('common.cancel')" :width="640"
    :keyboard="!composing" :closable="!composing" :mask-closable="!composing" :ok-button-props="{ disabled: composing }" :cancel-button-props="{ disabled: composing }"
    :focus-trigger-after-close="false" destroy-on-hidden
    :after-close="() => emit('closed')" :after-open-change="afterOpenChange"
    @ok="apply" @cancel="() => { if (!composing) emit('update:open', false) }">
    <Form layout="vertical" class="code-block-panel" :style="{ '--code-panel-color': token.colorTextSecondary }"
      @compositionstart="composing = true" @compositionend="composing = false" @keydown.capture="guardKey">
      <FormItem :label="t('document.code-block.language')"><Select ref="languageSelect" v-model:value="language" :options="languages" :aria-label="t('document.code-block.language-label')" show-search option-filter-prop="label" /></FormItem>
      <FormItem label="等宽字体"><Select ref="fontSelect" v-model:value="fontFamily" :options="fonts" show-search option-filter-prop="label" aria-label="代码块等宽字体" /></FormItem>
      <FormItem class="code-block-panel__toggle" :label="t('document.code-block.line-numbers')"><Switch v-model:checked="lineNumbers" :aria-label="t('document.code-block.show-line-numbers')" /></FormItem>
      <FormItem label="缩进空格数"><InputNumber ref="indentationInput" v-model:value="tabSize" :min="1" :max="16" :precision="0" aria-label="代码块缩进空格数" /></FormItem>
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
