<script setup lang="ts">
import { computed, nextTick, ref, watch } from 'vue'
import { useI18n } from 'vue-i18n'
import { Alert, Form, FormItem, Input, Modal, Select, theme, type TextAreaRef } from 'antdv-next'
import { LatexFormulaError, layoutLatexFormula, type LatexFormula, type LatexFormulaMode } from '../../features/document/latex/latex-formula.ts'
import { getDocumentLatexTemplate, insertDocumentLatexTemplate } from '../../features/document/latex/latex-templates'
import DocumentLatexTemplateGallery from './DocumentLatexTemplateGallery.vue'
import DocumentLatexPreview from './DocumentLatexPreview.vue'

const props = defineProps<{ open: boolean; initial?: LatexFormula; initialSource?: string; editorKey?: number; compact?: boolean; readonly?: boolean; error?: string }>()
const emit = defineEmits<{ 'update:open': [value: boolean]; apply: [value: { latex: string; mode: LatexFormulaMode }]; 'template-request': [templateId: string]; 'draft-change': []; 'composition-change': [value: boolean]; closed: [] }>()
const { t } = useI18n({ useScope: 'global' })
const { token } = theme.useToken()
const source = ref('')
const mode = ref<LatexFormulaMode>('inline')
const composing = ref(false)
const sourceInput = ref<TextAreaRef>()
const sourceSelection = ref({ start: 0, end: 0 })
let draftGeneration = 0
let compositionGeneration = 0
const modeOptions = computed(() => [
  { label: t('document.latex.inline'), value: 'inline' },
  { label: t('document.latex.block'), value: 'block' },
])
const validationError = computed(() => {
  try { layoutLatexFormula(source.value); return '' }
  catch (cause) {
    const message = cause instanceof Error ? cause.message : t('document.latex.invalid')
    return cause instanceof LatexFormulaError ? t('document.latex.invalid-offset', { message, offset: cause.offset + 1 }) : message
  }
})
const canApply = computed(() => props.open && !props.readonly && !composing.value && !validationError.value)

watch(() => [props.open, props.editorKey] as const, ([open]) => {
  draftGeneration++
  compositionGeneration++
  composing.value = false
  if (!open) return
  source.value = props.initialSource ?? props.initial?.latex ?? ''
  mode.value = props.initial?.mode ?? 'inline'
  sourceSelection.value = { start: source.value.length, end: source.value.length }
}, { immediate: true })

watch([source, mode], () => { draftGeneration++; emit('draft-change') }, { flush: 'sync' })
watch(composing, value => emit('composition-change', value), { flush: 'sync' })
/** TextArea exposes its native element, but has no public source-range setter. */
function sourceElement(): HTMLTextAreaElement | undefined {
  const native = sourceInput.value?.nativeElement
  return native instanceof HTMLTextAreaElement ? native : native instanceof HTMLElement ? native.querySelector('textarea') ?? undefined : undefined
}
function rememberSourceRange() {
  const native = sourceElement()
  if (native) sourceSelection.value = { start: native.selectionStart, end: native.selectionEnd }
}
function insertTemplate(templateId: string): boolean {
  if (!props.open || props.readonly || composing.value || !getDocumentLatexTemplate(templateId)) return false
  rememberSourceRange()
  const result = insertDocumentLatexTemplate(source.value, templateId, sourceSelection.value.start, sourceSelection.value.end)
  source.value = result.value
  sourceSelection.value = { start: result.selectionStart, end: result.selectionEnd }
  const generation = draftGeneration
  void nextTick(() => {
    if (!props.open || props.readonly || composing.value || generation !== draftGeneration || source.value !== result.value) return
    const native = sourceElement()
    sourceInput.value?.focus()
    native?.setSelectionRange(result.selectionStart, result.selectionEnd)
  })
  return true
}
defineExpose({ insertTemplate })

function apply() {
  if (canApply.value) emit('apply', { latex: source.value, mode: mode.value })
}
function guardKey(event: KeyboardEvent) {
  if (composing.value || event.isComposing || event.keyCode === 229) {
    event.stopPropagation()
    if (event.key === 'Escape') event.preventDefault()
    return
  }
  if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) { event.preventDefault(); event.stopPropagation(); apply() }
}
function compositionStart() { compositionGeneration++; composing.value = true }
async function compositionEnd() {
  const generation = compositionGeneration
  await nextTick()
  if (!props.open || generation !== compositionGeneration) return
  composing.value = false; rememberSourceRange()
}
function afterOpenChange(open: boolean) {
  if (open && props.open && !composing.value) { sourceInput.value?.focus(); sourceElement()?.setSelectionRange(sourceSelection.value.start, sourceSelection.value.end) }
}
</script>

<template>
  <Modal :open="open" :title="t(initial ? 'document.latex.editTitle' : 'document.latex.title')" :ok-text="t('common.apply')" :cancel-text="t('common.cancel')" :width="compact ? 'calc(100vw - 24px)' : 600"
    :ok-button-props="{ disabled: !canApply, size: compact ? 'large' : 'middle' }" :cancel-button-props="{ size: compact ? 'large' : 'middle' }"
    :keyboard="!composing" :mask-closable="false" :focus-trigger-after-close="false" :mask="false" :scroll-lock="false"
    :focusable="{ trap: false, focusTriggerAfterClose: false }" :wrap-props="{ style: { pointerEvents: 'none' } }" :style="{ pointerEvents: 'auto', top: compact ? '12px' : '24px', maxWidth: 'calc(100vw - 24px)' }" destroy-on-hidden
    @ok="apply" @cancel="emit('update:open', false)" @after-close="emit('closed')" @after-open-change="afterOpenChange">
    <div class="latex-panel-body" :class="{ 'latex-panel-body--compact': compact }" :data-document-latex-mobile-popup="compact || undefined">
    <Form layout="vertical" class="latex-panel" :style="{ '--latex-panel-color': token.colorTextSecondary, '--latex-panel-border': token.colorBorder, '--latex-panel-background': token.colorBgContainer }" @keydown.capture="guardKey">
      <Alert v-if="error" :title="error" type="error" role="alert" show-icon />
      <Alert v-else-if="readonly" :title="t('document.latex.read-only')" type="info" show-icon />
      <FormItem class="latex-panel__mode" :label="t('document.latex.mode')">
        <Select v-model:value="mode" :options="modeOptions" :disabled="!!initial || readonly" :aria-label="t('document.latex.mode')" />
      </FormItem>
      <FormItem class="latex-panel__source" :label="t('document.latex.source')" :validate-status="validationError && !composing ? 'error' : undefined" :help="!composing ? validationError || undefined : undefined">
        <Input.TextArea ref="sourceInput" v-model:value="source" :readonly="readonly" :aria-label="t('document.latex.source')" :auto-size="{ minRows: 3, maxRows: 8 }" spellcheck="false"
          @select="rememberSourceRange" @mouseup="rememberSourceRange" @keyup="rememberSourceRange" @change="rememberSourceRange"
          @compositionstart="compositionStart" @compositionend="compositionEnd" />
      </FormItem>
      <div class="latex-panel__preview">
        <span>{{ t('document.latex.preview') }}</span>
        <DocumentLatexPreview :latex="source" :label="t('document.latex.preview')" />
      </div>
    </Form>
    <DocumentLatexTemplateGallery class="latex-panel__templates" :disabled="readonly || composing" @insert="request => emit('template-request', request.templateId)" />
    </div>
  </Modal>
</template>

<style scoped>
.latex-panel { display: grid; gap: 14px; }
.latex-panel-body { max-height: max(240px, calc(100dvh - 220px)); overflow-y: auto; overscroll-behavior: contain; }
.latex-panel-body--compact { max-height: max(64px, calc(100dvh - 240px)); }
.latex-panel-body--compact .latex-panel__mode :deep(.ant-form-item-row) { grid-template-columns: 86px minmax(0, 1fr); gap: 8px; }
.latex-panel-body--compact :deep(.latex-template-gallery__grid) { grid-template-columns: repeat(2, minmax(0, 1fr)); max-height: 280px; }
.latex-panel :deep(.ant-form-item) { margin-bottom: 0; color: var(--latex-panel-color); font-size: 12px; }
.latex-panel :deep(.ant-form-item-label) { padding-bottom: 6px; }
.latex-panel__mode :deep(.ant-form-item-row) { display: grid; grid-template-columns: 120px minmax(0, 1fr); align-items: center; gap: 10px; }
.latex-panel__mode :deep(.ant-form-item-label) { padding-bottom: 0; }
.latex-panel__preview { display: grid; gap: 6px; color: var(--latex-panel-color); font-size: 12px; }
.latex-panel__source :deep(textarea) { font: 13px/1.5 "IBM Plex Mono", Consolas, monospace; }
.latex-panel__preview :deep(canvas) { border: 1px solid var(--latex-panel-border); background: var(--latex-panel-background); }
.latex-panel__templates { width: 100%; margin-top: 14px; }
</style>
