<script setup lang="ts">
import { Checkbox, InputNumber, Modal, Select } from 'antdv-next'
import { computed, inject, ref, watch } from 'vue'
import type { DocumentParagraphStyle } from '../../domain/model.ts'
import { validateDocumentParagraphStylePatch } from '../../features/document/services/document-settings.ts'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

const props = defineProps<{ open: boolean; style: DocumentParagraphStyle }>()
const emit = defineEmits<{ 'update:open': [open: boolean]; apply: [patch: Partial<DocumentParagraphStyle>] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const draft = ref({ ...props.style })
const error = ref('')
const alignments = computed(() => (['left', 'center', 'right', 'justify'] as const).map(value => ({ value, label: t(`doc.paragraph-settings.align.${value}`, value) })))
const spacingRules = computed(() => (['auto', 'atLeast', 'exact'] as const).map(value => ({ value, label: t(`doc.paragraph-settings.spacing-rule.${value}`, value) })))

watch(() => props.open, open => { if (open) { draft.value = { ...props.style }; error.value = '' } })

function apply() {
  try {
    validateDocumentParagraphStylePatch(draft.value)
    emit('apply', { ...draft.value })
    emit('update:open', false)
  } catch (reason) {
    error.value = reason instanceof Error ? reason.message : t('doc.paragraph-settings.invalid', 'Invalid paragraph settings')
  }
}
</script>

<template>
  <Modal :open="open" :title="t('doc.paragraph-settings.title', 'Paragraph settings')" :ok-text="t('common.apply', 'Apply')" :cancel-text="t('common.cancel', 'Cancel')" :width="440" destroy-on-hidden @ok="apply" @cancel="emit('update:open', false)">
    <div class="paragraph-settings">
      <label>{{ t('doc.paragraph-settings.alignment', 'Alignment') }}<Select v-model:value="draft.align" :options="alignments" /></label>
      <fieldset>
        <legend>{{ t('doc.paragraph-settings.indentation', 'Indentation') }}</legend>
        <label>{{ t('doc.paragraph-settings.left-indent', 'Left indent') }}<InputNumber v-model:value="draft.indentStart" :min="0" :precision="1" /></label>
        <label>{{ t('doc.paragraph-settings.right-indent', 'Right indent') }}<InputNumber v-model:value="draft.indentEnd" :min="0" :precision="1" /></label>
        <label>{{ t('doc.paragraph-settings.first-line-indent', 'First-line indent') }}<InputNumber v-model:value="draft.indentFirstLine" :min="0" :precision="1" /></label>
        <label>{{ t('doc.paragraph-settings.hanging-indent', 'Hanging indent') }}<InputNumber v-model:value="draft.hanging" :min="0" :precision="1" /></label>
      </fieldset>
      <fieldset>
        <legend>{{ t('doc.paragraph-settings.spacing', 'Spacing') }}</legend>
        <label>{{ t('doc.paragraph-settings.before', 'Before') }}<InputNumber v-model:value="draft.spaceBefore" :min="0" :precision="1" /></label>
        <label>{{ t('doc.paragraph-settings.after', 'After') }}<InputNumber v-model:value="draft.spaceAfter" :min="0" :precision="1" /></label>
        <label>{{ t('doc.paragraph-settings.line-rule', 'Line spacing rule') }}<Select v-model:value="draft.spacingRule" :options="spacingRules" /></label>
        <label>{{ t('doc.paragraph-settings.line-spacing', 'Line spacing') }}<InputNumber v-model:value="draft.lineHeight" :min="1" :step="0.1" :precision="2" /></label>
      </fieldset>
      <Checkbox v-model:checked="draft.wordWrap">{{ t('doc.paragraph-settings.word-wrap', 'Wrap text') }}</Checkbox>
      <Checkbox v-model:checked="draft.snapToGrid">{{ t('doc.paragraph-settings.snap-to-grid', 'Snap to document grid') }}</Checkbox>
      <p v-if="error" class="settings-error" role="alert">{{ error }}</p>
    </div>
  </Modal>
</template>

<style scoped>
.paragraph-settings { display: grid; gap: 14px; }
.paragraph-settings > label, fieldset label { display: grid; grid-template-columns: minmax(0, 1fr) 180px; align-items: center; gap: 10px; min-width: 0; }
fieldset { display: grid; gap: 9px; min-width: 0; margin: 0; padding: 10px 0 0; border: 0; border-top: 1px solid #e5e8ed; }
legend { padding-inline-end: 8px; font-size: 13px; }
fieldset label :deep(.ant-input-number), .paragraph-settings > label :deep(.ant-select) { width: 100%; }
.settings-error { margin: 0; color: #b42318; overflow-wrap: anywhere; }
</style>
