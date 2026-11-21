<script setup lang="ts">
import { InputNumber, Modal, Select } from 'antdv-next'
import { computed, inject, ref, watch } from 'vue'
import type { DocumentSectionSetup } from '../../domain/model.ts'
import { validateDocumentSectionSetupPatch } from '../../features/document/services/document-settings.ts'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

const props = defineProps<{ open: boolean; sectionId: string; setup: DocumentSectionSetup }>()
const emit = defineEmits<{ 'update:open': [open: boolean]; apply: [sectionId: string, patch: Partial<DocumentSectionSetup>] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const draft = ref({ ...props.setup })
const error = ref('')
const orientations = computed(() => (['portrait', 'landscape'] as const).map(value => ({ value, label: t(`doc.section-settings.orientation.${value}`, value) })))
const separators = computed(() => (['none', 'line'] as const).map(value => ({ value, label: t(`doc.section-settings.separator.${value}`, value) })))

watch(() => [props.open, props.sectionId] as const, ([open]) => { if (open) { draft.value = { ...props.setup }; error.value = '' } })

function apply() {
  try {
    validateDocumentSectionSetupPatch(draft.value)
    emit('apply', props.sectionId, { ...draft.value })
    emit('update:open', false)
  } catch (reason) {
    error.value = reason instanceof Error ? reason.message : t('doc.section-settings.invalid', 'Invalid section settings')
  }
}
</script>

<template>
  <Modal :open="open" :title="t('doc.section-settings.title', 'Section settings')" :ok-text="t('common.apply', 'Apply')" :cancel-text="t('common.cancel', 'Cancel')" :width="460" destroy-on-hidden @ok="apply" @cancel="emit('update:open', false)">
    <div class="section-settings">
      <label>{{ t('doc.section-settings.orientation', 'Orientation') }}<Select v-model:value="draft.orientation" :options="orientations" /></label>
      <fieldset>
        <legend>{{ t('doc.section-settings.page', 'Page size and margins') }}</legend>
        <label>{{ t('doc.section-settings.width', 'Width') }}<InputNumber v-model:value="draft.width" :min="200" :precision="1" /></label>
        <label>{{ t('doc.section-settings.height', 'Height') }}<InputNumber v-model:value="draft.height" :min="240" :precision="1" /></label>
        <label>{{ t('doc.section-settings.margin-top', 'Top margin') }}<InputNumber v-model:value="draft.marginTop" :min="0" :precision="1" /></label>
        <label>{{ t('doc.section-settings.margin-right', 'Right margin') }}<InputNumber v-model:value="draft.marginRight" :min="0" :precision="1" /></label>
        <label>{{ t('doc.section-settings.margin-bottom', 'Bottom margin') }}<InputNumber v-model:value="draft.marginBottom" :min="0" :precision="1" /></label>
        <label>{{ t('doc.section-settings.margin-left', 'Left margin') }}<InputNumber v-model:value="draft.marginLeft" :min="0" :precision="1" /></label>
      </fieldset>
      <fieldset>
        <legend>{{ t('doc.section-settings.columns', 'Columns') }}</legend>
        <label>{{ t('doc.section-settings.column-count', 'Column count') }}<InputNumber v-model:value="draft.columnCount" :min="1" :max="12" :precision="0" /></label>
        <label>{{ t('doc.section-settings.column-gap', 'Column gap') }}<InputNumber v-model:value="draft.columnGap" :min="0" :precision="1" /></label>
        <label>{{ t('doc.section-settings.separator', 'Separator') }}<Select v-model:value="draft.columnSeparator" :options="separators" /></label>
      </fieldset>
      <label>{{ t('doc.section-settings.page-number-start', 'Page number starts at') }}<InputNumber v-model:value="draft.pageNumberStart" :min="1" :precision="0" /></label>
      <p v-if="error" class="settings-error" role="alert">{{ error }}</p>
    </div>
  </Modal>
</template>

<style scoped>
.section-settings { display: grid; gap: 14px; }
.section-settings > label, fieldset label { display: grid; grid-template-columns: minmax(0, 1fr) 180px; align-items: center; gap: 10px; min-width: 0; }
fieldset { display: grid; gap: 9px; min-width: 0; margin: 0; padding: 10px 0 0; border: 0; border-top: 1px solid #e5e8ed; }
legend { padding-inline-end: 8px; font-size: 13px; }
fieldset label :deep(.ant-input-number), .section-settings > label :deep(.ant-select) { width: 100%; }
.settings-error { margin: 0; color: #b42318; overflow-wrap: anywhere; }
</style>
