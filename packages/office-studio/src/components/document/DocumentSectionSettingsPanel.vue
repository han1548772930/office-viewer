<script setup lang="ts">
import { Drawer, Form, FormItem, InputNumber, Select } from 'antdv-next'
import { computed, inject, ref, watch } from 'vue'
import type { DocumentSectionBreakType, DocumentSectionSetup } from '../../domain/model.ts'
import { validateDocumentSectionSetupPatch } from '../../features/document/services/document-settings.ts'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

const props = withDefaults(defineProps<{ open: boolean; sectionId: string; setup: DocumentSectionSetup; sectionOptions?: Array<{ label: string; value: string }> }>(), { sectionOptions: () => [] })
const emit = defineEmits<{ 'update:open': [open: boolean]; apply: [sectionId: string, patch: Partial<DocumentSectionSetup>]; selectSection: [sectionId: string] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const draft = ref({ ...props.setup })
const error = ref('')
const orientations = computed(() => (['portrait', 'landscape'] as const).map(value => ({ value, label: t(`doc.section-settings.orientation.${value}`, value) })))
const separators = computed(() => (['none', 'line'] as const).map(value => ({ value, label: t(`doc.section-settings.separator.${value}`, value) })))
const sectionStarts = computed(() => ([
  ['unspecified', 'unspecified'],
  ['nextPage', 'next-page'], ['continuous', 'continuous'], ['nextColumn', 'next-column'], ['evenPage', 'even-page'], ['oddPage', 'odd-page'],
] as const).map(([value, label]) => ({ value, label: t(`doc.section-settings.start.${label}`, label) })))
const sectionStartValue = computed({
  get: () => draft.value.sectionStart ?? 'unspecified',
  set: (value: string) => { draft.value.sectionStart = value as DocumentSectionBreakType },
})
const orientationValue = computed({
  get: () => draft.value.orientation ?? ((draft.value.width ?? 0) > (draft.value.height ?? 0) ? 'landscape' : 'portrait'),
  set: (orientation: 'portrait' | 'landscape') => {
    const previous = draft.value.orientation ?? ((draft.value.width ?? 0) > (draft.value.height ?? 0) ? 'landscape' : 'portrait')
    if (orientation === previous) return
    draft.value = { ...draft.value, orientation, width: draft.value.height, height: draft.value.width }
  },
})

watch(() => [props.open, props.sectionId] as const, ([open]) => { if (open) { draft.value = { ...props.setup }; error.value = '' } })
watch(() => props.setup, setup => { if (props.open) draft.value = { ...setup } }, { deep: true })
watch(draft, value => {
  if (!props.open) return
  const changed = (Object.keys(value) as Array<keyof DocumentSectionSetup>).some(key => value[key] !== props.setup[key])
  if (!changed) return
  try {
    validateDocumentSectionSetupPatch(value)
    error.value = ''
    emit('apply', props.sectionId, { ...value })
  } catch (reason) {
    error.value = reason instanceof Error ? reason.message : t('doc.section-settings.invalid', 'Invalid section settings')
  }
}, { deep: true })
</script>

<template>
  <Drawer :open="open" :title="t('doc.section-settings.title', 'Section settings')" :size="300" placement="right" :mask="false" :destroy-on-hidden="true" @close="emit('update:open', false)">
    <Form :model="draft" layout="horizontal" class="section-settings">
      <FormItem v-if="sectionOptions.length" :label="t('doc.section-settings.section', 'Section')">
        <Select :value="sectionId" :options="sectionOptions" @change="emit('selectSection', String($event))" />
      </FormItem>
      <FormItem name="columnCount" :label="t('doc.section-settings.column-count', 'Column count')">
        <InputNumber v-model:value="draft.columnCount" :min="1" :max="12" :step="1" :precision="0" />
      </FormItem>
      <FormItem name="columnGap" :label="`${t('doc.section-settings.column-gap', 'Column gap')} (px)`">
        <InputNumber v-model:value="draft.columnGap" :min="0" :max="1000" :step="1" :precision="1" :disabled="draft.columnCount === 1" />
      </FormItem>
      <FormItem name="columnSeparator" :label="t('doc.section-settings.separator', 'Separator')">
        <Select v-model:value="draft.columnSeparator" :options="separators" />
      </FormItem>
      <FormItem v-if="sectionId !== 'first'" name="sectionStart" :label="t('doc.section-settings.section-start', 'Section starts')">
        <Select v-model:value="sectionStartValue" :options="sectionStarts" />
      </FormItem>
      <div class="settings-heading">{{ t('doc.section-settings.page', 'Page size and margins') }}</div>
      <FormItem name="width" :label="`${t('doc.section-settings.width', 'Width')} (px)`">
        <InputNumber v-model:value="draft.width" :min="1" :step="1" :precision="1" />
      </FormItem>
      <FormItem name="height" :label="`${t('doc.section-settings.height', 'Height')} (px)`">
        <InputNumber v-model:value="draft.height" :min="1" :step="1" :precision="1" />
      </FormItem>
      <FormItem name="orientation" :label="t('doc.section-settings.orientation', 'Orientation')">
        <Select v-model:value="orientationValue" :options="orientations" />
      </FormItem>
      <FormItem name="marginTop" :label="`${t('doc.section-settings.margin-top', 'Top margin')} (px)`">
        <InputNumber v-model:value="draft.marginTop" :min="0" :step="1" :precision="1" />
      </FormItem>
      <FormItem name="marginBottom" :label="`${t('doc.section-settings.margin-bottom', 'Bottom margin')} (px)`">
        <InputNumber v-model:value="draft.marginBottom" :min="0" :step="1" :precision="1" />
      </FormItem>
      <FormItem name="marginLeft" :label="`${t('doc.section-settings.margin-left', 'Left margin')} (px)`">
        <InputNumber v-model:value="draft.marginLeft" :min="0" :step="1" :precision="1" />
      </FormItem>
      <FormItem name="marginRight" :label="`${t('doc.section-settings.margin-right', 'Right margin')} (px)`">
        <InputNumber v-model:value="draft.marginRight" :min="0" :step="1" :precision="1" />
      </FormItem>
      <FormItem name="pageNumberStart" :label="t('doc.section-settings.page-number-start', 'Page number starts at')">
        <InputNumber v-model:value="draft.pageNumberStart" :min="1" :step="1" :precision="0" />
      </FormItem>
      <p v-if="error" class="settings-error" role="alert">{{ error }}</p>
    </Form>
  </Drawer>
</template>

<style scoped>
.section-settings { display: grid; gap: 10px; max-height: calc(100vh - 120px); overflow-y: auto; padding-inline-end: 2px; }
.section-settings :deep(.ant-form-item) { margin-bottom: 0; }
.section-settings :deep(.ant-form-item-row) { display: grid; grid-template-columns: minmax(0, 1fr) minmax(100px, 130px); align-items: center; gap: 8px; }
.section-settings :deep(.ant-form-item-label) { text-align: start; }
.section-settings :deep(.ant-input-number), .section-settings :deep(.ant-select) { width: 100%; }
.settings-heading { padding-top: 8px; font-size: 13px; font-weight: 500; }
.settings-error { margin: 0; color: #b42318; overflow-wrap: anywhere; }
</style>
