<script setup lang="ts">
import { Button, InputNumber, Modal, Select } from 'antdv-next'
import { computed, inject, ref, watch } from 'vue'
import type { DocumentPageSetup } from '../../domain/model'
import { DOCUMENT_MODERN_WIDTHS, DOCUMENT_PAPER_SIZES, documentModernWidth, documentPageSettingsError, documentPaperSize } from './document-page-settings'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

const props = withDefaults(defineProps<{ open: boolean; page: DocumentPageSetup; mode?: 'traditional' | 'modern' }>(), { mode: 'traditional' })
const emit = defineEmits<{ 'update:open': [open: boolean]; apply: [page: DocumentPageSetup, mode: 'traditional' | 'modern'] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const draft = ref({ ...props.page })
const error = ref('')
const paperOptions = computed(() => [
  ...(['A3', 'A4', 'A5', 'B4', 'B5', 'Executive', 'Folio', 'Legal', 'Letter', 'Statement', 'Tabloid'] as const).map(value => ({ value, label: t(`doc.page-settings.paper-size.${value}`, value) })),
  { value: 'Custom', label: t('doc.page-settings.paper-size.custom', 'Custom') },
])
const paper = computed(() => documentPaperSize(draft.value))
const modernWidth = computed(() => documentModernWidth(draft.value))
watch(() => props.open, (open) => { if (open) { draft.value = { ...props.page }; error.value = '' } })
function setPaper(value: string) {
  if (value === 'Custom' || !(value in DOCUMENT_PAPER_SIZES)) return
  const size = DOCUMENT_PAPER_SIZES[value as keyof typeof DOCUMENT_PAPER_SIZES]
  const landscape = draft.value.width > draft.value.height
  draft.value = { ...draft.value, width: landscape ? size.height : size.width, height: landscape ? size.width : size.height }
}
function setModernWidth(value: keyof typeof DOCUMENT_MODERN_WIDTHS) {
  draft.value = { ...draft.value, width: DOCUMENT_MODERN_WIDTHS[value], height: 1124 }
}
function apply() {
  const message = documentPageSettingsError(draft.value)
  if (message) { error.value = message; return }
  emit('apply', { ...draft.value }, props.mode)
  emit('update:open', false)
}
</script>

<template>
  <Modal :open="open" :title="t('doc.page-settings.title', '页面设置')" :ok-text="t('common.apply', '应用')" :cancel-text="t('common.cancel', '取消')" :width="528" destroy-on-hidden @ok="apply" @cancel="emit('update:open', false)">
    <div class="page-settings-form">
      <div v-if="props.mode === 'traditional'">
      <fieldset>
        <legend>{{ t('doc.page-settings.paper', '纸张尺寸') }}</legend>
        <label class="paper-choice">{{ t('doc.page-settings.paper-size', '纸张') }}<Select :value="paper" :options="paperOptions" :aria-label="t('doc.page-settings.paper-size', '纸张')" @change="setPaper" /></label>
        <label>{{ t('doc.page-settings.width', '宽度') }}<InputNumber v-model:value="draft.width" :aria-label="t('doc.page-settings.width-aria', '纸张宽度')" :min="200" :precision="1" /></label>
        <label>{{ t('doc.page-settings.height', '高度') }}<InputNumber v-model:value="draft.height" :aria-label="t('doc.page-settings.height-aria', '纸张高度')" :min="240" :precision="1" /></label>
      </fieldset>
      </div>
      <fieldset v-else class="modern-widths">
        <legend>{{ t('doc.page-settings.modern-width', '内容宽度') }}</legend>
        <Button v-for="(width, value) in DOCUMENT_MODERN_WIDTHS" :key="value" :type="modernWidth === value ? 'primary' : 'default'" @click="setModernWidth(value)">
          {{ t(`doc.page-settings.modern-width-${value}`, value) }} <span>{{ width }}px</span>
        </Button>
      </fieldset>
      <fieldset v-if="props.mode === 'traditional'">
        <legend>{{ t('doc.page-settings.margins', '页边距') }}</legend>
        <label>{{ t('doc.page-settings.margin-top', '上') }}<InputNumber v-model:value="draft.marginTop" :aria-label="t('doc.page-settings.margin-top-aria', '上边距')" :min="0" :max="draft.height / 2" :precision="1" /></label>
        <label>{{ t('doc.page-settings.margin-bottom', '下') }}<InputNumber v-model:value="draft.marginBottom" :aria-label="t('doc.page-settings.margin-bottom-aria', '下边距')" :min="0" :max="draft.height / 2" :precision="1" /></label>
        <label>{{ t('doc.page-settings.margin-left', '左') }}<InputNumber v-model:value="draft.marginLeft" :aria-label="t('doc.page-settings.margin-left-aria', '左边距')" :min="0" :max="draft.width / 2" :precision="1" /></label>
        <label>{{ t('doc.page-settings.margin-right', '右') }}<InputNumber v-model:value="draft.marginRight" :aria-label="t('doc.page-settings.margin-right-aria', '右边距')" :min="0" :max="draft.width / 2" :precision="1" /></label>
      </fieldset>
      <p v-if="error" class="page-settings-error" role="alert">{{ error }}</p>
    </div>
  </Modal>
</template>

<style scoped>
.page-settings-form { padding-top: 8px; color: #4f5b69; }
.page-settings-form p { margin: 0 0 16px; color: #7b8592; font-size: 12px; }
fieldset { display: grid; grid-template-columns: 1fr 1fr; gap: 14px 18px; min-width: 0; margin: 0 0 20px; padding: 12px 0 0; border: 0; border-top: 1px solid #e5e8ed; }
.paper-choice { grid-column: 1 / -1; }
legend { padding-right: 12px; font-size: 13px; }
label { display: flex; align-items: center; justify-content: space-between; gap: 12px; }
label :deep(.ant-input-number), .paper-choice :deep(.ant-select) { width: 156px; }
.modern-widths { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); }
.modern-widths legend { grid-column: 1 / -1; }
.modern-widths button { height: 64px; display: grid; align-content: center; gap: 4px; }
.modern-widths span { font-size: 12px; color: #73808e; }
.page-settings-form .page-settings-error { margin-bottom: 0; color: #c14444; }
</style>
