<script setup lang="ts">
import { Alert, Button, Form, FormItem, InputNumber, Modal, Select, theme } from 'antdv-next'
import { computed, inject, ref, watch } from 'vue'
import type { DocumentPageSetup } from '../../domain/model'
import { DOCUMENT_MODERN_WIDTHS, DOCUMENT_PAPER_SIZES, documentModernWidth, documentPageSettingsError, documentPaperSize } from './document-page-settings'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

const props = withDefaults(defineProps<{ open: boolean; page: DocumentPageSetup; mode?: 'traditional' | 'modern'; disabled?: boolean }>(), { mode: 'traditional', disabled: false })
const emit = defineEmits<{ 'update:open': [open: boolean]; apply: [page: DocumentPageSetup, mode: 'traditional' | 'modern'] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const { token } = theme.useToken()
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
  if (props.disabled) return
  const message = documentPageSettingsError(draft.value)
  if (message) { error.value = message; return }
  emit('apply', { ...draft.value }, props.mode)
}
</script>

<template>
  <Modal :open="open" :title="t('doc.page-settings.title', '页面设置')" :ok-text="t('common.apply', '应用')" :cancel-text="t('common.cancel', '取消')" :ok-button-props="{ disabled }" :width="528" destroy-on-hidden @ok="apply" @cancel="emit('update:open', false)">
    <Form :model="draft" :disabled="disabled" layout="horizontal" class="page-settings-form">
      <div v-if="props.mode === 'traditional'">
      <fieldset>
        <legend>{{ t('doc.page-settings.paper', '纸张尺寸') }}</legend>
        <FormItem class="paper-choice" :label="t('doc.page-settings.paper-size', '纸张')">
          <Select :value="paper" :options="paperOptions" :aria-label="t('doc.page-settings.paper-size', '纸张')" @change="setPaper" />
        </FormItem>
        <FormItem name="width" :label="t('doc.page-settings.width', '宽度')">
          <InputNumber v-model:value="draft.width" :aria-label="t('doc.page-settings.width-aria', '纸张宽度')" :min="200" :precision="1" />
        </FormItem>
        <FormItem name="height" :label="t('doc.page-settings.height', '高度')">
          <InputNumber v-model:value="draft.height" :aria-label="t('doc.page-settings.height-aria', '纸张高度')" :min="240" :precision="1" />
        </FormItem>
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
        <FormItem name="marginTop" :label="t('doc.page-settings.margin-top', '上')">
          <InputNumber v-model:value="draft.marginTop" :aria-label="t('doc.page-settings.margin-top-aria', '上边距')" :min="0" :max="draft.height / 2" :precision="1" />
        </FormItem>
        <FormItem name="marginBottom" :label="t('doc.page-settings.margin-bottom', '下')">
          <InputNumber v-model:value="draft.marginBottom" :aria-label="t('doc.page-settings.margin-bottom-aria', '下边距')" :min="0" :max="draft.height / 2" :precision="1" />
        </FormItem>
        <FormItem name="marginLeft" :label="t('doc.page-settings.margin-left', '左')">
          <InputNumber v-model:value="draft.marginLeft" :aria-label="t('doc.page-settings.margin-left-aria', '左边距')" :min="0" :max="draft.width / 2" :precision="1" />
        </FormItem>
        <FormItem name="marginRight" :label="t('doc.page-settings.margin-right', '右')">
          <InputNumber v-model:value="draft.marginRight" :aria-label="t('doc.page-settings.margin-right-aria', '右边距')" :min="0" :max="draft.width / 2" :precision="1" />
        </FormItem>
      </fieldset>
      <Alert v-if="error" :title="error" type="error" role="alert" show-icon />
    </Form>
  </Modal>
</template>

<style scoped>
.page-settings-form { padding-top: 8px; color: v-bind('token.colorText'); }
.page-settings-form :deep(.ant-form-item) { min-width: 0; margin-bottom: 0; }
.page-settings-form :deep(.ant-form-item-row) { display: flex; align-items: center; justify-content: space-between; gap: 12px; }
.page-settings-form :deep(.ant-form-item-label) { text-align: start; }
.page-settings-form :deep(.ant-form-item-control) { flex: 0 0 156px; }
.page-settings-form :deep(.ant-form-item-control-input-content), .page-settings-form :deep(.ant-input-number), .page-settings-form :deep(.ant-select) { width: 100%; }
fieldset { display: grid; grid-template-columns: 1fr 1fr; gap: 14px 18px; min-width: 0; margin: 0 0 20px; padding: 12px 0 0; border: 0; border-top: 1px solid v-bind('token.colorBorderSecondary'); }
.paper-choice { grid-column: 1 / -1; }
legend { padding-right: 12px; font-size: 13px; }
.modern-widths { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); }
.modern-widths legend { grid-column: 1 / -1; }
.modern-widths button { height: 64px; display: grid; align-content: center; gap: 4px; }
.modern-widths span { font-size: 12px; color: v-bind('token.colorTextSecondary'); }
</style>
