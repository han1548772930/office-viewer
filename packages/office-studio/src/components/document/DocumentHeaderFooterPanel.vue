<script setup lang="ts">
import { computed, inject } from 'vue'
import { Button, Form, FormItem, Input, InputNumber, Modal, Switch } from 'antdv-next'
import type { DocumentModel, DocumentSectionSetup } from '../../domain/model'
import { documentSectionIds, resolveSectionHeaderFooter, type SectionHeaderFooterPatch } from '../../features/document/services/section-header-footer'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

const props = defineProps<{ open: boolean; model: DocumentModel; sectionId: string; setup: DocumentSectionSetup }>()
const emit = defineEmits<{ 'update:open': [value: boolean]; closed: []; position: [patch: Pick<DocumentSectionSetup, 'marginHeader' | 'marginFooter'>]; apply: [patch: SectionHeaderFooterPatch]; 'insert-page-field': [story: 'header' | 'footer', field: 'PAGE' | 'NUMPAGES'] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const section = computed(() => resolveSectionHeaderFooter(props.model, props.sectionId))
const sections = computed(() => documentSectionIds(props.model))

function update(key: keyof typeof section.value, value: string | boolean) {
  emit('apply', { [key]: value })
}
function setDistance(key: 'marginHeader' | 'marginFooter', value: unknown) {
  if (typeof value === 'number' && Number.isFinite(value) && value >= 0 && value <= 200 && value !== (props.setup[key] ?? 48)) emit('position', { [key]: value })
}
</script>

<template>
  <Modal :open="open" :title="t('doc.header-footer.title', '页眉和页脚')" :ok-text="t('common.apply', '应用')" :cancel-text="t('common.close', '关闭')" :width="520" @ok="emit('update:open', false)" @cancel="emit('update:open', false)" @after-close="emit('closed')">
    <Form layout="vertical" class="header-footer-panel">
      <FormItem class="header-footer-panel__row" :label="t('doc.header-footer.current-section', '当前节')"><Input :value="`${sections.indexOf(sectionId) + 1}`" readonly /></FormItem>
      <FormItem class="header-footer-panel__toggle" :label="t('doc.header-footer.link-previous', '链接到前一节')"><Switch :checked="section.linkToPrevious" :disabled="sections.indexOf(sectionId) === 0" @change="update('linkToPrevious', $event)" /></FormItem>
      <FormItem class="header-footer-panel__row" :label="t('doc.header-footer.header', '页眉')"><Input.TextArea class="header-footer-panel__textarea" :value="section.header" :disabled="section.linkToPrevious" :maxlength="500" :auto-size="{ minRows: 1, maxRows: 4 }" @update:value="update('header', String($event ?? ''))" /></FormItem>
      <FormItem class="header-footer-panel__row" :label="t('doc.header-footer.footer', '页脚')"><Input.TextArea class="header-footer-panel__textarea" :value="section.footer" :disabled="section.linkToPrevious" :maxlength="500" :auto-size="{ minRows: 1, maxRows: 4 }" @update:value="update('footer', String($event ?? ''))" /></FormItem>
      <div class="header-footer-panel__fields">
        <span>{{ t('doc.header-footer.fields', '页码字段') }}</span>
        <div class="header-footer-panel__actions">
          <Button size="small" :disabled="section.linkToPrevious" @click="emit('insert-page-field', 'header', 'PAGE')">{{ t('doc.header-footer.insert-page-header', '在页眉插入当前页码') }}</Button>
          <Button size="small" :disabled="section.linkToPrevious" @click="emit('insert-page-field', 'footer', 'PAGE')">{{ t('doc.header-footer.insert-page-footer', '在页脚插入当前页码') }}</Button>
          <Button size="small" :disabled="section.linkToPrevious" @click="emit('insert-page-field', 'footer', 'NUMPAGES')">{{ t('doc.header-footer.insert-total-pages', '在页脚插入总页数') }}</Button>
        </div>
        <span class="header-footer-panel__field-list">{{ (section.pageFields ?? []).map(item => `${item.story}: ${item.field}`).join(' · ') }}</span>
      </div>
      <FormItem class="header-footer-panel__toggle" :label="t('doc.header-footer.different-first', '首页不同')"><Switch :checked="section.differentFirstPage" :disabled="section.linkToPrevious" @change="update('differentFirstPage', $event)" /></FormItem>
      <template v-if="section.differentFirstPage">
        <FormItem class="header-footer-panel__row" :label="t('doc.header-footer.first-header', '首页页眉')"><Input.TextArea class="header-footer-panel__textarea" :value="section.firstPageHeader" :disabled="section.linkToPrevious" :maxlength="500" :auto-size="{ minRows: 1, maxRows: 4 }" @update:value="update('firstPageHeader', String($event ?? ''))" /></FormItem>
        <FormItem class="header-footer-panel__row" :label="t('doc.header-footer.first-footer', '首页页脚')"><Input.TextArea class="header-footer-panel__textarea" :value="section.firstPageFooter" :disabled="section.linkToPrevious" :maxlength="500" :auto-size="{ minRows: 1, maxRows: 4 }" @update:value="update('firstPageFooter', String($event ?? ''))" /></FormItem>
      </template>
      <FormItem class="header-footer-panel__toggle" :label="t('doc.header-footer.different-odd-even', '奇偶页不同')"><Switch :checked="section.differentOddEven" :disabled="section.linkToPrevious" @change="update('differentOddEven', $event)" /></FormItem>
      <template v-if="section.differentOddEven">
        <FormItem class="header-footer-panel__row" :label="t('doc.header-footer.even-header', '偶数页页眉')"><Input.TextArea class="header-footer-panel__textarea" :value="section.evenPageHeader" :disabled="section.linkToPrevious" :maxlength="500" :auto-size="{ minRows: 1, maxRows: 4 }" @update:value="update('evenPageHeader', String($event ?? ''))" /></FormItem>
        <FormItem class="header-footer-panel__row" :label="t('doc.header-footer.even-footer', '偶数页页脚')"><Input.TextArea class="header-footer-panel__textarea" :value="section.evenPageFooter" :disabled="section.linkToPrevious" :maxlength="500" :auto-size="{ minRows: 1, maxRows: 4 }" @update:value="update('evenPageFooter', String($event ?? ''))" /></FormItem>
      </template>
      <FormItem :label="t('doc.header-footer.margin-header', '页眉距顶端')"><InputNumber :value="setup.marginHeader ?? 48" :min="0" :max="200" :precision="1" @change="setDistance('marginHeader', $event)" /></FormItem>
      <FormItem :label="t('doc.header-footer.margin-footer', '页脚距底端')"><InputNumber :value="setup.marginFooter ?? 48" :min="0" :max="200" :precision="1" @change="setDistance('marginFooter', $event)" /></FormItem>
      <Button type="text" :disabled="section.linkToPrevious" @click="emit('apply', { header: '', footer: '', firstPageHeader: '', firstPageFooter: '', evenPageHeader: '', evenPageFooter: '', pageFields: [] })">{{ t('doc.header-footer.clear-section', '清除本节页眉页脚内容') }}</Button>
    </Form>
  </Modal>
</template>

<style scoped>
.header-footer-panel { display: grid; gap: 12px; }
.header-footer-panel :deep(.ant-form-item) { display: grid; grid-template-columns: 140px minmax(0, 1fr); align-items: center; gap: 12px; margin-bottom: 0; font-size: 13px; }
.header-footer-panel :deep(.ant-form-item-label), .header-footer-panel :deep(.ant-form-item-control) { min-width: 0; padding: 0; }
.header-footer-panel :deep(.header-footer-panel__toggle .ant-form-item-control-input-content) { display: flex; justify-content: flex-start; }
.header-footer-panel__toggle :deep(.ant-switch) { justify-self: start; }
.header-footer-panel__textarea { width: 100%; min-width: 0; }
.header-footer-panel :deep(textarea.ant-input) { min-height: 32px; font: inherit; }
@media (max-width: 560px) { .header-footer-panel :deep(.ant-form-item) { grid-template-columns: minmax(0, 1fr); gap: 5px; } }
</style>
