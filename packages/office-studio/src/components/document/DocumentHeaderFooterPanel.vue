<script setup lang="ts">
import { computed, inject } from 'vue'
import { Button, Input, Modal, Switch } from 'antdv-next'
import type { DocumentModel } from '../../domain/model'
import { documentSectionIds, resolveSectionHeaderFooter, type SectionHeaderFooterPatch } from '../../features/document/services/section-header-footer'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

const props = defineProps<{ open: boolean; model: DocumentModel; sectionId: string }>()
const emit = defineEmits<{ 'update:open': [value: boolean]; apply: [patch: SectionHeaderFooterPatch]; 'insert-page-field': [story: 'header' | 'footer', field: 'PAGE' | 'NUMPAGES'] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const section = computed(() => resolveSectionHeaderFooter(props.model, props.sectionId))
const sections = computed(() => documentSectionIds(props.model))

function update(key: keyof typeof section.value, value: string | boolean) {
  emit('apply', { [key]: value })
}
</script>

<template>
  <Modal :open="open" :title="t('doc.header-footer.title', '页眉和页脚')" :ok-text="t('common.apply', '应用')" :cancel-text="t('common.close', '关闭')" :width="520" @ok="emit('update:open', false)" @cancel="emit('update:open', false)">
    <div class="header-footer-panel">
      <label class="header-footer-panel__row"><span>{{ t('doc.header-footer.current-section', '当前节') }}</span><Input :value="`${sections.indexOf(sectionId) + 1}`" readonly /></label>
      <label class="header-footer-panel__toggle"><span>{{ t('doc.header-footer.link-previous', '链接到前一节') }}</span><Switch :checked="section.linkToPrevious" :disabled="sections.indexOf(sectionId) === 0" @change="update('linkToPrevious', $event)" /></label>
      <label class="header-footer-panel__row"><span>{{ t('doc.header-footer.header', '页眉') }}</span><Input.TextArea class="header-footer-panel__textarea" :value="section.header" :disabled="section.linkToPrevious" :maxlength="500" :auto-size="{ minRows: 1, maxRows: 4 }" @update:value="update('header', String($event ?? ''))" /></label>
      <label class="header-footer-panel__row"><span>{{ t('doc.header-footer.footer', '页脚') }}</span><Input.TextArea class="header-footer-panel__textarea" :value="section.footer" :disabled="section.linkToPrevious" :maxlength="500" :auto-size="{ minRows: 1, maxRows: 4 }" @update:value="update('footer', String($event ?? ''))" /></label>
      <div class="header-footer-panel__fields">
        <span>{{ t('doc.header-footer.fields', '页码字段') }}</span>
        <div class="header-footer-panel__actions">
          <Button size="small" :disabled="section.linkToPrevious" @click="emit('insert-page-field', 'header', 'PAGE')">{{ t('doc.header-footer.insert-page-header', '在页眉插入当前页码') }}</Button>
          <Button size="small" :disabled="section.linkToPrevious" @click="emit('insert-page-field', 'footer', 'PAGE')">{{ t('doc.header-footer.insert-page-footer', '在页脚插入当前页码') }}</Button>
          <Button size="small" :disabled="section.linkToPrevious" @click="emit('insert-page-field', 'footer', 'NUMPAGES')">{{ t('doc.header-footer.insert-total-pages', '在页脚插入总页数') }}</Button>
        </div>
        <span class="header-footer-panel__field-list">{{ (section.pageFields ?? []).map(item => `${item.story}: ${item.field}`).join(' · ') }}</span>
      </div>
      <label class="header-footer-panel__toggle"><span>{{ t('doc.header-footer.different-first', '首页不同') }}</span><Switch :checked="section.differentFirstPage" :disabled="section.linkToPrevious" @change="update('differentFirstPage', $event)" /></label>
      <template v-if="section.differentFirstPage">
        <label class="header-footer-panel__row"><span>{{ t('doc.header-footer.first-header', '首页页眉') }}</span><Input.TextArea class="header-footer-panel__textarea" :value="section.firstPageHeader" :disabled="section.linkToPrevious" :maxlength="500" :auto-size="{ minRows: 1, maxRows: 4 }" @update:value="update('firstPageHeader', String($event ?? ''))" /></label>
        <label class="header-footer-panel__row"><span>{{ t('doc.header-footer.first-footer', '首页页脚') }}</span><Input.TextArea class="header-footer-panel__textarea" :value="section.firstPageFooter" :disabled="section.linkToPrevious" :maxlength="500" :auto-size="{ minRows: 1, maxRows: 4 }" @update:value="update('firstPageFooter', String($event ?? ''))" /></label>
      </template>
      <label class="header-footer-panel__toggle"><span>{{ t('doc.header-footer.different-odd-even', '奇偶页不同') }}</span><Switch :checked="section.differentOddEven" :disabled="section.linkToPrevious" @change="update('differentOddEven', $event)" /></label>
      <template v-if="section.differentOddEven">
        <label class="header-footer-panel__row"><span>{{ t('doc.header-footer.even-header', '偶数页页眉') }}</span><Input.TextArea class="header-footer-panel__textarea" :value="section.evenPageHeader" :disabled="section.linkToPrevious" :maxlength="500" :auto-size="{ minRows: 1, maxRows: 4 }" @update:value="update('evenPageHeader', String($event ?? ''))" /></label>
        <label class="header-footer-panel__row"><span>{{ t('doc.header-footer.even-footer', '偶数页页脚') }}</span><Input.TextArea class="header-footer-panel__textarea" :value="section.evenPageFooter" :disabled="section.linkToPrevious" :maxlength="500" :auto-size="{ minRows: 1, maxRows: 4 }" @update:value="update('evenPageFooter', String($event ?? ''))" /></label>
      </template>
      <Button type="text" @click="emit('apply', { header: '', footer: '', firstPageHeader: '', firstPageFooter: '', evenPageHeader: '', evenPageFooter: '' })">{{ t('doc.header-footer.clear-section', '清除本节页眉页脚内容') }}</Button>
    </div>
  </Modal>
</template>

<style scoped>
.header-footer-panel { display: grid; gap: 12px; }
.header-footer-panel__row, .header-footer-panel__toggle { display: grid; grid-template-columns: 140px minmax(0, 1fr); align-items: center; gap: 12px; color: #465365; font-size: 13px; }
.header-footer-panel__toggle :deep(.ant-switch) { justify-self: start; }
.header-footer-panel__textarea { width: 100%; min-width: 0; }
.header-footer-panel :deep(textarea.ant-input) { min-height: 32px; font: inherit; }
@media (max-width: 560px) { .header-footer-panel__row, .header-footer-panel__toggle { grid-template-columns: minmax(0, 1fr); gap: 5px; } }
</style>
