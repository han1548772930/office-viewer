<script setup lang="ts">
import { computed, inject } from 'vue'
import { Button, Form, FormItem, Input, InputNumber, Modal, Switch } from 'antdv-next'
import type { DocumentModel, DocumentSectionSetup } from '../../domain/model'
import { DOCUMENT_STORY_KEYS, type DocumentStoryKey } from '../../features/document/model/document-story'
import { documentSectionIds, resolveSectionHeaderFooter, sectionStoryLinkedToPrevious, type SectionHeaderFooterPatch } from '../../features/document/services/section-header-footer'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

const props = defineProps<{ open: boolean; model: DocumentModel; sectionId: string; setup: DocumentSectionSetup; readOnly?: boolean; editableStories?: readonly DocumentStoryKey[] }>()
const emit = defineEmits<{ 'update:open': [value: boolean]; closed: []; position: [patch: Pick<DocumentSectionSetup, 'marginHeader' | 'marginFooter'>]; apply: [patch: SectionHeaderFooterPatch]; 'insert-page-field': [story: DocumentStoryKey, field: 'PAGE' | 'NUMPAGES'] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const section = computed(() => resolveSectionHeaderFooter(props.model, props.sectionId))
const sections = computed(() => documentSectionIds(props.model))

const storyLabels: Record<DocumentStoryKey, { key: string; label: string }> = {
  header: { key: 'header', label: '页眉' }, footer: { key: 'footer', label: '页脚' },
  firstPageHeader: { key: 'first-header', label: '首页页眉' }, firstPageFooter: { key: 'first-footer', label: '首页页脚' },
  evenPageHeader: { key: 'even-header', label: '偶数页页眉' }, evenPageFooter: { key: 'even-footer', label: '偶数页页脚' },
}
const visibleStories = computed(() => DOCUMENT_STORY_KEYS.filter(story => !story.startsWith('first') && !story.startsWith('even')
  || story.startsWith('first') && section.value.differentFirstPage || story.startsWith('even') && section.value.differentOddEven))
const storyLabel = (story: DocumentStoryKey) => t(`doc.header-footer.${storyLabels[story].key}`, storyLabels[story].label)
const linked = (story: DocumentStoryKey) => sectionStoryLinkedToPrevious(section.value, story)
const canEdit = (story: DocumentStoryKey) => !props.readOnly && (props.editableStories === undefined || props.editableStories.includes(story))
const editableUnlinkedStories = computed(() => DOCUMENT_STORY_KEYS.filter(story => canEdit(story) && !linked(story)))

function updateStory(story: DocumentStoryKey, value: unknown) {
  if (canEdit(story) && !linked(story)) emit('apply', { [story]: String(value ?? '') })
}
function updateOption(key: 'differentFirstPage' | 'differentOddEven', value: unknown) {
  if (!props.readOnly && typeof value === 'boolean') emit('apply', { [key]: value })
}
function setLinked(story: DocumentStoryKey, value: unknown) {
  if (canEdit(story) && sections.value.indexOf(props.sectionId) > 0 && typeof value === 'boolean') emit('apply', { storyLinkToPrevious: { [story]: value } })
}
function setDistance(key: 'marginHeader' | 'marginFooter', value: unknown) {
  if (!props.readOnly && typeof value === 'number' && Number.isFinite(value) && value >= 0 && value <= 200 && value !== (props.setup[key] ?? 48)) emit('position', { [key]: value })
}
function clearUnlinkedStories() {
  const stories = editableUnlinkedStories.value
  if (!stories.length) return
  const patch: SectionHeaderFooterPatch = { pageFields: section.value.pageFields?.filter(field => !stories.includes(field.story)) }
  for (const story of stories) patch[story] = ''
  emit('apply', patch)
}
</script>

<template>
  <Modal :open="open" :title="t('doc.header-footer.title', '页眉和页脚')" :ok-text="t('common.apply', '应用')" :cancel-text="t('common.close', '关闭')" :width="560" @ok="emit('update:open', false)" @cancel="emit('update:open', false)" @after-close="emit('closed')">
    <Form layout="horizontal" label-align="start" :colon="false" class="header-footer-panel">
      <FormItem class="header-footer-panel__row" :label="t('doc.header-footer.current-section', '当前节')"><Input :value="`${sections.indexOf(sectionId) + 1}`" readonly /></FormItem>
      <FormItem class="header-footer-panel__toggle" :label="t('doc.header-footer.different-first', '首页不同')"><Switch :checked="section.differentFirstPage" :disabled="readOnly" @change="updateOption('differentFirstPage', $event)" /></FormItem>
      <FormItem class="header-footer-panel__toggle" :label="t('doc.header-footer.different-odd-even', '奇偶页不同')"><Switch :checked="section.differentOddEven" :disabled="readOnly" @change="updateOption('differentOddEven', $event)" /></FormItem>
      <FormItem v-for="story in visibleStories" :key="story" :label="storyLabel(story)">
        <div class="header-footer-panel__story">
          <Input.TextArea :value="section[story]" :disabled="!canEdit(story) || linked(story)" :auto-size="{ minRows: 1, maxRows: 4 }" @update:value="updateStory(story, $event)" />
          <div class="header-footer-panel__actions">
            <Switch :checked="linked(story)" :disabled="!canEdit(story) || sections.indexOf(sectionId) === 0" :aria-label="`${storyLabel(story)}：链接到前一节`" @change="setLinked(story, $event)" />
            <span>{{ t('doc.header-footer.link-previous', '链接到前一节') }}</span>
            <Button size="small" :disabled="!canEdit(story) || linked(story)" @click="emit('insert-page-field', story, 'PAGE')">{{ t('doc.header-footer.page-number', '当前页码') }}</Button>
            <Button size="small" :disabled="!canEdit(story) || linked(story)" @click="emit('insert-page-field', story, 'NUMPAGES')">{{ t('doc.header-footer.total-pages', '总页数') }}</Button>
          </div>
        </div>
      </FormItem>
      <span class="header-footer-panel__field-list">{{ (section.pageFields ?? []).map(item => `${storyLabel(item.story)}：${item.field === 'PAGE' ? '当前页码' : item.field === 'NUMPAGES' ? '总页数' : item.field === 'DATE' ? '日期' : '时间'}`).join(' · ') }}</span>
      <FormItem :label="t('doc.header-footer.margin-header', '页眉距顶端')"><InputNumber :value="setup.marginHeader ?? 48" :min="0" :max="200" :precision="1" :disabled="readOnly" @change="setDistance('marginHeader', $event)" /></FormItem>
      <FormItem :label="t('doc.header-footer.margin-footer', '页脚距底端')"><InputNumber :value="setup.marginFooter ?? 48" :min="0" :max="200" :precision="1" :disabled="readOnly" @change="setDistance('marginFooter', $event)" /></FormItem>
      <Button type="text" :disabled="!editableUnlinkedStories.length" @click="clearUnlinkedStories">{{ t('doc.header-footer.clear-unlinked', '清除本节未关联的页眉页脚内容') }}</Button>
    </Form>
  </Modal>
</template>

<style scoped>
.header-footer-panel { display: grid; gap: 12px; }
.header-footer-panel :deep(.ant-form-item) { margin-bottom: 0; font-size: 13px; }
.header-footer-panel :deep(.ant-form-item-row) {
  display: grid;
  grid-template-columns: 100px minmax(0, 1fr);
  align-items: center;
  gap: 12px;
}
.header-footer-panel :deep(.ant-form-item-label), .header-footer-panel :deep(.ant-form-item-control) { min-width: 0; padding: 0; }
.header-footer-panel :deep(.header-footer-panel__toggle .ant-form-item-control-input-content) { display: flex; justify-content: flex-start; }
.header-footer-panel__story { display: grid; gap: 8px; min-width: 0; }
.header-footer-panel__actions { display: flex; align-items: center; flex-wrap: wrap; gap: 6px; }
.header-footer-panel__field-list { font-size: 12px; color: var(--document-secondary-text, inherit); }
.header-footer-panel :deep(textarea.ant-input) { min-height: 32px; font: inherit; }
@media (max-width: 560px) { .header-footer-panel :deep(.ant-form-item-row) { grid-template-columns: minmax(0, 1fr); gap: 5px; } }
</style>
