<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { useI18n } from 'vue-i18n'
import { Button, Form, FormItem, Select, Tooltip, theme } from 'antdv-next'
import { DOCUMENT_LATEX_TEMPLATE_CATEGORIES, getDocumentLatexTemplates, type DocumentLatexTemplate, type DocumentLatexTemplateCategory } from '../../features/document/latex/latex-templates'
import DocumentLatexPreview from './DocumentLatexPreview.vue'

const props = defineProps<{ category?: DocumentLatexTemplateCategory; disabled?: boolean }>()
const emit = defineEmits<{ insert: [request: { templateId: string }] }>()
const { t } = useI18n({ useScope: 'global' }), { token } = theme.useToken()
const selected = ref<DocumentLatexTemplateCategory>('fractions')
watch(() => props.category, category => { if (category) selected.value = category }, { immediate: true })
const options = computed(() => DOCUMENT_LATEX_TEMPLATE_CATEGORIES.map(value => ({ value, label: t('document.latex.category.' + value) })))
const templates = computed(() => getDocumentLatexTemplates(props.category ?? selected.value))
const label = (entry: DocumentLatexTemplate) => t('document.latex.template.' + entry.id)
function insert(entry: DocumentLatexTemplate) { if (!props.disabled) emit('insert', { templateId: entry.id }) }
</script>

<template>
  <div class="latex-template-gallery" :aria-label="t('document.latex.templates')" :style="{ '--latex-template-border': token.colorBorder, '--latex-template-muted': token.colorTextSecondary }">
    <Form v-if="!category" layout="vertical" class="latex-template-gallery__category">
      <FormItem :label="t('document.latex.template-category')"><Select v-model:value="selected" :options="options" :disabled="disabled" :aria-label="t('document.latex.template-category')" /></FormItem>
    </Form>
    <div class="latex-template-gallery__grid">
      <Tooltip v-for="entry in templates" :key="entry.id" :title="entry.latex">
        <Button class="latex-template-gallery__item" :data-latex-template="entry.id" :disabled="disabled" :aria-label="t('document.latex.insert-template', { name: label(entry) })" @mousedown.prevent @click="insert(entry)">
          <DocumentLatexPreview :latex="entry.previewLatex" :label="label(entry)" :width="132" :height="58" :font-size="19" />
          <span class="latex-template-gallery__label">{{ label(entry) }}</span>
        </Button>
      </Tooltip>
    </div>
  </div>
</template>
<style scoped>
.latex-template-gallery { width: 440px; max-width: 100%; }
.latex-template-gallery__category :deep(.ant-form-item) { margin-bottom: 8px; }
.latex-template-gallery__grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 6px; max-height: 330px; overflow-y: auto; padding: 2px; }
.latex-template-gallery__grid :deep(.ant-btn) { display: flex; flex-direction: column; align-items: stretch; height: 90px; min-width: 0; padding: 3px 5px; border-color: var(--latex-template-border); }
.latex-template-gallery__label { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-size: 11px; color: var(--latex-template-muted); }
@media (max-width: 640px) { .latex-template-gallery__grid { grid-template-columns: repeat(2, minmax(0, 1fr)); max-height: min(330px, 45dvh); } }
</style>
