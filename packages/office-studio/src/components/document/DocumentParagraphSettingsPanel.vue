<script setup lang="ts">
import { Button, Checkbox, Drawer, InputNumber, Select, Tooltip } from 'antdv-next'
import { computed, inject, ref, watch } from 'vue'
import { AlignCenter, AlignJustify, AlignLeft, AlignRight } from 'lucide-vue-next'
import type { DocumentParagraphStyle } from '../../domain/model.ts'
import { validateDocumentParagraphStylePatch } from '../../features/document/services/document-settings.ts'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

const props = withDefaults(defineProps<{ open: boolean; style: DocumentParagraphStyle; fontSize?: number }>(), { fontSize: 11 })
const emit = defineEmits<{ 'update:open': [open: boolean]; apply: [patch: Partial<DocumentParagraphStyle>] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
function paragraphDraft(style: DocumentParagraphStyle): DocumentParagraphStyle {
  return { indentStart: 0, indentEnd: 0, indentFirstLine: 0, hanging: 0, spacingRule: 'auto', wordWrap: false, ...style }
}
const draft = ref(paragraphDraft(props.style))
const error = ref('')
const alignments = computed(() => ([
  { value: 'left', label: t('doc.paragraph-settings.align.left', 'Left'), icon: AlignLeft },
  { value: 'center', label: t('doc.paragraph-settings.align.center', 'Center'), icon: AlignCenter },
  { value: 'right', label: t('doc.paragraph-settings.align.right', 'Right'), icon: AlignRight },
  { value: 'justify', label: t('doc.paragraph-settings.align.justify', 'Justify'), icon: AlignJustify },
] as const))
const spacingRules = computed(() => (['auto', 'atLeast', 'exact'] as const).map(value => ({ value, label: t(`doc.paragraph-settings.spacing-rule.${value}`, value) })))
const lineSpacingRange = computed(() => (draft.value.spacingRule ?? 'auto') === 'auto' ? { min: 1, max: 5, step: 0.1, precision: 1 } : { min: 1, max: 100, step: 0.1, precision: 1 })
function setSpacingRule(rule: 'auto' | 'atLeast' | 'exact') {
  const currentRule = draft.value.spacingRule ?? 'auto'
  if (currentRule === rule) return
  draft.value.lineHeight = rule === 'auto'
    ? Math.max(1, Math.min(5, Number((draft.value.lineHeight / props.fontSize).toFixed(2))))
    : Number((currentRule === 'auto' ? draft.value.lineHeight * props.fontSize : draft.value.lineHeight).toFixed(2))
  draft.value.spacingRule = rule
}

watch(() => props.open, open => { if (open) { draft.value = paragraphDraft(props.style); error.value = '' } })
watch(() => props.style, style => { if (props.open) draft.value = paragraphDraft(style) }, { deep: true })
watch(draft, value => {
  if (!props.open) return
  const source = paragraphDraft(props.style)
  const defaults: Partial<DocumentParagraphStyle> = { indentStart: 0, indentEnd: 0, indentFirstLine: 0, hanging: 0, spacingRule: 'auto', wordWrap: false }
  const patch = Object.fromEntries(Object.entries(value).filter(([key, field]) => {
    const styleKey = key as keyof DocumentParagraphStyle
    return field !== source[styleKey] && !(props.style[styleKey] === undefined && field === defaults[styleKey])
  })) as Partial<DocumentParagraphStyle>
  if (!Object.keys(patch).length) return
  try {
    validateDocumentParagraphStylePatch(value)
    error.value = ''
    emit('apply', patch)
  } catch (reason) {
    error.value = reason instanceof Error ? reason.message : t('doc.paragraph-settings.invalid', 'Invalid paragraph settings')
  }
}, { deep: true })

</script>

<template>
  <Drawer :open="open" :title="t('doc.paragraph-settings.title', 'Paragraph settings')" :width="300" placement="right" :mask="false" :destroy-on-close="true" @close="emit('update:open', false)">
    <div class="paragraph-settings">
      <section>
        <div class="settings-heading">{{ t('doc.paragraph-settings.alignment', 'Alignment') }}</div>
        <div class="alignment-buttons" role="group" :aria-label="t('doc.paragraph-settings.alignment', 'Alignment')">
          <Tooltip v-for="item in alignments" :key="item.value" :title="item.label" placement="bottom">
            <Button :aria-label="item.label" :aria-pressed="draft.align === item.value" :type="draft.align === item.value ? 'primary' : 'text'" @click="draft.align = item.value">
              <component :is="item.icon" :size="18" aria-hidden="true" />
            </Button>
          </Tooltip>
        </div>
      </section>
      <fieldset>
        <legend>{{ t('doc.paragraph-settings.indentation', 'Indentation') }}</legend>
        <label>{{ t('doc.paragraph-settings.left-indent', 'Left indent') }} (px)<InputNumber v-model:value="draft.indentStart" :min="0" :max="100" :step="0.1" :precision="1" /></label>
        <label>{{ t('doc.paragraph-settings.right-indent', 'Right indent') }} (px)<InputNumber v-model:value="draft.indentEnd" :min="0" :max="100" :step="0.1" :precision="1" /></label>
        <label>{{ t('doc.paragraph-settings.first-line-indent', 'First-line indent') }} (px)<InputNumber v-model:value="draft.indentFirstLine" :min="0" :max="100" :step="0.1" :precision="1" /></label>
        <label>{{ t('doc.paragraph-settings.hanging-indent', 'Hanging indent') }} (px)<InputNumber v-model:value="draft.hanging" :min="0" :max="100" :step="0.1" :precision="1" /></label>
      </fieldset>
      <fieldset>
        <legend>{{ t('doc.paragraph-settings.spacing', 'Spacing') }}</legend>
        <label>{{ t('doc.paragraph-settings.before', 'Before') }} (px)<InputNumber v-model:value="draft.spaceBefore" :min="0" :max="100" :step="0.1" :precision="1" /></label>
        <label>{{ t('doc.paragraph-settings.after', 'After') }} (px)<InputNumber v-model:value="draft.spaceAfter" :min="0" :max="100" :step="0.1" :precision="1" /></label>
        <label>{{ t('doc.paragraph-settings.line-rule', 'Line spacing rule') }}<Select :value="draft.spacingRule ?? 'auto'" :options="spacingRules" @change="setSpacingRule" /></label>
        <label>{{ t('doc.paragraph-settings.line-spacing', 'Line spacing') }}<InputNumber v-model:value="draft.lineHeight" v-bind="lineSpacingRange" /></label>
      </fieldset>
      <Checkbox v-model:checked="draft.wordWrap">{{ t('doc.paragraph-settings.word-wrap', 'Wrap text') }}</Checkbox>
      <p v-if="error" class="settings-error" role="alert">{{ error }}</p>
    </div>
  </Drawer>
</template>

<style scoped>
.paragraph-settings { display: grid; gap: 14px; }
.settings-heading { margin-bottom: 10px; font-size: 13px; font-weight: 500; }
.alignment-buttons { display: grid; grid-template-columns: repeat(4, 1fr); gap: 8px; border: 1px solid #e5e8ed; padding: 4px; }
.alignment-buttons :deep(.ant-btn) { width: 100%; aspect-ratio: 1; padding: 0; }
.paragraph-settings > label, fieldset label { display: grid; grid-template-columns: minmax(0, 1fr) 140px; align-items: center; gap: 10px; min-width: 0; }
fieldset { display: grid; gap: 9px; min-width: 0; margin: 0; padding: 10px 0 0; border: 0; border-top: 1px solid #e5e8ed; }
legend { padding-inline-end: 8px; font-size: 13px; }
fieldset label :deep(.ant-input-number), fieldset label :deep(.ant-select) { width: 100%; }
.settings-error { margin: 0; color: #b42318; overflow-wrap: anywhere; }
</style>
