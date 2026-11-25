<script setup lang="ts">
import { Alert, Button, Card, Checkbox, Form, FormItem, InputNumber, Select, Tooltip, theme } from 'antdv-next'
import { computed, inject, ref, watch } from 'vue'
import { AlignCenter, AlignJustify, AlignLeft, AlignRight, X } from 'lucide-vue-next'
import type { DocumentParagraphStyle } from '../../domain/model'
import { validateDocumentParagraphStylePatch } from '../../features/document/services/document-settings'
import { OfficeTranslatorKey } from '../../i18n/office-i18n'
import {
  changeDocumentParagraphSpacingRule, documentParagraphLineSpacingInput, documentParagraphSettingsStyle,
  isDocumentParagraphSpacingRule, type DocumentParagraphLineMetrics, type DocumentParagraphNumericField,
  type DocumentParagraphSettingField,
} from './document-paragraph-settings-state'

const props = withDefaults(defineProps<{
  open: boolean
  style: Partial<DocumentParagraphStyle>
  mixed?: readonly DocumentParagraphSettingField[]
  fontSize?: number
  lineMetrics?: DocumentParagraphLineMetrics
  selectionKey?: string
  readOnly?: boolean
}>(), { fontSize: 11, selectionKey: '', readOnly: false })
const emit = defineEmits<{ 'update:open': [open: boolean]; apply: [patch: Partial<DocumentParagraphStyle>] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const { token } = theme.useToken()
const draft = ref(documentParagraphSettingsStyle(props.style))
const mixedFields = ref<readonly DocumentParagraphSettingField[]>(props.mixed ?? [])
const error = ref('')
const composing = ref(false)
const spacingMenuOpen = ref(false)
let compositionPatch: Partial<DocumentParagraphStyle> | undefined

const panelStyle = computed(() => ({
  '--paragraph-border': token.value.colorBorderSecondary,
  '--paragraph-background': token.value.colorBgContainer,
  '--paragraph-muted': token.value.colorTextSecondary,
}))
const mixedLabel = computed(() => t('doc.paragraph-settings.mixed', 'Mixed'))
const alignments = computed(() => ([
  { value: 'left', label: t('doc.paragraph-settings.align.left', 'Left'), icon: AlignLeft },
  { value: 'center', label: t('doc.paragraph-settings.align.center', 'Center'), icon: AlignCenter },
  { value: 'right', label: t('doc.paragraph-settings.align.right', 'Right'), icon: AlignRight },
  { value: 'justify', label: t('doc.paragraph-settings.align.justify', 'Justify'), icon: AlignJustify },
] as const))
const indentationFields = [
  { field: 'indentStart', key: 'left-indent', label: 'Left indent' },
  { field: 'indentEnd', key: 'right-indent', label: 'Right indent' },
  { field: 'indentFirstLine', key: 'first-line-indent', label: 'First-line indent' },
  { field: 'hanging', key: 'hanging-indent', label: 'Hanging indent' },
] as const
const spacingFields = [
  { field: 'spaceBefore', key: 'before', label: 'Before' },
  { field: 'spaceAfter', key: 'after', label: 'After' },
] as const
const paginationFields = [
  { field: 'keepLines', key: 'keep-lines', label: 'Keep lines together' },
  { field: 'keepNext', key: 'keep-next', label: 'Keep with next paragraph' },
  { field: 'widowControl', key: 'widow-control', label: 'Widow and orphan control' },
  { field: 'pageBreakBefore', key: 'page-break-before', label: 'Page break before' },
] as const
const spacingRules = computed(() => (['auto', 'atLeast', 'exact'] as const).map(value => ({ value, label: t('doc.paragraph-settings.spacing-rule.' + value, value) })))
const lineSpacingRange = computed(() => documentParagraphLineSpacingInput(draft.value.spacingRule))
const lineSpacingLabel = computed(() => t('doc.paragraph-settings.line-spacing', 'Line spacing') + (isMixed('spacingRule') ? '' : draft.value.spacingRule === 'auto' ? ' (×)' : ' (px)'))

function isMixed(field: DocumentParagraphSettingField) { return mixedFields.value.includes(field) }
function numericValue(field: DocumentParagraphNumericField) { return isMixed(field) ? null : draft.value[field] }
function numberLabel(item: { key: string; label: string }) { return t('doc.paragraph-settings.' + item.key, item.label) + ' (px)' }

watch([() => props.open, () => props.style, () => props.mixed, () => props.selectionKey, () => props.readOnly], () => {
  draft.value = documentParagraphSettingsStyle(props.style)
  mixedFields.value = props.mixed ?? []
  error.value = ''
  compositionPatch = undefined
  composing.value = false
}, { deep: true })

/** Each control submits only its own changed fields; receiving state never dispatches an edit. */
function applyPatch(patch: Partial<DocumentParagraphStyle>) {
  if (!props.open || props.readOnly) return
  if (composing.value) { compositionPatch = { ...compositionPatch, ...patch }; return }
  try {
    validateDocumentParagraphStylePatch(patch, draft.value.spacingRule)
    const changed: Partial<DocumentParagraphStyle> = {}
    for (const [key, value] of Object.entries(patch)) {
      const field = key as DocumentParagraphSettingField
      if (isMixed(field) || value !== draft.value[field]) Object.assign(changed, { [field]: value })
    }
    error.value = ''
    if (!Object.keys(changed).length) return
    Object.assign(draft.value, changed)
    mixedFields.value = mixedFields.value.filter(field => !Object.hasOwn(changed, field))
    emit('apply', changed)
  } catch (reason) {
    error.value = reason instanceof Error ? reason.message : t('doc.paragraph-settings.invalid', 'Invalid paragraph settings')
  }
}

function setNumber(field: DocumentParagraphNumericField, value: unknown) {
  if (value === null || value === undefined || value === '') return
  const number = typeof value === 'number' ? value : typeof value === 'string' ? Number(value) : Number.NaN
  applyPatch({ [field]: number })
}
function setSpacingRule(value: unknown) {
  if (!isDocumentParagraphSpacingRule(value) || (!isMixed('spacingRule') && value === draft.value.spacingRule)) return
  const lineHeight = changeDocumentParagraphSpacingRule(draft.value.lineHeight, draft.value.spacingRule, value, props.lineMetrics, props.fontSize)
  applyPatch({ spacingRule: value, lineHeight })
}
function finishComposition() {
  composing.value = false
  const patch = compositionPatch
  compositionPatch = undefined
  if (patch) applyPatch(patch)
}
function closePanel() { compositionPatch = undefined; emit('update:open', false) }
function onPanelKeydown(event: KeyboardEvent) {
  if (event.key !== 'Escape' || event.isComposing || event.keyCode === 229 || composing.value || spacingMenuOpen.value || event.defaultPrevented) return
  event.preventDefault()
  event.stopPropagation()
  closePanel()
}
</script>

<template>
  <Card v-if="open" class="paragraph-settings-panel" data-document-paragraph-settings :title="t('doc.paragraph-settings.title', 'Paragraph settings')" size="small" variant="borderless" :style="panelStyle"
    @compositionstart="composing = true" @compositionend="finishComposition" @keydown="onPanelKeydown">
    <template #extra>
      <Tooltip :title="t('doc.paragraph-settings.close', 'Close paragraph settings')">
        <Button type="text" size="small" :aria-label="t('doc.paragraph-settings.close', 'Close paragraph settings')" @click="closePanel"><X :size="16" /></Button>
      </Tooltip>
    </template>
    <Form :model="draft" :disabled="readOnly" layout="horizontal" class="paragraph-settings" @submit.prevent>
      <Alert v-if="readOnly" type="info" show-icon :message="t('doc.paragraph-settings.read-only', 'The selection is read-only')" />
      <section>
        <div class="settings-heading">{{ t('doc.paragraph-settings.alignment', 'Alignment') }}<span v-if="isMixed('align')"> · {{ mixedLabel }}</span></div>
        <div class="alignment-buttons" role="group" :aria-label="t('doc.paragraph-settings.alignment', 'Alignment')">
          <Tooltip v-for="item in alignments" :key="item.value" :title="item.label" placement="bottom">
            <Button :disabled="readOnly" :aria-label="item.label" :aria-pressed="!isMixed('align') && draft.align === item.value" :type="!isMixed('align') && draft.align === item.value ? 'primary' : 'text'" @click="applyPatch({ align: item.value })">
              <component :is="item.icon" :size="18" aria-hidden="true" />
            </Button>
          </Tooltip>
        </div>
      </section>
      <fieldset>
        <legend>{{ t('doc.paragraph-settings.indentation', 'Indentation') }}</legend>
        <FormItem v-for="item in indentationFields" :key="item.field" :name="item.field" :label="numberLabel(item)">
          <InputNumber :value="numericValue(item.field)" :aria-label="numberLabel(item)" :placeholder="isMixed(item.field) ? mixedLabel : undefined" :min="0" :max="100" :step="0.1" :precision="1" @change="value => setNumber(item.field, value)" />
        </FormItem>
      </fieldset>
      <fieldset>
        <legend>{{ t('doc.paragraph-settings.spacing', 'Spacing') }}</legend>
        <FormItem v-for="item in spacingFields" :key="item.field" :name="item.field" :label="numberLabel(item)">
          <InputNumber :value="numericValue(item.field)" :aria-label="numberLabel(item)" :placeholder="isMixed(item.field) ? mixedLabel : undefined" :min="0" :max="100" :step="0.1" :precision="1" @change="value => setNumber(item.field, value)" />
        </FormItem>
        <FormItem name="spacingRule" :label="t('doc.paragraph-settings.line-rule', 'Line spacing rule')">
          <Select :value="isMixed('spacingRule') ? undefined : draft.spacingRule" :aria-label="t('doc.paragraph-settings.line-rule', 'Line spacing rule')" :placeholder="mixedLabel" :options="spacingRules" @change="setSpacingRule" @open-change="spacingMenuOpen = $event" />
        </FormItem>
        <FormItem name="lineHeight" :label="lineSpacingLabel" :extra="isMixed('spacingRule') ? t('doc.paragraph-settings.choose-line-rule', 'Choose a line spacing rule first') : undefined">
          <InputNumber :value="numericValue('lineHeight')" :disabled="readOnly || isMixed('spacingRule')" :aria-label="lineSpacingLabel" :placeholder="isMixed('lineHeight') ? mixedLabel : undefined" v-bind="lineSpacingRange" @change="value => setNumber('lineHeight', value)" />
        </FormItem>
      </fieldset>
      <FormItem name="wordWrap" class="paragraph-wrap-item">
        <Checkbox :checked="draft.wordWrap" :indeterminate="isMixed('wordWrap')" :aria-checked="isMixed('wordWrap') ? 'mixed' : draft.wordWrap" @update:checked="checked => applyPatch({ wordWrap: checked })">{{ t('doc.paragraph-settings.word-wrap', 'Allow Latin text to wrap inside words') }}</Checkbox>
      </FormItem>
      <fieldset>
        <legend>{{ t('doc.paragraph-settings.pagination', 'Pagination') }}</legend>
        <FormItem v-for="item in paginationFields" :key="item.field" :name="item.field" class="paragraph-wrap-item">
          <Checkbox :checked="draft[item.field]" :indeterminate="isMixed(item.field)" :aria-checked="isMixed(item.field) ? 'mixed' : draft[item.field]"
            :data-document-pagination-field="item.field" @update:checked="checked => applyPatch({ [item.field]: checked })">{{ t('doc.paragraph-settings.' + item.key, item.label) }}</Checkbox>
        </FormItem>
        <div class="pagination-description">{{ t('doc.paragraph-settings.pagination-description', 'Applied in traditional paginated documents. Values are retained in other modes.') }}</div>
      </fieldset>
      <Alert v-if="error" type="error" show-icon :message="error" role="alert" />
    </Form>
  </Card>
</template>

<style scoped>
.paragraph-settings-panel { height: 100%; min-height: 0; display: flex; flex-direction: column; border-radius: 0; background: var(--paragraph-background); }
.paragraph-settings-panel :deep(.ant-card-head) { flex: 0 0 auto; }
.paragraph-settings-panel :deep(.ant-card-body) { flex: 1; min-height: 0; overflow: auto; }
.paragraph-settings { display: grid; gap: 14px; }
.paragraph-settings :deep(.ant-form-item) { margin-bottom: 0; }
.paragraph-settings :deep(.ant-form-item-row) { display: grid; grid-template-columns: minmax(0, 1fr) 116px; align-items: center; gap: 10px; }
.paragraph-settings :deep(.ant-form-item-label) { text-align: start; white-space: normal; }
.paragraph-settings :deep(.paragraph-wrap-item .ant-form-item-row) { display: block; }
.settings-heading { margin-bottom: 10px; font-size: 13px; font-weight: 500; }
.pagination-description { color: var(--paragraph-muted); font-size: 12px; }
.alignment-buttons { display: grid; grid-template-columns: repeat(4, 1fr); gap: 8px; border: 1px solid var(--paragraph-border); padding: 4px; }
.alignment-buttons :deep(.ant-btn) { width: 100%; height: 32px; padding: 0; }
fieldset { display: grid; gap: 9px; min-width: 0; margin: 0; padding: 10px 0 0; border: 0; border-top: 1px solid var(--paragraph-border); }
legend { padding-inline-end: 8px; font-size: 13px; }
.paragraph-settings :deep(.ant-input-number), .paragraph-settings :deep(.ant-select) { width: 100%; }
</style>
