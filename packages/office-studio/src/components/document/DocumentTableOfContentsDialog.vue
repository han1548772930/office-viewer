<script setup lang="ts">
import { Checkbox, Modal, Select } from 'antdv-next'
import { computed, inject, ref, watch } from 'vue'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'
import type { DocumentTocOptions, DocumentTocUpdateMode } from '../../features/document/toc/document-toc.ts'

const props = defineProps<{
  open: boolean
  mode: 'insert' | 'update'
  options: DocumentTocOptions
  updateMode?: DocumentTocUpdateMode
}>()

const emit = defineEmits<{
  'update:open': [open: boolean]
  apply: [payload: { options: DocumentTocOptions; updateMode?: DocumentTocUpdateMode }]
}>()

const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const draft = ref<DocumentTocOptions>({ ...props.options })
const draftUpdateMode = ref<DocumentTocUpdateMode>(props.updateMode ?? 'pageNumbersOnly')
const levelOptions = computed(() => Array.from({ length: 9 }, (_, index) => ({ value: index + 1, label: String(index + 1) })))
const formatOptions = computed(() => ([
  ['fromTemplate', 'document.toc.format-template'],
  ['classic', 'document.toc.format-classic'],
  ['modern', 'document.toc.format-modern'],
  ['simple', 'document.toc.format-simple'],
] as const).map(([value, key]) => ({ value, label: t(key, value) })))
const leaderOptions = computed(() => ([
  ['none', 'document.toc.leader-none'],
  ['dots', 'document.toc.leader-dots'],
  ['dashes', 'document.toc.leader-dashes'],
  ['underline', 'document.toc.leader-underline'],
] as const).map(([value, key]) => ({ value, label: t(key, value) })))
const updateOptions = computed(() => ([
  ['pageNumbersOnly', 'document.toc.update-pages'],
  ['all', 'document.toc.update-all'],
] as const).map(([value, key]) => ({ value, label: t(key, value) })))

watch(() => props.open, (open) => {
  if (!open) return
  draft.value = { ...props.options }
  draftUpdateMode.value = props.updateMode ?? 'pageNumbersOnly'
})

function apply() {
  emit('apply', {
    options: { ...draft.value },
    ...(props.mode === 'update' ? { updateMode: draftUpdateMode.value } : {}),
  })
  emit('update:open', false)
}
</script>

<template>
  <Modal
    :open="open"
    :title="t(mode === 'insert' ? 'document.toc.insert' : 'document.toc.update', mode === 'insert' ? 'Table of Contents' : 'Update Table of Contents')"
    :ok-text="t('common.apply', 'Apply')"
    :cancel-text="t('common.cancel', 'Cancel')"
    :width="500"
    destroy-on-hidden
    @ok="apply"
    @cancel="emit('update:open', false)"
  >
    <div class="toc-dialog">
      <div v-if="mode === 'insert'" class="toc-grid">
        <label>
          <span>{{ t('document.toc.levels', 'Show levels') }}</span>
          <Select v-model:value="draft.levels" :options="levelOptions" />
        </label>
        <label>
          <span>{{ t('document.toc.format', 'Format') }}</span>
          <Select v-model:value="draft.format" :options="formatOptions" />
        </label>
      </div>
      <div v-if="mode === 'insert'" class="toc-options">
        <Checkbox v-model:checked="draft.showPageNumbers">
          {{ t('document.toc.show-page-numbers', 'Show page numbers') }}
        </Checkbox>
        <Checkbox v-model:checked="draft.rightAlignPageNumbers" :disabled="!draft.showPageNumbers">
          {{ t('document.toc.right-align-page-numbers', 'Right align page numbers') }}
        </Checkbox>
        <label class="toc-leader">
          <span>{{ t('document.toc.tab-leader', 'Tab leader') }}</span>
          <Select v-model:value="draft.tabLeader" :options="leaderOptions" :disabled="!draft.showPageNumbers || !draft.rightAlignPageNumbers" />
        </label>
      </div>
      <label v-else class="toc-update-mode">
        <span>{{ t('document.toc.update', 'Update Table of Contents') }}</span>
        <Select v-model:value="draftUpdateMode" :options="updateOptions" />
      </label>
      <section v-if="mode === 'insert'" class="toc-preview" :aria-label="t('document.toc.preview', 'Preview')">
        <div v-for="level in Math.min(draft.levels, 3)" :key="level" class="toc-preview-row" :style="{ paddingInlineStart: `${(level - 1) * 18}px` }">
          <span>{{ t('document.toc.preview-heading', 'Heading') }} {{ level }}</span>
          <span v-if="draft.showPageNumbers && draft.rightAlignPageNumbers" class="toc-preview-leader" aria-hidden="true">
            {{ draft.tabLeader === 'dots' ? '··········' : draft.tabLeader === 'dashes' ? '----------' : draft.tabLeader === 'underline' ? '__________' : '' }}
          </span>
          <span v-if="draft.showPageNumbers">{{ level }}</span>
        </div>
      </section>
    </div>
  </Modal>
</template>

<style scoped>
.toc-dialog { display: grid; gap: 18px; padding-top: 8px; }
.toc-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 14px; }
.toc-grid label, .toc-leader, .toc-update-mode { display: grid; gap: 6px; min-width: 0; }
.toc-options { display: grid; gap: 10px; }
.toc-leader { grid-template-columns: minmax(0, 1fr) 190px; align-items: center; }
.toc-update-mode { grid-template-columns: minmax(0, 1fr) 240px; align-items: center; }
.toc-preview { display: grid; gap: 8px; padding: 14px 16px; border: 1px solid #e5e8ed; background: #f7f8fa; }
.toc-preview-row { display: flex; min-width: 0; align-items: baseline; gap: 8px; }
.toc-preview-leader { flex: 1; overflow: hidden; color: #87919d; white-space: nowrap; }
@media (max-width: 520px) {
  .toc-grid { grid-template-columns: minmax(0, 1fr); }
  .toc-leader, .toc-update-mode { grid-template-columns: minmax(0, 1fr); }
}
</style>
