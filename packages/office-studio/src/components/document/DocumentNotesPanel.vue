<script setup lang="ts">
import { computed, inject, ref, watch } from 'vue'
import { Button, Form, FormItem, Input, Modal, Tabs, theme } from 'antdv-next'
import { Trash2 } from 'lucide-vue-next'
import type { DocumentNote } from '../../domain/model'
import type { DocumentNoteKind } from '../../features/document/notes/document-notes'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

const props = defineProps<{ open: boolean; footnotes: Record<string, DocumentNote>; endnotes: Record<string, DocumentNote> }>()
const emit = defineEmits<{ 'update:open': [value: boolean]; apply: [kind: DocumentNoteKind, values: Record<string, string>]; remove: [kind: DocumentNoteKind, id: string]; navigate: [kind: DocumentNoteKind, id: string]; edit: [kind: DocumentNoteKind, id: string] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const { token } = theme.useToken()
const active = ref<DocumentNoteKind>('footnote')
const drafts = ref<Record<string, string>>({})
const notes = computed(() => active.value === 'footnote' ? props.footnotes : props.endnotes)
function hasStructuredContent(note: DocumentNote) {
  return !!note.content?.blocks?.some(block => !['paragraph', 'heading', 'quote', 'bullet', 'numbered', 'checklist'].includes(block.type)
    || block.metadata.table || block.metadata.image || block.metadata.codeBlock)
}
watch(() => [props.open, props.footnotes, props.endnotes], () => {
  drafts.value = Object.fromEntries([...Object.values(props.footnotes), ...Object.values(props.endnotes)].map(note => [note.id, note.text]))
}, { immediate: true })
function apply() {
  emit('apply', active.value, drafts.value)
  emit('update:open', false)
}
</script>

<template>
  <Modal :open="open" :title="t('doc.notes.title', '脚注和尾注')" :ok-text="t('common.apply', '应用')" :cancel-text="t('common.close', '关闭')" :width="620" @ok="apply" @cancel="emit('update:open', false)">
    <Tabs v-model:active-key="active" :items="[{ key: 'footnote', label: `${t('doc.notes.footnotes', '脚注')} (${Object.keys(footnotes).length})` }, { key: 'endnote', label: `${t('doc.notes.endnotes', '尾注')} (${Object.keys(endnotes).length})` }]" />
    <Form class="notes-list" layout="vertical" @submit.prevent>
      <p v-if="!Object.keys(notes).length" class="notes-empty" :style="{ color: token.colorTextSecondary }">{{ t('doc.notes.empty', '当前文档没有此类注释。') }}</p>
      <section v-for="(note, index) in Object.values(notes)" :key="note.id" class="note-row">
        <Button type="text" class="note-number" @click="emit('navigate', active, note.id)">{{ index + 1 }}</Button>
        <FormItem class="note-field" :label="`${active} ${index + 1}`">
          <Input v-model:value="drafts[note.id]" :readonly="hasStructuredContent(note)" :aria-label="`${active} ${index + 1}`" />
        </FormItem>
        <Button type="link" size="small" :data-document-note-edit="note.id" @click="emit('edit', active, note.id)">{{ t('doc.notes.edit-canvas', '在画布中编辑') }}</Button>
        <Button type="text" danger :aria-label="t('doc.notes.delete', '删除注释')" @click="emit('remove', active, note.id)"><Trash2 :size="15" /></Button>
      </section>
    </Form>
  </Modal>
</template>

<style scoped>
.notes-list { display: grid; gap: 8px; max-height: 52vh; overflow: auto; }
.note-row { display: grid; grid-template-columns: 40px minmax(0, 1fr) auto 32px; gap: 8px; align-items: center; }
.note-field { min-width: 0; margin-bottom: 0; }
.note-field :deep(.ant-form-item-label) { position: absolute; width: 1px; height: 1px; overflow: hidden; clip-path: inset(50%); }
.note-number { width: 36px; }
.notes-empty { margin: 8px 0; }
</style>
