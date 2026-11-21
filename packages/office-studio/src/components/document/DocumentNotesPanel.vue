<script setup lang="ts">
import { computed, inject, ref, watch } from 'vue'
import { Button, Input, Modal, Tabs } from 'antdv-next'
import type { DocumentNote } from '../../domain/model'
import type { DocumentNoteKind } from '../../features/document/notes/document-notes'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

const props = defineProps<{ open: boolean; footnotes: Record<string, DocumentNote>; endnotes: Record<string, DocumentNote> }>()
const emit = defineEmits<{ 'update:open': [value: boolean]; apply: [kind: DocumentNoteKind, values: Record<string, string>]; remove: [kind: DocumentNoteKind, id: string]; navigate: [kind: DocumentNoteKind, id: string] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const active = ref<DocumentNoteKind>('footnote')
const drafts = ref<Record<string, string>>({})
const notes = computed(() => active.value === 'footnote' ? props.footnotes : props.endnotes)
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
    <div class="notes-list">
      <p v-if="!Object.keys(notes).length" class="notes-empty">{{ t('doc.notes.empty', '当前文档没有此类注释。') }}</p>
      <section v-for="(note, index) in Object.values(notes)" :key="note.id" class="note-row">
        <Button type="text" class="note-number" @click="emit('navigate', active, note.id)">{{ index + 1 }}</Button>
        <Input v-model:value="drafts[note.id]" :aria-label="`${active} ${index + 1}`" />
        <Button type="text" danger :aria-label="t('doc.notes.delete', '删除注释')" @click="emit('remove', active, note.id)">×</Button>
      </section>
    </div>
  </Modal>
</template>

<style scoped>
.notes-list { display: grid; gap: 8px; max-height: 52vh; overflow: auto; }
.note-row { display: grid; grid-template-columns: 40px minmax(0, 1fr) 32px; gap: 8px; align-items: center; }
.note-number { width: 36px; }
.notes-empty { color: #697586; }
</style>
