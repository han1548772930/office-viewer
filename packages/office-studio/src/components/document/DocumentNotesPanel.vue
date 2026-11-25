<script setup lang="ts">
import { computed, inject, ref, watch } from 'vue'
import { Alert, Button, Form, FormItem, Input, Modal, Tabs, Tooltip, theme } from 'antdv-next'
import { Trash2 } from 'lucide-vue-next'
import type { DocumentModel } from '../../domain/model'
import type { DocumentNoteKind } from '../../features/document/notes/document-note-types'
import { documentNotePanelEntries } from '../../features/document/controllers/document-note-panel'
import type { DocumentPermissionContext } from '../../features/document/permissions/document-permissions'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

const props = defineProps<{ open: boolean; model: DocumentModel; labels?: Record<string, string>; permission?: DocumentPermissionContext; disabled?: boolean; error?: string }>()
const emit = defineEmits<{ 'update:open': [value: boolean]; apply: [kind: DocumentNoteKind, values: Record<string, string>]; remove: [kind: DocumentNoteKind, id: string]; navigate: [kind: DocumentNoteKind, id: string]; edit: [kind: DocumentNoteKind, id: string]; closed: [] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const { token } = theme.useToken()
const active = ref<DocumentNoteKind>('footnote')
const drafts = ref<Record<string, string>>({}), originals = ref<Record<string, string>>({})
const entries = computed(() => props.open ? documentNotePanelEntries(props.model, props.labels, props.permission, props.disabled) : [])
const notes = computed(() => entries.value.filter(entry => entry.kind === active.value))
const conflicts = computed(() => notes.value.filter(entry => drafts.value[entry.key] !== originals.value[entry.key] && entry.text !== originals.value[entry.key]))
const changedNotes = computed(() => notes.value.filter(entry => entry.editable && entry.textEditable && drafts.value[entry.key] !== entry.text))
watch(() => [props.open, props.model.id] as const, ([open]) => {
  if (!open) return
  const values = Object.fromEntries(entries.value.map(entry => [entry.key, entry.text]))
  drafts.value = { ...values }; originals.value = { ...values }
}, { immediate: true, flush: 'post' })
watch(entries, (current) => {
  for (const entry of current) {
    if (!entry.viewable || drafts.value[entry.key] === undefined || drafts.value[entry.key] === originals.value[entry.key] || drafts.value[entry.key] === entry.text) {
      drafts.value[entry.key] = entry.text; originals.value[entry.key] = entry.text
    }
  }
})
function apply() {
  if (props.disabled || conflicts.value.length || !changedNotes.value.length) return
  emit('apply', active.value, Object.fromEntries(changedNotes.value.map(entry => [entry.note.id, drafts.value[entry.key]!])))
}
</script>

<template>
  <Modal :open="open" :title="t('doc.notes.title', '脚注和尾注')" :width="720" :focus-trigger-after-close="false" :after-close="() => emit('closed')" @cancel="emit('update:open', false)">
    <Tabs v-model:active-key="active" :items="[{ key: 'footnote', label: `${t('doc.notes.footnotes', '脚注')} (${Object.keys(model.footnotes ?? {}).length})` }, { key: 'endnote', label: `${t('doc.notes.endnotes', '尾注')} (${Object.keys(model.endnotes ?? {}).length})` }]" />
    <Alert v-if="error || conflicts.length" type="error" show-icon :message="error || '注释已被其他操作修改，请在画布中核对后再应用。'" />
    <Form :model="drafts" class="notes-list" layout="vertical" @submit.prevent>
      <p v-if="!notes.length" class="notes-empty" :style="{ color: token.colorTextSecondary }">{{ t('doc.notes.empty', '当前文档没有此类注释。') }}</p>
      <section v-for="entry in notes" :key="entry.key" class="note-row" :data-document-note="entry.key">
        <Tooltip :title="entry.referenceAvailable ? '返回引用' : entry.viewable ? '此注释没有可导航的引用' : entry.reason">
          <Button type="text" class="note-number" :disabled="disabled || !entry.referenceAvailable" @click="emit('navigate', active, entry.note.id)">{{ entry.viewable ? entry.label || '空标记' : '受限' }}</Button>
        </Tooltip>
        <FormItem class="note-field" :label="`${active === 'footnote' ? '脚注' : '尾注'} ${entry.label}`">
          <Input.TextArea v-model:value="drafts[entry.key]" :auto-size="{ minRows: 1, maxRows: 4 }" :readonly="!entry.textEditable" :disabled="disabled || !entry.editable" :aria-label="`${active === 'footnote' ? '脚注' : '尾注'} ${entry.label}`" />
          <span v-if="!entry.viewable" class="note-hint" :style="{ color: token.colorTextSecondary }">{{ entry.reason }}</span>
          <span v-else-if="!entry.textEditable" class="note-hint" :style="{ color: token.colorTextSecondary }">此注释包含结构化内容，请在画布中编辑。</span>
        </FormItem>
        <Button type="link" size="small" :disabled="disabled || !entry.viewable" :data-document-note-edit="entry.key" @click="emit('edit', active, entry.note.id)">{{ t('doc.notes.edit-canvas', '在画布中编辑') }}</Button>
        <Tooltip :title="entry.reason || '删除注释及其引用'">
          <Button type="text" danger :disabled="disabled || !entry.removable" :aria-label="t('doc.notes.delete', '删除注释')" @click="emit('remove', active, entry.note.id)"><Trash2 :size="15" /></Button>
        </Tooltip>
      </section>
    </Form>
    <template #footer>
      <Button @click="emit('update:open', false)">{{ t('common.close', '关闭') }}</Button>
      <Button type="primary" :disabled="disabled || !changedNotes.length || !!conflicts.length" @click="apply">{{ t('common.apply', '应用') }}</Button>
    </template>
  </Modal>
</template>

<style scoped>
.notes-list { display: grid; gap: 12px; max-height: 52vh; overflow: auto; padding-top: 8px; }
.note-row { display: grid; grid-template-columns: 56px minmax(0, 1fr) auto 32px; gap: 8px; align-items: center; }
.note-field { min-width: 0; margin-bottom: 0; }
.note-field :deep(.ant-form-item-label) { position: absolute; width: 1px; height: 1px; overflow: hidden; clip-path: inset(50%); }
.note-number { min-width: 44px; max-width: 56px; overflow: hidden; }
.note-hint { font-size: 12px; }
.notes-empty { margin: 8px 0; }
</style>
