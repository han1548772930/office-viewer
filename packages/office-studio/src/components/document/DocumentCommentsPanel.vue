<script setup lang="ts">
import { Button, Input, Select, Tooltip } from 'antdv-next'
import { Check, MessageSquare, Pencil, RotateCcw, Send, Trash2, X } from 'lucide-vue-next'
import { computed, inject, ref, watch } from 'vue'
import type { DocumentCommentMessage, DocumentThreadComment } from '../../domain/model.ts'
import { OfficeLocaleKey, OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

type CommentFilter = 'all' | 'open' | 'resolved'

const props = withDefaults(defineProps<{
  comments: Readonly<Record<string, DocumentThreadComment>>
  currentUserId: string
  canComment: boolean
  canCreate?: boolean
  deletableCommentIds?: readonly string[]
  selectedCommentId?: string
  locale?: string
  error?: string
  busy?: boolean
}>(), { canCreate: false, deletableCommentIds: () => [], selectedCommentId: '', locale: '', error: '', busy: false })

const emit = defineEmits<{
  create: [text: string]
  reply: [commentId: string, text: string]
  edit: [commentId: string, messageId: string, text: string]
  delete: [commentId: string]
  resolve: [commentId: string, resolved: boolean]
  select: [commentId: string]
  close: []
}>()

const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const localeRef = inject(OfficeLocaleKey)
const filter = ref<CommentFilter>('all')
const newComment = ref('')
const replyDrafts = ref<Record<string, string>>({})
const editingMessage = ref<{ commentId: string; messageId: string } | null>(null)
const editDraft = ref('')
const activeLocale = computed(() => props.locale || localeRef?.value || 'en-US')
const filterOptions = computed(() => [
  { value: 'all', label: t('doc.comments.filter.all', 'All comments') },
  { value: 'open', label: t('doc.comments.filter.open', 'Open') },
  { value: 'resolved', label: t('doc.comments.filter.resolved', 'Resolved') },
])
const visibleComments = computed(() => Object.values(props.comments)
  .filter(comment => filter.value === 'all' || comment.resolved === (filter.value === 'resolved'))
  .sort((left, right) => Date.parse(left.createdAt) - Date.parse(right.createdAt)))

watch(() => props.selectedCommentId, id => {
  if (!id) return
  requestAnimationFrame(() => document.getElementById(`document-comment-${id}`)?.scrollIntoView({ block: 'nearest' }))
})

function formatDate(value: string): string {
  const timestamp = Date.parse(value)
  if (!Number.isFinite(timestamp)) return value
  return new Intl.DateTimeFormat(activeLocale.value, { dateStyle: 'medium', timeStyle: 'short' }).format(timestamp)
}

function submitCreate() {
  const text = newComment.value.trim()
  if (!text || !props.canComment || !props.canCreate || props.busy) return
  emit('create', text)
  newComment.value = ''
}

function submitReply(commentId: string) {
  const text = (replyDrafts.value[commentId] ?? '').trim()
  if (!text || !props.canComment || props.busy) return
  emit('reply', commentId, text)
  replyDrafts.value[commentId] = ''
}

function beginEdit(comment: DocumentThreadComment, message: DocumentCommentMessage) {
  if (!props.canComment || message.authorId !== props.currentUserId) return
  editingMessage.value = { commentId: comment.id, messageId: message.id }
  editDraft.value = message.text
}

function submitEdit(comment: DocumentThreadComment, message: DocumentCommentMessage) {
  const text = editDraft.value.trim()
  if (!text || !props.canComment || message.authorId !== props.currentUserId || props.busy) return
  emit('edit', comment.id, message.id, text)
  editingMessage.value = null
  editDraft.value = ''
}
</script>

<template>
  <section class="document-comments" :aria-label="t('doc.comments.title', 'Comments')">
    <header class="document-comments__header">
      <div class="document-comments__heading">
        <MessageSquare :size="16" aria-hidden="true" />
        <h2>{{ t('doc.comments.title', 'Comments') }}</h2>
        <span class="document-comments__count">{{ visibleComments.length }}</span>
      </div>
      <div class="document-comments__header-actions">
        <Select v-model:value="filter" class="document-comments__filter" :options="filterOptions" :aria-label="t('doc.comments.filter.label', 'Filter comments')" />
        <Tooltip :title="t('doc.comments.close', 'Close comments')">
          <Button type="text" :aria-label="t('doc.comments.close', 'Close comments')" @click="emit('close')"><X :size="16" /></Button>
        </Tooltip>
      </div>
    </header>

    <div v-if="error" class="document-comments__error" role="alert">{{ error }}</div>
    <div v-if="visibleComments.length" class="document-comments__threads">
      <article
        v-for="comment in visibleComments"
        :id="`document-comment-${comment.id}`"
        :key="comment.id"
        class="document-comments__thread"
        :class="{ 'is-selected': selectedCommentId === comment.id, 'is-resolved': comment.resolved }"
        @click="emit('select', comment.id)"
      >
        <button class="document-comments__anchor" type="button" :aria-label="t('doc.comments.jump-to-anchor', 'Select comment anchor')" @click.stop="emit('select', comment.id)">
          <span>{{ comment.referenceText || t('doc.comments.anchor-unavailable', 'Text anchor unavailable') }}</span>
        </button>
        <div v-for="message in comment.messages" :key="message.id" class="document-comments__message">
          <div class="document-comments__message-meta">
            <strong>{{ message.authorName }}</strong>
            <time :datetime="message.createdAt">{{ formatDate(message.createdAt) }}</time>
            <Tooltip v-if="canComment && message.authorId === currentUserId && !comment.resolved" :title="t('doc.comments.edit', 'Edit')">
              <Button type="text" size="small" :aria-label="t('doc.comments.edit', 'Edit')" :disabled="busy" @click.stop="beginEdit(comment, message)"><Pencil :size="14" />{{ t('doc.comments.edit', 'Edit') }}</Button>
            </Tooltip>
          </div>
          <template v-if="editingMessage?.commentId === comment.id && editingMessage.messageId === message.id">
            <Input.TextArea v-model:value="editDraft" :aria-label="t('doc.comments.edit', 'Edit comment')" :auto-size="{ minRows: 2, maxRows: 5 }" />
            <div class="document-comments__actions">
              <Button size="small" :disabled="!editDraft.trim() || busy" @click.stop="submitEdit(comment, message)">{{ t('common.save', 'Save') }}</Button>
              <Button size="small" type="text" @click.stop="editingMessage = null">{{ t('common.cancel', 'Cancel') }}</Button>
            </div>
          </template>
          <p v-else class="document-comments__text">{{ message.text }}</p>
        </div>

        <div v-if="!comment.resolved && canComment" class="document-comments__reply">
          <Input.TextArea v-model:value="replyDrafts[comment.id]" :placeholder="t('doc.comments.reply-placeholder', 'Reply')" :aria-label="t('doc.comments.reply-placeholder', 'Reply')" :auto-size="{ minRows: 1, maxRows: 4 }" :disabled="busy" />
          <Button type="text" :aria-label="t('doc.comments.reply', 'Send reply')" :disabled="!replyDrafts[comment.id]?.trim() || busy" @click.stop="submitReply(comment.id)"><Send :size="15" /></Button>
        </div>
        <footer class="document-comments__footer">
          <span v-if="comment.resolved" class="document-comments__resolved"><Check :size="13" />{{ t('doc.comments.status.resolved', 'Resolved') }}</span>
          <span v-else class="document-comments__open">{{ t('doc.comments.status.open', 'Open') }}</span>
          <div class="document-comments__thread-actions">
            <Button v-if="canComment" type="text" size="small" :disabled="busy" @click.stop="emit('resolve', comment.id, !comment.resolved)">
              <RotateCcw v-if="comment.resolved" :size="14" />
              <Check v-else :size="14" />
              {{ comment.resolved ? t('doc.comments.reopen', 'Reopen') : t('doc.comments.resolve', 'Resolve') }}
            </Button>
            <Tooltip v-if="deletableCommentIds.includes(comment.id)" :title="t('doc.comments.delete', 'Delete thread')">
              <Button type="text" size="small" danger :aria-label="t('doc.comments.delete', 'Delete thread')" :disabled="busy" @click.stop="emit('delete', comment.id)"><Trash2 :size="14" /></Button>
            </Tooltip>
          </div>
        </footer>
      </article>
    </div>
    <div v-else class="document-comments__empty">
      <MessageSquare :size="22" aria-hidden="true" />
      <span>{{ filter === 'all' ? t('doc.comments.empty', 'No comments yet') : t('doc.comments.empty-filter', 'No comments match this filter') }}</span>
    </div>

    <form v-if="canComment && canCreate" class="document-comments__composer" @submit.prevent="submitCreate">
      <Input.TextArea v-model:value="newComment" :placeholder="t('doc.comments.new-placeholder', 'Add a comment')" :aria-label="t('doc.comments.new-placeholder', 'Add a comment')" :auto-size="{ minRows: 2, maxRows: 5 }" :disabled="busy" />
      <div class="document-comments__composer-actions">
        <Button type="primary" :loading="busy" :disabled="!newComment.trim()" html-type="submit"><Send :size="14" />{{ t('doc.comments.add', 'Comment') }}</Button>
      </div>
    </form>
  </section>
</template>

<style scoped>
.document-comments { display: flex; min-height: 0; height: 100%; flex-direction: column; color: #202124; background: #fff; }
.document-comments__header { display: flex; min-height: 48px; align-items: center; justify-content: space-between; gap: 10px; padding: 0 14px; border-bottom: 1px solid #e5e7eb; }
.document-comments__heading { display: flex; min-width: 0; align-items: center; gap: 8px; }
.document-comments__header-actions { display: flex; min-width: 0; align-items: center; gap: 4px; }
.document-comments__heading h2 { margin: 0; font-size: 14px; font-weight: 600; }
.document-comments__count { color: #667085; font-size: 12px; }
.document-comments__filter { width: 136px; flex: 0 0 136px; }
.document-comments__error { margin: 10px 14px 0; color: #b42318; font-size: 12px; overflow-wrap: anywhere; }
.document-comments__threads { min-height: 0; flex: 1; overflow: auto; padding: 12px; }
.document-comments__thread { margin-bottom: 10px; padding: 10px; border: 1px solid #e3e6eb; border-radius: 4px; background: #fff; cursor: pointer; }
.document-comments__thread.is-selected { border-color: #3977d5; box-shadow: 0 0 0 1px #3977d5; }
.document-comments__thread.is-resolved { background: #f8fafb; }
.document-comments__anchor { display: block; width: 100%; margin: 0 0 10px; padding: 0 0 8px; border: 0; border-bottom: 1px solid #eaecf0; color: #475467; background: transparent; text-align: left; cursor: pointer; }
.document-comments__anchor span { display: -webkit-box; overflow: hidden; -webkit-box-orient: vertical; -webkit-line-clamp: 2; font-size: 12px; font-style: italic; }
.document-comments__message + .document-comments__message { margin-top: 10px; }
.document-comments__message-meta { display: flex; min-height: 24px; align-items: center; gap: 6px; }
.document-comments__message-meta strong { min-width: 0; overflow: hidden; font-size: 12px; font-weight: 600; text-overflow: ellipsis; white-space: nowrap; }
.document-comments__message-meta time { margin-inline-start: auto; color: #667085; font-size: 10px; white-space: nowrap; }
.document-comments__text { margin: 3px 0 0; color: #344054; font-size: 13px; line-height: 1.5; overflow-wrap: anywhere; white-space: pre-wrap; }
.document-comments__reply { display: flex; align-items: flex-end; gap: 4px; margin-top: 10px; }
.document-comments__reply :deep(.ant-input) { min-width: 0; }
.document-comments__actions, .document-comments__composer-actions, .document-comments__thread-actions { display: flex; align-items: center; gap: 4px; }
.document-comments__actions { margin-top: 6px; }
.document-comments__footer { display: flex; min-height: 28px; align-items: center; justify-content: space-between; gap: 8px; margin-top: 8px; }
.document-comments__resolved, .document-comments__open { display: inline-flex; align-items: center; gap: 4px; color: #667085; font-size: 11px; }
.document-comments__resolved { color: #16804a; }
.document-comments__empty { display: flex; min-height: 120px; flex: 1; flex-direction: column; align-items: center; justify-content: center; gap: 8px; color: #667085; font-size: 12px; }
.document-comments__composer { display: grid; gap: 8px; padding: 12px; border-top: 1px solid #e5e7eb; background: #fff; }
.document-comments__composer-actions { justify-content: flex-end; }
@media (max-width: 520px) { .document-comments__header { padding-inline: 10px; } .document-comments__threads { padding: 8px; } .document-comments__composer { padding: 10px; } .document-comments__filter { width: 120px; flex-basis: 120px; } }
</style>
