<script setup lang="ts">
import { Button, Dropdown, Input, Select, Tooltip } from 'antdv-next'
import { CircleCheck, MoreHorizontal, Plus, Reply, RotateCcw, Trash2, X } from 'lucide-vue-next'
import { computed, inject, nextTick, ref, watch } from 'vue'
import type { DocumentCommentMessage, DocumentThreadComment } from '../../domain/model.ts'
import { OfficeLocaleKey, OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

type CommentFilter = 'all' | 'resolved' | 'open' | 'concern-me'

const props = withDefaults(defineProps<{
  comments: Readonly<Record<string, DocumentThreadComment>>
  currentUserId: string
  canComment: boolean
  canCreate?: boolean
  canStartCreate?: boolean
  deletableCommentIds?: readonly string[]
  selectedCommentId?: string
  locale?: string
  error?: string
  busy?: boolean
}>(), { canCreate: false, canStartCreate: false, deletableCommentIds: () => [], selectedCommentId: '', locale: '', error: '', busy: false })

const emit = defineEmits<{
  create: [text: string]
  reply: [commentId: string, text: string, mentions: readonly { userId: string; label: string }[]]
  edit: [commentId: string, messageId: string, text: string]
  delete: [commentId: string]
  'delete-message': [commentId: string, messageId: string]
  resolve: [commentId: string, resolved: boolean]
  select: [commentId: string]
  close: []
  add: []
}>()

const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const localeRef = inject(OfficeLocaleKey)
const filter = ref<CommentFilter>('all')
const newComment = ref('')
const replyDrafts = ref<Record<string, string>>({})
const replyingTo = ref('')
const replyMentions = ref<Record<string, { userId: string; label: string }[]>>({})
const editingMessage = ref<{ commentId: string; messageId: string } | null>(null)
const editDraft = ref('')
const activeLocale = computed(() => props.locale || localeRef?.value || 'en-US')
const filterOptions = computed(() => [
  { value: 'all', label: t('doc.comments.filter.all', 'All comments') },
  { value: 'resolved', label: t('doc.comments.filter.resolved', 'Resolved') },
  { value: 'open', label: t('doc.comments.filter.open', 'Not resolved') },
  { value: 'concern-me', label: t('doc.comments.filter.concern-me', 'Concern me') },
])
const filteredComments = computed(() => Object.values(props.comments)
  .filter(comment => filter.value === 'all'
    || (filter.value === 'concern-me' && (!props.currentUserId || comment.messages.some(message => message.authorId === props.currentUserId)))
    || (filter.value !== 'concern-me' && comment.resolved === (filter.value === 'resolved')))
)
const openComments = computed(() => filteredComments.value.filter(comment => !comment.resolved))
const resolvedComments = computed(() => filteredComments.value.filter(comment => comment.resolved))
const visibleComments = computed(() => [...openComments.value, ...resolvedComments.value])

watch(() => props.selectedCommentId, id => {
  if (!id) return
  requestAnimationFrame(() => document.getElementById(`document-comment-${id}`)?.scrollIntoView({ block: 'nearest' }))
})

function formatDate(value: string): string {
  const timestamp = Date.parse(value)
  if (!Number.isFinite(timestamp)) return value
  return new Intl.DateTimeFormat(activeLocale.value, { dateStyle: 'medium', timeStyle: 'short' }).format(timestamp)
}

function commentMention(message: DocumentCommentMessage) {
  const mention = message.mentions?.[0]
  return mention && message.text.startsWith(mention.label) ? mention : undefined
}

function commentBody(message: DocumentCommentMessage): string {
  const mention = commentMention(message)
  return mention ? message.text.slice(mention.label.length) : message.text
}

function submitCreate() {
  const text = newComment.value.trim()
  if (!text || !props.canComment || !props.canCreate || props.busy) return
  emit('create', text)
  newComment.value = ''
}

function resetFilter() {
  filter.value = 'all'
}

async function beginReply(commentId: string, message: DocumentCommentMessage) {
  if (!props.canComment || props.busy) return
  replyingTo.value = commentId
  const label = `@${message.authorName}`
  replyDrafts.value[commentId] = `${label} `
  replyMentions.value[commentId] = message.authorId ? [{ userId: message.authorId, label }] : []
  await nextTick()
  document.getElementById(`document-comment-reply-${commentId}`)?.focus()
}

function submitReply(commentId: string) {
  const text = (replyDrafts.value[commentId] ?? '').trim()
  if (!text || !props.canComment || props.busy) return
  emit('reply', commentId, text, replyMentions.value[commentId] ?? [])
  replyDrafts.value[commentId] = ''
  replyMentions.value[commentId] = []
  replyingTo.value = ''
}

function messageMenu(comment: DocumentThreadComment, message: DocumentCommentMessage) {
  return {
    items: [
      { key: 'edit', label: t('doc.comments.edit', 'Edit') },
      { key: 'delete', label: t('doc.comments.delete-message', 'Delete comment') },
    ],
    onClick: ({ key }: { key: string | number }) => {
      if (key === 'edit') beginEdit(comment, message)
      if (key === 'delete') emit('delete-message', comment.id, message.id)
    },
  }
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
        <h2>{{ t('doc.comments.title', 'Comments') }}</h2>
      </div>
      <div class="document-comments__header-actions">
        <Select v-model:value="filter" class="document-comments__filter" :options="filterOptions" :aria-label="t('doc.comments.filter.label', 'Filter comments')" />
        <Tooltip :title="t('doc.comments.close', 'Close comments')">
          <Button type="text" :aria-label="t('doc.comments.close', 'Close comments')" @click="emit('close')"><X :size="16" aria-hidden="true" /></Button>
        </Tooltip>
      </div>
    </header>

    <div v-if="error" class="document-comments__error" role="alert">{{ error }}</div>
    <div v-if="visibleComments.length" class="document-comments__threads">
      <template v-for="(comment, index) in visibleComments" :key="comment.id">
      <div v-if="comment.resolved && index === openComments.length" class="document-comments__solved-label">{{ t('doc.comments.status.resolved', 'Resolved') }}</div>
      <article
        :id="`document-comment-${comment.id}`"
        class="document-comments__thread"
        :class="{ 'is-selected': selectedCommentId === comment.id, 'is-resolved': comment.resolved }"
        @click="emit('select', comment.id)"
      >
        <div class="document-comments__thread-header">
          <button class="document-comments__anchor" type="button" :aria-label="t('doc.comments.jump-to-anchor', 'Select comment anchor')" @click.stop="emit('select', comment.id)">
            <span>{{ comment.referenceText || t('doc.comments.anchor-unavailable', 'Text anchor unavailable') }}</span>
          </button>
          <div class="document-comments__thread-actions">
            <Tooltip v-if="canComment" :title="t(comment.resolved ? 'doc.comments.reopen' : 'doc.comments.resolve', comment.resolved ? 'Reopen' : 'Resolve')">
              <Button type="text" size="small" :aria-label="t(comment.resolved ? 'doc.comments.reopen' : 'doc.comments.resolve', comment.resolved ? 'Reopen' : 'Resolve')" :disabled="busy" @click.stop="emit('resolve', comment.id, !comment.resolved)"><component :is="comment.resolved ? RotateCcw : CircleCheck" :size="16" aria-hidden="true" /></Button>
            </Tooltip>
            <Tooltip v-if="deletableCommentIds.includes(comment.id)" :title="t('doc.comments.delete', 'Delete thread')">
              <Button type="text" size="small" danger :aria-label="t('doc.comments.delete', 'Delete thread')" :disabled="busy" @click.stop="emit('delete', comment.id)"><Trash2 :size="16" aria-hidden="true" /></Button>
            </Tooltip>
          </div>
        </div>
        <div v-for="message in comment.messages" :key="message.id" class="document-comments__message">
          <div class="document-comments__message-meta">
            <strong>{{ message.authorName }}</strong>
            <time :datetime="message.createdAt">{{ formatDate(message.createdAt) }}</time>
            <Tooltip v-if="canComment && !comment.resolved" :title="t('doc.comments.reply', 'Reply')">
              <Button class="document-comments__reply-trigger" type="text" size="small" :aria-label="t('doc.comments.reply', 'Reply')" :disabled="busy" @click.stop="beginReply(comment.id, message)"><Reply :size="16" aria-hidden="true" /></Button>
            </Tooltip>
            <Dropdown v-if="canComment && message.authorId === currentUserId && !comment.resolved" :trigger="['click']" placement="bottomRight" :menu="messageMenu(comment, message)">
              <Button class="document-comments__more-trigger" type="text" size="small" :aria-label="t('doc.comments.more', 'More')" @click.stop><MoreHorizontal :size="16" aria-hidden="true" /></Button>
            </Dropdown>
          </div>
          <template v-if="editingMessage?.commentId === comment.id && editingMessage.messageId === message.id">
            <Input.TextArea v-model:value="editDraft" :aria-label="t('doc.comments.edit', 'Edit comment')" :auto-size="{ minRows: 2, maxRows: 5 }" />
            <div class="document-comments__actions">
              <Button size="small" :disabled="!editDraft.trim() || busy" @click.stop="submitEdit(comment, message)">{{ t('common.save', 'Save') }}</Button>
              <Button size="small" type="text" @click.stop="editingMessage = null; editDraft = ''">{{ t('common.cancel', 'Cancel') }}</Button>
            </div>
          </template>
          <p v-else class="document-comments__text"><span v-if="commentMention(message)" class="document-comments__mention">{{ commentMention(message)?.label }}</span>{{ commentBody(message) }}</p>
        </div>

        <div v-if="canComment && replyingTo === comment.id" class="document-comments__reply">
          <Input.TextArea :id="`document-comment-reply-${comment.id}`" v-model:value="replyDrafts[comment.id]" :placeholder="t('doc.comments.reply-placeholder', 'Reply')" :aria-label="t('doc.comments.reply-placeholder', 'Reply')" :auto-size="{ minRows: 1, maxRows: 4 }" :disabled="busy" @keydown.ctrl.enter.prevent="submitReply(comment.id)" @keydown.meta.enter.prevent="submitReply(comment.id)" />
          <div class="document-comments__actions">
            <Button size="small" :disabled="!replyDrafts[comment.id]?.trim() || busy" @click.stop="submitReply(comment.id)">{{ t('doc.comments.reply-submit', 'Comment') }}</Button>
            <Button size="small" type="text" @click.stop="replyingTo = ''; replyDrafts[comment.id] = ''; replyMentions[comment.id] = []">{{ t('common.cancel', 'Cancel') }}</Button>
          </div>
        </div>
      </article>
      </template>
    </div>
    <div v-else class="document-comments__empty">
      <span>{{ t('doc.comments.empty', 'No comments yet') }}</span>
      <Button v-if="filter !== 'all'" size="small" @click="resetFilter">{{ t('doc.comments.filter.reset', 'Show all comments') }}</Button>
      <Button v-if="filter === 'all' && canStartCreate" size="small" @click="emit('add')"><Plus :size="16" aria-hidden="true" />{{ t('doc.comments.panel-add', 'Add Comment') }}</Button>
    </div>

    <form v-if="canComment && canCreate" class="document-comments__composer" @submit.prevent="submitCreate">
      <Input.TextArea v-model:value="newComment" :placeholder="t('doc.comments.new-placeholder', 'Add a comment')" :aria-label="t('doc.comments.new-placeholder', 'Add a comment')" :auto-size="{ minRows: 2, maxRows: 5 }" :disabled="busy" />
      <div class="document-comments__composer-actions">
        <Button type="primary" :loading="busy" :disabled="!newComment.trim()" html-type="submit">{{ t('doc.comments.add', 'Comment') }}</Button>
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
.document-comments__filter { width: 136px; flex: 0 0 136px; }
.document-comments__error { margin: 10px 14px 0; color: #b42318; font-size: 12px; overflow-wrap: anywhere; }
.document-comments__threads { min-height: 0; flex: 1; overflow: auto; padding: 12px; }
.document-comments__thread { margin-bottom: 10px; padding: 10px; border: 1px solid #e3e6eb; border-radius: 4px; background: #fff; cursor: pointer; }
.document-comments__thread.is-selected { border-color: #3977d5; box-shadow: 0 0 0 1px #3977d5; }
.document-comments__thread.is-resolved { background: #f8fafb; }
.document-comments__thread-header { display: flex; min-width: 0; align-items: center; justify-content: space-between; gap: 8px; margin-bottom: 10px; }
.document-comments__anchor { display: block; min-width: 0; flex: 1; padding: 0 0 8px; border: 0; border-bottom: 1px solid #eaecf0; color: #475467; background: transparent; text-align: left; cursor: pointer; }
.document-comments__anchor span { display: -webkit-box; overflow: hidden; -webkit-box-orient: vertical; -webkit-line-clamp: 2; font-size: 12px; font-style: italic; }
.document-comments__message + .document-comments__message { margin-top: 10px; }
.document-comments__message-meta { display: flex; min-height: 24px; align-items: center; gap: 6px; }
.document-comments__reply-trigger, .document-comments__more-trigger { opacity: 0; }
.document-comments__message:hover .document-comments__reply-trigger, .document-comments__message:focus-within .document-comments__reply-trigger, .document-comments__message:hover .document-comments__more-trigger, .document-comments__message:focus-within .document-comments__more-trigger { opacity: 1; }
.document-comments__message-meta strong { min-width: 0; overflow: hidden; font-size: 12px; font-weight: 600; text-overflow: ellipsis; white-space: nowrap; }
.document-comments__message-meta time { margin-inline-start: auto; color: #667085; font-size: 10px; white-space: nowrap; }
.document-comments__text { margin: 3px 0 0; color: #344054; font-size: 13px; line-height: 1.5; overflow-wrap: anywhere; white-space: pre-wrap; }
.document-comments__mention { color: #2563eb; }
.document-comments__reply { display: flex; align-items: flex-end; gap: 4px; margin-top: 10px; }
.document-comments__reply :deep(.ant-input) { min-width: 0; }
.document-comments__actions, .document-comments__composer-actions, .document-comments__thread-actions { display: flex; align-items: center; gap: 4px; }
.document-comments__actions { margin-top: 6px; }
.document-comments__solved-label { margin: 4px 0 -4px; color: #667085; font-size: 12px; }
.document-comments__empty { display: flex; min-height: 120px; flex: 1; flex-direction: column; align-items: center; justify-content: center; gap: 8px; color: #667085; font-size: 12px; }
.document-comments__composer { display: grid; gap: 8px; padding: 12px; border-top: 1px solid #e5e7eb; background: #fff; }
.document-comments__composer-actions { justify-content: flex-end; }
@media (max-width: 520px) { .document-comments__header { padding-inline: 10px; } .document-comments__threads { padding: 8px; } .document-comments__composer { padding: 10px; } .document-comments__filter { width: 120px; flex-basis: 120px; } }
@media (hover: none) { .document-comments__reply-trigger, .document-comments__more-trigger { opacity: 1; } }
</style>
