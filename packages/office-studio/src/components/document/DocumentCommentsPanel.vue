<script setup lang="ts">
import { Button, Card, Dropdown, Form, FormItem, Input, Select, Tooltip, theme } from 'antdv-next'
import { CircleCheck, MoreHorizontal, Plus, Reply, RotateCcw, Trash2, X } from 'lucide-vue-next'
import { computed, inject, nextTick, ref, watch } from 'vue'
import type { DocumentCommentMessage, DocumentThreadComment } from '../../domain/model.ts'
import { isDocumentThreadCommentRelevantToUser } from '../../features/document/comments/thread-comments.ts'
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
  onCreateComment?: (text: string) => boolean | Promise<boolean>
  onReplyComment?: (commentId: string, text: string, mentions: readonly { userId: string; label: string }[]) => boolean | Promise<boolean>
  onEditComment?: (commentId: string, messageId: string, text: string) => boolean | Promise<boolean>
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
const { token } = theme.useToken()
const filter = ref<CommentFilter>('all')
const newComment = ref('')
const replyDrafts = ref<Record<string, string>>({})
const replyingTo = ref('')
const replyMentions = ref<Record<string, { userId: string; label: string }[]>>({})
const editingMessage = ref<{ commentId: string; messageId: string } | null>(null)
const editDraft = ref('')
const submitting = ref(false)
const controlsBusy = computed(() => props.busy || submitting.value)
const activeLocale = computed(() => props.locale || localeRef?.value || 'en-US')
const panelStyle = computed(() => ({
  '--comment-background': token.value.colorBgContainer,
  '--comment-subtle-background': token.value.colorFillAlter,
  '--comment-border': token.value.colorBorderSecondary,
  '--comment-text': token.value.colorText,
  '--comment-secondary-text': token.value.colorTextSecondary,
  '--comment-primary': token.value.colorPrimary,
  '--comment-error': token.value.colorError,
}))
const filterOptions = computed(() => [
  { value: 'all', label: t('doc.comments.filter.all', 'All comments') },
  { value: 'resolved', label: t('doc.comments.filter.resolved', 'Resolved') },
  { value: 'open', label: t('doc.comments.filter.open', 'Not resolved') },
  { value: 'concern-me', label: t('doc.comments.filter.concern-me', 'Concern me') },
])
const filteredComments = computed(() => Object.values(props.comments)
  .filter(comment => filter.value === 'all'
    || (filter.value === 'concern-me' && isDocumentThreadCommentRelevantToUser(comment, props.currentUserId))
    || (filter.value !== 'concern-me' && comment.resolved === (filter.value === 'resolved')))
)
const openComments = computed(() => filteredComments.value.filter(comment => !comment.resolved))
const resolvedComments = computed(() => filteredComments.value.filter(comment => comment.resolved))
const visibleComments = computed(() => [...openComments.value, ...resolvedComments.value])

watch(() => props.selectedCommentId, id => {
  if (!id) return
  requestAnimationFrame(() => document.getElementById(`document-comment-${id}`)?.scrollIntoView({ block: 'nearest' }))
})
watch(() => [props.canComment, props.comments] as const, () => {
  const editing = editingMessage.value
  if (editing && (!props.canComment || props.comments[editing.commentId]?.resolved
    || !props.comments[editing.commentId]?.messages.some(message => message.id === editing.messageId && message.authorId === props.currentUserId))) {
    editingMessage.value = null
  }
}, { deep: true })

function selectThread(event: MouseEvent, comment: DocumentThreadComment) {
  const selection = window.getSelection()
  if (selection && !selection.isCollapsed && selection.anchorNode && (event.currentTarget as HTMLElement).contains(selection.anchorNode)) return
  emit('select', comment.resolved ? '' : comment.id)
}

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

async function submitCreate() {
  const text = newComment.value.trim()
  if (!text || !props.canComment || !props.canCreate || controlsBusy.value) return
  submitting.value = true
  try {
    if (props.onCreateComment) {
      if (!await props.onCreateComment(text)) return
    } else {
      const previousIds = new Set(Object.keys(props.comments))
      emit('create', text)
      await nextTick()
      if (!Object.values(props.comments).some(comment => !previousIds.has(comment.id)
        && comment.messages[0]?.text === text && comment.messages[0]?.authorId === props.currentUserId)) return
    }
    if (newComment.value.trim() === text) newComment.value = ''
  } catch {
    // Keep the draft so a failed mutation can be retried.
  } finally {
    submitting.value = false
  }
}

function resetFilter() {
  filter.value = 'all'
}

async function beginReply(commentId: string, message: DocumentCommentMessage) {
  if (!props.canComment || controlsBusy.value || props.comments[commentId]?.resolved || !props.comments[commentId]?.messages.some(item => item.id === message.id)) return
  editingMessage.value = null
  emit('select', commentId)
  replyingTo.value = commentId
  const label = `@${message.authorName}`
  replyDrafts.value[commentId] = `${label} `
  replyMentions.value[commentId] = message.authorId ? [{ userId: message.authorId, label }] : []
  await nextTick()
  document.getElementById(`document-comment-reply-${commentId}`)?.focus()
}

async function submitReply(commentId: string) {
  const text = (replyDrafts.value[commentId] ?? '').trim()
  const comment = props.comments[commentId]
  if (!text || !props.canComment || controlsBusy.value || !comment || comment.resolved) return
  submitting.value = true
  try {
    const mentions = replyMentions.value[commentId] ?? []
    if (props.onReplyComment) {
      if (!await props.onReplyComment(commentId, text, mentions)) return
    } else {
      const previousIds = new Set(comment.messages.map(message => message.id))
      emit('reply', commentId, text, mentions)
      await nextTick()
      if (!props.comments[commentId]?.messages.some(message => !previousIds.has(message.id)
        && message.text === text && message.authorId === props.currentUserId)) return
    }
    replyDrafts.value[commentId] = ''
    replyMentions.value[commentId] = []
    replyingTo.value = ''
  } catch {
    // Keep text and mention identities together when retrying.
  } finally {
    submitting.value = false
  }
}

function messageMenu(comment: DocumentThreadComment, message: DocumentCommentMessage) {
  return {
    items: [
      { key: 'edit', label: t('doc.comments.edit', 'Edit') },
      { key: 'delete', label: t('doc.comments.delete-message', 'Delete comment') },
    ],
    onClick: ({ key }: { key: string | number }) => {
      if (!props.canComment || controlsBusy.value || props.comments[comment.id]?.resolved || message.authorId !== props.currentUserId) return
      if (key === 'edit') beginEdit(comment, message)
      if (key === 'delete') emit('delete-message', comment.id, message.id)
    },
  }
}

async function beginEdit(comment: DocumentThreadComment, message: DocumentCommentMessage) {
  if (!props.canComment || controlsBusy.value || props.comments[comment.id]?.resolved || message.authorId !== props.currentUserId) return
  editingMessage.value = { commentId: comment.id, messageId: message.id }
  editDraft.value = message.text
  await nextTick()
  document.getElementById(`document-comment-edit-${message.id}`)?.focus()
}

async function submitEdit(comment: DocumentThreadComment, message: DocumentCommentMessage) {
  const text = editDraft.value.trim()
  if (!text || !props.canComment || message.authorId !== props.currentUserId || controlsBusy.value || props.comments[comment.id]?.resolved) return
  submitting.value = true
  try {
    if (props.onEditComment) {
      if (!await props.onEditComment(comment.id, message.id, text)) return
    } else {
      emit('edit', comment.id, message.id, text)
      await nextTick()
      if (props.comments[comment.id]?.messages.find(item => item.id === message.id)?.text !== text) return
    }
    editingMessage.value = null
    editDraft.value = ''
  } catch {
    // The parent mutation reports its error while the editor retains input.
  } finally {
    submitting.value = false
  }
}
</script>

<template>
  <section class="document-comments" :style="panelStyle" :aria-label="t('doc.comments.title', 'Comments')">
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
      <Card
        :id="`document-comment-${comment.id}`"
        class="document-comments__thread"
        :class="{ 'is-selected': selectedCommentId === comment.id, 'is-resolved': comment.resolved }"
        role="article"
        :styles="{ body: { padding: '10px' } }"
        @click="selectThread($event, comment)"
      >
        <div class="document-comments__thread-header">
          <Button class="document-comments__anchor" type="text" :aria-label="t('doc.comments.jump-to-anchor', 'Select comment anchor')" @click.stop="emit('select', comment.resolved ? '' : comment.id)">
            <span>{{ comment.referenceText || t('doc.comments.anchor-unavailable', 'Text anchor unavailable') }}</span>
          </Button>
          <div class="document-comments__thread-actions">
            <Tooltip v-if="canComment" :title="t(comment.resolved ? 'doc.comments.reopen' : 'doc.comments.resolve', comment.resolved ? 'Reopen' : 'Resolve')">
              <Button type="text" size="small" :aria-label="t(comment.resolved ? 'doc.comments.reopen' : 'doc.comments.resolve', comment.resolved ? 'Reopen' : 'Resolve')" :disabled="controlsBusy" @click.stop="emit('resolve', comment.id, !comment.resolved)"><component :is="comment.resolved ? RotateCcw : CircleCheck" :size="16" aria-hidden="true" /></Button>
            </Tooltip>
            <Tooltip v-if="deletableCommentIds.includes(comment.id)" :title="t('doc.comments.delete', 'Delete thread')">
              <Button type="text" size="small" danger :aria-label="t('doc.comments.delete', 'Delete thread')" :disabled="controlsBusy" @click.stop="emit('delete', comment.id)"><Trash2 :size="16" aria-hidden="true" /></Button>
            </Tooltip>
          </div>
        </div>
        <div v-for="message in comment.messages" :key="message.id" class="document-comments__message">
          <div class="document-comments__message-meta">
            <strong>{{ message.authorName }}</strong>
            <time :datetime="message.createdAt">{{ formatDate(message.createdAt) }}</time>
            <Tooltip v-if="canComment && !comment.resolved" :title="t('doc.comments.reply', 'Reply')">
              <Button class="document-comments__reply-trigger" type="text" size="small" :aria-label="t('doc.comments.reply', 'Reply')" :disabled="controlsBusy" @click.stop="beginReply(comment.id, message)"><Reply :size="16" aria-hidden="true" /></Button>
            </Tooltip>
            <Dropdown v-if="canComment && message.authorId === currentUserId && !comment.resolved" :trigger="['click']" placement="bottomRight" :menu="messageMenu(comment, message)">
              <Button class="document-comments__more-trigger" type="text" size="small" :aria-label="t('doc.comments.more', 'More')" :disabled="controlsBusy" @click.stop><MoreHorizontal :size="16" aria-hidden="true" /></Button>
            </Dropdown>
          </div>
          <Form v-if="editingMessage?.commentId === comment.id && editingMessage.messageId === message.id" class="document-comments__edit" layout="vertical" @click.stop @finish="submitEdit(comment, message)">
            <FormItem class="document-comments__field">
              <Input.TextArea :id="`document-comment-edit-${message.id}`" v-model:value="editDraft" :aria-label="t('doc.comments.edit', 'Edit comment')" :auto-size="{ minRows: 2, maxRows: 5 }" :disabled="controlsBusy" @keydown.ctrl.enter.prevent="submitEdit(comment, message)" @keydown.meta.enter.prevent="submitEdit(comment, message)" />
            </FormItem>
            <div class="document-comments__actions">
              <Button size="small" type="primary" html-type="submit" :disabled="!editDraft.trim() || controlsBusy">{{ t('common.save', 'Save') }}</Button>
              <Button size="small" type="text" :disabled="controlsBusy" @click.stop="editingMessage = null; editDraft = ''">{{ t('common.cancel', 'Cancel') }}</Button>
            </div>
          </Form>
          <p v-else class="document-comments__text"><span v-if="commentMention(message)" class="document-comments__mention">{{ commentMention(message)?.label }}</span>{{ commentBody(message) }}</p>
        </div>

        <Form v-if="canComment && !comment.resolved && !editingMessage && selectedCommentId === comment.id" class="document-comments__reply" layout="vertical" @click.stop @finish="submitReply(comment.id)">
          <FormItem class="document-comments__field">
            <Input.TextArea :id="`document-comment-reply-${comment.id}`" v-model:value="replyDrafts[comment.id]" :placeholder="t('doc.comments.reply-placeholder', 'Reply')" :aria-label="t('doc.comments.reply-placeholder', 'Reply')" :auto-size="{ minRows: 1, maxRows: 4 }" :disabled="controlsBusy" @focus="replyingTo = comment.id" @keydown.ctrl.enter.prevent="submitReply(comment.id)" @keydown.meta.enter.prevent="submitReply(comment.id)" />
          </FormItem>
          <div v-if="replyingTo === comment.id || replyDrafts[comment.id]?.trim()" class="document-comments__actions">
            <Button size="small" type="primary" html-type="submit" :disabled="!replyDrafts[comment.id]?.trim() || controlsBusy">{{ t('doc.comments.reply-submit', 'Comment') }}</Button>
            <Button size="small" type="text" :disabled="controlsBusy" @click.stop="replyingTo = ''; replyDrafts[comment.id] = ''; replyMentions[comment.id] = []">{{ t('common.cancel', 'Cancel') }}</Button>
          </div>
        </Form>
      </Card>
      </template>
    </div>
    <div v-else class="document-comments__empty">
      <span>{{ t('doc.comments.empty', 'No comments yet') }}</span>
      <Button v-if="filter !== 'all'" size="small" @click="resetFilter">{{ t('doc.comments.filter.reset', 'Show all comments') }}</Button>
      <Button v-if="filter === 'all' && canStartCreate" size="small" @click="emit('add')"><Plus :size="16" aria-hidden="true" />{{ t('doc.comments.panel-add', 'Add Comment') }}</Button>
    </div>

    <Form v-if="canComment && canCreate" class="document-comments__composer" layout="vertical" @finish="submitCreate">
      <FormItem class="document-comments__field">
      <Input.TextArea v-model:value="newComment" :placeholder="t('doc.comments.new-placeholder', 'Add a comment')" :aria-label="t('doc.comments.new-placeholder', 'Add a comment')" :auto-size="{ minRows: 2, maxRows: 5 }" :disabled="controlsBusy" @keydown.ctrl.enter.prevent="submitCreate" @keydown.meta.enter.prevent="submitCreate" />
      </FormItem>
      <div class="document-comments__composer-actions">
        <Button type="primary" :loading="controlsBusy" :disabled="!newComment.trim() || controlsBusy" html-type="submit">{{ t('doc.comments.add', 'Comment') }}</Button>
      </div>
    </Form>
  </section>
</template>

<style scoped>
.document-comments { display: flex; min-height: 0; height: 100%; flex-direction: column; color: var(--comment-text); background: var(--comment-background); }
.document-comments__header { display: flex; min-height: 48px; align-items: center; justify-content: space-between; gap: 10px; padding: 0 14px; border-bottom: 1px solid var(--comment-border); }
.document-comments__heading { display: flex; min-width: 0; align-items: center; gap: 8px; }
.document-comments__header-actions { display: flex; min-width: 0; align-items: center; gap: 4px; }
.document-comments__heading h2 { margin: 0; font-size: 14px; font-weight: 600; }
.document-comments__filter { width: 136px; flex: 0 0 136px; }
.document-comments__error { margin: 10px 14px 0; color: var(--comment-error); font-size: 12px; overflow-wrap: anywhere; }
.document-comments__threads { min-height: 0; flex: 1; overflow: auto; padding: 12px; }
.document-comments__thread { margin-bottom: 10px; border-color: var(--comment-border); border-radius: 4px; background: var(--comment-background); cursor: pointer; }
.document-comments__thread.is-selected { border-color: var(--comment-primary); box-shadow: 0 0 0 1px var(--comment-primary); }
.document-comments__thread.is-resolved { background: var(--comment-subtle-background); }
.document-comments__thread-header { display: flex; min-width: 0; align-items: center; justify-content: space-between; gap: 8px; margin-bottom: 10px; }
.document-comments__anchor { display: block; min-width: 0; height: auto; flex: 1; padding: 0 0 8px; border: 0; border-bottom: 1px solid var(--comment-border); border-radius: 0; color: var(--comment-secondary-text); text-align: left; white-space: normal; }
.document-comments__anchor span { display: -webkit-box; overflow: hidden; -webkit-box-orient: vertical; -webkit-line-clamp: 2; font-size: 12px; font-style: italic; }
.document-comments__message + .document-comments__message { margin-top: 10px; }
.document-comments__message-meta { display: flex; min-height: 24px; align-items: center; gap: 6px; }
.document-comments__reply-trigger, .document-comments__more-trigger { opacity: 0; }
.document-comments__message:hover .document-comments__reply-trigger, .document-comments__message:focus-within .document-comments__reply-trigger, .document-comments__message:hover .document-comments__more-trigger, .document-comments__message:focus-within .document-comments__more-trigger { opacity: 1; }
.document-comments__message-meta strong { min-width: 0; overflow: hidden; font-size: 12px; font-weight: 600; text-overflow: ellipsis; white-space: nowrap; }
.document-comments__message-meta time { margin-inline-start: auto; color: var(--comment-secondary-text); font-size: 10px; white-space: nowrap; }
.document-comments__text { margin: 3px 0 0; color: var(--comment-text); font-size: 13px; line-height: 1.5; overflow-wrap: anywhere; white-space: pre-wrap; }
.document-comments__mention { color: var(--comment-primary); }
.document-comments__reply { margin-top: 10px; }
.document-comments__field { margin-bottom: 0; }
.document-comments__reply :deep(.ant-input) { min-width: 0; }
.document-comments__actions, .document-comments__composer-actions, .document-comments__thread-actions { display: flex; align-items: center; gap: 4px; }
.document-comments__actions { justify-content: flex-end; margin-top: 6px; }
.document-comments__solved-label { margin: 4px 0 8px; color: var(--comment-secondary-text); font-size: 12px; }
.document-comments__empty { display: flex; min-height: 120px; flex: 1; flex-direction: column; align-items: center; justify-content: center; gap: 8px; color: var(--comment-secondary-text); font-size: 12px; }
.document-comments__composer { display: grid; gap: 8px; padding: 12px; border-top: 1px solid var(--comment-border); background: var(--comment-background); }
.document-comments__composer-actions { justify-content: flex-end; }
@media (max-width: 520px) { .document-comments__header { padding-inline: 10px; } .document-comments__threads { padding: 8px; } .document-comments__composer { padding: 10px; } .document-comments__filter { width: 120px; flex-basis: 120px; } }
@media (hover: none) { .document-comments__reply-trigger, .document-comments__more-trigger { opacity: 1; } }
</style>
