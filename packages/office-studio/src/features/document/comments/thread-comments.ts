import type {
  DocumentCommentAnchorSegment,
  DocumentCommentMessage,
  DocumentModel,
  DocumentThreadComment,
} from '../../../domain/model'
import {
  assertDocumentPermission,
  permissionTargetsForRange,
  type DocumentPermissionContext,
} from '../permissions/document-permissions'
import { documentId } from '../model/document-model'
import { flattenDocumentBody } from '../services/document-facade'
import type { DocumentSelection } from '../model/document-model'

export interface DocumentCommentActor extends DocumentPermissionContext {
  displayName: string
}

export type DocumentCommentAnchorStatus = 'current' | 'stale' | 'orphaned'

function assertMessageAuthor(message: DocumentCommentMessage | undefined, actor: DocumentCommentActor): void {
  if (message?.authorId !== actor.userId) throw new Error('Only the message author can edit or delete this comment.')
}

function commentsFor(model: DocumentModel): Record<string, DocumentThreadComment> {
  return model.comments ??= {}
}

function commentText(value: string): string {
  const text = value.trim()
  if (!text) throw new RangeError('Comment text cannot be empty.')
  return text
}

function assertCanComment(model: DocumentModel, actor: DocumentCommentActor, startOffset = 0, endOffset = 0): void {
  const targets = startOffset === endOffset ? [] : permissionTargetsForRange(model, startOffset, endOffset)
  assertDocumentPermission(model.permissionRules ?? [], actor, 'comment', targets)
}

function anchorForRange(model: DocumentModel, startOffset: number, endOffset: number): DocumentCommentAnchorSegment[] {
  const body = flattenDocumentBody(model)
  if (!Number.isInteger(startOffset) || !Number.isInteger(endOffset) || startOffset < 0 || endOffset > body.text.length || startOffset >= endOffset) {
    throw new RangeError('Comment anchor must be a non-empty range within the document body.')
  }
  const anchor: DocumentCommentAnchorSegment[] = []
  for (const segment of body.segments) {
    const start = Math.max(startOffset, segment.startOffset)
    const end = Math.min(endOffset, segment.endOffset)
    if (start >= end) continue
    anchor.push({
      blockId: segment.cellBlockId ?? segment.blockId,
      startOffset: start - segment.startOffset,
      endOffset: end - segment.startOffset,
      text: segment.text.slice(start - segment.startOffset, end - segment.startOffset),
    })
  }
  if (!anchor.length) throw new RangeError('Comment anchor must include document text.')
  return anchor
}

function createMessage(actor: DocumentCommentActor, text: string, now: string): DocumentCommentMessage {
  return { id: documentId('comment-message'), authorId: actor.userId, authorName: actor.displayName, text: commentText(text), createdAt: now }
}

function updateTimestamp(comment: DocumentThreadComment, now: string): void {
  comment.updatedAt = now
}

export function createDocumentThreadComment(
  model: DocumentModel,
  selection: DocumentSelection,
  text: string,
  actor: DocumentCommentActor,
  now = new Date().toISOString(),
): DocumentThreadComment {
  const body = flattenDocumentBody(model)
  const ordered = body.segments.map(segment => ({ segment, start: segment.startOffset, end: segment.endOffset }))
  const pointOffset = (point: DocumentSelection['anchor']): number => {
    const found = ordered.find(({ segment }) => segment.blockId === point.blockId
      && segment.row === point.row && segment.col === point.col
      && (segment.cellBlockId ?? segment.blockId) === (point.cellBlockId ?? point.blockId))
    if (!found) throw new RangeError('Comment selection is not in the document body.')
    if (!Number.isInteger(point.offset) || point.offset < 0 || point.offset > found.end - found.start) {
      throw new RangeError('Comment selection offset is outside the paragraph.')
    }
    return found.start + point.offset
  }
  const anchorOffset = pointOffset(selection.anchor)
  const focusOffset = pointOffset(selection.focus)
  const start = Math.min(anchorOffset, focusOffset)
  const end = Math.max(anchorOffset, focusOffset)
  assertCanComment(model, actor, start, end)
  const anchor = anchorForRange(model, start, end)
  const comment: DocumentThreadComment = {
    id: documentId('comment'),
    anchor,
    referenceText: anchor.map(segment => segment.text).join('\n'),
    messages: [createMessage(actor, text, now)],
    resolved: false,
    createdAt: now,
    updatedAt: now,
  }
  commentsFor(model)[comment.id] = comment
  model.modified = true
  return comment
}

export function replyToDocumentThreadComment(
  model: DocumentModel,
  commentId: string,
  text: string,
  actor: DocumentCommentActor,
  now = new Date().toISOString(),
  mentions: readonly { userId: string; label: string }[] = [],
): DocumentCommentMessage {
  assertCanComment(model, actor)
  const comment = requireComment(model, commentId)
  const message = createMessage(actor, text, now)
  const retainedMentions = mentions.filter(mention => mention.userId && mention.label && message.text.startsWith(mention.label))
  if (retainedMentions.length) message.mentions = retainedMentions.map(mention => ({ ...mention }))
  comment.messages.push(message)
  updateTimestamp(comment, now)
  model.modified = true
  return message
}

export function editDocumentCommentMessage(
  model: DocumentModel,
  commentId: string,
  messageId: string,
  text: string,
  actor: DocumentCommentActor,
  now = new Date().toISOString(),
): void {
  assertCanComment(model, actor)
  const comment = requireComment(model, commentId)
  const message = comment.messages.find(candidate => candidate.id === messageId)
  if (!message) throw new RangeError(`Unknown comment message id: ${messageId}`)
  assertMessageAuthor(message, actor)
  const nextText = commentText(text)
  if (message.text === nextText) return
  message.text = nextText
  if (message.mentions) {
    const mentions = message.mentions.filter(mention => mention.userId && mention.label && nextText.startsWith(mention.label))
    if (mentions.length) message.mentions = mentions
    else delete message.mentions
  }
  message.updatedAt = now
  updateTimestamp(comment, now)
  model.modified = true
}

export function setDocumentThreadCommentResolved(
  model: DocumentModel,
  commentId: string,
  resolved: boolean,
  actor: DocumentCommentActor,
  now = new Date().toISOString(),
): void {
  assertCanComment(model, actor)
  const comment = requireComment(model, commentId)
  if (comment.resolved === resolved) return
  comment.resolved = resolved
  updateTimestamp(comment, now)
  model.modified = true
}

export function deleteDocumentThreadComment(model: DocumentModel, commentId: string, actor: DocumentCommentActor): boolean {
  assertCanComment(model, actor)
  const comments = model.comments
  if (!comments?.[commentId]) return false
  assertMessageAuthor(comments[commentId].messages[0], actor)
  delete comments[commentId]
  model.modified = true
  return true
}

export function deleteDocumentCommentMessage(
  model: DocumentModel,
  commentId: string,
  messageId: string,
  actor: DocumentCommentActor,
  now = new Date().toISOString(),
): boolean {
  assertCanComment(model, actor)
  const comment = requireComment(model, commentId)
  const index = comment.messages.findIndex(message => message.id === messageId)
  if (index < 0) return false
  assertMessageAuthor(comment.messages[index], actor)
  if (index === 0) return deleteDocumentThreadComment(model, commentId, actor)
  comment.messages.splice(index, 1)
  updateTimestamp(comment, now)
  model.modified = true
  return true
}

export function relocateDocumentThreadComment(
  model: DocumentModel,
  commentId: string,
  selection: DocumentSelection,
  actor: DocumentCommentActor,
  now = new Date().toISOString(),
): void {
  const body = flattenDocumentBody(model)
  const pointOffset = (point: DocumentSelection['anchor']): number => {
    const segment = body.segments.find(candidate => candidate.blockId === point.blockId
      && candidate.row === point.row && candidate.col === point.col
      && (candidate.cellBlockId ?? candidate.blockId) === (point.cellBlockId ?? point.blockId))
    if (!segment || point.offset < 0 || point.offset > segment.endOffset - segment.startOffset) {
      throw new RangeError('Comment selection is not in the document body.')
    }
    return segment.startOffset + point.offset
  }
  const first = pointOffset(selection.anchor)
  const second = pointOffset(selection.focus)
  const start = Math.min(first, second)
  const end = Math.max(first, second)
  assertCanComment(model, actor, start, end)
  const comment = requireComment(model, commentId)
  const anchor = anchorForRange(model, start, end)
  comment.anchor = anchor
  comment.referenceText = anchor.map(segment => segment.text).join('\n')
  updateTimestamp(comment, now)
  model.modified = true
}

export function documentThreadCommentAnchorStatus(model: DocumentModel, comment: DocumentThreadComment): DocumentCommentAnchorStatus {
  const body = flattenDocumentBody(model)
  let resolved = 0
  for (const anchor of comment.anchor) {
    const segment = body.segments.find(candidate => (candidate.cellBlockId ?? candidate.blockId) === anchor.blockId)
    if (!segment) return 'orphaned'
    const current = segment.text.slice(anchor.startOffset, anchor.endOffset)
    if (current !== anchor.text) return 'stale'
    resolved += current.length
  }
  return resolved === comment.referenceText.replaceAll('\n', '').length ? 'current' : 'stale'
}

export function isDocumentThreadCommentRelevantToUser(comment: DocumentThreadComment, userId: string): boolean {
  if (!userId) return false
  return comment.messages.some(message => message.authorId === userId
    || message.mentions?.some(mention => mention.userId === userId) === true)
}

export function transformDocumentCommentAnchorsForTextEdit(
  model: DocumentModel,
  blockId: string,
  startOffset: number,
  endOffset: number,
  insertedText: string,
): void {
  const paragraph = findParagraph(model.blocks, blockId)
  if (!paragraph) return
  const text = paragraph.runs.map(run => run.text).join('')
  const insertedLength = insertedText.length
  const delta = insertedLength - (endOffset - startOffset)
  for (const comment of Object.values(model.comments ?? {})) {
    let changed = false
    for (const anchor of comment.anchor) {
      if (anchor.blockId !== blockId) continue
      if (endOffset <= anchor.startOffset) {
        anchor.startOffset += delta
        anchor.endOffset += delta
      } else if (startOffset < anchor.endOffset && endOffset > anchor.startOffset) {
        const nextStart = startOffset <= anchor.startOffset ? startOffset : anchor.startOffset
        const nextEnd = endOffset >= anchor.endOffset
          ? startOffset + insertedLength
          : anchor.endOffset + delta
        anchor.startOffset = nextStart
        anchor.endOffset = Math.max(nextStart, nextEnd)
      } else {
        continue
      }
      anchor.text = text.slice(anchor.startOffset, anchor.endOffset)
      changed = true
    }
    if (changed) {
      comment.referenceText = comment.anchor.map(anchor => anchor.text).join('\n')
      comment.updatedAt = new Date().toISOString()
    }
  }
}

function findParagraph(blocks: DocumentModel['blocks'], id: string): DocumentModel['blocks'][number] | undefined {
  for (const block of blocks) {
    if (block.id === id) return block
    const table = block.metadata.table as { rows?: Array<{ cells?: Array<{ blocks?: DocumentModel['blocks'] }> }> } | undefined
    for (const row of table?.rows ?? []) for (const cell of row.cells ?? []) {
      const nested = findParagraph(cell.blocks ?? [], id)
      if (nested) return nested
    }
  }
  return undefined
}

function requireComment(model: DocumentModel, id: string): DocumentThreadComment {
  const comment = model.comments?.[id]
  if (!comment) throw new RangeError(`Unknown document comment id: ${id}`)
  return comment
}
