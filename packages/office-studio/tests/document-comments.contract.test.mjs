import assert from 'node:assert/strict'
import { test, afterAll } from 'vitest'
import { createServer } from 'vite'
import { createBlankDocument } from '../src/domain/model.ts'
import { createParagraph } from '../src/features/document/model/document-model.ts'
import {
  createDocumentThreadComment,
  deleteDocumentThreadComment,
  documentThreadCommentAnchorStatus,
  editDocumentCommentMessage,
  replyToDocumentThreadComment,
  setDocumentThreadCommentResolved,
} from '../src/features/document/comments/thread-comments.ts'

const vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom' })
afterAll(() => vite.close())

function setup() {
  const model = createBlankDocument()
  const paragraph = createParagraph('Highlighted phrase')
  model.blocks = [paragraph]
  const selection = {
    anchor: { blockId: paragraph.id, offset: 0 },
    focus: { blockId: paragraph.id, offset: 11 },
  }
  const author = { documentId: 'doc-1', userId: 'author', displayName: 'Author' }
  return { model, paragraph, selection, author }
}

test('creates a thread anchored to the exact selected text and marks the document modified', () => {
  const { model, selection, author } = setup()
  const comment = createDocumentThreadComment(model, selection, 'Please review', author, '2026-11-21T00:00:00.000Z')

  assert.equal(comment.referenceText, 'Highlighted')
  assert.equal(comment.anchor.length, 1)
  assert.equal(comment.anchor[0].startOffset, 0)
  assert.equal(comment.anchor[0].endOffset, 11)
  assert.equal(comment.messages[0].text, 'Please review')
  assert.equal(documentThreadCommentAnchorStatus(model, comment), 'current')
  assert.equal(model.modified, true)
  assert.throws(() => createDocumentThreadComment(model, { anchor: selection.anchor, focus: selection.anchor }, 'No range', author), /non-empty range/)
})

test('enforces comment permission without creating partial thread state', () => {
  const { model, selection, author } = setup()
  model.permissionRules = [{ id: 'deny-comments', scope: 'document', action: 'comment', effect: 'deny', userIds: ['author'] }]

  assert.throws(() => createDocumentThreadComment(model, selection, 'Blocked', author), /permission denied/)
  assert.equal(model.comments, undefined)
})

test('supports replies, author-only edits, resolve and reopen transitions, and deletion', () => {
  const { model, selection, author } = setup()
  const comment = createDocumentThreadComment(model, selection, 'Initial', author, '2026-11-21T00:00:00.000Z')
  const other = { documentId: 'doc-1', userId: 'other', displayName: 'Other' }
  const reply = replyToDocumentThreadComment(model, comment.id, 'A reply', other, '2026-11-21T00:01:00.000Z')

  assert.equal(comment.messages.length, 2)
  assert.equal(reply.authorName, 'Other')
  assert.throws(() => editDocumentCommentMessage(model, comment.id, comment.messages[0].id, 'Changed by other', other), /Only the message author/)
  editDocumentCommentMessage(model, comment.id, comment.messages[0].id, 'Updated', author, '2026-11-21T00:02:00.000Z')
  assert.equal(comment.messages[0].text, 'Updated')
  assert.throws(() => editDocumentCommentMessage(model, comment.id, reply.id, ' ', other), /cannot be empty/)

  setDocumentThreadCommentResolved(model, comment.id, true, other, '2026-11-21T00:03:00.000Z')
  assert.equal(comment.resolved, true)
  setDocumentThreadCommentResolved(model, comment.id, false, author, '2026-11-21T00:04:00.000Z')
  assert.equal(comment.resolved, false)
  assert.equal(deleteDocumentThreadComment(model, comment.id, author), true)
  assert.equal(deleteDocumentThreadComment(model, comment.id, author), false)
})
