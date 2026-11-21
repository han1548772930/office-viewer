import assert from 'node:assert/strict'
import { test, afterAll } from 'vitest'
import { createServer } from 'vite'
import { createBlankDocument } from '../src/domain/model.ts'
import { createParagraph } from '../src/features/document/model/document-model.ts'

const vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom' })
const comments = await vite.ssrLoadModule('/packages/office-studio/src/features/document/comments/thread-comments.ts')
const commands = await vite.ssrLoadModule('/packages/office-studio/src/features/document/commands/document-commands.ts')
afterAll(() => vite.close())

const actor = { documentId: 'doc-1', userId: 'alice', displayName: 'Alice' }

test('thread comments preserve range anchors, replies, edits, resolution and JSON snapshots', () => {
  const model = createBlankDocument()
  const first = model.blocks[0]
  first.runs[0].text = 'First paragraph'
  const second = createParagraph('Second paragraph')
  model.blocks.push(second)

  const thread = comments.createDocumentThreadComment(model, {
    anchor: { blockId: first.id, offset: 6 },
    focus: { blockId: second.id, offset: 6 },
  }, 'Review this passage', actor, '2026-01-01T00:00:00.000Z')

  assert.deepEqual(thread.anchor.map(({ blockId, text }) => ({ blockId, text })), [
    { blockId: first.id, text: 'paragraph' },
    { blockId: second.id, text: 'Second' },
  ])
  assert.equal(thread.referenceText, 'paragraph\nSecond')
  const reply = comments.replyToDocumentThreadComment(model, thread.id, 'I will revise it.', { ...actor, userId: 'bob', displayName: 'Bob' }, '2026-01-02T00:00:00.000Z')
  comments.editDocumentCommentMessage(model, thread.id, reply.id, 'Revised.', { ...actor, userId: 'bob', displayName: 'Bob' }, '2026-01-03T00:00:00.000Z')
  comments.setDocumentThreadCommentResolved(model, thread.id, true, actor, '2026-01-04T00:00:00.000Z')

  const restored = JSON.parse(JSON.stringify(model))
  assert.equal(restored.comments[thread.id].messages[1].text, 'Revised.')
  assert.equal(restored.comments[thread.id].resolved, true)
  assert.equal(restored.comments[thread.id].updatedAt, '2026-01-04T00:00:00.000Z')
  assert.equal(comments.documentThreadCommentAnchorStatus(restored, restored.comments[thread.id]), 'current')
})

test('stale and orphaned anchors are reported after content changes', () => {
  const model = createBlankDocument()
  const paragraph = model.blocks[0]
  paragraph.runs[0].text = 'Anchor text'
  const thread = comments.createDocumentThreadComment(model, {
    anchor: { blockId: paragraph.id, offset: 0 },
    focus: { blockId: paragraph.id, offset: 6 },
  }, 'Comment', actor, '2026-01-01T00:00:00.000Z')

  paragraph.runs[0].text = 'Changed text'
  assert.equal(comments.documentThreadCommentAnchorStatus(model, thread), 'stale')
  model.blocks = []
  assert.equal(comments.documentThreadCommentAnchorStatus(model, thread), 'orphaned')
})

test('single-paragraph text edits shift or expand matching comment anchors', () => {
  const model = createBlankDocument()
  const paragraph = model.blocks[0]
  paragraph.runs[0].text = 'Hello World'
  const thread = comments.createDocumentThreadComment(model, {
    anchor: { blockId: paragraph.id, offset: 6 },
    focus: { blockId: paragraph.id, offset: 11 },
  }, 'Review', actor, '2026-01-01T00:00:00.000Z')

  commands.replaceSelectionText(model, {
    anchor: { blockId: paragraph.id, offset: 0 },
    focus: { blockId: paragraph.id, offset: 0 },
  }, 'New ')
  assert.deepEqual(thread.anchor[0], { blockId: paragraph.id, startOffset: 10, endOffset: 15, text: 'World' })

  commands.replaceSelectionText(model, {
    anchor: { blockId: paragraph.id, offset: 11 },
    focus: { blockId: paragraph.id, offset: 12 },
  }, 'A')
  assert.equal(thread.anchor[0].text, 'WArld')
  assert.equal(thread.referenceText, 'WArld')
  assert.equal(comments.documentThreadCommentAnchorStatus(model, thread), 'current')
})

test('comment permission denial rejects creation before mutating the model', () => {
  const model = createBlankDocument()
  const paragraph = model.blocks[0]
  paragraph.runs[0].text = 'Protected text'
  model.permissionRules = [{ id: 'deny-comments', scope: 'document', action: 'comment', effect: 'deny', userIds: ['alice'] }]
  const before = structuredClone(model)

  assert.throws(() => comments.createDocumentThreadComment(model, {
    anchor: { blockId: paragraph.id, offset: 0 },
    focus: { blockId: paragraph.id, offset: 9 },
  }, 'Comment', actor), { code: 'DOCUMENT_PERMISSION_DENIED' })
  assert.deepEqual(model, before)
})

test('empty ranges and empty comment messages are rejected without partial mutation', () => {
  const model = createBlankDocument()
  const paragraph = model.blocks[0]
  paragraph.runs[0].text = 'Text'
  const before = structuredClone(model)

  assert.throws(() => comments.createDocumentThreadComment(model, {
    anchor: { blockId: paragraph.id, offset: 1 },
    focus: { blockId: paragraph.id, offset: 1 },
  }, 'Comment', actor))
  assert.throws(() => comments.createDocumentThreadComment(model, {
    anchor: { blockId: paragraph.id, offset: 0 },
    focus: { blockId: paragraph.id, offset: 1 },
  }, '  ', actor))
  assert.deepEqual(model, before)
})
