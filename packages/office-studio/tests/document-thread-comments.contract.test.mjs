import assert from 'node:assert/strict'
import { test, afterAll } from 'vitest'
import { createServer } from 'vite'
import { createBlankDocument } from '../src/domain/model.ts'
import { createParagraph } from '../src/features/document/model/document-model.ts'

const vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom' })
const comments = await vite.ssrLoadModule('/packages/office-studio/src/features/document/comments/thread-comments.ts')
const commands = await vite.ssrLoadModule('/packages/office-studio/src/features/document/commands/document-commands.ts')
const facade = await vite.ssrLoadModule('/packages/office-studio/src/features/document/services/document-facade.ts')
const fields = await vite.ssrLoadModule('/packages/office-studio/src/features/document/model/document-text-fields.ts')
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

test('a paragraph split expands one comment into ordered segments and joining it coalesces the original thread', () => {
  const model = createBlankDocument(), first = model.blocks[0]
  first.runs[0].text = 'AoldZ'
  const thread = comments.createDocumentThreadComment(model, {
    anchor: { blockId: first.id, offset: 1 }, focus: { blockId: first.id, offset: 4 },
  }, 'Keep the selected result', actor, '2026-01-01T00:00:00.000Z')
  const messages = structuredClone(thread.messages), originalId = thread.id
  commands.replaceSelectionText(model, { anchor: { blockId: first.id, offset: 2 }, focus: { blockId: first.id, offset: 2 } }, '\n')
  const last = model.blocks[1]
  assert.deepEqual(thread.anchor, [
    { blockId: first.id, startOffset: 1, endOffset: 2, text: 'o' },
    { blockId: last.id, startOffset: 0, endOffset: 2, text: 'ld' },
  ])
  assert.equal(thread.referenceText, 'o\nld')
  assert.equal(comments.documentThreadCommentAnchorStatus(model, thread), 'current')
  commands.replaceSelectionText(model, { anchor: { blockId: first.id, offset: 2 }, focus: { blockId: last.id, offset: 0 } }, '')
  assert.deepEqual(thread.anchor, [{ blockId: first.id, startOffset: 1, endOffset: 4, text: 'old' }])
  assert.equal(thread.referenceText, 'old'); assert.equal(thread.id, originalId)
  assert.deepEqual(thread.messages, messages)
  assert.equal(comments.documentThreadCommentAnchorStatus(model, thread), 'current')
})

test('a multiline facade replacement coalesces overlapping mapped segments and preserves the unselected suffix thread', () => {
  const model = createBlankDocument(), first = model.blocks[0], last = createParagraph('cacheZ')
  first.runs[0].text = 'ABold'; model.blocks.push(last)
  const thread = comments.createDocumentThreadComment(model, {
    anchor: { blockId: first.id, offset: 2 }, focus: { blockId: last.id, offset: 5 },
  }, 'Keep the surviving result', actor, '2026-01-01T00:00:00.000Z')
  const suffix = comments.createDocumentThreadComment(model, {
    anchor: { blockId: last.id, offset: 5 }, focus: { blockId: last.id, offset: 6 },
  }, 'Keep Z', actor, '2026-01-01T00:00:00.000Z')
  const before = structuredClone(model), next = facade.replaceDocumentRange(model, 1, 8, 'Q\nR')
  assert.deepEqual(model, before)
  assert.deepEqual(next.blocks.map(block => block.runs.map(run => run.text).join('')), ['AQ', 'RcheZ'])
  assert.deepEqual(next.comments[thread.id].anchor, [
    { blockId: first.id, startOffset: 1, endOffset: 2, text: 'Q' },
    { blockId: next.blocks[1].id, startOffset: 0, endOffset: 4, text: 'Rche' },
  ])
  assert.equal(next.comments[thread.id].referenceText, 'Q\nRche')
  assert.deepEqual(next.comments[suffix.id].anchor, [{ blockId: next.blocks[1].id, startOffset: 4, endOffset: 5, text: 'Z' }])
  assert.deepEqual(next.comments[thread.id].messages, before.comments[thread.id].messages)
  assert.equal(comments.documentThreadCommentAnchorStatus(next, next.comments[thread.id]), 'current')
})

test('a later invalid comment anchor rejects the complete anchor plan before updating an earlier thread', () => {
  const model = createBlankDocument(), paragraph = model.blocks[0]
  paragraph.runs[0].text = 'AB'
  const thread = comments.createDocumentThreadComment(model, {
    anchor: { blockId: paragraph.id, offset: 0 }, focus: { blockId: paragraph.id, offset: 2 },
  }, 'Valid', actor, '2026-01-01T00:00:00.000Z')
  const source = fields.documentTextFieldParagraphs(model)
  model.comments.invalid = { ...structuredClone(thread), id: 'invalid', anchor: [{ blockId: paragraph.id, startOffset: 0, endOffset: 99, text: 'stale' }] }
  paragraph.runs[0].text = 'AXB'
  const before = structuredClone(model.comments)
  assert.throws(() => comments.transformDocumentCommentAnchorsForBodyRangeEdit(model, source, fields.documentTextFieldParagraphs(model), 1, 1, 'X'), /源 UTF-16/)
  assert.deepEqual(model.comments, before)
})

test('comment creation and relocation reject UTF-16 positions inside a surrogate pair', () => {
  const model = createBlankDocument(), paragraph = model.blocks[0]
  paragraph.runs[0].text = 'A😀B'
  const before = structuredClone(model)
  assert.throws(() => comments.createDocumentThreadComment(model, {
    anchor: { blockId: paragraph.id, offset: 1 }, focus: { blockId: paragraph.id, offset: 2 },
  }, 'Invalid half character', actor), /offset/)
  assert.deepEqual(model, before)
  const thread = comments.createDocumentThreadComment(model, {
    anchor: { blockId: paragraph.id, offset: 1 }, focus: { blockId: paragraph.id, offset: 3 },
  }, 'Whole character', actor, '2026-01-01T00:00:00.000Z')
  const anchor = structuredClone(thread)
  assert.throws(() => comments.relocateDocumentThreadComment(model, thread.id, {
    anchor: { blockId: paragraph.id, offset: 2 }, focus: { blockId: paragraph.id, offset: 3 },
  }, actor), /selection/)
  assert.deepEqual(thread, anchor)
})

test('mapped replacement anchors and empty deleted points stay on complete UTF-16 characters', () => {
  const model = createBlankDocument(), paragraph = model.blocks[0]
  paragraph.runs[0].text = 'AoldZ'
  const thread = comments.createDocumentThreadComment(model, {
    anchor: { blockId: paragraph.id, offset: 2 }, focus: { blockId: paragraph.id, offset: 3 },
  }, 'Keep the replacement', actor, '2026-01-01T00:00:00.000Z')
  model.comments.empty = { ...structuredClone(thread), id: 'empty', anchor: [{ blockId: paragraph.id, startOffset: 2, endOffset: 2, text: '' }], referenceText: '' }
  commands.replaceSelectionText(model, { anchor: { blockId: paragraph.id, offset: 1 }, focus: { blockId: paragraph.id, offset: 4 } }, '😀')
  assert.deepEqual(thread.anchor, [{ blockId: paragraph.id, startOffset: 1, endOffset: 3, text: '😀' }])
  assert.deepEqual(model.comments.empty.anchor, [{ blockId: paragraph.id, startOffset: 1, endOffset: 1, text: '' }])
  commands.replaceSelectionText(model, { anchor: { blockId: paragraph.id, offset: 1 }, focus: { blockId: paragraph.id, offset: 3 } }, '')
  assert.deepEqual(thread.anchor, [{ blockId: paragraph.id, startOffset: 1, endOffset: 1, text: '' }])
  assert.equal(thread.referenceText, '')
  assert.equal(comments.documentThreadCommentAnchorStatus(model, thread), 'current')
})

test('an unresolved source anchor remains a group boundary when a new paragraph reuses its ID', () => {
  const model = createBlankDocument(), paragraph = model.blocks[0]
  paragraph.runs[0].text = 'AB'
  const thread = comments.createDocumentThreadComment(model, {
    anchor: { blockId: paragraph.id, offset: 0 }, focus: { blockId: paragraph.id, offset: 2 },
  }, 'Preserve orphan provenance', actor, '2026-01-01T00:00:00.000Z')
  thread.anchor = [
    { blockId: paragraph.id, startOffset: 0, endOffset: 1, text: 'A' },
    { blockId: 'new-paragraph', startOffset: 0, endOffset: 1, text: 'unresolved' },
    { blockId: paragraph.id, startOffset: 1, endOffset: 2, text: 'B' },
  ]
  const source = fields.documentTextFieldParagraphs(model)
  paragraph.runs[0].text = 'A'
  const last = createParagraph('B'); last.id = 'new-paragraph'; model.blocks.push(last)
  comments.transformDocumentCommentAnchorsForBodyRangeEdit(model, source, fields.documentTextFieldParagraphs(model), 1, 1, '\n')
  assert.deepEqual(thread.anchor, [
    { blockId: paragraph.id, startOffset: 0, endOffset: 1, text: 'A' },
    { blockId: last.id, startOffset: 0, endOffset: 1, text: 'unresolved' },
    { blockId: last.id, startOffset: 0, endOffset: 1, text: 'B' },
  ])
  assert.equal(thread.referenceText, 'A\nunresolved\nB')
  assert.equal(comments.documentThreadCommentAnchorStatus(model, thread), 'stale')
})
