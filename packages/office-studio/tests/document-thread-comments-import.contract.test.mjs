import assert from 'node:assert/strict'
import { test, afterAll } from 'vitest'
import { createServer } from 'vite'

const vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom' })
const { normalizeDocument } = await vite.ssrLoadModule('/packages/office-studio/src/parser/documentNormalizer.ts')
const { documentThreadCommentAnchorStatus } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/comments/thread-comments.ts')
afterAll(() => vite.close())

test('DOCX thread comment payloads map source range targets into the normalized document model', () => {
  const model = normalizeDocument({
    schemaVersion: 1,
    parserVersion: 'test',
    format: 'docx',
    fileName: 'comments.docx',
    size: 1,
    model: 'document',
    payload: {
      modelType: 'document',
      data: {
        documentId: 'doc-1',
        blocks: [{
          id: 'normalized-paragraph-1',
          type: 'text',
          content: 'Hello World',
          sourceRefs: [{ nodeId: 'source-paragraph-1', story: 'body' }],
          details: {
            inlines: [{ type: 'text', text: 'Hello World', style: {} }],
            style: {},
            numbering: {},
          },
        }],
        comments: [
          {
            id: '7',
            author: 'Alice',
            date: '2026-01-01T00:00:00Z',
            text: 'Please review this word.',
            resolved: true,
            targets: [{ nodeId: 'source-paragraph-1', story: 'body', start: 6, end: 11, offsetEncoding: 'utf-16' }],
          },
          {
            id: '8',
            parentId: '7',
            author: 'Bob',
            date: '2026-01-02T00:00:00Z',
            text: 'Updated.',
            targets: [],
          },
        ],
      },
    },
    diagnostics: { warnings: [], errors: [], stats: {}, capabilities: [] },
  })

  const thread = model.comments['docx-comment-7']
  assert.equal(thread.referenceText, 'World')
  assert.deepEqual(thread.anchor[0], { blockId: 'normalized-paragraph-1', startOffset: 6, endOffset: 11, text: 'World' })
  assert.equal(thread.messages.length, 2)
  assert.equal(thread.messages[1].authorName, 'Bob')
  assert.equal(thread.resolved, true)
  assert.equal(documentThreadCommentAnchorStatus(model, thread), 'current')
})
