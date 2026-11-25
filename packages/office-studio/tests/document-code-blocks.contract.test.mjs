import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import {
  cancelEmptyCodeParagraph,
  codeBlockData,
  codeBlockPlainText,
  codeBlocks,
  createCodeBlock,
  exitCodeBlock,
  editCodeBlockText,
  insertCodeBlock,
  normalizeCodeLanguage,
  tokenizeCode,
  unwrapCodeBlock,
  updateCodeBlock,
} from '../src/features/document/code-blocks/code-blocks.ts'
import { collapsedSelection } from '../src/features/document/model/document-model.ts'
import { createDocumentContentControl } from '../src/features/document/services/content-controls.ts'
import { createDocumentThreadComment, documentThreadCommentAnchorStatus } from '../src/features/document/comments/thread-comments.ts'

test('code block preserves source, language and line-number metadata across JSON round trips', () => {
  const block = createCodeBlock('const value = 2\r\nreturn value', 'typescript', true)
  const restored = JSON.parse(JSON.stringify(block))

  assert.deepEqual(codeBlockData(restored), {
    source: 'const value = 2\nreturn value', language: 'typescript', lineNumbers: true, fontFamily: 'Consolas',
  })
  assert.equal(codeBlockPlainText(restored), 'const value = 2\nreturn value')
})

test('inserting, querying and updating code blocks mutates the document model', () => {
  const document = createBlankDocument()
  const anchor = document.blocks[0]
  const block = insertCodeBlock(document, anchor.id, '{"ok":true}', 'json')

  assert.equal(codeBlocks(document).length, 1)
  assert.equal(document.blocks[1].id, block.id)
  assert.equal(updateCodeBlock(document, block.id, { source: '{\n  "ok": false\n}', lineNumbers: true }), true)
  assert.deepEqual(codeBlockData(block), { language: 'json', source: '{\n  "ok": false\n}', lineNumbers: true, fontFamily: 'Consolas' })
  assert.equal(updateCodeBlock(document, 'missing', { source: 'x' }), false)
})

test('tokenization returns source-preserving ranges instead of generated HTML', () => {
  const source = 'const answer = "42" // note'
  const tokens = tokenizeCode(source, 'javascript')

  assert.equal(tokens.map(token => token.text).join(''), source)
  assert.equal(tokens[0].kind, 'keyword')
  assert.ok(tokens.some(token => token.kind === 'string' && token.text === '"42"'))
  assert.ok(tokens.some(token => token.kind === 'comment' && token.text === '// note'))
  assert.ok(tokens.every(token => source.slice(token.start, token.end) === token.text))
})

test('author language values are preserved while aliases and unknown grammars resolve only in the view', () => {
  assert.equal(normalizeCodeLanguage(' TypeScript '), ' TypeScript ')
  assert.equal(normalizeCodeLanguage('unknown-language'), 'unknown-language')
  assert.deepEqual(tokenizeCode('<tag>', 'plain'), [{ kind: 'plain', text: '<tag>', start: 0, end: 5 }])
  assert.deepEqual(tokenizeCode('<tag>', 'unknown-language'), [{ kind: 'plain', text: '<tag>', start: 0, end: 5 }])
})

test('exit adds a normal paragraph and cancel removes only an empty code block', () => {
  const document = createBlankDocument()
  const block = insertCodeBlock(document, document.blocks[0].id, 'x', 'javascript')
  const paragraph = exitCodeBlock(document, block.id)
  assert.ok(paragraph)
  assert.equal(document.blocks[document.blocks.indexOf(block) + 1].id, paragraph.id)

  const empty = insertCodeBlock(document, paragraph.id)
  assert.equal(cancelEmptyCodeParagraph(document, empty.id), true)
  assert.equal(cancelEmptyCodeParagraph(document, block.id), false)
})

test('unwrap converts a code block back into ordinary editable text', () => {
  const block = createCodeBlock('print(1)', 'python')
  assert.equal(unwrapCodeBlock(block), true)
  assert.equal(codeBlockData(block), undefined)
  assert.equal(codeBlockPlainText(block), 'print(1)')
})

test('changing only code presentation preserves source runs and an identical update is a no-op', () => {
  const document = createBlankDocument()
  const block = insertCodeBlock(document, document.blocks[0].id, 'let answer = 1', 'plain')
  const runs = structuredClone(block.runs)
  assert.equal(updateCodeBlock(document, block.id, { language: 'typescript', lineNumbers: true }), true)
  assert.deepEqual(block.runs, runs)
  const before = structuredClone(document)
  assert.equal(updateCodeBlock(document, block.id, { language: 'typescript', lineNumbers: true, source: 'let answer = 1' }), false)
  editCodeBlockText(document, collapsedSelection({ blockId: block.id, offset: 3 }), { text: '' })
  assert.deepEqual(document, before)
})

test('code edits shift content controls and comment anchors through multiline insertion and grapheme deletion', () => {
  const document = createBlankDocument()
  const block = insertCodeBlock(document, document.blocks[0].id, 'a😀b', 'typescript')
  const control = createDocumentContentControl(document, { id: 'code-control', type: 'text', blockId: block.id, startOffset: 3, endOffset: 4 })
  const comment = createDocumentThreadComment(document, {
    anchor: { blockId: block.id, offset: 3 }, focus: { blockId: block.id, offset: 4 },
  }, 'Review b', { documentId: document.id, userId: 'local-user', displayName: 'Editor' }, '2026-01-01T00:00:00.000Z')
  const inserted = editCodeBlockText(document, collapsedSelection({ blockId: block.id, offset: 1 }), { text: '\n\t' })
  assert.equal(inserted.focus.offset, 3)
  assert.equal(codeBlockData(block).source, 'a\n\t😀b')
  assert.deepEqual(document.contentControls.map(({ startOffset, endOffset }) => ({ startOffset, endOffset })), [{ startOffset: 5, endOffset: 6 }])
  assert.deepEqual(comment.anchor[0], { blockId: block.id, startOffset: 5, endOffset: 6, text: 'b' })
  assert.equal(documentThreadCommentAnchorStatus(document, comment), 'current')
  editCodeBlockText(document, inserted, { deleteDirection: 1 })
  assert.equal(codeBlockData(block).source, 'a\n\tb')
  assert.equal(document.contentControls[0].id, control.id)
  assert.deepEqual(comment.anchor[0], { blockId: block.id, startOffset: 3, endOffset: 4, text: 'b' })
  assert.equal(documentThreadCommentAnchorStatus(document, comment), 'current')
})

test('code source and all anchors remain unchanged when a content or wrapper lock rejects the edit', () => {
  for (const lock of ['contentLocked', 'sdtContentLocked', 'controlLocked']) {
    const document = createBlankDocument()
    const block = insertCodeBlock(document, document.blocks[0].id, 'abc', 'plain')
    createDocumentContentControl(document, { type: 'text', blockId: block.id, startOffset: 1, endOffset: 2, lock })
    const before = structuredClone(document)
    const selection = { anchor: { blockId: block.id, offset: 1 }, focus: { blockId: block.id, offset: 2 } }
    assert.throws(() => editCodeBlockText(document, selection, { text: lock === 'controlLocked' ? '' : 'X' }), /locked/)
    assert.deepEqual(document, before)
  }
})

test('full-source updates retain anchors outside the changed substring and enforce locks on the changed substring', () => {
  const document = createBlankDocument()
  const block = insertCodeBlock(document, document.blocks[0].id, 'let name = 1', 'typescript')
  createDocumentContentControl(document, { type: 'text', blockId: block.id, startOffset: 4, endOffset: 8, lock: 'contentLocked' })
  assert.equal(updateCodeBlock(document, block.id, { source: 'let name = 22' }), true)
  assert.deepEqual(document.contentControls.map(({ startOffset, endOffset }) => [startOffset, endOffset]), [[4, 8]])
  const before = structuredClone(document)
  assert.throws(() => updateCodeBlock(document, block.id, { source: 'let changed = 22' }), /locked/)
  assert.deepEqual(document, before)
})
