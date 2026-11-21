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
  insertCodeBlock,
  normalizeCodeLanguage,
  tokenizeCode,
  unwrapCodeBlock,
  updateCodeBlock,
} from '../src/features/document/code-blocks/code-blocks.ts'

test('code block preserves source, language and line-number metadata across JSON round trips', () => {
  const block = createCodeBlock('const value = 2\r\nreturn value', 'typescript', true)
  const restored = JSON.parse(JSON.stringify(block))

  assert.deepEqual(codeBlockData(restored), {
    source: 'const value = 2\nreturn value', language: 'typescript', lineNumbers: true,
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
  assert.deepEqual(codeBlockData(block), { language: 'json', source: '{\n  "ok": false\n}', lineNumbers: true })
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

test('language aliases normalize and unsupported values remain editable as plain source', () => {
  assert.equal(normalizeCodeLanguage(' TypeScript '), 'typescript')
  assert.equal(normalizeCodeLanguage('unknown-language'), 'plain')
  assert.deepEqual(tokenizeCode('<tag>', 'plain'), [{ kind: 'plain', text: '<tag>', start: 0, end: 5 }])
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
