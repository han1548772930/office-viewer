import assert from 'node:assert/strict'
import { test } from 'vitest'
import Prism from 'prismjs'
import { DOCUMENT_STRUCTURED_CODE_LANGUAGES } from '../src/features/document/code-blocks/document-code-languages.ts'
import { documentCodeGrammarLanguage, documentCodeHasGrammar, tokenizeCode } from '../src/features/document/code-blocks/document-code-tokenizer.ts'
import { codeBlockData, createCodeBlock, updateCodeBlock } from '../src/features/document/code-blocks/code-blocks.ts'
import { createBlankDocument } from '../src/domain/model.ts'

// Deferred lexical contracts, never executed in this implementation batch.
// Each catalogue entry has a language-specific sample as well as source partition assertions.
const samples = [
  ['plaintext', 'const plain = "text" // still plain', 'plain', 'const plain'],
  ['javascript', 'const value = 42; // js', 'keyword', 'const'],
  ['typescript', 'interface Item { value: number }; const x: Item = { value: 1 }', 'keyword', 'interface'],
  ['json', '{ "active": true, "value": 42 }', 'keyword', 'true'],
  ['markup', '<div class="item">text</div>', 'keyword', 'div'],
  ['css', '/* css */\na { color: red; width: 10px }', 'comment', '/* css */'],
  ['markdown', '# Heading\n\n**bold** and `code`', 'keyword', 'Heading'],
  ['yaml', 'enabled: true\nvalue: 42 # yaml', 'comment', '# yaml'],
  ['sql', 'SELECT value FROM items WHERE id = 42;', 'keyword', 'SELECT'],
  ['bash', '# bash\nif true; then echo "$HOME"; fi', 'keyword', 'if'],
  ['python', 'def value():\n    return "python" # py', 'keyword', 'def'],
  ['java', 'public class Item { int value = 42; }', 'keyword', 'public'],
  ['c', '#include <stdio.h>\nint main(void) { return 42; }', 'keyword', 'int'],
  ['cpp', 'class Item { public: int value = 42; };', 'keyword', 'class'],
  ['csharp', 'namespace Demo { public class Item { int value = 42; } }', 'keyword', 'namespace'],
  ['go', 'package main\nfunc main() { value := 42 }', 'keyword', 'package'],
  ['rust', 'fn main() { let value: i32 = 42; }', 'keyword', 'fn'],
  ['ruby', 'def value\n  puts "ruby"\nend', 'keyword', 'def'],
  ['php', '<?php function value() { return "php"; } ?>', 'keyword', 'function'],
  ['swift', 'func value() -> Int { let x = 42; return x }', 'keyword', 'func'],
  ['kotlin', 'fun value(): Int { val x = 42; return x }', 'keyword', 'fun'],
]
function partition(source, language) {
  const normalized = source.replace(/\r\n?/g, '\n'), tokens = tokenizeCode(source, language)
  assert.equal(tokens.map(token => token.text).join(''), normalized)
  let offset = 0
  for (const token of tokens) {
    assert.equal(token.start, offset); assert.ok(token.end > token.start)
    assert.equal(normalized.slice(token.start, token.end), token.text)
    assert.ok(['plain', 'keyword', 'comment', 'string', 'number', 'operator', 'punctuation'].includes(token.kind))
    offset = token.end
  }
  assert.equal(offset, normalized.length)
  return tokens
}
for (const [language, source, kind, signature] of samples) test('fixed Code grammar recognizes a real ' + language + ' sample', () => {
  assert.equal(documentCodeHasGrammar(language), true)
  const tokens = partition(source, language)
  assert.ok(tokens.some(token => token.kind === kind && token.text.includes(signature)), language + ' signature: ' + signature)
  if (language === 'plaintext') assert.ok(tokens.every(token => token.kind === 'plain'))
})

test('all fixed language entries have independent semantic samples and explicit lexical-mode configuration', () => {
  assert.deepEqual(samples.map(sample => sample[0]), DOCUMENT_STRUCTURED_CODE_LANGUAGES.map(language => language.value))
  assert.equal(Prism.manual, true); assert.equal(Prism.disableWorkerMessageHandler, true)
})

test('nested JavaScript interpolation flattens expression tokens while retaining surrounding string syntax', () => {
  const source = 'const message = `hello ${1 + value} 😀`', tokens = partition(source, 'javascript')
  assert.ok(tokens.some(token => token.kind === 'string' && token.text.includes('hello')))
  assert.ok(tokens.some(token => token.kind === 'number' && token.text.includes('1')))
  assert.ok(tokens.some(token => token.kind === 'operator' && token.text.includes('+')))
})

test('markup delegates embedded script and style syntax without turning authoritative source into HTML', () => {
  const source = '<style>/* css */ a { color: red }</style><script>const value = 42;</script>', tokens = partition(source, 'markup')
  assert.ok(tokens.some(token => token.kind === 'comment' && token.text.includes('/* css */')))
  assert.ok(tokens.some(token => token.kind === 'keyword' && token.text.includes('const')))
  assert.ok(tokens.some(token => token.kind === 'number' && token.text.includes('42')))
})

test('multiline and escaped literals retain every character after line ending normalization', () => {
  const source = 'value = """first\r\nsecond 😀"""\r\n# last', tokens = partition(source, 'python')
  assert.ok(tokens.some(token => token.kind === 'string' && token.text.includes('first\nsecond 😀')))
  assert.ok(tokens.some(token => token.kind === 'comment' && token.text.includes('# last')))
  partition('const value = "a\\\"b\\nc";', 'javascript')
})

test('UTF-16, combining marks, tabs and empty source remain a complete partition in every language', () => {
  for (const language of DOCUMENT_STRUCTURED_CODE_LANGUAGES) {
    const source = '😀e\u0301\t"中😀"\r\n'
    const tokens = partition(source, language.value)
    const normalized = source.replace(/\r\n?/g, '\n')
    assert.ok(tokens.every(token => !/[\uD800-\uDBFF]$/.test(normalized.slice(0, token.end))))
    assert.deepEqual(tokenizeCode('', language.value), [])
  }
})

test('language aliases are lexical-only and unsupported author values survive editing and serialization', () => {
  for (const [alias, language] of [[' TypeScript ', 'typescript'], ['HTML', 'markup'], ['C++', 'cpp'], ['c#', 'csharp'], ['yml', 'yaml']]) {
    assert.equal(documentCodeGrammarLanguage(alias), language)
    assert.deepEqual(tokenizeCode('x = 42;', alias), tokenizeCode('x = 42;', language))
  }
  for (const author of ['constructor', 'toString', '__proto__']) {
    assert.equal(documentCodeGrammarLanguage(author), undefined)
    assert.equal(documentCodeHasGrammar(author), false)
    assert.deepEqual(tokenizeCode('x = 42;', author), [{ kind: 'plain', text: 'x = 42;', start: 0, end: 7 }])
  }
  const model = createBlankDocument(), author = ' Vendor.Language / 2026 ', block = createCodeBlock('a😀b', author)
  block.metadata.codeBlock.authorExtension = { vendor: 'retained' }; model.blocks = [block]
  assert.equal(documentCodeHasGrammar(author), false)
  assert.deepEqual(partition('a😀b', author), [{ kind: 'plain', text: 'a😀b', start: 0, end: 4 }])
  assert.equal(updateCodeBlock(model, block.id, { lineNumbers: true }), true)
  assert.equal(codeBlockData(model.blocks[0]).language, author)
  assert.deepEqual(model.blocks[0].metadata.codeBlock.authorExtension, { vendor: 'retained' })
  assert.equal(codeBlockData(JSON.parse(JSON.stringify(model)).blocks[0]).language, author)
})
