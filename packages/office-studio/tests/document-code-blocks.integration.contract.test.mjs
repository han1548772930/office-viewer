import assert from 'node:assert/strict'
import { test, afterAll } from 'vitest'
import { createServer } from 'vite'
import { createBlankDocument } from '../src/domain/model.ts'
import { collapsedSelection } from '../src/features/document/model/document-model.ts'
import { createCodeBlock, editCodeBlockText, insertCodeBlock, codeBlockData } from '../src/features/document/code-blocks/code-blocks.ts'

const vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom' })
const { layoutDocument } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/layout/document-layout.ts')
const { sceneFromDocument } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/rendering/document-scene.ts')
afterAll(() => vite.close())

const measurer = { measure: (text, font) => text.length * (Number(/([\d.]+)px/.exec(font)?.[1] ?? 12) * 0.58) }

test('code editing preserves source offsets, inserts newlines and deletes graphemes atomically', () => {
  const document = createBlankDocument()
  const block = insertCodeBlock(document, document.blocks[0].id, 'a😀b', 'typescript')
  const inserted = editCodeBlockText(document, collapsedSelection({ blockId: block.id, offset: 1 }), { text: '\nconst value = 1' })
  assert.equal(inserted.focus.offset, 17)
  assert.equal(codeBlockData(block).source, 'a\nconst value = 1😀b')

  const deleted = editCodeBlockText(document, collapsedSelection({ blockId: block.id, offset: 17 }), { deleteDirection: -1 })
  assert.equal(deleted.focus.offset, 16)
  assert.equal(codeBlockData(block).source, 'a\nconst value = 😀b')
})

test('code editing refuses ranges crossing block boundaries', () => {
  const document = createBlankDocument()
  const paragraph = document.blocks[0]
  const code = insertCodeBlock(document, paragraph.id, 'const value = 1', 'typescript')
  const selection = { anchor: { blockId: paragraph.id, offset: 0 }, focus: { blockId: code.id, offset: 3 } }
  assert.equal(editCodeBlockText(document, selection, { text: 'x' }), undefined)
  assert.equal(codeBlockData(code).source, 'const value = 1')
})

test('document layout and Canvas scene render code rows, line numbers and syntax token colors', () => {
  const document = createBlankDocument()
  const block = insertCodeBlock(document, document.blocks[0].id, 'const answer = 42\nreturn answer', 'typescript')
  block.metadata.codeBlock.lineNumbers = true
  const layout = layoutDocument(document, measurer)
  const codeLines = layout.lines.filter((line) => line.target.blockId === block.id)
  assert.deepEqual(codeLines.map((line) => line.codeBlock?.lineNumber), [1, 2])

  const selection = collapsedSelection({ blockId: block.id, offset: 0 })
  const scene = sceneFromDocument(document, layout, { width: 1200, height: 900, scrollLeft: 0, scrollTop: 0, zoom: 100 }, {
    selection, focused: false, caretVisible: false, composing: false, images: new Map(),
  })
  assert.ok(scene.commands.some((command) => command.kind === 'rect' && command.fill === '#f1f4f8'))
  assert.ok(scene.commands.some((command) => command.kind === 'text' && command.text === '1' && command.color === '#8a96a5'))
  const keywordGlyphs = scene.commands
    .filter((command) => command.kind === 'text' && command.color === '#7c3aed')
    .map((command) => command.text)
    .join('')
  assert.ok(keywordGlyphs.includes('const'))
  const numberGlyphs = scene.commands
    .filter((command) => command.kind === 'text' && command.color === '#b45309')
    .map((command) => command.text)
    .join('')
  assert.ok(numberGlyphs.includes('42'))
})
