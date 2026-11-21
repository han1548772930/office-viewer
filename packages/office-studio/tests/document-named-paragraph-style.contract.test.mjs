import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

let server
let commands
let modelUtils
let domain

beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  commands = await server.ssrLoadModule('/packages/office-studio/src/features/document/commands/document-commands.ts')
  modelUtils = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
  domain = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
})

afterAll(async () => { await server?.close() })

function paragraph(id, text) { return { ...modelUtils.createParagraph(text), id } }
function documentWith(blocks) { return { ...domain.createBlankDocument(), id: 'named-style-doc', blocks } }
function selection(from, to) { return { anchor: { blockId: from, offset: 0 }, focus: { blockId: to, offset: 1 } } }

test('named paragraph styles update semantic heading levels, visual runs, and spacing across selected paragraphs', () => {
  const first = paragraph('first', 'Title')
  const second = paragraph('second', 'Section')
  const model = documentWith([first, second])

  commands.executeDocumentCommand(model, selection('first', 'second'), {
    id: 'doc.command.set-paragraph-named-style', namedStyle: 'heading2',
  })

  assert.deepEqual(model.blocks.map((block) => ({ type: block.type, level: block.level, namedStyle: block.metadata.namedStyleType })), [
    { type: 'heading', level: 2, namedStyle: 'heading2' },
    { type: 'heading', level: 2, namedStyle: 'heading2' },
  ])
  assert.deepEqual(first.runs.map(({ text, bold, fontSize }) => ({ text, bold, fontSize })), [{ text: 'Title', bold: true, fontSize: 18 }])
  assert.equal(first.style.spaceBefore, 24)
  assert.equal(first.style.spaceAfter, 13.3)
})

test('title and subtitle remain outside the heading outline and persist in JSON snapshots', () => {
  const model = documentWith([paragraph('title', 'Document'), paragraph('subtitle', 'Summary')])
  commands.executeDocumentCommand(model, selection('title', 'title'), {
    id: 'doc.command.set-paragraph-named-style', namedStyle: 'title',
  })
  commands.executeDocumentCommand(model, selection('subtitle', 'subtitle'), {
    id: 'doc.command.set-paragraph-named-style', namedStyle: 'subtitle',
  })

  const restored = JSON.parse(JSON.stringify(model))
  assert.deepEqual(restored.blocks.map((block) => [block.type, block.level, block.metadata.namedStyleType]), [
    ['paragraph', 0, 'title'],
    ['paragraph', 0, 'subtitle'],
  ])
  assert.equal(restored.blocks[0].runs[0].fontSize, 26)
  assert.equal(restored.blocks[1].runs[0].color, '#999999')
})

test('normal style resets heading semantics and visual defaults', () => {
  const heading = paragraph('heading', 'Body again')
  commands.executeDocumentCommand(documentWith([heading]), selection('heading', 'heading'), {
    id: 'doc.command.set-paragraph-named-style', namedStyle: 'heading1',
  })
  commands.executeDocumentCommand(documentWith([heading]), selection('heading', 'heading'), {
    id: 'doc.command.set-paragraph-named-style', namedStyle: 'normal',
  })

  assert.equal(heading.type, 'paragraph')
  assert.equal(heading.level, 0)
  assert.equal(heading.metadata.namedStyleType, 'normal')
  assert.deepEqual(heading.runs.map(({ bold, fontSize, color }) => ({ bold, fontSize, color })), [
    { bold: false, fontSize: 11, color: '#1f2937' },
  ])
})
