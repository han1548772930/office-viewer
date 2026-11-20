import assert from 'node:assert/strict'
import { test, after } from 'node:test'
import { createServer } from 'vite'

const vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true }, appType: 'custom' })
const { layoutDocument } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/layout/document-layout.ts')
const { sceneFromDocument } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/rendering/document-scene.ts')

after(() => vite.close())

function paragraph(id, text) {
  return {
    id,
    type: 'paragraph',
    level: 0,
    runs: [{ id: `${id}-run`, text, bold: false, italic: false, underline: false, strike: false, color: '#172033', fontFamily: 'Arial', fontSize: 11 }],
    style: { align: 'left', indent: 0, lineHeight: 1.5, spaceBefore: 0, spaceAfter: 8 },
    metadata: {},
  }
}

function model() {
  return {
    kind: 'document',
    id: 'scene-document',
    name: 'Scene contract',
    blocks: [paragraph('first', 'First page text'), { ...paragraph('break', ''), type: 'pageBreak', runs: [] }, paragraph('second', 'Second page text')],
    page: { width: 794, height: 1123, marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72 },
    modified: false,
    wordCount: 29,
  }
}

function sceneAt(documentModel, scrollTop) {
  const layout = layoutDocument(documentModel, { measure: text => text.length * 8 })
  const selection = { anchor: { blockId: 'first', offset: 2 }, focus: { blockId: 'first', offset: 2 } }
  return sceneFromDocument(documentModel, layout, { width: 900, height: 700, scrollLeft: 0, scrollTop, zoom: 100 }, {
    selection,
    focused: true,
    caretVisible: true,
    composing: false,
    images: new Map(),
  })
}

test('document scene emits Canvas commands for visible page text and caret', () => {
  const scene = sceneAt(model(), 0)
  const text = scene.commands.filter(command => command.kind === 'text').map(command => command.text)
  assert.ok(text.includes('F'))
  assert.ok(!text.includes('S'))
  assert.ok(scene.commands.some(command => command.kind === 'rect' && command.fill === '#24569c'))
  assert.equal(scene.background, '#eef1f5')
})

test('document scene switches page content with scroll without changing the model', () => {
  const documentModel = model()
  const before = JSON.stringify(documentModel)
  const scene = sceneAt(documentModel, 1160)
  const text = scene.commands.filter(command => command.kind === 'text').map(command => command.text)
  assert.ok(text.includes('S'))
  assert.equal(JSON.stringify(documentModel), before)
})
