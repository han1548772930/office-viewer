import assert from 'node:assert/strict'
import { test, afterAll } from 'vitest'
import { createServer } from 'vite'

const vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom' })
const { hitTestChecklist, layoutDocument } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/layout/document-layout.ts')
const { sceneFromDocument } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/rendering/document-scene.ts')

afterAll(() => vite.close())

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

test('document scene highlights every search match on visible lines', () => {
  const documentModel = model()
  const layout = layoutDocument(documentModel, { measure: text => text.length * 8 })
  const selection = { anchor: { blockId: 'first', offset: 2 }, focus: { blockId: 'first', offset: 2 } }
  const scene = sceneFromDocument(documentModel, layout, { width: 900, height: 700, scrollLeft: 0, scrollTop: 0, zoom: 100 }, {
    selection,
    focused: false,
    caretVisible: false,
    composing: false,
    images: new Map(),
    searchMatches: [
      { target: { blockId: 'first' }, start: 0, end: 5 },
      { target: { blockId: 'first' }, start: 6, end: 10 },
    ],
  })
  assert.equal(scene.commands.filter(command => command.kind === 'rect' && command.fill === '#ffe58f').length, 9)
})

test('document scene switches page content with scroll without changing the model', () => {
  const documentModel = model()
  const before = JSON.stringify(documentModel)
  const scene = sceneAt(documentModel, 1160)
  const text = scene.commands.filter(command => command.kind === 'text').map(command => command.text)
  assert.ok(text.includes('S'))
  assert.equal(JSON.stringify(documentModel), before)
})

test('section and column break types place following content on the required page', () => {
  const cases = [
    ['nextPage', 1],
    ['continuous', 0],
    ['nextColumn', 1],
    ['evenPage', 1],
    ['oddPage', 2],
  ]
  for (const [type, expectedPage] of cases) {
    const documentModel = model()
    documentModel.blocks = [
      paragraph('before', 'Before'),
      { ...paragraph('section-break', ''), type: 'sectionBreak', runs: [], sectionBreak: { type } },
      paragraph('after', 'After'),
    ]
    const layout = layoutDocument(documentModel, { measure: text => text.length * 8 })
    assert.equal(layout.lines.find(line => line.target.blockId === 'after').page, expectedPage, type)
  }

  const documentModel = model()
  documentModel.blocks = [paragraph('before-column', 'Before'), { ...paragraph('column-break', ''), type: 'columnBreak', runs: [] }, paragraph('after-column', 'After')]
  const layout = layoutDocument(documentModel, { measure: text => text.length * 8 })
  assert.equal(layout.lines.find(line => line.target.blockId === 'after-column').page, 1)
})

test('checklist state is laid out as a hit-testable Canvas control and changes only its scene commands', () => {
  const documentModel = model()
  const checklist = paragraph('task', 'Review the changes')
  checklist.type = 'checklist'
  checklist.checklist = { listId: 'review', checked: false }
  documentModel.blocks = [checklist]
  const selection = { anchor: { blockId: 'task', offset: 0 }, focus: { blockId: 'task', offset: 0 } }
  const layout = layoutDocument(documentModel, { measure: text => text.length * 8 })
  const line = layout.lines.find(candidate => candidate.target.blockId === 'task')
  assert.equal(line.checklist.checked, false)
  assert.equal(hitTestChecklist(layout, line.x - 15, line.y + line.height / 2), 'task')
  const scene = sceneFromDocument(documentModel, layout, { width: 900, height: 700, scrollLeft: 0, scrollTop: 0, zoom: 100 }, { selection, focused: false, caretVisible: false, composing: false, images: new Map() })
  assert.ok(scene.commands.some(command => command.kind === 'rect' && command.stroke === '#718096'))
  documentModel.blocks[0].checklist.checked = true
  const checkedScene = sceneFromDocument(documentModel, layoutDocument(documentModel, { measure: text => text.length * 8 }), { width: 900, height: 700, scrollLeft: 0, scrollTop: 0, zoom: 100 }, { selection, focused: false, caretVisible: false, composing: false, images: new Map() })
  assert.ok(checkedScene.commands.some(command => command.kind === 'rect' && command.stroke === '#28744b'))
  assert.equal(checkedScene.commands.filter(command => command.kind === 'line' && command.color === '#28744b').length, 2)
})

test('subscript and superscript use reduced glyph metrics and distinct Canvas baselines without changing line height', () => {
  const documentModel = model()
  const formula = paragraph('formula', 'H2O')
  formula.runs = [
    { ...formula.runs[0], id: 'h', text: 'H', verticalAlign: 'baseline' },
    { ...formula.runs[0], id: 'two', text: '2', verticalAlign: 'subscript' },
    { ...formula.runs[0], id: 'o', text: 'O', verticalAlign: 'superscript' },
  ]
  documentModel.blocks = [formula]
  const measure = (text, font) => text.length * Number(font.match(/([\d.]+)px/)[1])
  const layout = layoutDocument(documentModel, { measure })
  const glyphs = layout.lines[0].glyphs
  const fontSize = glyph => Number(glyph.font.match(/([\d.]+)px/)[1])
  assert.ok(Math.abs(fontSize(glyphs[0]) - glyphs[0].run.fontSize * 4 / 3) < 0.001)
  assert.ok(fontSize(glyphs[1]) < fontSize(glyphs[0]))
  assert.equal(layout.lines[0].height, layoutDocument({ ...documentModel, blocks: [paragraph('plain', 'H2O')] }, { measure }).lines[0].height)
  const selection = { anchor: { blockId: 'formula', offset: 0 }, focus: { blockId: 'formula', offset: 0 } }
  const scene = sceneFromDocument(documentModel, layout, { width: 900, height: 700, scrollLeft: 0, scrollTop: 0, zoom: 100 }, { selection, focused: false, caretVisible: false, composing: false, images: new Map() })
  const textCommands = scene.commands.filter(command => command.kind === 'text' && ['H', '2', 'O'].includes(command.text))
  assert.equal(textCommands.length, 3)
  assert.ok(textCommands[1].y > textCommands[0].y)
  assert.ok(textCommands[2].y < textCommands[0].y)
})
