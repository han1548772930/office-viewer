import assert from 'node:assert/strict'
import { test, after } from 'node:test'
import { createServer } from 'vite'
import { createBlankDocument } from '../src/domain/model.ts'
import { collapsedSelection } from '../src/features/document/model/document-model.ts'
import { insertLatexFormula } from '../src/features/document/latex/latex-formula.ts'

const vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom' })
const { layoutDocument } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/layout/document-layout.ts')
const { sceneFromDocument } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/rendering/document-scene.ts')
after(() => vite.close())

const measurer = { measure: (text, font) => text.length * (Number(/([\d.]+)px/.exec(font)?.[1] ?? 12) * 0.58) }

test('LaTeX formulas participate in document layout and render math primitives on Canvas', () => {
  const document = createBlankDocument()
  const target = document.blocks[0]
  insertLatexFormula(document, collapsedSelection({ blockId: target.id, offset: 0 }), 'x^2+\\frac{1}{y}')
  const layout = layoutDocument(document, measurer)
  const formulaGlyph = layout.lines.flatMap((line) => line.glyphs).find((glyph) => glyph.latex)
  assert.ok(formulaGlyph)
  assert.ok(formulaGlyph.width > measurer.measure('\uFFFC', formulaGlyph.font))
  assert.ok(formulaGlyph.latex.primitives.some((primitive) => primitive.kind === 'line'))

  const scene = sceneFromDocument(document, layout, { width: 1200, height: 900, scrollLeft: 0, scrollTop: 0, zoom: 100 }, {
    selection: collapsedSelection({ blockId: target.id, offset: 0 }), focused: false, caretVisible: false, composing: false, images: new Map(),
  })
  assert.ok(scene.commands.some((command) => command.kind === 'text' && command.text === 'x'))
  assert.ok(scene.commands.some((command) => command.kind === 'text' && command.text === '2'))
  assert.ok(scene.commands.some((command) => command.kind === 'line'))
  assert.equal(scene.commands.some((command) => command.kind === 'text' && command.text === '\uFFFC'), false)
})
