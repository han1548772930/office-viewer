import assert from 'node:assert/strict'
import { test } from 'node:test'
import { createBlankDocument } from '../src/domain/model.ts'
import { collapsedSelection, createParagraph } from '../src/features/document/model/document-model.ts'
import {
  createLatexFormulaRun, getLatexFormula, getLatexFormulas, insertLatexFormula,
  latexFormulaDrawCommands, layoutLatexFormula, LatexFormulaError, parseLatexFormula, updateLatexFormula,
} from '../src/features/document/latex/latex-formula.ts'

test('LaTeX parser supports grouped fractions, roots, scripts, and common symbols', () => {
  const tree = parseLatexFormula('\\frac{x_1}{\\sqrt{y}} + \\alpha')
  assert.equal(tree.type, 'sequence')
  assert.throws(() => parseLatexFormula('\\unknown{x}'), LatexFormulaError)
  assert.throws(() => parseLatexFormula('\\frac{a}{b'), /未闭合/)
})

test('unsupported syntax reports its source offset instead of silently rendering source text', () => {
  assert.throws(() => parseLatexFormula('x + \\begin{matrix}'), (error) => error instanceof LatexFormulaError && error.offset === 4)
  assert.throws(() => parseLatexFormula('x^'), /缺少内容/)
})

test('formula renderer produces measurable text and line primitives for nested expressions', () => {
  const rendered = layoutLatexFormula('\\frac{a+b}{x_2} + \\sqrt{n}', { fontSize: 18, measure: (text, font) => text.length * Number(/([\d.]+)px/.exec(font)?.[1] ?? 18) / 2 })
  assert.ok(rendered.width > 0)
  assert.ok(rendered.height > 18)
  assert.ok(rendered.primitives.some((primitive) => primitive.kind === 'line'))
  assert.ok(rendered.primitives.some((primitive) => primitive.kind === 'text' && primitive.text === '√'))
  const drawing = latexFormulaDrawCommands('a^2', 12, 30, { color: '#223344' })
  assert.ok(drawing.commands.every((command) => command.kind === 'text' && command.color === '#223344'))
})

test('inline formula insertion is a serializable document run and can be updated', () => {
  const document = createBlankDocument()
  const paragraph = createParagraph('E = ')
  document.blocks = [paragraph]
  const point = { blockId: paragraph.id, offset: 4 }
  const caret = insertLatexFormula(document, collapsedSelection(point), 'mc^2')
  const formulas = getLatexFormulas(document)
  assert.equal(formulas.length, 1)
  assert.equal(formulas[0].formula.mode, 'inline')
  assert.deepEqual(getLatexFormula(paragraph.runs.at(-1)), formulas[0].formula)
  assert.equal(caret.offset, 5)
  assert.equal(updateLatexFormula(document, formulas[0].runId, 'E=mc^2'), true)
  assert.equal(getLatexFormulas(document)[0].formula.latex, 'E=mc^2')
  const restored = JSON.parse(JSON.stringify(document))
  assert.equal(getLatexFormulas(restored)[0].formula.latex, 'E=mc^2')
})

test('block formula insertion isolates the selected text in surrounding paragraphs', () => {
  const document = createBlankDocument()
  const paragraph = createParagraph('before selected after')
  document.blocks = [paragraph]
  insertLatexFormula(document, { anchor: { blockId: paragraph.id, offset: 7 }, focus: { blockId: paragraph.id, offset: 15 } }, 'x^2', 'block')
  assert.deepEqual(document.blocks.map((block) => block.metadata.latexFormulaBlock === true), [false, true, false])
  assert.equal(getLatexFormulas(document)[0].formula.mode, 'block')
  assert.equal(document.blocks[0].runs.map((run) => run.text).join(''), 'before ')
  assert.equal(document.blocks[2].runs.map((run) => run.text).join(''), ' after')
})

test('formula run validates source before entering document data', () => {
  const run = createLatexFormulaRun('\\sum_{i=1}^{n} i')
  assert.equal(run.text, '\uFFFC')
  assert.equal(getLatexFormula(run).latex, '\\sum_{i=1}^{n} i')
  assert.throws(() => createLatexFormulaRun('\\badcommand'), /暂不支持/)
})
