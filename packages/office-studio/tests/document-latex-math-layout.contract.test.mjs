import assert from 'node:assert/strict'
import { test } from 'vitest'
import { LatexFormulaError, parseLatexFormula } from '../src/features/document/latex/latex-parser.ts'
import { latexFormulaDrawCommands, layoutLatexFormula } from '../src/features/document/latex/latex-layout.ts'

// Deferred mathematical layout contracts; no execution/benchmark evidence.
const size = 24
const measure = (text, font) => Array.from(text).length * Number(/([\d.]+)px/.exec(font)[1]) * 0.5
const layout = source => layoutLatexFormula(source, { fontSize: size, measure })
const texts = rendered => rendered.primitives.filter(item => item.kind === 'text')
const glyph = (rendered, value) => texts(rendered).find(item => item.text === value)
const fontSize = item => Number(/([\d.]+)px/.exec(item.font)[1])
function assertEnclosed(rendered) {
  const top = -rendered.baseline, bottom = rendered.height - rendered.baseline
  for (const item of rendered.primitives) {
    if (item.kind === 'text') {
      assert.ok(item.y - fontSize(item) * 0.9 >= top - 1e-8)
      assert.ok(item.y + fontSize(item) * 0.3 <= bottom + 1e-8)
      assert.ok(item.x >= -1e-8 && item.x + measure(item.text, item.font) <= rendered.width + 1e-8)
    } else {
      assert.ok(Math.min(item.y1, item.y2) - item.width / 2 >= top - 1e-8)
      assert.ok(Math.max(item.y1, item.y2) + item.width / 2 <= bottom + 1e-8)
      assert.ok(Math.min(item.x1, item.x2) >= -1e-8 && Math.max(item.x1, item.x2) <= rendered.width + 1e-8)
    }
  }
}

test('matrix rows and columns occupy distinct geometric positions around one baseline', () => {
  const tree = parseLatexFormula('\\begin{matrix}a&b\\\\c&d\\end{matrix}')
  assert.equal(tree.type, 'matrix'); assert.deepEqual(tree.rows.map(row => row.map(cell => cell.value)), [['a', 'b'], ['c', 'd']])
  const rendered = layout('\\begin{matrix}a&b\\\\c&d\\end{matrix}')
  assert.equal(glyph(rendered, 'a').y, glyph(rendered, 'b').y); assert.equal(glyph(rendered, 'c').y, glyph(rendered, 'd').y)
  assert.ok(glyph(rendered, 'a').y < glyph(rendered, 'c').y); assert.ok(glyph(rendered, 'a').x < glyph(rendered, 'b').x)
  assert.equal(glyph(rendered, 'a').x, glyph(rendered, 'c').x); assertEnclosed(rendered)
})

test('ragged matrices and nested fraction cells preserve every cell and derive widths from the largest row', () => {
  const source = '\\begin{bmatrix}\\frac{a}{b}&c\\\\d\\end{bmatrix}', tree = parseLatexFormula(source), rendered = layout(source)
  assert.deepEqual(tree.rows.map(row => row.length), [2, 1]); assert.equal(tree.rows[0][0].type, 'fraction')
  assert.ok(glyph(rendered, '[') && glyph(rendered, ']'))
  for (const value of ['a', 'b', 'c', 'd']) assert.ok(glyph(rendered, value))
  assertEnclosed(rendered)
})

test('paired delimiters grow to enclose tall expressions and invisible sides add no text glyph', () => {
  const tall = layout('\\left(\\frac{a}{\\frac{b}{c}}\\right)'), small = layout('\\left(a\\right)')
  assert.ok(fontSize(glyph(tall, '(')) > fontSize(glyph(small, '('))); assertEnclosed(tall)
  const invisible = layout('\\left.\\frac{a}{b}\\right|')
  assert.deepEqual(texts(invisible).map(item => item.text), ['a', 'b', '|']); assertEnclosed(invisible)
})

test('indexed radicals retain a smaller root index and the same measurable root/rule primitives', () => {
  const tree = parseLatexFormula('\\sqrt[n]{\\frac{a}{b}}'), rendered = layout('\\sqrt[n]{\\frac{a}{b}}')
  assert.equal(tree.type, 'sqrt'); assert.equal(tree.index.value, 'n'); assert.equal(tree.value.type, 'fraction')
  assert.ok(fontSize(glyph(rendered, 'n')) < fontSize(glyph(rendered, 'a')))
  assert.ok(glyph(rendered, '√') && rendered.primitives.some(item => item.kind === 'line')); assertEnclosed(rendered)
})

test('large operator upper and lower limits are centered above and below the base', () => {
  const rendered = layout('\\sum_{i=1}^{n}a_i'), base = glyph(rendered, '∑'), upper = glyph(rendered, 'n'), lower = glyph(rendered, 'i')
  assert.ok(upper.y < base.y && lower.y > base.y)
  assert.ok(Math.abs(upper.x + measure(upper.text, upper.font) / 2 - base.x - measure(base.text, base.font) / 2) < 1e-8)
  assertEnclosed(rendered)
})

test('integral bounds stay beside the integral while differentials retain roman text and explicit spacing', () => {
  const tree = parseLatexFormula('\\int_{a}^{b}f(x)\\,\\mathrm{d}x'), rendered = layout('\\int_{a}^{b}f(x)\\,\\mathrm{d}x')
  assert.equal(tree.children[0].base.limits, false)
  const base = glyph(rendered, '∫'); assert.ok(glyph(rendered, 'a').x >= base.x + measure(base.text, base.font))
  assert.ok(glyph(rendered, 'b').x >= base.x + measure(base.text, base.font)); assert.ok(!glyph(rendered, 'd').font.startsWith('italic'))
  assert.ok(tree.children.some(node => node.type === 'space')); assertEnclosed(rendered)
})

test('limit and conditioned min/max expressions place their complete lower ranges beneath roman operators', () => {
  for (const operator of ['lim', 'min', 'max']) {
    const rendered = layout('\\' + operator + '_{x \\in S}f(x)'), base = glyph(rendered, operator)
    assert.ok(glyph(rendered, 'S').y > base.y); assert.equal(fontSize(base), size); assertEnclosed(rendered)
  }
})

test('overbrace annotations are placed above the brace and leave their source operands intact', () => {
  const rendered = layout('\\overbrace{a+b+\\cdots+z}^{n}'), ruleTop = Math.min(...rendered.primitives.filter(item => item.kind === 'line').flatMap(item => [item.y1, item.y2]))
  assert.ok(glyph(rendered, 'n').y + fontSize(glyph(rendered, 'n')) * 0.3 < ruleTop)
  for (const value of ['a', 'b', '⋯', 'z']) assert.ok(glyph(rendered, value))
  assertEnclosed(rendered)
})

test('all eight accent shapes change geometry without losing the nested operand', () => {
  for (const accent of ['hat', 'bar', 'vec', 'dot', 'ddot', 'overline', 'underline', 'overbrace']) {
    const rendered = layout('\\' + accent + '{a+b}'), bare = layout('a+b')
    assert.ok(rendered.height > bare.height)
    assert.ok(glyph(rendered, 'a') && glyph(rendered, 'b'))
    if (accent === 'dot' || accent === 'ddot') assert.equal(texts(rendered).filter(item => item.text === '·').length, accent === 'dot' ? 1 : 2)
    else assert.ok(rendered.primitives.some(item => item.kind === 'line'))
    assertEnclosed(rendered)
    assertEnclosed(layout('\\' + accent + '{a}'))
  }
})

test('nested font styles and text-group spaces remain distinct from ordinary math spacing', () => {
  const tree = parseLatexFormula('\\mathbf{a+\\mathit{b}}+\\text{two words}'), rendered = layout('\\mathbf{a+\\mathit{b}}+\\text{two words}')
  assert.equal(tree.children[0].style, 'bold'); assert.equal(tree.children.at(-1).style, 'roman')
  assert.equal(tree.children.at(-1).value.children.map(node => node.value).join(''), 'two words')
  assert.ok(glyph(rendered, 'a').font.startsWith('bold ')); assert.ok(glyph(rendered, 'b').font.startsWith('italic ')); assertEnclosed(rendered)
})

test('unpaired delimiters, mismatched environments and matrix separators fail at their source positions', () => {
  for (const source of ['  \\right)', '  \\end{matrix}', '  \\left(a', '  \\begin{bmatrix}a\\end{pmatrix}']) {
    assert.throws(() => parseLatexFormula(source), error => error instanceof LatexFormulaError && error.offset === 2)
  }
  assert.throws(() => parseLatexFormula('a&b'), error => error instanceof LatexFormulaError && error.offset === 1)
  assert.throws(() => parseLatexFormula('a\\\\b'), error => error instanceof LatexFormulaError && error.offset === 1)
  assert.equal(parseLatexFormula('\\&').value, '&')
})

test('source, node and nesting budgets reject oversized input with explicit parse errors', () => {
  assert.throws(() => parseLatexFormula('a'.repeat(65_537)), error => error instanceof LatexFormulaError && error.offset === 65_536)
  assert.throws(() => parseLatexFormula('a'.repeat(16_385)), error => error instanceof LatexFormulaError && error.offset === 16_384)
  assert.throws(() => parseLatexFormula('{'.repeat(65) + 'a' + '}'.repeat(65)), error => error instanceof LatexFormulaError && /嵌套.*预算/.test(error.message))
})

test('invalid font sizes and invalid measurements reject instead of producing non-finite Canvas geometry', () => {
  for (const fontSize of [0, -1, NaN, Infinity]) assert.throws(() => layoutLatexFormula('a', { fontSize, measure }), RangeError)
  for (const value of [NaN, Infinity, -1]) assert.throws(() => layoutLatexFormula('a', { measure: () => value }), RangeError)
  assertEnclosed(layout('\\begin{Vmatrix}a&b\\\\c&d\\end{Vmatrix}'))
})

test('complex formulas share text/line drawing commands with explicit translation and theme color', () => {
  const source = '\\hat{a}+\\begin{pmatrix}b&c\\\\d&e\\end{pmatrix}', base = layout(source)
  const result = latexFormulaDrawCommands(source, 100, 200, { fontSize: size, measure, color: '#123456' })
  assert.equal(result.commands.length, base.primitives.length)
  result.commands.forEach((command, index) => {
    const primitive = base.primitives[index]; assert.equal(command.kind, primitive.kind); assert.equal(command.color, '#123456')
    if (primitive.kind === 'text') assert.deepEqual([command.x, command.y, command.text], [100 + primitive.x, 200 + primitive.y, primitive.text])
    else assert.deepEqual([command.x1, command.y1, command.x2, command.y2], [100 + primitive.x1, 200 + primitive.y1, 100 + primitive.x2, 200 + primitive.y2])
  })
})
