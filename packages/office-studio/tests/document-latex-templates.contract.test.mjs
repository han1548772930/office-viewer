import assert from 'node:assert/strict'
import { test } from 'vitest'
import { DOCUMENT_LATEX_CATEGORY_LABELS, DOCUMENT_LATEX_TEMPLATE_CATEGORIES, DOCUMENT_LATEX_TEMPLATES,
  getDocumentLatexTemplate, getDocumentLatexTemplates, insertDocumentLatexTemplate, isDocumentLatexTemplateCategory } from '../src/features/document/latex/latex-templates.ts'
import { parseLatexFormula } from '../src/features/document/latex/latex-parser.ts'
import { layoutLatexFormula } from '../src/features/document/latex/latex-layout.ts'
import { documentLatexEnUS, documentLatexLocaleMessages } from '../src/i18n/docs/locales/document-latex.ts'

// Deferred drafts: 71 catalog cases and 10 independent source/catalog contracts.
// Nothing in this batch has executed the parser, test runner or browser.
const measure = (text, font) => Array.from(text).length * Number(/([\d.]+)px/.exec(font)[1]) * 0.5
const categories = ['fractions', 'scripts', 'radicals', 'integrals', 'large-operators', 'brackets', 'functions', 'accents', 'limits-logarithms', 'operators', 'matrices']

for (const entry of DOCUMENT_LATEX_TEMPLATES) {
  test('fixed catalog template ' + entry.id + ' parses, draws and selects its first source placeholder', () => {
    assert.ok(parseLatexFormula(entry.latex))
    const rendered = layoutLatexFormula(entry.previewLatex, { fontSize: 22, measure })
    assert.ok(Number.isFinite(rendered.width) && rendered.width > 0)
    assert.ok(Number.isFinite(rendered.height) && rendered.height > 0)
    assert.ok(rendered.baseline > 0 && rendered.baseline < rendered.height)
    assert.ok(rendered.primitives.length > 0)
    const result = insertDocumentLatexTemplate('', entry.id, 0, 0)
    assert.equal(result.value, entry.latex)
    assert.equal(result.selectionStart, entry.placeholderStart); assert.equal(result.selectionEnd, entry.placeholderEnd)
    assert.ok(result.selectionEnd > result.selectionStart)
    assert.ok(!result.value.includes('⟦') && !result.value.includes('⟧'))
    assert.ok(result.value.slice(result.selectionStart, result.selectionEnd).length > 0)
  })
}

test('catalog category order, counts and immutable identities match the fixed published directory', () => {
  assert.deepEqual(DOCUMENT_LATEX_TEMPLATE_CATEGORIES, categories)
  assert.deepEqual(categories.map(category => getDocumentLatexTemplates(category).length), [4, 4, 3, 6, 6, 8, 8, 8, 6, 12, 6])
  assert.equal(DOCUMENT_LATEX_TEMPLATES.length, 71)
  assert.equal(new Set(DOCUMENT_LATEX_TEMPLATES.map(entry => entry.id)).size, 71)
  assert.ok(Object.isFrozen(DOCUMENT_LATEX_TEMPLATES) && DOCUMENT_LATEX_TEMPLATES.every(Object.isFrozen))
  assert.ok(categories.every(isDocumentLatexTemplateCategory)); assert.equal(isDocumentLatexTemplateCategory('unknown'), false)
  assert.equal(getDocumentLatexTemplate({ id: 'fraction-basic' }), undefined)
  assert.ok(categories.every(category => DOCUMENT_LATEX_CATEGORY_LABELS[category].every(Boolean)))
})

test('representative sources and first placeholders preserve independently recorded upstream examples', () => {
  const expected = [
    ['fraction-derivative', '\\frac{\\mathrm{d}y}{\\mathrm{d}x}', 'y'],
    ['script-left-right', '{}_{a}^{b}c_{d}^{e}', 'c'], ['radical-nth', '\\sqrt[n]{a}', 'a'],
    ['integral-double', '\\iint f(x,y)\\,\\mathrm{d}x\\,\\mathrm{d}y', 'f(x,y)'],
    ['large-sum', '\\sum_{i=1}^{n} a_i', 'a_i'], ['bracket-norm', '\\left\\|a\\right\\|', 'a'],
    ['function-minimum', '\\min\\left(a,b\\right)', 'a'], ['accent-overbrace', '\\overbrace{a+b+\\cdots+z}^{n}', 'a+b+\\cdots+z'],
    ['limit-one-sided', '\\lim_{x \\to a^{+}} f(x)', 'f(x)'], ['operator-not-element', 'a \\notin A', 'a'],
    ['matrix-parenthesized', '\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}', 'a'],
  ]
  for (const [id, source, placeholder] of expected) {
    const entry = getDocumentLatexTemplate(id)
    assert.equal(entry.latex, source); assert.equal(entry.previewLatex, source)
    assert.equal(source.slice(entry.placeholderStart, entry.placeholderEnd), placeholder)
  }
})

test('empty source selection inserts at the real caret and selects the first placeholder', () => {
  const result = insertDocumentLatexTemplate('AB', 'fraction-basic', 1, 1)
  assert.deepEqual(result, { value: 'A\\frac{a}{b}B', selectionStart: 7, selectionEnd: 8 })
})

test('selected source replaces the first placeholder and leaves the cursor after the whole insertion', () => {
  const result = insertDocumentLatexTemplate('Lx+yR', 'fraction-basic', 1, 4)
  const inserted = '\\frac{x+y}{b}'
  assert.deepEqual(result, { value: 'L' + inserted + 'R', selectionStart: 1 + inserted.length, selectionEnd: 1 + inserted.length })
  const derivative = insertDocumentLatexTemplate('f(x)', 'fraction-derivative', 0, 4)
  assert.equal(derivative.value, '\\frac{\\mathrm{d}f(x)}{\\mathrm{d}x}')
})

test('finite reversed and out-of-bounds selections retain the upstream clamping behavior', () => {
  const reversed = insertDocumentLatexTemplate('ABCD', 'fraction-basic', 2, 1)
  assert.equal(reversed.value, 'AB\\frac{a}{b}CD'); assert.deepEqual([reversed.selectionStart, reversed.selectionEnd], [8, 9])
  assert.equal(insertDocumentLatexTemplate('AB', 'operator-plus-minus', -5, 99).value, 'AB \\pm b')
  assert.equal(insertDocumentLatexTemplate('AB', 'operator-plus-minus', 99, 99).value, 'ABa \\pm b')
})

test('invalid template IDs and non-integer source ranges reject before producing a draft', () => {
  assert.throws(() => insertDocumentLatexTemplate('x', 'missing'), RangeError)
  for (const start of [NaN, Infinity, 0.5, Number.MAX_SAFE_INTEGER + 1, '0']) {
    assert.throws(() => insertDocumentLatexTemplate('x', 'fraction-basic', start), TypeError)
  }
  assert.throws(() => insertDocumentLatexTemplate('x', 'fraction-basic', 0, NaN), TypeError)
  assert.throws(() => insertDocumentLatexTemplate(null, 'fraction-basic', 0, 0), TypeError)
})

test('source ranges reject split surrogate pairs and accept whole astral characters', () => {
  assert.throws(() => insertDocumentLatexTemplate('😀x', 'fraction-basic', 1, 2), /UTF-16/)
  assert.throws(() => insertDocumentLatexTemplate('😀x', 'fraction-basic', 0, 1), /UTF-16/)
  const result = insertDocumentLatexTemplate('😀x', 'radical-square', 0, 2)
  assert.equal(result.value, '\\sqrt{😀}x'); assert.equal(result.selectionStart, '\\sqrt{😀}'.length)
  assert.equal(result.selectionStart, result.selectionEnd)
})

test('template replacement keeps selected tabs and newlines and surrounding whitespace unchanged', () => {
  const result = insertDocumentLatexTemplate(' L\r\nx\t+ y\nR ', 'fraction-basic', 4, 9)
  assert.equal(result.value, ' L\r\n\\frac{x\t+ y}{b}\nR ')
  assert.ok(parseLatexFormula(result.value))
})

test('repeated placeholder insertion builds an editable nested source without touching other operands', () => {
  const first = insertDocumentLatexTemplate('', 'fraction-basic', 0, 0)
  const second = insertDocumentLatexTemplate(first.value, 'radical-nth', first.selectionStart, first.selectionEnd)
  assert.equal(second.value, '\\frac{\\sqrt[n]{a}}{b}')
  assert.equal(second.selectionStart, '\\frac{\\sqrt[n]{a}'.length)
  assert.equal(second.selectionStart, second.selectionEnd)
  assert.ok(parseLatexFormula(second.value))
})

test('all category and template names are available in English and three Chinese locale variants', () => {
  for (const messages of [documentLatexEnUS, documentLatexLocaleMessages['zh-CN'], documentLatexLocaleMessages['zh-TW'], documentLatexLocaleMessages['zh-HK']]) {
    for (const category of categories) assert.ok(messages['document.latex.category.' + category])
    for (const entry of DOCUMENT_LATEX_TEMPLATES) assert.ok(messages['document.latex.template.' + entry.id])
    for (const key of ['templates', 'template-category', 'formula-tab', 'insert-template']) assert.ok(messages['document.latex.' + key])
  }
  assert.equal(documentLatexLocaleMessages['zh-CN']['document.latex.template.matrix-parenthesized'], '圆括号矩阵')
  assert.equal(documentLatexLocaleMessages['zh-TW']['document.latex.template.matrix-parenthesized'], '圓括號矩陣')
  assert.equal(documentLatexEnUS['document.latex.template.matrix-parenthesized'], 'Parenthesized matrix')
})
