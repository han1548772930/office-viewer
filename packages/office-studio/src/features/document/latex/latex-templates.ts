import { isDocumentTextBoundary } from '../model/document-story'

export const DOCUMENT_LATEX_TEMPLATE_CATEGORIES = [
  'fractions', 'scripts', 'radicals', 'integrals', 'large-operators', 'brackets', 'functions', 'accents', 'limits-logarithms', 'operators', 'matrices',
] as const
export type DocumentLatexTemplateCategory = typeof DOCUMENT_LATEX_TEMPLATE_CATEGORIES[number]
export const DOCUMENT_LATEX_TEMPLATE_OPERATION = 'docs-latex.operation.insert-formula-template'
export const DOCUMENT_LATEX_TEMPLATE_LOCAL_OPERATION = 'doc.operation.insert-latex-template'
export interface DocumentLatexTemplate {
  readonly id: string
  readonly category: DocumentLatexTemplateCategory
  readonly label: string
  readonly englishLabel: string
  readonly latex: string
  readonly previewLatex: string
  readonly placeholderStart: number
  readonly placeholderEnd: number
}
export interface DocumentLatexTemplateInsertion { value: string; selectionStart: number; selectionEnd: number }
export const DOCUMENT_LATEX_CATEGORY_LABELS: Record<DocumentLatexTemplateCategory, readonly [string, string]> = {
  fractions: ['分式', 'Fractions'], scripts: ['上下标', 'Scripts'], radicals: ['根式', 'Radicals'], integrals: ['积分', 'Integrals'],
  'large-operators': ['大型运算符', 'Large operators'], brackets: ['括号', 'Brackets'], functions: ['函数', 'Functions'],
  accents: ['重音', 'Accents'], 'limits-logarithms': ['极限与对数', 'Limits and logarithms'], operators: ['运算符', 'Operators'], matrices: ['矩阵', 'Matrices'],
}
function template(category: DocumentLatexTemplateCategory, id: string, marked: string, label: string, englishLabel: string): DocumentLatexTemplate {
  const start = marked.indexOf('⟦'), end = marked.indexOf('⟧')
  if (start < 0 || end <= start + 1 || marked.indexOf('⟦', start + 1) >= 0 || marked.indexOf('⟧', end + 1) >= 0) throw new TypeError('公式模板必须包含一个非空占位范围。')
  const latex = marked.slice(0, start) + marked.slice(start + 1, end) + marked.slice(end + 1)
  return Object.freeze({ category, id, label, englishLabel, latex, previewLatex: latex, placeholderStart: start, placeholderEnd: end - 1 })
}
/** Observable v1.6.0 catalog: same IDs, category order, source and first placeholder. */
export const DOCUMENT_LATEX_TEMPLATES: readonly DocumentLatexTemplate[] = Object.freeze([
  template('fractions', 'fraction-basic', '\\frac{⟦a⟧}{b}', '基本分式', 'Basic fraction'),
  template('fractions', 'fraction-expression', '\\frac{⟦a+b⟧}{c+d}', '表达式分式', 'Expression fraction'),
  template('fractions', 'fraction-derivative', '\\frac{\\mathrm{d}⟦y⟧}{\\mathrm{d}x}', '导数分式', 'Derivative fraction'),
  template('fractions', 'fraction-partial', '\\frac{\\partial ⟦f⟧}{\\partial x}', '偏导数分式', 'Partial derivative'),
  template('scripts', 'script-superscript', '⟦a⟧^{b}', '上标', 'Superscript'),
  template('scripts', 'script-subscript', '⟦a⟧_{b}', '下标', 'Subscript'),
  template('scripts', 'script-combined', '⟦a⟧_{b}^{c}', '上下标', 'Combined scripts'),
  template('scripts', 'script-left-right', '{}_{a}^{b}⟦c⟧_{d}^{e}', '左右上下标', 'Left and right scripts'),
  template('radicals', 'radical-square', '\\sqrt{⟦a⟧}', '平方根', 'Square root'),
  template('radicals', 'radical-nth', '\\sqrt[n]{⟦a⟧}', 'n 次方根', 'Nth root'),
  template('radicals', 'radical-nested', '\\sqrt{\\sqrt{⟦a⟧}}', '嵌套根式', 'Nested radical'),
  template('integrals', 'integral-basic', '\\int ⟦f(x)⟧\\,\\mathrm{d}x', '不定积分', 'Integral'),
  template('integrals', 'integral-bounded', '\\int_{a}^{b} ⟦f(x)⟧\\,\\mathrm{d}x', '定积分', 'Bounded integral'),
  template('integrals', 'integral-double', '\\iint ⟦f(x,y)⟧\\,\\mathrm{d}x\\,\\mathrm{d}y', '二重积分', 'Double integral'),
  template('integrals', 'integral-triple', '\\iiint ⟦f(x,y,z)⟧\\,\\mathrm{d}V', '三重积分', 'Triple integral'),
  template('integrals', 'integral-contour', '\\oint ⟦f(z)⟧\\,\\mathrm{d}z', '环路积分', 'Contour integral'),
  template('integrals', 'integral-contour-bounded', '\\oint_{C} ⟦f(z)⟧\\,\\mathrm{d}z', '带路径积分', 'Contour integral with path'),
  template('large-operators', 'large-sum', '\\sum_{i=1}^{n} ⟦a_i⟧', '求和', 'Summation'),
  template('large-operators', 'large-product', '\\prod_{i=1}^{n} ⟦a_i⟧', '乘积', 'Product'),
  template('large-operators', 'large-coproduct', '\\coprod_{i=1}^{n} ⟦A_i⟧', '余积', 'Coproduct'),
  template('large-operators', 'large-union', '\\bigcup_{i=1}^{n} ⟦A_i⟧', '大型并集', 'Large union'),
  template('large-operators', 'large-intersection', '\\bigcap_{i=1}^{n} ⟦A_i⟧', '大型交集', 'Large intersection'),
  template('large-operators', 'large-logical-product', '\\bigwedge_{i=1}^{n} ⟦P_i⟧', '大型逻辑与', 'Large logical conjunction'),
  template('brackets', 'bracket-parentheses', '\\left(⟦a⟧\\right)', '圆括号', 'Parentheses'),
  template('brackets', 'bracket-square', '\\left[⟦a⟧\\right]', '方括号', 'Square brackets'),
  template('brackets', 'bracket-braces', '\\left\\{⟦a⟧\\right\\}', '花括号', 'Braces'),
  template('brackets', 'bracket-angle', '\\left\\langle ⟦a⟧\\right\\rangle', '尖括号', 'Angle brackets'),
  template('brackets', 'bracket-absolute', '\\left|⟦a⟧\\right|', '绝对值', 'Absolute value'),
  template('brackets', 'bracket-norm', '\\left\\|⟦a⟧\\right\\|', '范数', 'Norm'),
  template('brackets', 'bracket-floor', '\\left\\lfloor ⟦a⟧\\right\\rfloor', '向下取整', 'Floor'),
  template('brackets', 'bracket-ceiling', '\\left\\lceil ⟦a⟧\\right\\rceil', '向上取整', 'Ceiling'),
  template('functions', 'function-sine', '\\sin\\left(⟦a⟧\\right)', '正弦', 'Sine'),
  template('functions', 'function-cosine', '\\cos\\left(⟦a⟧\\right)', '余弦', 'Cosine'),
  template('functions', 'function-tangent', '\\tan\\left(⟦a⟧\\right)', '正切', 'Tangent'),
  template('functions', 'function-logarithm', '\\log\\left(⟦a⟧\\right)', '对数', 'Logarithm'),
  template('functions', 'function-natural-logarithm', '\\ln\\left(⟦a⟧\\right)', '自然对数', 'Natural logarithm'),
  template('functions', 'function-exponential', '\\exp\\left(⟦a⟧\\right)', '指数函数', 'Exponential'),
  template('functions', 'function-minimum', '\\min\\left(⟦a⟧,b\\right)', '最小值', 'Minimum'),
  template('functions', 'function-maximum', '\\max\\left(⟦a⟧,b\\right)', '最大值', 'Maximum'),
  template('accents', 'accent-hat', '\\hat{⟦a⟧}', '帽重音', 'Hat accent'),
  template('accents', 'accent-bar', '\\bar{⟦a⟧}', '横线重音', 'Bar accent'),
  template('accents', 'accent-vector', '\\vec{⟦a⟧}', '向量箭头', 'Vector accent'),
  template('accents', 'accent-dot', '\\dot{⟦a⟧}', '点重音', 'Dot accent'),
  template('accents', 'accent-double-dot', '\\ddot{⟦a⟧}', '双点重音', 'Double dot accent'),
  template('accents', 'accent-overline', '\\overline{⟦a+b⟧}', '上划线', 'Overline'),
  template('accents', 'accent-underline', '\\underline{⟦a+b⟧}', '下划线', 'Underline'),
  template('accents', 'accent-overbrace', '\\overbrace{⟦a+b+\\cdots+z⟧}^{n}', '上花括号', 'Overbrace'),
  template('limits-logarithms', 'limit-basic', '\\lim_{x \\to a} ⟦f(x)⟧', '基本极限', 'Limit'),
  template('limits-logarithms', 'limit-one-sided', '\\lim_{x \\to a^{+}} ⟦f(x)⟧', '单侧极限', 'One-sided limit'),
  template('limits-logarithms', 'limit-infinity', '\\lim_{x \\to \\infty} ⟦f(x)⟧', '无穷极限', 'Limit at infinity'),
  template('limits-logarithms', 'logarithm-base', '\\log_{a}\\left(⟦b⟧\\right)', '指定底数对数', 'Logarithm with base'),
  template('limits-logarithms', 'minimum-conditioned', '\\min_{x \\in S} ⟦f(x)⟧', '条件最小值', 'Conditioned minimum'),
  template('limits-logarithms', 'maximum-conditioned', '\\max_{x \\in S} ⟦f(x)⟧', '条件最大值', 'Conditioned maximum'),
  template('operators', 'operator-plus-minus', '⟦a⟧ \\pm b', '正负号', 'Plus or minus'),
  template('operators', 'operator-multiply', '⟦a⟧ \\times b', '乘号', 'Multiplication'),
  template('operators', 'operator-divide', '⟦a⟧ \\div b', '除号', 'Division'),
  template('operators', 'operator-not-equal', '⟦a⟧ \\ne b', '不等于', 'Not equal'),
  template('operators', 'operator-approximately', '⟦a⟧ \\approx b', '约等于', 'Approximately equal'),
  template('operators', 'operator-less-equal', '⟦a⟧ \\le b', '小于等于', 'Less than or equal'),
  template('operators', 'operator-greater-equal', '⟦a⟧ \\ge b', '大于等于', 'Greater than or equal'),
  template('operators', 'operator-element', '⟦a⟧ \\in A', '属于', 'Element of'),
  template('operators', 'operator-not-element', '⟦a⟧ \\notin A', '不属于', 'Not an element of'),
  template('operators', 'operator-subset', '⟦A⟧ \\subseteq B', '子集', 'Subset'),
  template('operators', 'operator-union', '⟦A⟧ \\cup B', '并集', 'Union'),
  template('operators', 'operator-intersection', '⟦A⟧ \\cap B', '交集', 'Intersection'),
  template('matrices', 'matrix-2x2', '\\begin{matrix} ⟦a⟧ & b \\\\ c & d \\end{matrix}', '2×2 矩阵', '2 by 2 matrix'),
  template('matrices', 'matrix-3x3', '\\begin{matrix} ⟦a⟧ & b & c \\\\ d & e & f \\\\ g & h & i \\end{matrix}', '3×3 矩阵', '3 by 3 matrix'),
  template('matrices', 'matrix-parenthesized', '\\begin{pmatrix} ⟦a⟧ & b \\\\ c & d \\end{pmatrix}', '圆括号矩阵', 'Parenthesized matrix'),
  template('matrices', 'matrix-bracketed', '\\begin{bmatrix} ⟦a⟧ & b \\\\ c & d \\end{bmatrix}', '方括号矩阵', 'Bracketed matrix'),
  template('matrices', 'matrix-braced', '\\begin{Bmatrix} ⟦a⟧ & b \\\\ c & d \\end{Bmatrix}', '花括号矩阵', 'Braced matrix'),
  template('matrices', 'matrix-determinant', '\\begin{vmatrix} ⟦a⟧ & b \\\\ c & d \\end{vmatrix}', '行列式', 'Determinant'),
])
export function isDocumentLatexTemplateCategory(value: unknown): value is DocumentLatexTemplateCategory {
  return DOCUMENT_LATEX_TEMPLATE_CATEGORIES.some(category => category === value)
}
export function getDocumentLatexTemplate(id: unknown): DocumentLatexTemplate | undefined {
  return typeof id === 'string' ? DOCUMENT_LATEX_TEMPLATES.find(value => value.id === id) : undefined
}
export function getDocumentLatexTemplates(category: DocumentLatexTemplateCategory): readonly DocumentLatexTemplate[] {
  return DOCUMENT_LATEX_TEMPLATES.filter(value => value.category === category)
}
/** Selected source replaces the first placeholder; an empty source range selects that placeholder. */
export function insertDocumentLatexTemplate(value: string, id: string, start = value.length, end = start): DocumentLatexTemplateInsertion {
  const entry = getDocumentLatexTemplate(id)
  if (!entry) throw new RangeError('公式模板不存在。')
  if (typeof value !== 'string' || !Number.isSafeInteger(start) || !Number.isSafeInteger(end)) throw new TypeError('公式模板输入范围无效。')
  const from = Math.max(0, Math.min(start, value.length)), to = Math.max(from, Math.min(end, value.length))
  if (!isDocumentTextBoundary(value, from) || !isDocumentTextBoundary(value, to)) throw new RangeError('公式模板范围不能切开 UTF-16 字符。')
  const selected = value.slice(from, to)
  const insertion = selected ? entry.latex.slice(0, entry.placeholderStart) + selected + entry.latex.slice(entry.placeholderEnd) : entry.latex
  const next = value.slice(0, from) + insertion + value.slice(to), cursor = from + insertion.length
  return { value: next, selectionStart: selected ? cursor : from + entry.placeholderStart, selectionEnd: selected ? cursor : from + entry.placeholderEnd }
}
