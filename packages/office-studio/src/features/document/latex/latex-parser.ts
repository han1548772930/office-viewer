import type { LatexAccent, LatexNode, LatexTextStyle } from './latex-types'

const SYMBOLS: Record<string, string> = {
  alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε', varepsilon: 'ϵ', zeta: 'ζ', eta: 'η',
  theta: 'θ', vartheta: 'ϑ', iota: 'ι', kappa: 'κ', lambda: 'λ', mu: 'μ', nu: 'ν', xi: 'ξ',
  pi: 'π', rho: 'ρ', sigma: 'σ', tau: 'τ', upsilon: 'υ', phi: 'φ', varphi: 'ϕ', chi: 'χ', psi: 'ψ', omega: 'ω',
  Delta: 'Δ', Gamma: 'Γ', Theta: 'Θ', Lambda: 'Λ', Xi: 'Ξ', Pi: 'Π', Sigma: 'Σ', Phi: 'Φ', Psi: 'Ψ', Omega: 'Ω',
  times: '×', cdot: '·', div: '÷', pm: '±', mp: '∓', le: '≤', leq: '≤', ge: '≥', geq: '≥',
  neq: '≠', ne: '≠', approx: '≈', equiv: '≡', infty: '∞', partial: '∂', nabla: '∇',
  rightarrow: '→', to: '→', leftarrow: '←', leftrightarrow: '↔', Rightarrow: '⇒', Leftarrow: '⇐',
  in: '∈', notin: '∉', subset: '⊂', subseteq: '⊆', supset: '⊃', supseteq: '⊇', cup: '∪', cap: '∩',
  emptyset: '∅', forall: '∀', exists: '∃', neg: '¬', land: '∧', lor: '∨',
  cdots: '⋯', ldots: '…', vdots: '⋮', ddots: '⋱', ell: 'ℓ',
}
const LARGE_OPERATORS: Record<string, { value: string; limits: boolean }> = {
  sum: { value: '∑', limits: true }, prod: { value: '∏', limits: true }, coprod: { value: '∐', limits: true },
  bigcup: { value: '⋃', limits: true }, bigcap: { value: '⋂', limits: true }, bigwedge: { value: '⋀', limits: true },
  bigvee: { value: '⋁', limits: true }, int: { value: '∫', limits: false }, iint: { value: '∬', limits: false },
  iiint: { value: '∭', limits: false }, oint: { value: '∮', limits: false },
}
const FUNCTIONS = new Set(['sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'sinh', 'cosh', 'tanh', 'log', 'ln', 'exp', 'min', 'max', 'lim', 'det'])
const ACCENTS = new Set<LatexAccent>(['hat', 'bar', 'vec', 'dot', 'ddot', 'overline', 'underline', 'overbrace'])
const STYLES: Record<string, LatexTextStyle> = {
  text: 'roman', mathrm: 'roman', operatorname: 'roman', mathbf: 'bold', mathit: 'italic', mathsf: 'sans', mathtt: 'monospace',
}
const MATRIX_DELIMITERS: Record<string, readonly [string, string]> = {
  matrix: ['', ''], pmatrix: ['(', ')'], bmatrix: ['[', ']'], Bmatrix: ['{', '}'], vmatrix: ['|', '|'], Vmatrix: ['‖', '‖'],
}
const DELIMITERS: Record<string, string> = {
  '{': '{', '}': '}', '|': '‖', langle: '⟨', rangle: '⟩', lfloor: '⌊', rfloor: '⌋', lceil: '⌈', rceil: '⌉',
  lvert: '|', rvert: '|', vert: '|', lVert: '‖', rVert: '‖', Vert: '‖',
}
const SPACING: Record<string, number> = { ',': 1 / 6, ':': 2 / 9, ';': 5 / 18, ' ': 1 / 3 }

export class LatexFormulaError extends Error {
  readonly offset: number
  constructor(message: string, offset: number) { super(message); this.offset = offset; this.name = 'LatexFormulaError' }
}

/** Source remains unchanged; only document structural control markers are excluded. */
export function assertLatexFormulaSource(source: unknown): asserts source is string {
  if (typeof source !== 'string') throw new TypeError('公式源文本必须是字符串。')
  if (!source.trim()) throw new LatexFormulaError('公式不能为空。', 0)
  const marker = /[\x00-\x08\x0b\x0c\x0e-\x1f]/.exec(source)
  if (marker) throw new LatexFormulaError('公式源文本包含文档结构标记。', marker.index)
}

function sequence(children: LatexNode[]): LatexNode {
  return children.length === 1 ? children[0]! : children.length ? { type: 'sequence', children } : { type: 'text', value: '' }
}

/** Typed nodes cover the fixed 71-template catalog; this is not a complete TeX engine. */
export function parseLatexFormula(source: string): LatexNode {
  assertLatexFormulaSource(source)
  return new LatexParser(source).parse()
}

class LatexParser {
  private index = 0
  private depth = 0
  private nodes = 0
  constructor(private readonly source: string) {}
  parse(): LatexNode {
    if (this.source.length > 65_536) throw this.error('公式超过当前单次排版文本预算。', 65_536)
    return this.readSequence(() => false)
  }
  private error(message: string, offset = this.index): LatexFormulaError { return new LatexFormulaError(message, offset) }
  private space() { while (/\s/.test(this.source[this.index] ?? '') && this.index < this.source.length) this.index++ }
  private commandAhead(command: string): boolean {
    const prefix = '\\' + command
    return this.source.startsWith(prefix, this.index) && !/[A-Za-z]/.test(this.source[this.index + prefix.length] ?? '')
  }
  private readSequence(stop: () => boolean, preserveSpace = false): LatexNode {
    if (++this.depth > 64) throw this.error('公式嵌套超过当前排版预算。')
    try {
      const children: LatexNode[] = []
      while (this.index < this.source.length) {
        if (stop()) break
        if (!preserveSpace) { this.space(); if (this.index >= this.source.length || stop()) break }
        if (++this.nodes > 16_384) throw this.error('公式项超过当前排版预算。')
        const char = this.source[this.index]
        if (char === '}') throw this.error('存在多余的右花括号。')
        if (char === '&' || this.source.startsWith('\\\\', this.index)) throw this.error('矩阵分隔符必须位于矩阵环境内。')
        let base = this.readAtom(preserveSpace), sup: LatexNode | undefined, sub: LatexNode | undefined
        if (!preserveSpace) this.space()
        while (this.source[this.index] === '^' || this.source[this.index] === '_') {
          const offset = this.index, operator = this.source[this.index++]
          if (operator === '^' && sup || operator === '_' && sub) throw this.error(operator === '^' ? '一个公式项不能有两个上标。' : '一个公式项不能有两个下标。', offset)
          const script = this.readScript()
          if (operator === '^') sup = script
          else sub = script
          if (!preserveSpace) this.space()
        }
        if (sup || sub) base = { type: 'script', base, sup, sub }
        children.push(base)
      }
      return sequence(children)
    } finally { this.depth-- }
  }
  private readAtom(preserveSpace = false): LatexNode {
    const start = this.index, char = this.source[this.index]
    if (char === '{') return this.readGroup('分组', preserveSpace)
    if (char === '\\') {
      this.index++
      const word = /^[A-Za-z]+/.exec(this.source.slice(this.index))?.[0]
      if (!word) {
        const escaped = this.source[this.index++]
        const spacing = SPACING[escaped ?? '']
        if (spacing !== undefined) return { type: 'space', em: spacing }
        if (!['{', '}', '_', '%', '$', '^', '&', '#', '|'].includes(escaped ?? '')) throw this.error('不支持该转义符。', start)
        return { type: 'text', value: escaped === '|' ? '‖' : escaped! }
      }
      this.index += word.length
      if (word === 'frac' || word === 'dfrac' || word === 'tfrac') return { type: 'fraction', numerator: this.readGroup(word), denominator: this.readGroup(word) }
      if (word === 'sqrt') {
        this.space()
        let index: LatexNode | undefined
        if (this.source[this.index] === '[') {
          this.index++
          index = this.readSequence(() => this.source[this.index] === ']')
          if (this.source[this.index] !== ']') throw this.error('根指数有未闭合的方括号。', start)
          this.index++
        }
        return { type: 'sqrt', value: this.readGroup(word), ...(index ? { index } : {}) }
      }
      if (word === 'left') {
        const left = this.readDelimiter()
        const value = this.readSequence(() => this.commandAhead('right'))
        if (!this.commandAhead('right')) throw this.error('定界符缺少 \\right。', start)
        this.index += '\\right'.length
        return { type: 'delimited', left, right: this.readDelimiter(), value }
      }
      if (word === 'begin') return this.readMatrix(start)
      if (word === 'right' || word === 'end') throw this.error('存在未配对的公式结束标记。', start)
      if (word === 'quad' || word === 'qquad') return { type: 'space', em: word === 'quad' ? 1 : 2 }
      if (word === 'displaystyle' || word === 'textstyle') { this.space(); return this.readAtom() }
      const style = STYLES[word]
      if (style) return { type: 'styled', style, value: this.readGroup(word, word === 'text') }
      if (ACCENTS.has(word as LatexAccent)) return { type: 'accent', accent: word as LatexAccent, value: this.readGroup(word) }
      const large = LARGE_OPERATORS[word]
      if (large) return { type: 'operator', ...large, large: true }
      if (FUNCTIONS.has(word)) return { type: 'operator', value: word, large: false, limits: ['lim', 'min', 'max'].includes(word) }
      const symbol = SYMBOLS[word]
      if (symbol) return { type: 'text', value: symbol }
      throw this.error('暂不支持 \\' + word + ' 命令。', start)
    }
    if (char === '^' || char === '_') throw this.error('上标或下标缺少基项。')
    if (!char || char === '}') throw this.error('公式项不完整。')
    const value = String.fromCodePoint(this.source.codePointAt(this.index)!)
    this.index += value.length
    return { type: 'text', value }
  }
  private readGroup(command: string, preserveSpace = false): LatexNode {
    this.space()
    if (this.source[this.index] !== '{') throw this.error('\\' + command + ' 后必须使用花括号参数。')
    this.index++
    const value = this.readSequence(() => this.source[this.index] === '}', preserveSpace)
    if (this.source[this.index] !== '}') throw this.error('公式中有未闭合的左花括号。')
    this.index++
    return value
  }
  private readScript(): LatexNode {
    this.space()
    if (this.source[this.index] === '{') return this.readGroup('上标或下标')
    if (this.index >= this.source.length || ['^', '_', '}'].includes(this.source[this.index]!)) throw this.error('上标或下标缺少内容。')
    return this.readAtom()
  }
  private readDelimiter(): string {
    this.space()
    const start = this.index
    if (this.source[this.index] === '\\') {
      this.index++
      const word = /^[A-Za-z]+/.exec(this.source.slice(this.index))?.[0] ?? this.source[this.index]
      if (!word || DELIMITERS[word] === undefined) throw this.error('不支持该定界符。', start)
      this.index += word.length
      return DELIMITERS[word]!
    }
    const char = this.source[this.index++]
    if (!char || !['(', ')', '[', ']', '|', '.'].includes(char)) throw this.error('公式定界符无效。', start)
    return char === '.' ? '' : char
  }
  private environmentName(): string {
    this.space()
    const match = /^\{([A-Za-z]+)\}/.exec(this.source.slice(this.index))
    if (!match) throw this.error('矩阵环境需要花括号名称。')
    this.index += match[0].length
    return match[1]!
  }
  private readMatrix(start: number): LatexNode {
    const environment = this.environmentName(), delimiters = MATRIX_DELIMITERS[environment]
    if (!delimiters) throw this.error('暂不支持 ' + environment + ' 数学环境。', start)
    const rows: LatexNode[][] = [], row: LatexNode[] = []
    const stop = () => this.source[this.index] === '&' || this.source.startsWith('\\\\', this.index) || this.commandAhead('end')
    while (this.index < this.source.length) {
      row.push(this.readSequence(stop))
      if (this.source[this.index] === '&') { this.index++; continue }
      if (this.source.startsWith('\\\\', this.index)) {
        this.index += 2; rows.push(row.splice(0)); this.space()
        if (this.commandAhead('end')) break
        continue
      }
      if (this.commandAhead('end')) break
      throw this.error('矩阵环境未闭合。', start)
    }
    if (!this.commandAhead('end')) throw this.error('矩阵环境未闭合。', start)
    this.index += '\\end'.length
    if (this.environmentName() !== environment) throw this.error('矩阵环境的结束名称不匹配。', start)
    if (row.length) rows.push(row)
    if (!rows.length) rows.push([{ type: 'text', value: '' }])
    return { type: 'matrix', rows, left: delimiters[0], right: delimiters[1] }
  }
}
