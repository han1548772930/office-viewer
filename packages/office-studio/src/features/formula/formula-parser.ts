import type { FormulaNode, FormulaRangeReference, FormulaReference } from './formula-types.ts'

type TokenKind = 'number' | 'string' | 'reference' | 'identifier' | 'operator' | 'leftParen' | 'rightParen' | 'separator' | 'colon' | 'bang' | 'percent' | 'error' | 'eof'
interface Token { kind: TokenKind; text: string; reference?: FormulaReference; sheetName?: string }

export class FormulaSyntaxError extends Error {
  readonly code: string
  readonly position: number
  constructor(code: string, message: string, position: number) {
    super(message)
    this.name = 'FormulaSyntaxError'
    this.code = code
    this.position = position
  }
}

const MAX_COLUMN = 16_383
const MAX_ROW = 1_048_575

export function parseFormula(formula: string): FormulaNode {
  if (typeof formula !== 'string' || !formula.trim()) throw new FormulaSyntaxError('FORMULA_EMPTY', 'Formula cannot be empty', 0)
  const source = formula.trim().startsWith('=') ? formula.trim().slice(1) : formula.trim()
  const parser = new Parser(tokenize(source))
  const node = parser.parseComparison()
  if (parser.peek().kind !== 'eof') throw new FormulaSyntaxError('FORMULA_SYNTAX', `Unexpected token ${parser.peek().text}`, parser.peek().position ?? source.length)
  return node
}

interface PositionedToken extends Token { position: number }

function tokenize(source: string): PositionedToken[] {
  const tokens: PositionedToken[] = []
  let index = 0
  const push = (kind: TokenKind, text: string, position: number, extra: Partial<Token> = {}) => tokens.push({ kind, text, position, ...extra })
  while (index < source.length) {
    const character = source[index]!
    if (/\s/.test(character)) { index += 1; continue }
    const start = index
    if (character === '"') {
      index += 1
      let value = ''
      let closed = false
      while (index < source.length) {
        const next = source[index++]!
        if (next === '"') {
          if (source[index] === '"') { value += '"'; index += 1 }
          else { closed = true; break }
        } else value += next
      }
      if (!closed) throw new FormulaSyntaxError('FORMULA_SYNTAX', 'Unclosed string literal', start)
      push('string', value, start)
      continue
    }
    if (character === "'") {
      index += 1
      let sheetName = ''
      let closed = false
      while (index < source.length) {
        const next = source[index++]!
        if (next === "'") {
          if (source[index] === "'") { sheetName += "'"; index += 1 }
          else { closed = true; break }
        } else sheetName += next
      }
      if (!closed || source[index] !== '!') throw new FormulaSyntaxError('FORMULA_SYNTAX', 'Quoted sheet names must be followed by !', start)
      index += 1
      const errorReference = /^#(?:DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!|CIRCULAR!)/i.exec(source.slice(index))
      if (errorReference) { index += errorReference[0].length; push('error', errorReference[0].toUpperCase(), start); continue }
      const parsed = readReference(source, index)
      if (!parsed) throw new FormulaSyntaxError('FORMULA_SYNTAX', 'Expected a cell reference after sheet name', index)
      index = parsed.end
      push('reference', source.slice(start, index), start, { reference: parsed.reference, sheetName })
      continue
    }
    if (character === '#') {
      const error = /^#(?:DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!|CIRCULAR!)/i.exec(source.slice(index))
      if (!error) throw new FormulaSyntaxError('FORMULA_SYNTAX', 'Unknown formula error token', start)
      index += error[0].length
      push('error', error[0].toUpperCase(), start)
      continue
    }
    const reference = readReference(source, index)
    if (reference) {
      index = reference.end
      push('reference', source.slice(start, index), start, { reference: reference.reference })
      continue
    }
    const number = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?/.exec(source.slice(index))
    if (number) {
      const value = Number(number[0])
      if (!Number.isFinite(value)) throw new FormulaSyntaxError('FORMULA_NUMBER', 'Numeric literal is out of range', start)
      index += number[0].length
      push('number', number[0], start)
      continue
    }
    const identifier = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(source.slice(index))
    if (identifier) {
      index += identifier[0].length
      push('identifier', identifier[0], start)
      continue
    }
    const twoCharacterOperator = source.slice(index, index + 2)
    if (['<=', '>=', '<>', '!='].includes(twoCharacterOperator)) {
      index += 2
      push('operator', twoCharacterOperator === '!=' ? '<>' : twoCharacterOperator, start)
      continue
    }
    if ('+-*/^&=<>'.includes(character)) { index += 1; push('operator', character, start); continue }
    if (character === '(') { index += 1; push('leftParen', character, start); continue }
    if (character === ')') { index += 1; push('rightParen', character, start); continue }
    if (character === ',' || character === ';') { index += 1; push('separator', character, start); continue }
    if (character === ':') { index += 1; push('colon', character, start); continue }
    if (character === '!') { index += 1; push('bang', character, start); continue }
    if (character === '%') { index += 1; push('percent', character, start); continue }
    throw new FormulaSyntaxError('FORMULA_TOKEN', `Unexpected character ${character}`, start)
  }
  tokens.push({ kind: 'eof', text: '', position: source.length })
  return tokens
}

function readReference(source: string, start: number): { reference: FormulaReference; end: number } | undefined {
  let index = start
  const absoluteColumn = source[index] === '$'
  if (absoluteColumn) index += 1
  const lettersStart = index
  while (/[A-Za-z]/.test(source[index] ?? '')) index += 1
  const letters = source.slice(lettersStart, index)
  if (!letters || letters.length > 3) return undefined
  const absoluteRow = source[index] === '$'
  if (absoluteRow) index += 1
  const digitsStart = index
  while (/\d/.test(source[index] ?? '')) index += 1
  const digits = source.slice(digitsStart, index)
  if (!digits || /^0+$/.test(digits)) return undefined
  const column = letters.toUpperCase().split('').reduce((value, letter) => value * 26 + letter.charCodeAt(0) - 64, 0) - 1
  const row = Number(digits) - 1
  if (column < 0 || column > MAX_COLUMN || row < 0 || row > MAX_ROW) return undefined
  if (/[A-Za-z0-9_.]/.test(source[index] ?? '')) return undefined
  return { reference: { kind: 'cell', row, column, absoluteRow, absoluteColumn }, end: index }
}

class Parser {
  private readonly tokens: PositionedToken[]
  private index = 0
  constructor(tokens: PositionedToken[]) { this.tokens = tokens }
  peek(): PositionedToken { return this.tokens[this.index] ?? this.tokens.at(-1)! }
  take(): PositionedToken { const token = this.peek(); this.index += 1; return token }
  accept(kind: TokenKind, text?: string): PositionedToken | undefined {
    const token = this.peek()
    if (token.kind !== kind || (text !== undefined && token.text !== text)) return undefined
    this.index += 1
    return token
  }
  expect(kind: TokenKind, text?: string): PositionedToken {
    const token = this.accept(kind, text)
    if (!token) throw new FormulaSyntaxError('FORMULA_SYNTAX', `Expected ${text ?? kind}`, this.peek().position)
    return token
  }
  parseComparison(): FormulaNode {
    let node = this.parseConcatenation()
    while (this.peek().kind === 'operator' && ['=', '<>', '<', '>', '<=', '>='].includes(this.peek().text)) {
      const operator = this.take().text
      node = { kind: 'binary', operator, left: node, right: this.parseConcatenation() }
    }
    return node
  }
  private parseConcatenation(): FormulaNode {
    let node = this.parseAddition()
    while (this.accept('operator', '&')) node = { kind: 'binary', operator: '&', left: node, right: this.parseAddition() }
    return node
  }
  private parseAddition(): FormulaNode {
    let node = this.parseMultiplication()
    while (this.peek().kind === 'operator' && ['+', '-'].includes(this.peek().text)) {
      const operator = this.take().text
      node = { kind: 'binary', operator, left: node, right: this.parseMultiplication() }
    }
    return node
  }
  private parseMultiplication(): FormulaNode {
    let node = this.parseUnary()
    while (this.peek().kind === 'operator' && ['*', '/'].includes(this.peek().text)) {
      const operator = this.take().text
      node = { kind: 'binary', operator, left: node, right: this.parseUnary() }
    }
    return node
  }
  private parseUnary(): FormulaNode {
    if (this.peek().kind === 'operator' && ['+', '-'].includes(this.peek().text)) {
      const operator = this.take().text as '+' | '-'
      return { kind: 'unary', operator, operand: this.parseUnary() }
    }
    return this.parsePower()
  }
  private parsePower(): FormulaNode {
    let node = this.parsePostfix()
    if (this.accept('operator', '^')) node = { kind: 'binary', operator: '^', left: node, right: this.parseUnary() }
    return node
  }
  private parsePostfix(): FormulaNode {
    let node = this.parsePrimary()
    while (this.accept('percent')) node = { kind: 'binary', operator: '/', left: node, right: { kind: 'literal', value: 100 } }
    if (this.accept('colon')) {
      const end = this.parsePrimary()
      if (node.kind !== 'reference' || end.kind !== 'reference') throw new FormulaSyntaxError('FORMULA_RANGE', 'Range endpoints must be cell references', this.peek().position)
      const firstSheet = node.reference.sheetName, lastSheet = end.reference.sheetName
      if (firstSheet && lastSheet && firstSheet.toLocaleLowerCase() !== lastSheet.toLocaleLowerCase()) throw new FormulaSyntaxError('FORMULA_RANGE', 'A range cannot span different sheets', this.peek().position)
      const reference: FormulaRangeReference = { kind: 'range', sheetName: firstSheet ?? lastSheet, start: node.reference, end: end.reference }
      node = { kind: 'range', reference }
    }
    return node
  }
  private parsePrimary(): FormulaNode {
    const token = this.take()
    if (token.kind === 'number') return { kind: 'literal', value: Number(token.text) }
    if (token.kind === 'string') return { kind: 'literal', value: token.text }
    if (token.kind === 'error') return { kind: 'literal', value: token.text }
    if (token.kind === 'reference') {
      if (!token.reference) throw new FormulaSyntaxError('FORMULA_REFERENCE', 'Invalid cell reference', token.position)
      return { kind: 'reference', reference: { ...token.reference, sheetName: token.sheetName } }
    }
    if (token.kind === 'leftParen') {
      const node = this.parseComparison()
      this.expect('rightParen')
      return node
    }
    if (token.kind === 'identifier') {
      if (this.accept('bang')) {
        const cell = this.take()
        if (cell.kind === 'error') return { kind: 'literal', value: cell.text }
        if (cell.kind !== 'reference' || !cell.reference) throw new FormulaSyntaxError('FORMULA_REFERENCE', 'Expected a cell reference after sheet name', cell.position)
        return { kind: 'reference', reference: { ...cell.reference, sheetName: token.text } }
      }
      if (this.accept('leftParen')) {
        const args: FormulaNode[] = []
        if (!this.accept('rightParen')) {
          do { args.push(this.parseComparison()) } while (this.accept('separator'))
          this.expect('rightParen')
        }
        return { kind: 'function', name: token.text.toLocaleUpperCase(), arguments: args }
      }
      const upper = token.text.toLocaleUpperCase()
      if (upper === 'TRUE') return { kind: 'literal', value: true }
      if (upper === 'FALSE') return { kind: 'literal', value: false }
      throw new FormulaSyntaxError('FORMULA_NAME', `Unknown name ${token.text}`, token.position)
    }
    throw new FormulaSyntaxError('FORMULA_SYNTAX', `Unexpected token ${token.text || 'end of formula'}`, token.position)
  }
}

