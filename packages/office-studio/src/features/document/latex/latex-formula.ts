import type { DocumentBlock, DocumentModel } from '../../../domain/model.ts'
import type { DrawCommand } from '../../../rendering/contracts.ts'
import type { DocumentPoint, DocumentRun, DocumentSelection } from '../model/document-model.ts'
import { createParagraph, createRun, orderedSelection, runsForTarget, sameTarget, setTargetRuns, sliceRuns } from '../model/document-model.ts'

export type LatexFormulaMode = 'inline' | 'block'
export interface LatexFormula { latex: string; mode: LatexFormulaMode; version: 1 }
export type LatexNode =
  | { type: 'text'; value: string }
  | { type: 'sequence'; children: LatexNode[] }
  | { type: 'fraction'; numerator: LatexNode; denominator: LatexNode }
  | { type: 'sqrt'; value: LatexNode }
  | { type: 'script'; base: LatexNode; sup?: LatexNode; sub?: LatexNode }
export interface LatexMetrics { width: number; height: number; baseline: number }
export type LatexPrimitive =
  | { kind: 'text'; x: number; y: number; text: string; font: string }
  | { kind: 'line'; x1: number; y1: number; x2: number; y2: number; width: number }
export interface LatexLayout extends LatexMetrics { primitives: LatexPrimitive[] }

const OBJECT_REPLACEMENT = '\uFFFC'
const SYMBOLS: Record<string, string> = {
  alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', theta: 'θ', lambda: 'λ', mu: 'μ', pi: 'π', sigma: 'σ',
  phi: 'φ', omega: 'ω', Delta: 'Δ', Gamma: 'Γ', Lambda: 'Λ', Omega: 'Ω',
  times: '×', cdot: '·', pm: '±', le: '≤', leq: '≤', ge: '≥', geq: '≥', neq: '≠', ne: '≠',
  approx: '≈', infty: '∞', sum: '∑', prod: '∏', int: '∫', rightarrow: '→', leftarrow: '←',
}

export class LatexFormulaError extends Error {
  readonly offset: number

  constructor(message: string, offset: number) { super(message); this.offset = offset; this.name = 'LatexFormulaError' }
}

export function parseLatexFormula(source: string): LatexNode {
  const input = source.trim()
  if (!input) throw new LatexFormulaError('公式不能为空。', 0)
  let index = 0
  const parseSequence = (inGroup: boolean): LatexNode => {
    const children: LatexNode[] = []
    while (index < input.length) {
      const char = input[index]
      if (char === '}') {
        if (!inGroup) throw new LatexFormulaError('存在多余的右花括号。', index)
        index += 1
        return sequence(children)
      }
      if (char === '{') {
        index += 1
        children.push(parseSequence(true))
        continue
      }
      let base = parseAtom()
      let sup: LatexNode | undefined
      let sub: LatexNode | undefined
      while (input[index] === '^' || input[index] === '_') {
        const operator = input[index++]
        const script = parseScript()
        if (operator === '^') {
          if (sup) throw new LatexFormulaError('一个公式项不能有两个上标。', index - 1)
          sup = script
        } else {
          if (sub) throw new LatexFormulaError('一个公式项不能有两个下标。', index - 1)
          sub = script
        }
      }
      if (sup || sub) base = { type: 'script', base, sup, sub }
      children.push(base)
    }
    if (inGroup) throw new LatexFormulaError('公式中有未闭合的左花括号。', input.length)
    return sequence(children)
  }
  const parseAtom = (): LatexNode => {
    const start = index
    const char = input[index]
    if (char === '\\') {
      index += 1
      const match = /^[A-Za-z]+/.exec(input.slice(index))
      if (!match) {
        const escaped = input[index++]
        if (!['{', '}', '_', '%', '$'].includes(escaped ?? '')) throw new LatexFormulaError('不支持该转义符。', start)
        return { type: 'text', value: escaped! }
      }
      const command = match[0]
      index += command.length
      if (command === 'frac') return { type: 'fraction', numerator: parseRequiredGroup(command), denominator: parseRequiredGroup(command) }
      if (command === 'sqrt') return { type: 'sqrt', value: parseRequiredGroup(command) }
      if (command === 'text' || command === 'mathrm' || command === 'mathbf') return parseRequiredGroup(command)
      const symbol = SYMBOLS[command]
      if (!symbol) throw new LatexFormulaError(`暂不支持 \\${command} 命令。`, start)
      return { type: 'text', value: symbol }
    }
    if (char === '^' || char === '_') throw new LatexFormulaError('上标或下标缺少基项。', index)
    if (!char) throw new LatexFormulaError('公式项不完整。', index)
    index += 1
    return { type: 'text', value: char }
  }
  const parseRequiredGroup = (command: string): LatexNode => {
    if (input[index] !== '{') throw new LatexFormulaError(`\\${command} 后必须使用花括号参数。`, index)
    index += 1
    return parseSequence(true)
  }
  const parseScript = (): LatexNode => {
    if (input[index] === '{') { index += 1; return parseSequence(true) }
    if (index >= input.length || input[index] === '^' || input[index] === '_') throw new LatexFormulaError('上标或下标缺少内容。', index)
    return parseAtom()
  }
  return parseSequence(false)
}

function sequence(children: LatexNode[]): LatexNode {
  if (!children.length) return { type: 'text', value: '' }
  return children.length === 1 ? children[0]! : { type: 'sequence', children }
}

export function layoutLatexFormula(source: string, options: { fontSize?: number; measure?: (text: string, font: string) => number } = {}): LatexLayout {
  const root = parseLatexFormula(source)
  const baseSize = Math.max(8, Math.min(96, options.fontSize ?? 16))
  const measure = options.measure ?? ((text: string, font: string) => text.length * Number(/([\d.]+)px/.exec(font)?.[1] ?? baseSize) * 0.58)
  return layoutNode(root, 0, 0, baseSize, measure)
}

export function latexFormulaDrawCommands(source: string, x: number, baseline: number, options: { fontSize?: number; color?: string; measure?: (text: string, font: string) => number } = {}): { layout: LatexLayout; commands: DrawCommand[] } {
  const layout = layoutLatexFormula(source, options)
  const color = options.color ?? '#1f2937'
  const commands: DrawCommand[] = layout.primitives.map((primitive) => primitive.kind === 'text'
    ? { kind: 'text', x: x + primitive.x, y: baseline + primitive.y, text: primitive.text, font: primitive.font, color, baseline: 'alphabetic' }
    : { kind: 'line', x1: x + primitive.x1, y1: baseline + primitive.y1, x2: x + primitive.x2, y2: baseline + primitive.y2, color, width: primitive.width })
  return { layout, commands }
}

function layoutNode(node: LatexNode, x: number, baseline: number, size: number, measure: (text: string, font: string) => number): LatexLayout {
  const font = `${size}px "Cambria Math", "STIX Two Math", "Times New Roman", serif`
  if (node.type === 'text') {
    const width = measure(node.value, font)
    return { width, height: size * 1.2, baseline: size * 0.9, primitives: node.value ? [{ kind: 'text', x, y: baseline, text: node.value, font }] : [] }
  }
  if (node.type === 'sequence') {
    let cursor = x
    let height = size * 1.2
    const primitives: LatexPrimitive[] = []
    for (const child of node.children) {
      const item = layoutNode(child, cursor, baseline, size, measure)
      cursor += item.width
      height = Math.max(height, item.height)
      primitives.push(...item.primitives)
    }
    return { width: cursor - x, height, baseline: size * 0.9, primitives }
  }
  if (node.type === 'sqrt') {
    const inside = layoutNode(node.value, x + size * 0.78, baseline, size, measure)
    const radical = layoutNode({ type: 'text', value: '\u221A' }, x, baseline, size * 1.25, measure)
    const top = baseline - inside.baseline - size * 0.08
    return { width: inside.width + size * 0.78, height: Math.max(inside.height, size * 1.35), baseline: size * 0.9, primitives: [...radical.primitives, ...inside.primitives, { kind: 'line', x1: x + size * 0.72, y1: top, x2: x + size * 0.78 + inside.width, y2: top, width: Math.max(0.8, size / 18) }] }
  }
  if (node.type === 'fraction') {
    const childSize = size * 0.78
    const numerator = layoutNode(node.numerator, 0, 0, childSize, measure)
    const denominator = layoutNode(node.denominator, 0, 0, childSize, measure)
    const width = Math.max(numerator.width, denominator.width) + size * 0.32
    const ruleY = baseline - size * 0.12
    const numeratorX = x + (width - numerator.width) / 2
    const denominatorX = x + (width - denominator.width) / 2
    const numeratorBaseline = ruleY - size * 0.12 - numerator.baseline
    const denominatorBaseline = ruleY + size * 0.2 + denominator.baseline
    return { width, height: numerator.height + denominator.height + size * 0.45, baseline: size * 0.9, primitives: [
      ...translate(numerator.primitives, numeratorX, numeratorBaseline),
      { kind: 'line', x1: x + size * 0.08, y1: ruleY, x2: x + width - size * 0.08, y2: ruleY, width: Math.max(0.8, size / 18) },
      ...translate(denominator.primitives, denominatorX, denominatorBaseline),
    ] }
  }
  const base = layoutNode(node.base, x, baseline, size, measure)
  const scriptSize = size * 0.68
  const scriptX = x + base.width
  const sup = node.sup ? layoutNode(node.sup, scriptX, baseline - size * 0.42, scriptSize, measure) : undefined
  const sub = node.sub ? layoutNode(node.sub, scriptX, baseline + size * 0.32, scriptSize, measure) : undefined
  return { width: base.width + Math.max(sup?.width ?? 0, sub?.width ?? 0), height: Math.max(base.height, (sup?.height ?? 0) + size * 0.4, (sub?.height ?? 0) + size * 0.3), baseline: base.baseline, primitives: [...base.primitives, ...(sup?.primitives ?? []), ...(sub?.primitives ?? [])] }
}

function translate(primitives: LatexPrimitive[], x: number, y: number): LatexPrimitive[] {
  return primitives.map((primitive) => primitive.kind === 'text'
    ? { ...primitive, x: primitive.x + x, y: primitive.y + y }
    : { ...primitive, x1: primitive.x1 + x, x2: primitive.x2 + x, y1: primitive.y1 + y, y2: primitive.y2 + y })
}

export function createLatexFormulaRun(latex: string, mode: LatexFormulaMode = 'inline', style: Partial<DocumentRun> = {}): DocumentRun {
  parseLatexFormula(latex)
  const run = createRun(OBJECT_REPLACEMENT, style)
  run.metadata = { ...run.metadata, latexFormula: { latex: latex.trim(), mode, version: 1 } satisfies LatexFormula }
  return run
}

export function getLatexFormula(run: DocumentRun): LatexFormula | undefined {
  const value = run.metadata?.latexFormula
  if (!value || typeof value !== 'object') return undefined
  const formula = value as Partial<LatexFormula>
  return typeof formula.latex === 'string' && (formula.mode === 'inline' || formula.mode === 'block') && formula.version === 1
    ? { latex: formula.latex, mode: formula.mode, version: 1 }
    : undefined
}

export function getLatexFormulas(model: DocumentModel): Array<{ blockId: string; runId: string; formula: LatexFormula }> {
  const result: Array<{ blockId: string; runId: string; formula: LatexFormula }> = []
  const visit = (blocks: readonly DocumentBlock[]) => {
    for (const block of blocks) {
      for (const run of block.runs as DocumentRun[]) {
        const formula = getLatexFormula(run)
        if (formula) result.push({ blockId: block.id, runId: run.id, formula })
      }
      const table = block.metadata.table as { rows?: Array<{ cells?: Array<{ blocks?: DocumentBlock[] }> }> } | undefined
      for (const row of table?.rows ?? []) for (const cell of row.cells ?? []) visit(cell.blocks ?? [])
    }
  }
  visit(model.blocks)
  return result
}

export function updateLatexFormula(model: DocumentModel, runId: string, latex: string): boolean {
  parseLatexFormula(latex)
  for (const block of model.blocks) {
    const run = (block.runs as DocumentRun[]).find((candidate) => candidate.id === runId)
    if (!run || !getLatexFormula(run)) continue
    run.metadata = { ...run.metadata, latexFormula: { ...getLatexFormula(run)!, latex: latex.trim() } }
    return true
  }
  return false
}

export function createLatexFormulaBlock(latex: string): DocumentBlock {
  const block = createParagraph()
  block.style = { ...block.style, align: 'center', spaceBefore: 8, spaceAfter: 8 }
  block.runs = [createLatexFormulaRun(latex, 'block')]
  block.metadata = { ...block.metadata, latexFormulaBlock: true }
  return block
}

export function insertLatexFormula(model: DocumentModel, selection: DocumentSelection, latex: string, mode: LatexFormulaMode = 'inline'): DocumentPoint {
  const { start, end } = orderedSelection(model, selection)
  if (!sameTarget(start, end)) throw new RangeError('公式插入暂时要求在同一段落内选择内容。')
  const block = model.blocks.find((item) => item.id === start.blockId)
  if (!block || start.row != null) throw new RangeError('公式只能插入文档正文段落。')
  if (model.contentControls?.some((control) => control.blockId === block.id && start.offset < control.endOffset && end.offset > control.startOffset)) {
    throw new RangeError('不能替换内容控件中的文字来插入公式。')
  }
  const before = sliceRuns(runsForTarget(model, start), 0, start.offset)
  const after = sliceRuns(runsForTarget(model, start), end.offset)
  if (mode === 'inline') {
    const formula = createLatexFormulaRun(latex, mode)
    setTargetRuns(model, start, [...before, formula, ...after])
    return { ...start, offset: start.offset + formula.text.length }
  }
  const index = model.blocks.findIndex((item) => item.id === block.id)
  const equation = createLatexFormulaBlock(latex)
  const trailing = createParagraph()
  trailing.style = { ...block.style }
  trailing.runs = after.length ? after : [createRun()]
  block.runs = before.length ? before : [createRun()]
  model.blocks.splice(index + 1, 0, equation, trailing)
  return { blockId: equation.id, offset: OBJECT_REPLACEMENT.length }
}
