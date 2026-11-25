import type { DrawCommand } from '../../../rendering/contracts'
import { parseLatexFormula } from './latex-parser'
import type { LatexLayout, LatexNode, LatexPrimitive, LatexTextStyle } from './latex-types'

type Measure = (text: string, font: string) => number
export interface LatexLayoutOptions { fontSize?: number; measure?: Measure }

export function layoutLatexFormula(source: string, options: LatexLayoutOptions = {}): LatexLayout {
  const root = parseLatexFormula(source)
  if (options.fontSize !== undefined && (!Number.isFinite(options.fontSize) || options.fontSize <= 0)) throw new RangeError('公式字号必须为正有限数。')
  const size = Math.max(8, Math.min(96, options.fontSize ?? 16))
  const measure: Measure = options.measure ?? ((text, font) => Array.from(text).length * Number(/([\d.]+)px/.exec(font)?.[1] ?? size) * 0.58)
  return nodeLayout(root, 0, 0, size, measure)
}

export function latexFormulaDrawCommands(source: string, x: number, baseline: number, options: LatexLayoutOptions & { color?: string } = {}): { layout: LatexLayout; commands: DrawCommand[] } {
  const layout = layoutLatexFormula(source, options), color = options.color ?? '#1f2937'
  const commands: DrawCommand[] = layout.primitives.map(primitive => primitive.kind === 'text'
    ? { kind: 'text', x: x + primitive.x, y: baseline + primitive.y, text: primitive.text, font: primitive.font, color, baseline: 'alphabetic' }
    : { kind: 'line', x1: x + primitive.x1, y1: baseline + primitive.y1, x2: x + primitive.x2, y2: baseline + primitive.y2, color, width: primitive.width })
  return { layout, commands }
}

function fontFor(size: number, style?: LatexTextStyle): string {
  const family = style === 'monospace' ? 'Consolas, monospace' : style === 'sans' ? 'Arial, sans-serif' : '"Cambria Math", "STIX Two Math", "Times New Roman", serif'
  return (style === 'bold' ? 'bold ' : style === 'italic' ? 'italic ' : '') + size + 'px ' + family
}
function textLayout(value: string, x: number, baseline: number, size: number, measure: Measure, style?: LatexTextStyle): LatexLayout {
  const font = fontFor(size, style), width = measure(value, font)
  if (!Number.isFinite(width) || width < 0) throw new RangeError('公式字体测量结果无效。')
  return metrics(width, value ? [{ kind: 'text', x, y: baseline, text: value, font }] : [], baseline, size)
}
function nodeLayout(node: LatexNode, x: number, baseline: number, size: number, measure: Measure, style?: LatexTextStyle): LatexLayout {
  if (node.type === 'text') return textLayout(node.value, x, baseline, size, measure, style)
  if (node.type === 'space') return metrics(size * node.em, [], baseline, size)
  if (node.type === 'styled') return nodeLayout(node.value, x, baseline, size, measure, node.style)
  if (node.type === 'operator') return textLayout(node.value, x, baseline, node.large ? size * 1.35 : size, measure, 'roman')
  if (node.type === 'sequence') {
    let cursor = x
    const primitives: LatexPrimitive[] = []
    for (const child of node.children) {
      const item = nodeLayout(child, cursor, baseline, size, measure, style)
      cursor += item.width; primitives.push(...item.primitives)
    }
    return metrics(cursor - x, primitives, baseline, size)
  }
  if (node.type === 'fraction') {
    const numerator = nodeLayout(node.numerator, 0, 0, size * 0.78, measure, style)
    const denominator = nodeLayout(node.denominator, 0, 0, size * 0.78, measure, style)
    const width = Math.max(numerator.width, denominator.width) + size * 0.32, ruleY = baseline - size * 0.12
    return metrics(width, [
      ...translate(numerator.primitives, x + (width - numerator.width) / 2, ruleY - size * 0.12 - descent(numerator)),
      line(x + size * 0.08, ruleY, x + width - size * 0.08, ruleY, size),
      ...translate(denominator.primitives, x + (width - denominator.width) / 2, ruleY + size * 0.2 + denominator.baseline),
    ], baseline, size)
  }
  if (node.type === 'sqrt') {
    const inner = nodeLayout(node.value, 0, 0, size, measure, style)
    const rootIndex = node.index && nodeLayout(node.index, 0, 0, size * 0.5, measure, style)
    const left = rootIndex ? Math.max(size * 0.2, rootIndex.width * 0.75) : 0
    const radicalSize = Math.max(size * 1.25, inner.height / 1.05)
    const radicalBaseline = baseline - inner.baseline - size * 0.08 + radicalSize * 0.9
    const radical = textLayout('√', x + left, radicalBaseline, radicalSize, measure)
    const insideX = x + left + Math.max(size * 0.78, radical.width * 0.88), top = baseline - inner.baseline - size * 0.08
    return metrics(insideX - x + inner.width + size * 0.05, [
      ...radical.primitives, ...translate(inner.primitives, insideX, baseline),
      line(insideX - size * 0.06, top, insideX + inner.width + size * 0.05, top, size),
      ...(rootIndex ? translate(rootIndex.primitives, x, top + rootIndex.baseline * 0.35) : []),
    ], baseline, size)
  }
  if (node.type === 'delimited') return delimit(nodeLayout(node.value, 0, 0, size, measure, style), x, baseline, size, measure, node.left, node.right)
  if (node.type === 'matrix') return matrix(node, x, baseline, size, measure, style)
  if (node.type === 'accent') return accent(node, x, baseline, size, measure, style)
  const base = nodeLayout(node.base, 0, 0, size, measure, style), scriptSize = size * 0.68
  const sup = node.sup && nodeLayout(node.sup, 0, 0, scriptSize, measure, style)
  const sub = node.sub && nodeLayout(node.sub, 0, 0, scriptSize, measure, style)
  const limits = node.base.type === 'operator' && node.base.limits || node.base.type === 'accent' && node.base.accent === 'overbrace'
  if (limits) {
    const width = Math.max(base.width, sup?.width ?? 0, sub?.width ?? 0)
    return metrics(width, [
      ...translate(base.primitives, x + (width - base.width) / 2, baseline),
      ...(sup ? translate(sup.primitives, x + (width - sup.width) / 2, baseline - base.baseline - size * 0.12 - descent(sup)) : []),
      ...(sub ? translate(sub.primitives, x + (width - sub.width) / 2, baseline + descent(base) + size * 0.12 + sub.baseline) : []),
    ], baseline, size)
  }
  const supBaseline = baseline - Math.max(size * 0.42, base.baseline * 0.55)
  const subBaseline = baseline + Math.max(size * 0.32, descent(base) * 0.75)
  return metrics(base.width + Math.max(sup?.width ?? 0, sub?.width ?? 0), [
    ...translate(base.primitives, x, baseline),
    ...(sup ? translate(sup.primitives, x + base.width, supBaseline) : []),
    ...(sub ? translate(sub.primitives, x + base.width, subBaseline) : []),
  ], baseline, size)
}

function matrix(node: Extract<LatexNode, { type: 'matrix' }>, x: number, baseline: number, size: number, measure: Measure, style?: LatexTextStyle): LatexLayout {
  const columns = Math.max(...node.rows.map(row => row.length)), widths = Array<number>(columns).fill(0)
  const rows = node.rows.map(row => row.map((cell, column) => {
    const value = nodeLayout(cell, 0, 0, size, measure, style)
    widths[column] = Math.max(widths[column]!, value.width); return value
  }))
  const rowMetrics = rows.map(row => ({ ascent: Math.max(size * 0.9, ...row.map(value => value.baseline)), descent: Math.max(size * 0.3, ...row.map(descent)) }))
  const columnGap = size * 0.7, rowGap = size * 0.35
  const width = widths.reduce((sum, value) => sum + value, 0) + (columns - 1) * columnGap
  const height = rowMetrics.reduce((sum, value) => sum + value.ascent + value.descent, 0) + (rows.length - 1) * rowGap
  let rowY = -size * 0.12 - height / 2
  const primitives: LatexPrimitive[] = []
  rows.forEach((row, index) => {
    const metric = rowMetrics[index]!
    let columnX = 0
    row.forEach((cell, column) => {
      primitives.push(...translate(cell.primitives, columnX + (widths[column]! - cell.width) / 2, rowY + metric.ascent))
      columnX += widths[column]! + columnGap
    })
    rowY += metric.ascent + metric.descent + rowGap
  })
  return delimit(metrics(width, primitives, 0, size), x, baseline, size, measure, node.left, node.right)
}

function delimit(inner: LatexLayout, x: number, baseline: number, size: number, measure: Measure, left: string, right: string): LatexLayout {
  const delimiterSize = Math.max(size, inner.height / 1.2), delimiterBaseline = baseline - inner.baseline + delimiterSize * 0.9
  const leftLayout = left ? textLayout(left, x, delimiterBaseline, delimiterSize, measure) : undefined
  const gap = size * 0.08, insideX = x + (leftLayout ? leftLayout.width + gap : 0)
  const rightX = insideX + inner.width + (right ? gap : 0)
  const rightLayout = right ? textLayout(right, rightX, delimiterBaseline, delimiterSize, measure) : undefined
  return metrics(rightX - x + (rightLayout?.width ?? 0), [
    ...(leftLayout?.primitives ?? []), ...translate(inner.primitives, insideX, baseline), ...(rightLayout?.primitives ?? []),
  ], baseline, size)
}

function accent(node: Extract<LatexNode, { type: 'accent' }>, x: number, baseline: number, size: number, measure: Measure, style?: LatexTextStyle): LatexLayout {
  const inner = nodeLayout(node.value, 0, 0, size, measure, style)
  const dot = node.accent === 'dot' || node.accent === 'ddot' ? textLayout('·', 0, 0, size * 0.65, measure, 'roman') : undefined
  const width = Math.max(inner.width, size * 0.25, dot ? dot.width + (node.accent === 'ddot' ? size * 0.32 : 0) : 0)
  const top = baseline - inner.baseline - size * 0.1, bottom = baseline + descent(inner) + size * 0.08, center = x + width / 2
  const primitives: LatexPrimitive[] = translate(inner.primitives, x + (width - inner.width) / 2, baseline)
  if (node.accent === 'hat') {
    primitives.push(line(x, top, center, top - size * 0.23, size), line(center, top - size * 0.23, x + width, top, size))
  } else if (node.accent === 'dot' || node.accent === 'ddot') {
    const dots = node.accent === 'dot' ? [center] : [center - size * 0.16, center + size * 0.16]
    for (const point of dots) {
      if (dot) primitives.push(...translate(dot.primitives, point - dot.width / 2, top + size * 0.03))
    }
  } else if (node.accent === 'overbrace') {
    const bump = size * 0.16, edge = Math.min(size * 0.18, width / 4)
    primitives.push(line(x, top, x + edge, top - bump, size), line(x + edge, top - bump, center - edge, top - bump, size),
      line(center - edge, top - bump, center, top - bump * 2, size), line(center, top - bump * 2, center + edge, top - bump, size),
      line(center + edge, top - bump, x + width - edge, top - bump, size), line(x + width - edge, top - bump, x + width, top, size))
  } else {
    const y = node.accent === 'underline' ? bottom : top
    primitives.push(line(x, y, x + width, y, size))
    if (node.accent === 'vec') primitives.push(line(x + width - size * 0.2, y - size * 0.12, x + width, y, size), line(x + width - size * 0.2, y + size * 0.12, x + width, y, size))
  }
  return metrics(width, primitives, baseline, size)
}
function descent(layout: LatexLayout): number { return layout.height - layout.baseline }
function line(x1: number, y1: number, x2: number, y2: number, size: number): Extract<LatexPrimitive, { kind: 'line' }> {
  return { kind: 'line', x1, y1, x2, y2, width: Math.max(0.8, size / 18) }
}
/** All primitive bounds use one alphabetic baseline and feed Canvas hit geometry. */
function metrics(width: number, primitives: LatexPrimitive[], baseline: number, fallback: number): LatexLayout {
  let top = baseline - fallback * 0.9, bottom = baseline + fallback * 0.3
  for (const primitive of primitives) {
    if (primitive.kind === 'text') {
      const size = Number(/([\d.]+)px/.exec(primitive.font)?.[1] ?? fallback)
      top = Math.min(top, primitive.y - size * 0.9); bottom = Math.max(bottom, primitive.y + size * 0.3)
    } else {
      top = Math.min(top, primitive.y1 - primitive.width / 2, primitive.y2 - primitive.width / 2)
      bottom = Math.max(bottom, primitive.y1 + primitive.width / 2, primitive.y2 + primitive.width / 2)
    }
  }
  return { width, height: bottom - top, baseline: baseline - top, primitives }
}
function translate(primitives: LatexPrimitive[], x: number, y: number): LatexPrimitive[] {
  return primitives.map(primitive => primitive.kind === 'text' ? { ...primitive, x: primitive.x + x, y: primitive.y + y }
    : { ...primitive, x1: primitive.x1 + x, x2: primitive.x2 + x, y1: primitive.y1 + y, y2: primitive.y2 + y })
}
