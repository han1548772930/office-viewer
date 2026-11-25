export type LatexFormulaMode = 'inline' | 'block'
export interface LatexFormula { latex: string; mode: LatexFormulaMode; version: 1 }
export type LatexTextStyle = 'roman' | 'bold' | 'italic' | 'sans' | 'monospace'
export type LatexAccent = 'hat' | 'bar' | 'vec' | 'dot' | 'ddot' | 'overline' | 'underline' | 'overbrace'
export type LatexNode =
  | { type: 'text'; value: string }
  | { type: 'sequence'; children: LatexNode[] }
  | { type: 'space'; em: number }
  | { type: 'styled'; style: LatexTextStyle; value: LatexNode }
  | { type: 'operator'; value: string; large: boolean; limits: boolean }
  | { type: 'fraction'; numerator: LatexNode; denominator: LatexNode }
  | { type: 'sqrt'; value: LatexNode; index?: LatexNode }
  | { type: 'script'; base: LatexNode; sup?: LatexNode; sub?: LatexNode }
  | { type: 'delimited'; left: string; right: string; value: LatexNode }
  | { type: 'accent'; accent: LatexAccent; value: LatexNode }
  | { type: 'matrix'; rows: LatexNode[][]; left: string; right: string }
export interface LatexMetrics { width: number; height: number; baseline: number }
export type LatexPrimitive =
  | { kind: 'text'; x: number; y: number; text: string; font: string }
  | { kind: 'line'; x1: number; y1: number; x2: number; y2: number; width: number }
export interface LatexLayout extends LatexMetrics { primitives: LatexPrimitive[] }
