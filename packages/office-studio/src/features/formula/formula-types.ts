import type { CellRange, CellScalar, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'

export interface FormulaReference {
  kind: 'cell'
  sheetName?: string
  row: number
  column: number
  absoluteRow: boolean
  absoluteColumn: boolean
}

export interface FormulaRangeReference {
  kind: 'range'
  sheetName?: string
  start: FormulaReference
  end: FormulaReference
}

export interface FormulaLiteralNode { kind: 'literal'; value: CellScalar }
export interface FormulaErrorNode { kind: 'error'; code: string }
export interface FormulaReferenceNode { kind: 'reference'; reference: FormulaReference }
export interface FormulaRangeNode { kind: 'range'; reference: FormulaRangeReference }
export interface FormulaNameNode { kind: 'name'; name: string; sheetName?: string }
export interface FormulaStructuredReferenceNode { kind: 'structured'; tableName?: string; sheetName?: string; specifier: string }
export interface FormulaUnaryNode { kind: 'unary'; operator: '+' | '-'; operand: FormulaNode }
export interface FormulaBinaryNode { kind: 'binary'; operator: string; left: FormulaNode; right: FormulaNode }
export interface FormulaFunctionNode { kind: 'function'; name: string; arguments: FormulaNode[] }
export type FormulaNode = FormulaLiteralNode | FormulaErrorNode | FormulaReferenceNode | FormulaRangeNode | FormulaNameNode | FormulaStructuredReferenceNode | FormulaUnaryNode | FormulaBinaryNode | FormulaFunctionNode

export interface FormulaCellAddress { sheetId: string; row: number; column: number }
export interface FormulaEngineDiagnostic {
  code: string
  message: string
  cell: FormulaCellAddress
  formula: string
}
export interface FormulaDependencyGraph {
  /** Sparse range-aware references; no million-cell expansion for whole-column ranges. */
  dependencies: Map<string, Array<FormulaReference | FormulaRangeReference>>
  dependents: Map<string, Set<string>>
  diagnostics: FormulaEngineDiagnostic[]
}
export interface FormulaRecalculationResult {
  workbook: WorkbookModel
  graph: FormulaDependencyGraph
  recalculatedCells: FormulaCellAddress[]
  diagnostics: FormulaEngineDiagnostic[]
}

export interface FormulaEvaluationOptions {
  /** Display formatting uses the host region; scalar formula results are locale independent. */
  numberFormatLocale?: string
  /** A fixed clock value for deterministic evaluation and consistent recalculation. */
  now?: Date
  /** Cell whose formula is evaluated, including conditional formulas outside the stored cell map. */
  formulaCell?: Pick<FormulaCellAddress, 'row' | 'column'>
  /** Optional host policy, applied before actual reads and nested formula/cache access. */
  canReadCell?: (sheet: SpreadsheetSheet, row: number, column: number) => boolean
  /** Full ranges, including blank cells and intersecting merged regions. Supply
   * this for sparse/whole-axis queries rather than enumerating cell permissions. */
  canReadRange?: (sheet: SpreadsheetSheet, range: CellRange) => boolean
}
export interface FormulaResolvedReference { kind: 'reference'; sheetId: string; sheetName: string; range: CellRange }
export type FormulaValueResolver = ((sheet: SpreadsheetSheet, row: number, column: number) => FormulaValue) & {
  canReadRange?: (sheet: SpreadsheetSheet, range: CellRange) => boolean
}

export interface FormulaErrorValue { kind: 'error'; code: string }
export type FormulaValue = CellScalar | FormulaErrorValue

export function formulaCellKey(address: FormulaCellAddress): string {
  return `${address.sheetId}!${address.row}:${address.column}`
}

export function isFormulaError(value: unknown): value is FormulaErrorValue {
  return Boolean(value && typeof value === 'object' && 'kind' in value && (value as { kind?: unknown }).kind === 'error')
}

export function formulaError(code: string): FormulaErrorValue {
  return { kind: 'error', code }
}

