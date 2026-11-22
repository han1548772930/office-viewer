import type {
  CellRange,
  CellScalar,
  SheetDataValidationErrorStyle,
  SheetDataValidationOperator,
  SheetDataValidationRule,
  SheetDataValidationType,
  SpreadsheetCell,
  SpreadsheetSheet,
  WorkbookModel,
} from '../../domain/model.ts'
import { evaluateFormula } from '../formula/formula-engine.ts'
import { parseFormula } from '../formula/formula-parser.ts'
import { isFormulaError, type FormulaNode } from '../formula/formula-types.ts'
import { translateFormulaReferences } from '../formula/formula-references.ts'
import { parseTimelineDate } from '../slicer/timeline-date.ts'

export interface SheetValidationPosition { row: number; column: number }

export interface SheetValidationIssue extends SheetValidationPosition {
  ruleId: string
  type: SheetDataValidationType
  value: CellScalar
  severity: SheetDataValidationErrorStyle
  title: string
  message: string
}

export interface SheetValidationResult {
  valid: boolean
  blocked: boolean
  issues: SheetValidationIssue[]
}

export interface SheetDataValidationOptions {
  allowBlank?: boolean
  showDropDown?: boolean
  showErrorMessage?: boolean
  error?: string
  errorTitle?: string
  errorStyle?: SheetDataValidationErrorStyle
  showInputMessage?: boolean
  prompt?: string
  promptTitle?: string
  renderMode?: SheetDataValidationRule['renderMode']
}

export type SheetDataValidationDraft = Omit<SheetDataValidationRule, 'id' | 'ranges'> & { id?: string }

const DATA_VALIDATION_TYPES: readonly SheetDataValidationType[] = [
  'whole', 'decimal', 'textLength', 'date', 'time', 'checkbox', 'list', 'listMultiple', 'custom',
]
const DATA_VALIDATION_OPERATORS: readonly SheetDataValidationOperator[] = [
  'between', 'equal', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual', 'notBetween', 'notEqual',
]

export function createSheetDataValidationRule(
  draft: SheetDataValidationDraft,
  ranges: readonly CellRange[],
): SheetDataValidationRule {
  if (!DATA_VALIDATION_TYPES.includes(draft.type)) throw new TypeError('数据验证类型无效。')
  if (!ranges.length || ranges.some(range => !isValidRange(range))) throw new RangeError('请选择有效的数据验证区域。')
  if (draft.operator !== undefined && !DATA_VALIDATION_OPERATORS.includes(draft.operator)) throw new TypeError('数据验证条件无效。')
  if (requiresFormula1(draft.type, draft.operator) && !draft.formula1?.trim()) throw new TypeError('此数据验证类型需要规则值。')
  if (requiresFormula2(draft.operator) && !draft.formula2?.trim()) throw new TypeError('此比较条件需要第二个规则值。')
  validateRuleSettings(draft)
  return {
    ...draft,
    id: draft.id?.trim() || crypto.randomUUID(),
    ranges: mergeRanges(ranges),
    allowBlank: draft.allowBlank ?? true,
    showDropDown: draft.showDropDown ?? true,
    showErrorMessage: draft.showErrorMessage ?? true,
    errorStyle: draft.errorStyle ?? 'stop',
    showInputMessage: draft.showInputMessage ?? false,
    renderMode: draft.renderMode ?? 'arrow',
  }
}

export function applySheetDataValidationRule(sheet: SpreadsheetSheet, rule: SheetDataValidationRule): void {
  const existingRules = sheet.dataValidations ?? []
  const sameRule = existingRules.find(existing => existing.id === rule.id)
  const retainedRanges = sameRule?.ranges.flatMap(range => subtractRanges(range, rule.ranges)) ?? []
  const updatedRule = { ...rule, ranges: mergeRanges([...retainedRanges, ...rule.ranges]) }
  const next = existingRules.flatMap(existing => {
    if (existing.id === rule.id) {
      return retainedRanges.length ? [{ ...existing, id: crypto.randomUUID(), ranges: retainedRanges }] : []
    }
    const ranges = existing.ranges.flatMap(range => subtractRanges(range, rule.ranges))
    return ranges.length ? [{ ...existing, ranges }] : []
  })
  sheet.dataValidations = [...next, updatedRule]
}

export function replaceSheetDataValidationRule(sheet: SpreadsheetSheet, rule: SheetDataValidationRule): void {
  sheet.dataValidations = (sheet.dataValidations ?? []).filter(existing => existing.id !== rule.id)
  applySheetDataValidationRule(sheet, rule)
}

export function removeSheetDataValidationRule(sheet: SpreadsheetSheet, ruleId: string): boolean {
  const existingRules = sheet.dataValidations ?? []
  const nextRules = existingRules.filter(rule => rule.id !== ruleId)
  if (nextRules.length === existingRules.length) return false
  sheet.dataValidations = nextRules
  return true
}

export function clearSheetDataValidationRanges(sheet: SpreadsheetSheet, ranges: readonly CellRange[]): void {
  clearSheetDataValidation(sheet, ranges)
}

export function clearSheetDataValidationTypes(
  sheet: SpreadsheetSheet,
  ranges: readonly CellRange[],
  types: readonly SheetDataValidationType[],
): void {
  if (types.length === 0) throw new RangeError('请选择要清除的数据验证类型。')
  clearSheetDataValidation(sheet, ranges, new Set(types))
}

function clearSheetDataValidation(
  sheet: SpreadsheetSheet,
  ranges: readonly CellRange[],
  types?: ReadonlySet<SheetDataValidationType>,
): void {
  sheet.dataValidations = (sheet.dataValidations ?? []).flatMap(rule => {
    if (types && !types.has(rule.type)) return [rule]
    const remaining = rule.ranges.flatMap(range => subtractRanges(range, ranges))
    return remaining.length ? [{ ...rule, ranges: remaining }] : []
  })
}

export function dataValidationsAt(sheet: SpreadsheetSheet, row: number, column: number): SheetDataValidationRule[] {
  return (sheet.dataValidations ?? []).filter(rule => rule.ranges.some(range => contains(range, row, column)))
}

export function validateSheetDataValidationChanges(
  workbook: WorkbookModel,
  sheetId: string,
  positions: readonly SheetValidationPosition[],
): SheetValidationResult {
  const sheet = workbook.sheets.find(item => item.id === sheetId)
  if (!sheet) throw new Error(`Worksheet not found: ${sheetId}`)
  const issues = positions.flatMap(position => {
    const cell = sheet.cells[`${position.row}:${position.column}`]
    return dataValidationsAt(sheet, position.row, position.column)
      .filter(rule => !isValidValue(rule, cell, workbook, sheet, position.row, position.column))
      .map(rule => toIssue(rule, cell?.value ?? null, position))
  })
  return { valid: issues.length === 0, blocked: issues.some(issue => issue.severity === 'stop'), issues }
}

export function sheetDataValidationOptions(
  rule: SheetDataValidationRule,
  workbook: WorkbookModel,
  sheetId: string,
): string[] {
  if (rule.type === 'checkbox') return [rule.formula1 ?? 'TRUE', rule.formula2 ?? 'FALSE']
  if (rule.type !== 'list' && rule.type !== 'listMultiple') return []
  const inline = inlineListValues(rule.formula1 ?? '')
  if (inline) return inline
  const source = validationListSource(rule.formula1 ?? '', workbook, sheetId)
  if (!source) return []
  const sheet = source.sheet
  const values: string[] = []
  const positions = Object.values(sheet.cells).filter(cell => contains(source.range, cell.row, cell.col))
    .sort((left, right) => left.row - right.row || left.col - right.col)
  for (const cell of positions) {
      const value = cell.formula ? evaluateFormula(cell.formula, workbook, sheet.id) : cell.value
      if (!isFormulaError(value) && value !== null && value !== '') values.push(String(value))
  }
  return [...new Set(values)]
}

export function validationPromptForCell(sheet: SpreadsheetSheet, row: number, column: number): { title: string; message: string } | undefined {
  const rule = dataValidationsAt(sheet, row, column).find(item => item.showInputMessage && (item.prompt || item.promptTitle))
  if (!rule) return undefined
  return { title: rule.promptTitle ?? '数据验证', message: rule.prompt ?? '' }
}

function isValidValue(
  rule: SheetDataValidationRule,
  cell: SpreadsheetCell | undefined,
  workbook: WorkbookModel,
  sheet: SpreadsheetSheet,
  row: number,
  column: number,
): boolean {
  const actual = cell?.formula ? evaluateFormula(cell.formula, workbook, sheet.id) : cell?.value ?? null
  if (isBlank(actual)) return rule.allowBlank !== false
  if (isFormulaError(actual)) return false

  if (rule.type === 'checkbox') {
    if (rule.formula1 === undefined && rule.formula2 === undefined) return typeof actual === 'boolean' || ['TRUE', 'FALSE', '1', '0'].includes(String(actual).toUpperCase())
    return [rule.formula1, rule.formula2].some(value => value !== undefined && String(actual) === value)
  }
  if (rule.type === 'list' || rule.type === 'listMultiple') {
    const options = sheetDataValidationOptions(rule, workbook, sheet.id).map(normalizeListValue)
    if (!options.length) return false
    const selected = rule.type === 'listMultiple' ? String(actual).split(',').map(item => item.trim()).filter(Boolean) : [String(actual)]
    return selected.length > 0 && selected.every(item => options.includes(normalizeListValue(item)))
  }
  if (rule.type === 'custom') {
    const formula = rule.formula1?.trim()
    if (!formula) return false
    const shifted = shiftFormulaForCell(formula, rule.ranges[0], row, column)
    const result = evaluateFormula(shifted, workbook, sheet.id)
    return !isFormulaError(result) && (result === true || (typeof result === 'number' && result !== 0))
  }

  const typedValue = validationValue(rule.type, actual, workbook.date1904)
  if (typedValue === undefined) return false
  if (!rule.operator) return true
  const first = evaluateOperand(rule.formula1 && shiftFormulaForCell(rule.formula1, rule.ranges[0], row, column), workbook, sheet.id, rule.type)
  const second = evaluateOperand(rule.formula2 && shiftFormulaForCell(rule.formula2, rule.ranges[0], row, column), workbook, sheet.id, rule.type)
  if (first === undefined || (requiresFormula2(rule.operator) && second === undefined)) return false
  return compareByOperator(typedValue, first, second, rule.operator ?? 'between')
}

function validationValue(type: SheetDataValidationType, value: CellScalar, date1904: boolean): number | undefined {
  if (type === 'textLength') return String(value).length
  if (type === 'date') return parseTimelineDate(value, date1904)?.serial
  if (type === 'time') return timeFraction(value)
  if (type === 'whole' || type === 'decimal') {
    const number = typeof value === 'number' ? value : typeof value === 'string' && value.trim() ? Number(value) : undefined
    if (typeof number !== 'number' || !Number.isFinite(number)) return undefined
    if (type === 'whole' && !Number.isInteger(number)) return undefined
    return number
  }
  return undefined
}

function evaluateOperand(value: string | undefined, workbook: WorkbookModel, sheetId: string, type: SheetDataValidationType): number | undefined {
  if (value === undefined || !value.trim()) return undefined
  const expression = value.trim()
  const evaluated = evaluateFormula(expression.startsWith('=') ? expression : `=${expression}`, workbook, sheetId)
  if (isFormulaError(evaluated)) {
    if (type === 'date') return parseTimelineDate(expression, workbook.date1904)?.serial
    if (type === 'time') return timeFraction(expression)
    return undefined
  }
  if (type === 'date') return parseTimelineDate(evaluated, workbook.date1904)?.serial
  if (type === 'time') return timeFraction(evaluated)
  return typeof evaluated === 'number' && Number.isFinite(evaluated) ? evaluated : undefined
}

function compareByOperator(actual: number, first: number, second: number | undefined, operator: SheetDataValidationOperator): boolean {
  switch (operator) {
    case 'between': return second !== undefined && actual >= Math.min(first, second) && actual <= Math.max(first, second)
    case 'notBetween': return second !== undefined && (actual < Math.min(first, second) || actual > Math.max(first, second))
    case 'equal': return actual === first
    case 'notEqual': return actual !== first
    case 'greaterThan': return actual > first
    case 'greaterThanOrEqual': return actual >= first
    case 'lessThan': return actual < first
    case 'lessThanOrEqual': return actual <= first
  }
}

function toIssue(rule: SheetDataValidationRule, value: CellScalar, position: SheetValidationPosition): SheetValidationIssue {
  return {
    ...position,
    ruleId: rule.id,
    type: rule.type,
    value,
    severity: rule.errorStyle ?? 'stop',
    title: rule.showErrorMessage !== false && rule.errorTitle ? rule.errorTitle : '数据验证未通过',
    message: rule.showErrorMessage !== false && rule.error ? rule.error : defaultErrorMessage(rule.type),
  }
}

function defaultErrorMessage(type: SheetDataValidationType): string {
  const labels: Record<SheetDataValidationType, string> = {
    whole: '请输入符合整数条件的值。', decimal: '请输入符合数字条件的值。', textLength: '输入的文本长度不符合条件。',
    date: '请输入符合日期条件的值。', time: '请输入符合时间条件的值。', checkbox: '请输入复选框允许的值。',
    list: '请从列表中选择一个值。', listMultiple: '请从列表中选择允许的值。', custom: '输入值未通过自定义公式验证。',
  }
  return labels[type]
}

function isBlank(value: unknown): boolean { return value === null || value === undefined || value === '' }
function isValidRange(range: CellRange): boolean {
  return Number.isInteger(range.startRow) && Number.isInteger(range.startCol) && Number.isInteger(range.endRow) && Number.isInteger(range.endCol)
    && range.startRow >= 0 && range.startCol >= 0 && range.endRow >= range.startRow && range.endCol >= range.startCol
}
function requiresFormula1(type: SheetDataValidationType, operator?: SheetDataValidationOperator): boolean {
  return type === 'list' || type === 'listMultiple' || type === 'custom' || (type !== 'checkbox' && operator !== undefined)
}
function requiresFormula2(operator?: SheetDataValidationOperator): boolean { return operator === 'between' || operator === 'notBetween' }
function contains(range: CellRange, row: number, column: number): boolean {
  return row >= range.startRow && row <= range.endRow && column >= range.startCol && column <= range.endCol
}
function normalizeListValue(value: string): string { return value.trim().toLocaleLowerCase() }

function timeFraction(value: CellScalar): number | undefined {
  if (typeof value === 'number') return Number.isFinite(value) && value >= 0 && value < 1 ? value : undefined
  if (typeof value !== 'string') return undefined
  const parts = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(value.trim())
  if (!parts) return undefined
  const hours = Number(parts[1]), minutes = Number(parts[2]), seconds = Number(parts[3] ?? 0)
  if (hours > 23 || minutes > 59 || seconds > 59) return undefined
  return (hours * 3600 + minutes * 60 + seconds) / 86400
}

function inlineListValues(formula?: string): string[] | undefined {
  const value = formula?.trim()
  if (!value || value.startsWith('=')) return undefined
  return parseValidationListItems(value)
}

export function serializeValidationListItems(values: readonly string[]): string {
  const items = values.map(value => /[,"\r\n]/.test(value) ? `"${value.replaceAll('"', '""')}"` : value)
  return `"${items.join(',')}"`
}

export function parseValidationListItems(value: string): string[] {
  const content = value.startsWith('"') && value.endsWith('"') ? value.slice(1, -1) : value
  const items: string[] = []
  let current = '', quoted = false
  for (let index = 0; index < content.length; index++) {
    const character = content[index]!
    if (character === '"' && content[index + 1] === '"' && (quoted || current.length > 0)) { current += '"'; index++; continue }
    if (character === '"' && (quoted || current.length === 0)) { quoted = !quoted; continue }
    if (character === ',' && !quoted) { items.push(current.trim()); current = ''; continue }
    current += character
  }
  items.push(current.trim())
  return items.filter(item => item.length > 0)
}

function validateRuleSettings(rule: SheetDataValidationDraft): void {
  if (rule.type === 'checkbox' && rule.formula1 !== undefined && rule.formula2 !== undefined && rule.formula1 === rule.formula2) {
    throw new TypeError('复选框的勾选值和未勾选值必须不同。')
  }
  if (rule.type === 'custom' || ((rule.type === 'list' || rule.type === 'listMultiple') && rule.formula1?.trim().startsWith('='))) {
    validateFormula(rule.formula1!, rule.type === 'list' || rule.type === 'listMultiple')
  }
  for (const formula of [rule.formula1, requiresFormula2(rule.operator) ? rule.formula2 : undefined]) {
    if (formula?.trim().startsWith('=')) validateFormula(formula, false)
  }
}

function validateFormula(formula: string, allowListReferenceFunctions: boolean): void {
  const node = parseFormula(formula)
  const functions = collectFormulaFunctions(node)
  if (allowListReferenceFunctions && functions.some(name => !['IF', 'INDIRECT', 'CHOOSE', 'OFFSET'].includes(name))) {
    throw new TypeError('下拉列表引用公式仅支持 IF、INDIRECT、CHOOSE、OFFSET 和区域引用。')
  }
}

function collectFormulaFunctions(node: FormulaNode): string[] {
  if (node.kind === 'function') return [node.name.toUpperCase(), ...node.arguments.flatMap(collectFormulaFunctions)]
  if (node.kind === 'unary') return collectFormulaFunctions(node.operand)
  if (node.kind === 'binary') return [...collectFormulaFunctions(node.left), ...collectFormulaFunctions(node.right)]
  return []
}

function validationListSource(formula: string, workbook: WorkbookModel, sheetId: string): { sheet: SpreadsheetSheet; range: CellRange } | undefined {
  const source = parseListRange(formula)
  if (!source) return undefined
  const sheet = source.sheetName
    ? workbook.sheets.find(item => item.name.toLocaleLowerCase() === source.sheetName!.toLocaleLowerCase())
    : workbook.sheets.find(item => item.id === sheetId)
  return sheet ? { sheet, range: source.range } : undefined
}

function parseListRange(formula: string): { sheetName?: string; range: CellRange } | undefined {
  const match = /^=?\s*(?:'((?:[^']|'')+)'|([\w ]+))?!?\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?\s*$/i.exec(formula)
  if (!match) return undefined
  const first = parseAddress(`${match[3]}${match[4]}`)
  const last = parseAddress(`${match[5] ?? match[3]}${match[6] ?? match[4]}`)
  if (!first || !last) return undefined
  const quotedName = match[1]?.replace(/''/g, "'")
  const simpleName = match[2]?.trim()
  return {
    sheetName: quotedName ?? (simpleName && simpleName.toUpperCase() !== 'TRUE' ? simpleName : undefined),
    range: { startRow: Math.min(first.row, last.row), startCol: Math.min(first.column, last.column), endRow: Math.max(first.row, last.row), endCol: Math.max(first.column, last.column) },
  }
}

function parseAddress(value: string): { row: number; column: number } | undefined {
  const match = /^([A-Z]{1,3})(\d+)$/i.exec(value)
  if (!match) return undefined
  let column = 0
  for (const character of match[1]!.toUpperCase()) column = column * 26 + character.charCodeAt(0) - 64
  const row = Number(match[2]) - 1
  return row >= 0 ? { row, column: column - 1 } : undefined
}

function shiftFormulaForCell(formula: string, anchor: CellRange | undefined, row: number, column: number): string {
  if (!anchor) return formula
  return translateFormulaReferences(formula, row - anchor.startRow, column - anchor.startCol)
}

function subtractRanges(source: CellRange, cuts: readonly CellRange[]): CellRange[] {
  return cuts.reduce<CellRange[]>((ranges, cut) => ranges.flatMap(range => subtractRange(range, cut)), [source])
}

function subtractRange(source: CellRange, cut: CellRange): CellRange[] {
  const top = Math.max(source.startRow, cut.startRow), left = Math.max(source.startCol, cut.startCol)
  const bottom = Math.min(source.endRow, cut.endRow), right = Math.min(source.endCol, cut.endCol)
  if (top > bottom || left > right) return [source]
  const result: CellRange[] = []
  if (source.startRow < top) result.push({ ...source, endRow: top - 1 })
  if (bottom < source.endRow) result.push({ ...source, startRow: bottom + 1 })
  if (source.startCol < left) result.push({ startRow: top, endRow: bottom, startCol: source.startCol, endCol: left - 1 })
  if (right < source.endCol) result.push({ startRow: top, endRow: bottom, startCol: right + 1, endCol: source.endCol })
  return result
}

function mergeRanges(ranges: readonly CellRange[]): CellRange[] {
  const result: CellRange[] = []
  for (const range of ranges) {
    if (!result.some(existing => sameRange(existing, range))) result.push({ ...range })
  }
  return result
}

function sameRange(left: CellRange, right: CellRange): boolean {
  return left.startRow === right.startRow && left.endRow === right.endRow && left.startCol === right.startCol && left.endCol === right.endCol
}
