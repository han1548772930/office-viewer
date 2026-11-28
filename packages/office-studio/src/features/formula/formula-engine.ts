import type { CellScalar, SpreadsheetCell, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import { normalizeSheetCellValueType } from '../../domain/sheet-cell-value.ts'
import { formulaCellKey, formulaError, isFormulaError } from './formula-types.ts'
import type {
  FormulaCellAddress, FormulaDependencyGraph, FormulaEngineDiagnostic, FormulaErrorValue, FormulaNode,
  FormulaEvaluationOptions, FormulaRangeReference, FormulaRecalculationResult, FormulaReference, FormulaValue,
  FormulaResolvedReference, FormulaValueResolver,
} from './formula-types.ts'
import { FormulaSyntaxError, parseFormula } from './formula-parser.ts'
import { formatSheetValue, isSheetTextNumberFormat } from '../sheet-workbench/sheet-number-format.ts'
import { parseDate, parseNumber, parseTime } from '../sheet-workbench/number-format-engine/parse-value.ts'
import { DateSystem } from '../sheet-workbench/number-format-engine/date-system.ts'
import { excelDateSerial, excelDateTimePartsToSerial, excelSerialToDateTimeParts, getDateSystemMaxSerial } from '../sheet-workbench/number-format-engine/serial-date.ts'
import { describeSheetNumberFormat } from '../sheet-workbench/number-format-render.ts'
import { indexFormulaDefinedNames } from './formula-defined-names.ts'
import { bindFormulaReferences, formulaRangeReference } from './formula-reference-binding.ts'
import { formulaCalendarSerial, formulaImportEnvironment, readFormulaImportSource, type FormulaCalendarValue, type FormulaImportAliases } from './formula-import-context.ts'
import { parseDynamicFormulaReference, remapDynamicFormulaReference } from './formula-dynamic-references.ts'
import { formulaReadCellAllowed, formulaReadRangeAllowed } from './formula-read-policy.ts'

interface RangeValue { kind: 'rangeValue'; sheet: SpreadsheetSheet; reference: FormulaRangeReference }
type RuntimeValue = FormulaValue | RangeValue

const ERROR_CODES = new Set(['#DIV/0!', '#N/A', '#NAME?', '#NULL!', '#NUM!', '#REF!', '#VALUE!', '#CIRCULAR!'])

const ERROR_TYPE_NUMBERS = new Map([
  ['#NULL!', 1], ['#DIV/0!', 2], ['#VALUE!', 3], ['#REF!', 4], ['#NAME?', 5],
  ['#NUM!', 6], ['#N/A', 7], ['#GETTING_DATA', 8], ['#CALC!', 14],
])

const FORMULA_TOP_LEVEL_DOMAINS = [
  'ac', 'ad', 'ae', 'aero', 'af', 'ag', 'ai', 'al', 'am', 'ao', 'aq', 'ar', 'arpa', 'as', 'asia', 'at', 'au', 'aw', 'ax', 'az', 'ba', 'bb',
  'bd', 'be', 'bf', 'bg', 'bh', 'bi', 'biz', 'bj', 'bm', 'bn', 'bo', 'br', 'bs', 'bt', 'bv', 'bw', 'by', 'bz', 'ca', 'cat', 'cc', 'cd',
  'cf', 'cg', 'ch', 'ci', 'ck', 'cl', 'cm', 'cn', 'co', 'com', 'coop', 'cr', 'cu', 'cv', 'cw', 'cx', 'cy', 'cz', 'de', 'dj', 'dk', 'dm',
  'do', 'dz', 'ec', 'edu', 'ee', 'eg', 'er', 'es', 'et', 'eu', 'fi', 'fj', 'fk', 'fm', 'fo', 'fr', 'ga', 'gb', 'gd', 'ge', 'gf', 'gg',
  'gh', 'gi', 'gl', 'gm', 'gn', 'gov', 'gp', 'gq', 'gr', 'gs', 'gt', 'gu', 'gw', 'gy', 'hk', 'hm', 'hn', 'hr', 'ht', 'hu', 'id', 'ie',
  'il', 'im', 'in', 'info', 'int', 'io', 'iq', 'ir', 'is', 'it', 'je', 'jm', 'jo', 'jobs', 'jp', 'ke', 'kg', 'kh', 'ki', 'km', 'kn', 'kp',
  'kr', 'kw', 'ky', 'kz', 'la', 'lb', 'lc', 'li', 'lk', 'lr', 'ls', 'lt', 'lu', 'lv', 'ly', 'ma', 'mc', 'md', 'me', 'mg', 'mh', 'mil',
  'mk', 'ml', 'mm', 'mn', 'mo', 'mobi', 'mp', 'mq', 'mr', 'ms', 'mt', 'mu', 'museum', 'mv', 'mw', 'mx', 'my', 'mz', 'na', 'name', 'nc', 'ne',
  'net', 'nf', 'ng', 'ni', 'nl', 'no', 'np', 'nr', 'nu', 'nz', 'om', 'onion', 'org', 'pa', 'pe', 'pf', 'pg', 'ph', 'pk', 'pl', 'pm', 'pn',
  'post', 'pr', 'pro', 'ps', 'pt', 'pw', 'py', 'qa', 're', 'ro', 'rs', 'ru', 'rw', 'sa', 'sb', 'sc', 'sd', 'se', 'sg', 'sh', 'si', 'sj',
  'sk', 'sl', 'sm', 'sn', 'so', 'sr', 'ss', 'st', 'su', 'sv', 'sx', 'sy', 'sz', 'tc', 'td', 'tel', 'tf', 'tg', 'th', 'tj', 'tk', 'tl',
  'tm', 'tn', 'to', 'tr', 'tt', 'tv', 'tw', 'tz', 'ua', 'ug', 'uk', 'us', 'uy', 'uz', 'va', 'vc', 've', 'vg', 'vi', 'vn', 'vu', 'wf',
  'ws', 'yt', 'za', 'zm', 'zw',
].join('|')
const FORMULA_EMAIL_PATTERN = new RegExp(String.raw`^(?:[\w+-]+\.)*[\w+-]+@[a-z0-9_-]+(?:\.[a-z0-9_-]+)*\.(?:${FORMULA_TOP_LEVEL_DOMAINS})$`, 'i')
const FORMULA_WEB_URL_PATH = String.raw`(?::[0-9]+)?(?:/(?:[A-Za-z0-9\-._~!$&'()*+,;=:@]|%[A-Fa-f0-9]{2})*)*/?(?:[?#]\S*)?`
const FORMULA_URL_PATTERN = new RegExp(`^(?:${[
  String.raw`(?:https?|s?ftp|ftps|nfs|ssh)://+[a-z0-9_-]+(?:\.[a-z0-9_-]+)*${FORMULA_WEB_URL_PATH}`,
  String.raw`[a-z0-9_-]+(?:\.[a-z0-9_-]+)*\.(?:${FORMULA_TOP_LEVEL_DOMAINS})${FORMULA_WEB_URL_PATH}`,
  String.raw`mailto:(?:[\w+-]+\.)*[\w+-]+@[a-z0-9_-]+(?:\.[a-z0-9_-]+)*\.(?:${FORMULA_TOP_LEVEL_DOMAINS})`,
  String.raw`(?:news|aim):[%a-z0-9$_\.+!*(),;/?#:@&~=-]+`,
].join('|')})$`, 'i')
const FORMULA_TEXT_LIGATURES = new Map([
  ['ß', 'ss'], ['æ', 'ae'], ['œ', 'oe'], ['ﬀ', 'ff'], ['ﬁ', 'fi'], ['ﬂ', 'fl'],
  ['ﬃ', 'ffi'], ['ﬄ', 'ffl'], ['ﬅ', 'st'], ['ﬆ', 'st'],
])

const ROMAN_DIGIT_VALUES = new Map<string, number>([
  ['I', 1], ['V', 5], ['X', 10], ['L', 50], ['C', 100], ['D', 500], ['M', 1000],
])
const ROMAN_NUMERAL_TEXT = new Map<number, string>([
  [1, 'I'], [4, 'IV'], [5, 'V'], [9, 'IX'], [10, 'X'], [40, 'XL'], [45, 'VL'], [49, 'IL'], [50, 'L'],
  [90, 'XC'], [95, 'VC'], [99, 'IC'], [100, 'C'], [400, 'CD'], [450, 'LD'], [490, 'XD'], [495, 'VD'],
  [499, 'ID'], [500, 'D'], [900, 'CM'], [950, 'LM'], [990, 'XM'], [995, 'VM'], [999, 'IM'], [1000, 'M'],
])
const ROMAN_FORM_VALUES = [
  [1, 4, 5, 9, 10, 40, 50, 90, 100, 400, 500, 900, 1000],
  [1, 4, 5, 9, 10, 40, 45, 50, 90, 95, 100, 400, 450, 500, 900, 950, 1000],
  [1, 4, 5, 9, 10, 40, 45, 49, 50, 90, 95, 99, 100, 400, 450, 490, 500, 900, 950, 990, 1000],
  [1, 4, 5, 9, 10, 40, 45, 49, 50, 90, 95, 99, 100, 400, 450, 490, 495, 500, 900, 950, 990, 995, 1000],
  [1, 4, 5, 9, 10, 40, 45, 49, 50, 90, 95, 99, 100, 400, 450, 490, 495, 499, 500, 900, 950, 990, 995, 999, 1000],
] as const

/** Parses references once and builds a sparse range-aware dependency index. */
export function createFormulaDependencyGraph(workbook: WorkbookModel): FormulaDependencyGraph {
  const dependencies = new Map<string, Array<FormulaReference | FormulaRangeReference>>()
  const dependents = new Map<string, Set<string>>()
  const diagnostics: FormulaEngineDiagnostic[] = []
  const sheetByName = new Map(workbook.sheets.map((sheet) => [sheet.name.toLocaleLowerCase(), sheet]))
  const names = indexFormulaDefinedNames(workbook)
  for (const sheet of workbook.sheets) for (const cell of Object.values(sheet.cells)) {
    if (!cell.formula) continue
    const owner: FormulaCellAddress = { sheetId: sheet.id, row: cell.row, column: cell.col }
    const ownerKey = formulaCellKey(owner)
    try {
      const node = bindFormulaReferences(parseFormula(cell.formula), workbook, sheet, names, owner)
      const refs = collectReferences(node)
      if (containsDynamicReference(node)) for (const source of workbook.sheets) refs.push(formulaRangeReference({ startRow: 0, startCol: 0,
        endRow: (source.rowCount ?? 1_048_576) - 1, endCol: (source.columnCount ?? 16_384) - 1 }, source.name))
      dependencies.set(ownerKey, refs)
      for (const reference of refs) {
        const targetSheet = resolveSheet(workbook, reference.sheetName, sheet, sheetByName)
        if (!targetSheet) continue
        const targets = reference.kind === 'cell'
          ? [formulaCellKey({ sheetId: targetSheet.id, row: reference.row, column: reference.column })]
          : getSparseRangeTargets(targetSheet, reference)
        for (const target of targets) {
          const formulas = dependents.get(target) ?? new Set<string>()
          formulas.add(ownerKey)
          dependents.set(target, formulas)
        }
      }
    } catch (error) {
      diagnostics.push(diagnostic(owner, cell.formula, error))
      dependencies.set(ownerKey, [])
    }
  }
  return { dependencies, dependents, diagnostics }
}

/**
 * Recalculates every supported formula in workbook order with memoized dependency
 * evaluation. Formula sources and formatting stay intact. Unsupported imported
 * formulas keep their previous cached result and receive a diagnostic instead
 * of being silently replaced by a false value.
 */
export function recalculateWorkbookFormulas(workbook: WorkbookModel, previous?: WorkbookModel, options: FormulaEvaluationOptions = {}): FormulaRecalculationResult {
  const evaluationTime = validEvaluationTime(options.now) ?? new Date()
  const graph = createFormulaDependencyGraph(workbook)
  const sheetsByName = new Map(workbook.sheets.map((sheet) => [sheet.name.toLocaleLowerCase(), sheet]))
  const memo = new Map<string, FormulaValue>()
  const calendarValues = new Map<string, FormulaCalendarValue>()
  const calculatedDate = (sheet: SpreadsheetSheet, row: number, column: number) => calendarValues.get(formulaCellKey({ sheetId: sheet.id, row, column }))
  const active: string[] = []
  const recalculatedCells: FormulaCellAddress[] = []
  const diagnostics = [...graph.diagnostics]
  const cells = new Map<string, { sheet: SpreadsheetSheet; cell: SpreadsheetCell }>()
  const names = indexFormulaDefinedNames(workbook)
  for (const sheet of workbook.sheets) for (const cell of Object.values(sheet.cells)) {
    cells.set(formulaCellKey({ sheetId: sheet.id, row: cell.row, column: cell.col }), { sheet, cell })
  }

  const evaluateCell: FormulaValueResolver = (sheet, row, column) => {
    if (!formulaReadCellAllowed(sheet, row, column, options)) return formulaError('#REF!')
    const address = { sheetId: sheet.id, row, column }
    const key = formulaCellKey(address)
    const cached = memo.get(key)
    if (cached !== undefined) return cached
    const current = sheet.cells[cellKey(row, column)]
    if (!current) return null
    if (!current.formula) return rawCellValue(current)
    const cycleStart = active.indexOf(key)
    if (cycleStart >= 0) {
      for (const cycleKey of active.slice(cycleStart)) memo.set(cycleKey, formulaError('#CIRCULAR!'))
      memo.set(key, formulaError('#CIRCULAR!'))
      return formulaError('#CIRCULAR!')
    }
    active.push(key)
    let result: FormulaValue
    let evaluationFailed = false
    const previousCell = previous?.sheets.find((candidate) => candidate.id === sheet.id)?.cells[cellKey(row, column)]
    try {
      const environment = formulaImportEnvironment(workbook, sheet, current, evaluateCell, calculatedDate)
      const node = bindFormulaReferences(parseFormula(current.formula), workbook, sheet, names, { row, column })
      result = normalizeResult(evaluateNode(node, sheet, environment.resolve, environment.workbook, sheetsByName, { row, column }, evaluationTime, environment.aliases), environment.resolve)
      if (environment.source?.dateResult && typeof result === 'number') {
        const targetValue = formulaCalendarSerial(result, environment.source.date1904, workbook.date1904)
        calendarValues.set(key, { date1904: environment.source.date1904, value: result, targetValue })
        result = targetValue
      }
    } catch (error) {
      result = toFormulaError(error)
      evaluationFailed = true
    }
    // Cached text and cached errors remain distinct when an imported formula is unsupported.
    const cacheSource = previousCell?.formula === current.formula ? previousCell : readFormulaImportSource(current) ? current : undefined
    if (!options.canReadCell && !options.canReadRange && (evaluationFailed || isFormulaError(result) && result.code === '#NAME?') && cacheSource && hasCachedValue(cacheSource)) {
      result = rawCellValue(cacheSource)
      calendarValues.delete(key)
      diagnostics.push({ code: 'FORMULA_UNSUPPORTED_CACHE_PRESERVED', message: `${current.formula} is not implemented; the imported cached value was retained`, cell: address, formula: current.formula })
    }
    if (!memo.has(key)) memo.set(key, result)
    const stored = memo.get(key)!
    if (active.at(-1) === key) active.pop()
    if (stored !== null || current.value !== null || current.display !== '') recalculatedCells.push(address)
    return stored
  }
  evaluateCell.canReadRange = (sheet, range) => formulaReadRangeAllowed(sheet, range, options)

  for (const { sheet, cell } of cells.values()) if (cell.formula) evaluateCell(sheet, cell.row, cell.col)

  const updatedSheets = workbook.sheets.map((sheet) => {
    let changed = false
    const updatedCells: Record<string, SpreadsheetCell> = { ...sheet.cells }
    for (const cell of Object.values(sheet.cells)) {
      if (!cell.formula) continue
      const value = memo.get(formulaCellKey({ sheetId: sheet.id, row: cell.row, column: cell.col }))
      if (value === undefined) continue
      const scalar: CellScalar = isFormulaError(value) ? value.code : value
      const display = isFormulaError(value) ? value.code
        : formatSheetValue(scalar, cell.style.numberFormat, workbook.date1904, options.numberFormatLocale) ?? displayFormulaValue(value)
      const errorCode = isFormulaError(value) ? value.code : undefined
      const calendar = calendarValues.get(formulaCellKey({ sheetId: sheet.id, row: cell.row, column: cell.col }))
      const dateSource = calendar && calendar.targetValue === value ? { version: 1, date1904: calendar.date1904, value: calendar.value,
        targetDate1904: workbook.date1904, targetValue: calendar.targetValue, targetFormula: cell.formula } : undefined
      if (cell.value === scalar && cell.display === display && cell.dataType === 'formula' && cell.cachedDataType === valueType(value)
        && cell.metadata?.formulaError === errorCode && (!dateSource || JSON.stringify(cell.metadata?.chartDependencyDateSource) === JSON.stringify(dateSource))) continue
      const metadata = { ...cell.metadata }
      if (dateSource) metadata.chartDependencyDateSource = dateSource
      if (errorCode) metadata.formulaError = errorCode
      else delete metadata.formulaError
      updatedCells[cellKey(cell.row, cell.col)] = {
        ...cell,
        value: scalar,
        display,
        dataType: 'formula',
        cachedDataType: valueType(value),
        metadata,
      }
      changed = true
    }
    return changed ? { ...sheet, cells: updatedCells } : sheet
  })
  const nextWorkbook = updatedSheets.some((sheet, index) => sheet !== workbook.sheets[index]) ? { ...workbook, sheets: updatedSheets } : workbook
  return { workbook: nextWorkbook, graph, recalculatedCells, diagnostics }
}

export function evaluateFormula(formula: string, workbook: WorkbookModel, currentSheetId: string, options: FormulaEvaluationOptions = {}): FormulaValue {
  return evaluateFormulaResult(formula, workbook, currentSheetId, options, 'scalar')
}
export function evaluateFormulaReference(formula: string, workbook: WorkbookModel, currentSheetId: string,
  options: FormulaEvaluationOptions = {}): FormulaResolvedReference | FormulaErrorValue {
  return evaluateFormulaResult(formula, workbook, currentSheetId, options, 'reference')
}
function evaluateFormulaResult(formula: string, workbook: WorkbookModel, currentSheetId: string, options: FormulaEvaluationOptions, mode: 'scalar'): FormulaValue
function evaluateFormulaResult(formula: string, workbook: WorkbookModel, currentSheetId: string, options: FormulaEvaluationOptions, mode: 'reference'): FormulaResolvedReference | FormulaErrorValue
function evaluateFormulaResult(formula: string, workbook: WorkbookModel, currentSheetId: string, options: FormulaEvaluationOptions,
  mode: 'scalar' | 'reference'): FormulaValue | FormulaResolvedReference {
  const evaluationTime = validEvaluationTime(options.now) ?? new Date()
  const sheet = workbook.sheets.find((candidate) => candidate.id === currentSheetId)
  if (!sheet) return formulaError('#REF!')
  if (options.formulaCell && (!Number.isSafeInteger(options.formulaCell.row) || !Number.isSafeInteger(options.formulaCell.column)
    || options.formulaCell.row < 0 || options.formulaCell.row >= 1_048_576 || options.formulaCell.column < 0 || options.formulaCell.column >= 16_384)) return formulaError('#REF!')
  const byName = new Map(workbook.sheets.map((item) => [item.name.toLocaleLowerCase(), item]))
  const memo = new Map<string, FormulaValue>(), active: string[] = []
  const calendarValues = new Map<string, FormulaCalendarValue>()
  const calculatedDate = (sheet: SpreadsheetSheet, row: number, column: number) => calendarValues.get(formulaCellKey({ sheetId: sheet.id, row, column }))
  const names = indexFormulaDefinedNames(workbook)
  const resolve: FormulaValueResolver = (targetSheet, row, column) => {
    if (!formulaReadCellAllowed(targetSheet, row, column, options)) return formulaError('#REF!')
    const key = formulaCellKey({ sheetId: targetSheet.id, row, column })
    if (memo.has(key)) return memo.get(key)!
    const cell = targetSheet.cells[cellKey(row, column)]
    if (!cell) return null
    if (!cell.formula) return rawCellValue(cell)
    if (active.includes(key)) return formulaError('#CIRCULAR!')
    active.push(key)
    let value: FormulaValue
    try {
      const environment = formulaImportEnvironment(workbook, targetSheet, cell, resolve, calculatedDate)
      const node = bindFormulaReferences(parseFormula(cell.formula), workbook, targetSheet, names, { row, column })
      value = normalizeResult(evaluateNode(node, targetSheet, environment.resolve, environment.workbook, byName, { row, column }, evaluationTime, environment.aliases), environment.resolve)
      if (environment.source?.dateResult && typeof value === 'number') {
        const targetValue = formulaCalendarSerial(value, environment.source.date1904, workbook.date1904)
        calendarValues.set(key, { date1904: environment.source.date1904, value, targetValue })
        value = targetValue
      }
    }
    catch (error) { value = toFormulaError(error) }
    active.pop()
    memo.set(key, value)
    return value
  }
  resolve.canReadRange = (sheet, range) => formulaReadRangeAllowed(sheet, range, options)
  try {
    const node = bindFormulaReferences(parseFormula(formula), workbook, sheet, names, options.formulaCell)
    const current = options.formulaCell ? sheet.cells[cellKey(options.formulaCell.row, options.formulaCell.column)] : undefined
    const environment = current?.formula === formula ? formulaImportEnvironment(workbook, sheet, current, resolve, calculatedDate) : { workbook, resolve, source: undefined, aliases: undefined }
    if (mode === 'reference') {
      const reference = evaluateReferenceOperand(node, sheet, environment.resolve, environment.workbook, byName, options.formulaCell, evaluationTime, environment.aliases)
      if (isFormulaError(reference)) return reference
      const range = referenceBounds(reference.reference)
      return formulaReadRangeAllowed(reference.sheet, range, options)
        ? { kind: 'reference', sheetId: reference.sheet.id, sheetName: reference.sheet.name, range } : formulaError('#REF!')
    }
    let result = normalizeResult(evaluateNode(node, sheet, environment.resolve, environment.workbook, byName, options.formulaCell, evaluationTime, environment.aliases), environment.resolve)
    if (environment.source?.dateResult && typeof result === 'number') result = formulaCalendarSerial(result, environment.source.date1904, workbook.date1904)
    return result
  }
  catch (error) { return toFormulaError(error) }
}

export function formulaValueToDisplay(value: FormulaValue): string { return displayFormulaValue(value) }
export function formulaErrorCode(error: FormulaErrorValue): string { return error.code }

function evaluateNode(
  node: FormulaNode,
  currentSheet: SpreadsheetSheet,
  resolveCell: FormulaValueResolver,
  workbook: WorkbookModel,
  sheetsByName: Map<string, SpreadsheetSheet>,
  formulaCell?: Pick<FormulaCellAddress, 'row' | 'column'>,
  evaluationTime: Date = new Date(),
  aliases?: FormulaImportAliases,
): RuntimeValue {
  switch (node.kind) {
    case 'literal': return node.value
    case 'reference': {
      const targetSheet = resolveSheet(workbook, node.reference.sheetName, currentSheet, sheetsByName)
      return targetSheet ? resolveCell(targetSheet, node.reference.row, node.reference.column) : formulaError('#REF!')
    }
    case 'range': {
      const targetSheet = resolveSheet(workbook, node.reference.sheetName, currentSheet, sheetsByName)
      return targetSheet ? checkedReference({ kind: 'rangeValue', sheet: targetSheet, reference: node.reference }, resolveCell) : formulaError('#REF!')
    }
    case 'unary': {
      const value = singleRangeValue(evaluateNode(node.operand, currentSheet, resolveCell, workbook, sheetsByName, formulaCell, evaluationTime, aliases), resolveCell)
      if (isFormulaError(value)) return value
      const number = numericCoercion(value)
      if (isFormulaError(number)) return number
      return finiteResult(node.operator === '-' ? -number : number)
    }
    case 'binary': {
      if (node.operator === ':') {
        const endpoint = (part: FormulaNode) => evaluateReferenceOperand(part, currentSheet, resolveCell, workbook, sheetsByName, formulaCell, evaluationTime, aliases)
        const left = endpoint(node.left)
        if (isFormulaError(left)) return left
        const right = endpoint(node.right)
        if (isFormulaError(right)) return right
        if (left.sheet.id !== right.sheet.id) return formulaError('#REF!')
        return checkedReference({ kind: 'rangeValue', sheet: left.sheet, reference: {
          kind: 'range', sheetName: left.reference.sheetName,
          start: { ...left.reference.start, row: Math.min(left.reference.start.row, left.reference.end.row, right.reference.start.row, right.reference.end.row),
            column: Math.min(left.reference.start.column, left.reference.end.column, right.reference.start.column, right.reference.end.column) },
          end: { ...right.reference.end, row: Math.max(left.reference.start.row, left.reference.end.row, right.reference.start.row, right.reference.end.row),
            column: Math.max(left.reference.start.column, left.reference.end.column, right.reference.start.column, right.reference.end.column) },
        } }, resolveCell)
      }
      const left = singleRangeValue(evaluateNode(node.left, currentSheet, resolveCell, workbook, sheetsByName, formulaCell, evaluationTime, aliases), resolveCell)
      if (isFormulaError(left)) return left
      const right = singleRangeValue(evaluateNode(node.right, currentSheet, resolveCell, workbook, sheetsByName, formulaCell, evaluationTime, aliases), resolveCell)
      if (isFormulaError(right)) return right
      return evaluateBinary(node.operator, scalarValue(left), scalarValue(right))
    }
    case 'function': return evaluateFunction(node.name, node.arguments, currentSheet, resolveCell, workbook, sheetsByName, formulaCell, evaluationTime, aliases)
    case 'name': return formulaError('#NAME?')
    case 'structured': return formulaError('#REF!')
    case 'error': return formulaError(node.code)
  }
  return formulaError('#VALUE!')
}

function evaluateReferenceOperand(
  node: FormulaNode, sheet: SpreadsheetSheet, resolve: FormulaValueResolver,
  workbook: WorkbookModel, byName: Map<string, SpreadsheetSheet>, formulaCell: Pick<FormulaCellAddress, 'row' | 'column'> | undefined, evaluationTime: Date,
  aliases?: FormulaImportAliases,
): RangeValue | FormulaErrorValue {
  if (node.kind === 'reference') {
    const target = resolveSheet(workbook, node.reference.sheetName, sheet, byName)
    return target ? checkedReference({ kind: 'rangeValue', sheet: target, reference: { kind: 'range', sheetName: node.reference.sheetName,
      start: { ...node.reference }, end: { ...node.reference } } }, resolve) : formulaError('#REF!')
  }
  const value = evaluateNode(node, sheet, resolve, workbook, byName, formulaCell, evaluationTime, aliases)
  return isFormulaError(value) ? value : isRangeValue(value) ? checkedReference(value, resolve) : formulaError('#VALUE!')
}

function evaluateFunction(
  name: string,
  args: FormulaNode[],
  sheet: SpreadsheetSheet,
  resolve: FormulaValueResolver,
  workbook: WorkbookModel,
  byName: Map<string, SpreadsheetSheet>,
  formulaCell?: Pick<FormulaCellAddress, 'row' | 'column'>,
  evaluationTime: Date = new Date(),
  aliases?: FormulaImportAliases,
): RuntimeValue {
  const evaluateArgument = (argument: FormulaNode): RuntimeValue => evaluateNode(argument, sheet, resolve, workbook, byName, formulaCell, evaluationTime, aliases)
  const evaluateArrayArgument = (argument: FormulaNode): RuntimeValue => argument.kind === 'reference' || argument.kind === 'range'
    || argument.kind === 'binary' && argument.operator === ':' || argument.kind === 'function' && ['INDEX', 'INDIRECT', 'OFFSET'].includes(argument.name)
    ? evaluateReferenceOperand(argument, sheet, resolve, workbook, byName, formulaCell, evaluationTime, aliases)
    : evaluateArgument(argument)
  if ((name === 'AND' || name === 'OR' || name === 'XOR') && (args.length < 1 || args.length > 255)) return formulaError('#N/A')
  if (name === 'NOT' && args.length !== 1) return formulaError('#N/A')
  if (name === 'INDIRECT') {
    if (args.length < 1 || args.length > 2) return formulaError('#N/A')
    const text = singleRangeValue(evaluateArgument(args[0]!), resolve)
    if (isFormulaError(text)) return text
    const mode = args.length === 2 ? singleRangeValue(evaluateArgument(args[1]!), resolve) : true
    if (isFormulaError(mode)) return mode
    const a1 = typeof mode === 'string' ? /^(?:TRUE|FALSE)$/i.test(mode) ? mode.toUpperCase() === 'TRUE'
      : Number.isFinite(Number(mode)) ? Number(mode) !== 0 : formulaError('#VALUE!') : booleanCoercion(mode)
    if (isFormulaError(a1)) return a1
    const refText = remapDynamicFormulaReference(scalarToText(text).trim(), aliases, a1)
    if (!refText || refText.startsWith('=')) return formulaError('#REF!')
    try {
      const parsed = parseDynamicFormulaReference(refText, a1, formulaCell)
      if (parsed.kind === 'error') return formulaError(parsed.code)
      const bound = bindFormulaReferences(parsed, workbook, sheet, indexFormulaDefinedNames(workbook), formulaCell)
      const result = evaluateReferenceOperand(bound, sheet, resolve, workbook, byName, formulaCell, evaluationTime, aliases)
      return isFormulaError(result) && (result.code === '#VALUE!' || result.code === '#NAME?') ? formulaError('#REF!') : result
    } catch { return formulaError('#REF!') }
  }
  if (name === 'OFFSET') {
    if (args.length < 3 || args.length > 5) return formulaError('#N/A')
    const reference = evaluateReferenceOperand(args[0]!, sheet, resolve, workbook, byName, formulaCell, evaluationTime, aliases)
    if (isFormulaError(reference)) return reference
    const numbers: number[] = []
    for (let index = 1; index < args.length; index++) {
      const value = singleRangeValue(evaluateArgument(args[index]!), resolve)
      if (isFormulaError(value)) return value
      if (value === null && index >= 3) { numbers.push(index === 3 ? Math.abs(reference.reference.end.row - reference.reference.start.row) + 1
        : Math.abs(reference.reference.end.column - reference.reference.start.column) + 1); continue }
      const number = numericTextCoercion(value, workbook, evaluationTime, false)
      if (isFormulaError(number)) return number
      if (!Number.isFinite(number)) return formulaError('#VALUE!')
      numbers.push(Math.trunc(number))
    }
    const first = reference.reference.start, last = reference.reference.end
    const height = numbers[2] ?? Math.abs(last.row - first.row) + 1, width = numbers[3] ?? Math.abs(last.column - first.column) + 1
    if (!height || !width) return formulaError('#REF!')
    const row = Math.min(first.row, last.row) + numbers[0]!, column = Math.min(first.column, last.column) + numbers[1]!
    const endRow = row + height - Math.sign(height), endColumn = column + width - Math.sign(width)
    if ([row, endRow].some(value => value < 0 || value > 1_048_575) || [column, endColumn].some(value => value < 0 || value > 16_383)) return formulaError('#REF!')
    return checkedReference({ ...reference, reference: formulaRangeReference({ startRow: Math.min(row, endRow), endRow: Math.max(row, endRow),
      startCol: Math.min(column, endColumn), endCol: Math.max(column, endColumn) }, reference.sheet.name) }, resolve)
  }
  if (name === 'INDEX') {
    if (args.length !== 2 && args.length !== 3) return formulaError('#VALUE!')
    const values = [evaluateArrayArgument(args[0]!), ...args.slice(1).map(argument => singleRangeValue(evaluateArgument(argument), resolve))]
    const reference = indexReference(values, args.length)
    return isFormulaError(reference) ? reference : checkedReference(reference, resolve)
  }
  if (name === 'CHOOSE') {
    if (args.length < 2 || args.length > 255) return formulaError('#VALUE!')
    const indexValue = numericCoercion(singleRangeValue(evaluateArgument(args[0]!), resolve))
    if (isFormulaError(indexValue)) return indexValue
    const index = Math.trunc(indexValue)
    return index < 1 || index >= args.length ? formulaError('#VALUE!') : evaluateArrayArgument(args[index]!)
  }
  if (name === 'ADDRESS') return evaluateAddress(args, evaluateArgument)
  if (name === 'ROW' || name === 'COLUMN') return evaluateReferenceIndex(name, args, evaluateArgument, formulaCell, workbook, sheet, byName)
  if (name === 'AREAS') return evaluateReferenceAreas(args, evaluateArgument)
  if (name === 'TRUE' || name === 'FALSE') return args.length === 0 ? name === 'TRUE' : formulaError('#N/A')
  if ((name === 'NOW' || name === 'TODAY') && args.length !== 0) return formulaError('#N/A')
  if (name === 'AVEDEV' || name === 'DEVSQ' || name === 'GEOMEAN' || name === 'HARMEAN' || name === 'MODE.SNGL') {
    if (args.length < 1 || args.length > 255) return formulaError('#N/A')
    const groups: Array<{ numbers: number[]; referenced: boolean }> = []
    let averageSum: number | FormulaErrorValue = 0
    const addNumbers = (left: number | FormulaErrorValue, right: number | FormulaErrorValue): number | FormulaErrorValue => {
      if (isFormulaError(left) || isFormulaError(right) || Number.isNaN(left) || Number.isNaN(right)) return formulaError('#VALUE!')
      if (!Number.isFinite(left) || !Number.isFinite(right)) return formulaError('#NUM!')
      const total = left + right
      return Number.isFinite(total) ? total : formulaError('#NUM!')
    }
    for (const argument of args) {
      const input = evaluateArrayArgument(argument)
      if (isFormulaError(input)) return input
      const referenced = isRangeValue(input)
      const values = isRangeValue(input)
        ? cellsInRange(input).sort((left, right) => left.row - right.row || left.col - right.col)
          .map(cell => formulaReferenceCellValue(input.sheet, cell.row, cell.col, resolve))
        : [input]
      const numbers: number[] = []
      for (const item of values) {
        if (isFormulaError(item)) return item
        if (item === null) continue
        if (typeof item === 'boolean') {
          if (name === 'AVEDEV' && !referenced) numbers.push(Number(item))
          continue
        }
        if (typeof item === 'string') {
          if (referenced && (name === 'AVEDEV' || name === 'MODE.SNGL')) continue
          const converted = !referenced && name !== 'DEVSQ' ? numericTextCoercion(item, workbook, evaluationTime) : Number(item)
          if (isFormulaError(converted)) return converted
          if (name === 'MODE.SNGL' || !item.trim() || name !== 'AVEDEV' && Number.isNaN(Number(item))) continue
          numbers.push(name === 'AVEDEV' ? converted : Number(item))
          continue
        }
        if (name === 'AVEDEV' || !Number.isNaN(item)) numbers.push(item)
      }
      groups.push({ numbers, referenced })
      if (name === 'AVEDEV' && (referenced || numbers.length > 0)) {
        averageSum = addNumbers(averageSum, referenced ? numbers.reduce<number | FormulaErrorValue>(addNumbers, 0) : numbers[0]!)
        if (referenced && isFormulaError(averageSum)) return averageSum
      }
    }
    const numbers = groups.flatMap(group => group.numbers)
    if (name === 'MODE.SNGL') {
      const counts = new Map<number, number>()
      let maximumCount = 1
      for (const number of numbers) {
        const count = (counts.get(number) ?? 0) + 1
        counts.set(number, count)
        maximumCount = Math.max(maximumCount, count)
      }
      if (maximumCount === 1) return formulaError('#N/A')
      return numbers.find(number => counts.get(number) === maximumCount)!
    }
    if (numbers.length === 0) return formulaError(name === 'HARMEAN' ? '#N/A' : '#NUM!')
    if (name === 'GEOMEAN' || name === 'HARMEAN') {
      if (numbers.some(number => number <= 0)) return formulaError('#NUM!')
      if (name === 'GEOMEAN') return numbers.reduce((product, number) => product * number, 1) ** (1 / numbers.length)
      return numbers.length / numbers.reduce((sum, number) => sum + 1 / number, 0)
    }
    if (name === 'DEVSQ') {
      const mean = numbers.reduce((sum, number) => sum + number, 0) / numbers.length
      return numbers.reduce((sum, number) => sum + (number - mean) ** 2, 0)
    }
    if (isFormulaError(averageSum)) return formulaError('#VALUE!')
    const mean = averageSum / numbers.length
    let deviationSum: number | FormulaErrorValue = 0
    for (const group of groups) {
      if (!group.referenced && group.numbers.length === 0) continue
      const deviations = group.numbers.map(number => {
        const difference = number - mean
        return Number.isFinite(difference) ? Math.abs(difference) : formulaError('#VALUE!')
      })
      deviationSum = addNumbers(deviationSum, group.referenced ? deviations.reduce<number | FormulaErrorValue>(addNumbers, 0) : deviations[0]!)
      if (group.referenced && isFormulaError(deviationSum)) return deviationSum
    }
    return isFormulaError(deviationSum) ? formulaError('#VALUE!') : deviationSum / numbers.length
  }
  if (name === 'PERCENTOF') {
    if (args.length !== 2) return formulaError('#N/A')
    const subsetTotal = evaluateFunction('SUM', [args[0]!], sheet, resolve, workbook, byName, formulaCell, evaluationTime)
    if (isFormulaError(subsetTotal)) return subsetTotal
    const allTotal = evaluateFunction('SUM', [args[1]!], sheet, resolve, workbook, byName, formulaCell, evaluationTime)
    if (isFormulaError(allTotal)) return allTotal
    return evaluateBinary('/', scalarValue(subsetTotal), scalarValue(allTotal))
  }
  if (name === 'NA') return formulaError('#N/A')
  if (name === 'PI') return args.length === 0 ? Math.PI : formulaError('#N/A')
  if (name === 'RAND') return args.length === 0 ? Math.random() : formulaError('#N/A')
  if (name === 'RANDBETWEEN') {
    if (args.length !== 2) return formulaError('#N/A')
    const bottomValue = singleRangeValue(evaluateArgument(args[0]!), resolve)
    if (isFormulaError(bottomValue)) return bottomValue
    const topValue = singleRangeValue(evaluateArgument(args[1]!), resolve)
    if (isFormulaError(topValue)) return topValue
    if (typeof bottomValue === 'boolean' || typeof topValue === 'boolean') return formulaError('#VALUE!')
    const bottom = Number(bottomValue), top = Number(topValue)
    if (Number.isNaN(bottom) || Number.isNaN(top)) return formulaError('#VALUE!')
    if (bottom > top) return formulaError('#NUM!')
    const lower = Math.ceil(bottom), upper = Math.floor(top)
    return Math.floor(Math.random() * (upper - lower + 1)) + lower
  }
  if (name === 'ARABIC') {
    if (args.length !== 1) return formulaError('#N/A')
    const value = singleRangeValue(evaluateArgument(args[0]!), resolve)
    if (isFormulaError(value)) return value
    if (value === null) return 0
    if (typeof value !== 'string' || value.length > 255) return formulaError('#VALUE!')
    const upperText = value.toLocaleUpperCase()
    const negative = upperText.startsWith('-')
    const text = negative ? upperText.slice(1) : upperText
    let result = 0
    for (let index = 0; index < text.length; index++) {
      const current = ROMAN_DIGIT_VALUES.get(text.charAt(index)) ?? 0
      const next = ROMAN_DIGIT_VALUES.get(text.charAt(index + 1)) ?? 0
      const nextAfter = ROMAN_DIGIT_VALUES.get(text.charAt(index + 2)) ?? 0
      const fourth = ROMAN_DIGIT_VALUES.get(text.charAt(index + 3)) ?? 0
      if (!current || nextAfter >= next && nextAfter > current
        || current === next && current === nextAfter && current === fourth || current === next / 2) return formulaError('#VALUE!')
      result += current < next ? -current : current
    }
    return negative ? -result : result
  }
  if (name === 'ROMAN') {
    if (args.length < 1 || args.length > 2) return formulaError('#N/A')
    const numberValue = singleRangeValue(evaluateArgument(args[0]!), resolve)
    const formValue = args.length === 2 ? singleRangeValue(evaluateArgument(args[1]!), resolve) : 0
    if (isFormulaError(numberValue)) return numberValue
    if (isFormulaError(formValue)) return formValue
    const number = typeof numberValue === 'string' ? numericTextCoercion(numberValue, workbook, evaluationTime) : Number(numberValue)
    if (isFormulaError(number)) return number
    const form = typeof formValue === 'string' ? numericTextCoercion(formValue, workbook, evaluationTime) : Number(formValue)
    if (isFormulaError(form)) return form
    const integer = Math.floor(number)
    const formIndex = typeof formValue === 'boolean' ? formValue ? 0 : 4 : Math.floor(form)
    if (integer < 0 || integer > 3999 || formIndex < 0 || formIndex > 4) return formulaError('#VALUE!')
    const formValues = ROMAN_FORM_VALUES[formIndex]!
    let remaining = integer
    let result = ''
    for (let index = formValues.length - 1; index >= 0 && remaining > 0; index--) {
      const value = formValues[index]!
      const count = Math.floor(remaining / value)
      result += ROMAN_NUMERAL_TEXT.get(value)!.repeat(count)
      remaining -= count * value
    }
    return result
  }
  if (name === 'BASE') {
    if (args.length < 2 || args.length > 3) return formulaError('#N/A')
    const values = args.map(argument => singleRangeValue(evaluateArgument(argument), resolve))
    const error = values.find(isFormulaError)
    if (error) return error
    const number = typeof values[0] === 'string' ? numericTextCoercion(values[0], workbook, evaluationTime) : Number(values[0])
    if (isFormulaError(number)) return number
    const radix = typeof values[1] === 'string' ? numericTextCoercion(values[1], workbook, evaluationTime) : Number(values[1])
    if (isFormulaError(radix)) return radix
    const minLengthValue = args.length === 3 ? values[2]! : 0
    const minLength = typeof minLengthValue === 'string' ? numericTextCoercion(minLengthValue, workbook, evaluationTime) : Number(minLengthValue)
    if (isFormulaError(minLength)) return minLength
    const integer = Math.floor(number), base = Math.floor(radix), length = Math.floor(minLength)
    if (integer < 0 || integer >= 2 ** 53 || base < 2 || base > 36 || length < 0) return formulaError('#NUM!')
    return integer.toString(base).padStart(length, '0').toLocaleUpperCase()
  }
  if (name === 'DECIMAL') {
    if (args.length !== 2) return formulaError('#N/A')
    const values = args.map(argument => singleRangeValue(evaluateArgument(argument), resolve))
    const error = values.find(isFormulaError)
    if (error) return error
    const radix = typeof values[1] === 'string' ? numericTextCoercion(values[1], workbook, evaluationTime) : Number(values[1])
    if (isFormulaError(radix)) return radix
    const text = values[0] === null ? '0' : String(values[0])
    const number = Number(text)
    if (text.trim() && !Number.isNaN(number) && (number < 0 || number >= 2 ** 53 || !Number.isInteger(number))) return formulaError('#NUM!')
    const lowerText = text.toLocaleLowerCase()
    if (lowerText === 'true' || lowerText === 'false') return formulaError('#NUM!')
    const base = Math.floor(radix)
    if (base < 2 || base > 36) return formulaError('#NUM!')
    if (text.replace(/\s/g, '') === '') return 0
    for (const character of text) {
      const code = character.toUpperCase().charCodeAt(0)
      if (base <= 10 && !(code >= 48 && code < 48 + base)) return formulaError('#NUM!')
      if (base > 10 && !(code >= 48 && code < 58 || code >= 65 && code < 65 + base - 10)) return formulaError('#NUM!')
    }
    const result = Number.parseInt(text, base)
    return Number.isNaN(result) ? formulaError('#NUM!') : result
  }
  if (name === 'ABS' || name === 'SQRT' || name === 'INT') {
    if (args.length !== 1) return formulaError('#N/A')
    const value = singleRangeValue(evaluateArgument(args[0]!), resolve)
    const number = typeof value === 'number' ? value : numericTextCoercion(value, workbook, evaluationTime)
    if (isFormulaError(number)) return number
    if (name === 'INT') {
      const integerDigits = Math.floor(Math.abs(number)).toString().length
      const rounded = integerDigits >= 12 ? number : Number.parseFloat(number.toPrecision(12))
      if (Math.abs(number - rounded) < 1e-10) return Math.floor(rounded)
      return Math.floor(integerDigits >= 15 ? number : Number.parseFloat(number.toPrecision(15)))
    }
    if (!Number.isFinite(number)) return formulaError('#NUM!')
    return name === 'ABS' ? Math.abs(number) : finiteResult(Math.sqrt(number))
  }
  if (name === 'MOD' || name === 'POWER') {
    if (args.length !== 2) return formulaError('#N/A')
    const numberValue = singleRangeValue(evaluateArgument(args[0]!), resolve)
    const number = typeof numberValue === 'number' ? numberValue : numericTextCoercion(numberValue, workbook, evaluationTime)
    if (isFormulaError(number)) return number
    const operandValue = singleRangeValue(evaluateArgument(args[1]!), resolve)
    const operand = typeof operandValue === 'number' ? operandValue : numericTextCoercion(operandValue, workbook, evaluationTime)
    if (isFormulaError(operand)) return operand
    if (name === 'MOD') {
      if (operand === 0) return formulaError('#DIV/0!')
      if (typeof operandValue === 'boolean') return number - Math.floor(number)
      if (!Number.isFinite(number) || !Number.isFinite(operand)
        || Math.abs(operand) * 1125900000000 <= Math.abs(number)) return formulaError('#NUM!')
      return finiteResult(number - operand * Math.floor(number / operand))
    }
    if (Number.isNaN(operand)) return formulaError('#VALUE!')
    if (!Number.isFinite(number) || !Number.isFinite(operand)) return formulaError('#NUM!')
    if (number === 0) {
      if (operand < 0) return formulaError('#DIV/0!')
      return operand === 0 ? formulaError('#NUM!') : 0
    }
    return finiteResult(number ** operand)
  }
  if (name === 'DEGREES' || name === 'RADIANS') {
    if (args.length !== 1) return formulaError('#N/A')
    const value = singleRangeValue(evaluateArgument(args[0]!), resolve)
    const angle = typeof value === 'number' ? value : numericTextCoercion(value, workbook, evaluationTime)
    if (isFormulaError(angle)) return angle
    if (!Number.isFinite(angle)) return formulaError('#VALUE!')
    return angle * (name === 'DEGREES' ? 180 / Math.PI : Math.PI / 180)
  }
  if (['ACOS', 'ACOSH', 'ACOT', 'ACOTH', 'ASIN', 'ASINH', 'ATAN', 'ATANH'].includes(name)) {
    if (args.length !== 1) return formulaError('#N/A')
    const value = singleRangeValue(evaluateArgument(args[0]!), resolve)
    const angle = typeof value === 'number' ? value : numericTextCoercion(value, workbook, evaluationTime)
    if (isFormulaError(angle)) return angle
    if (!Number.isFinite(angle)) return formulaError(name === 'ACOT' || name === 'ACOTH' ? '#VALUE!' : '#NUM!')
    if (name === 'ACOT') return Math.atan(1 / angle) + (angle < 0 ? Math.PI : 0)
    if (name === 'ACOTH') return Math.abs(angle) <= 1 ? formulaError('#NUM!') : 0.5 * Math.log((angle + 1) / (angle - 1))
    if (name === 'ACOS') return finiteResult(Math.acos(angle))
    if (name === 'ACOSH') return finiteResult(Math.acosh(angle))
    if (name === 'ASIN') return finiteResult(Math.asin(angle))
    if (name === 'ASINH') return finiteResult(Math.asinh(angle))
    if (name === 'ATAN') return finiteResult(Math.atan(angle))
    return finiteResult(Math.atanh(angle))
  }
  if (['SINH', 'COSH', 'TANH', 'COTH', 'SECH', 'CSCH'].includes(name)) {
    if (args.length !== 1) return formulaError('#N/A')
    const value = singleRangeValue(evaluateArgument(args[0]!), resolve)
    const number = typeof value === 'number' ? value : numericTextCoercion(value, workbook, evaluationTime)
    if (isFormulaError(number)) return number
    if (name === 'SECH') {
      const denominator = Math.cosh(number)
      return Number.isFinite(denominator) ? finiteResult(1 / denominator) : 0
    }
    if (name === 'CSCH') {
      if (number === 0) return formulaError('#DIV/0!')
      const denominator = Math.sinh(number)
      if (!Number.isNaN(number) && !Number.isFinite(denominator)) return 0
      return Number.isFinite(number) ? finiteResult(1 / denominator) : formulaError('#VALUE!')
    }
    if (!Number.isFinite(number)) return formulaError(name === 'COTH' ? '#VALUE!' : '#NUM!')
    if (name === 'SINH') return finiteResult(Math.sinh(number))
    if (name === 'COSH') return finiteResult(Math.cosh(number))
    if (name === 'TANH') return finiteResult(Math.tanh(number))
    return number === 0 ? formulaError('#DIV/0!') : finiteResult(1 / Math.tanh(number))
  }
  if (['EXP', 'LN', 'LOG10', 'SIGN'].includes(name)) {
    if (args.length !== 1) return formulaError('#N/A')
    const value = singleRangeValue(evaluateArgument(args[0]!), resolve)
    const number = typeof value === 'number' ? value : numericTextCoercion(value, workbook, evaluationTime)
    if (isFormulaError(number)) return number
    if (name === 'SIGN') return number > 0 ? 1 : number < 0 ? -1 : 0
    if (!Number.isFinite(number)) return formulaError('#NUM!')
    if (name === 'EXP') return finiteResult(Math.exp(number))
    if (name === 'LN') return finiteResult(Math.log(number))
    return finiteResult(Math.log10(number))
  }
  if (name === 'LOG') {
    if (args.length < 1 || args.length > 2) return formulaError('#N/A')
    const value = singleRangeValue(evaluateArgument(args[0]!), resolve)
    if (isFormulaError(value)) return value
    const baseValue = args.length === 2 ? singleRangeValue(evaluateArgument(args[1]!), resolve) : 10
    if (isFormulaError(baseValue)) return baseValue
    const number = typeof value === 'number' ? value : numericTextCoercion(value, workbook, evaluationTime)
    if (isFormulaError(number)) return number
    const base = typeof baseValue === 'number' ? baseValue : numericTextCoercion(baseValue, workbook, evaluationTime)
    if (isFormulaError(base)) return base
    if (number <= 0 || base <= 0) return formulaError('#NUM!')
    const baseLog = Math.log(base)
    return baseLog === 0 ? formulaError('#DIV/0!') : Math.log(number) / baseLog
  }
  if (name === 'ATAN2') {
    if (args.length !== 2) return formulaError('#N/A')
    const xValue = singleRangeValue(evaluateArgument(args[0]!), resolve)
    const x = typeof xValue === 'number' ? xValue : numericTextCoercion(xValue, workbook, evaluationTime)
    if (isFormulaError(x)) return x
    const yValue = singleRangeValue(evaluateArgument(args[1]!), resolve)
    const y = typeof yValue === 'number' ? yValue : numericTextCoercion(yValue, workbook, evaluationTime)
    if (isFormulaError(y)) return y
    if (typeof xValue === 'boolean') return Math.atan2(y, x)
    if (!Number.isFinite(x) || !Number.isFinite(y)) return formulaError('#NUM!')
    return x === 0 && y === 0 ? formulaError('#DIV/0!') : finiteResult(Math.atan2(y, x))
  }
  if (['EVEN', 'ODD', 'SQRTPI'].includes(name)) {
    if (args.length !== 1) return formulaError('#N/A')
    const value = singleRangeValue(evaluateArgument(args[0]!), resolve)
    const number = typeof value === 'number' ? value : numericTextCoercion(value, workbook, evaluationTime)
    if (isFormulaError(number)) return number
    if (name === 'SQRTPI') return number < 0 ? formulaError('#NUM!') : Math.sqrt(number * Math.PI)
    if (!Number.isFinite(number)) return formulaError('#VALUE!')
    const base = Math.abs(number) / (name === 'EVEN' ? 2 : 1)
    const rounded = Math.ceil(base - Number.EPSILON * Math.max(1, base))
    const magnitude = name === 'EVEN' ? rounded * 2 : rounded + (rounded % 2 === 0 ? 1 : 0)
    const result = number < 0 ? -magnitude : magnitude
    return result === 0 ? 0 : result
  }
  if (name === 'QUOTIENT') {
    if (args.length !== 2) return formulaError('#N/A')
    const numeratorValue = singleRangeValue(evaluateArgument(args[0]!), resolve)
    if (isFormulaError(numeratorValue)) return numeratorValue
    if (typeof numeratorValue === 'boolean') return formulaError('#VALUE!')
    const denominatorValue = singleRangeValue(evaluateArgument(args[1]!), resolve)
    if (isFormulaError(denominatorValue)) return denominatorValue
    if (typeof denominatorValue === 'boolean') return formulaError('#VALUE!')
    const numerator = Number(numeratorValue)
    const denominator = Number(denominatorValue)
    if (Number.isNaN(numerator) || Number.isNaN(denominator)) return formulaError('#VALUE!')
    return denominator === 0 ? formulaError('#DIV/0!') : Math.trunc(numerator / denominator)
  }
  if (name === 'FACT' || name === 'FACTDOUBLE') {
    if (args.length !== 1) return formulaError('#N/A')
    const value = singleRangeValue(evaluateArgument(args[0]!), resolve)
    const number = typeof value === 'number' ? value : numericTextCoercion(value, workbook, evaluationTime)
    if (isFormulaError(number)) return number
    const integer = Math.floor(number)
    if (integer < 0) return formulaError('#NUM!')
    return finiteResult(factorial(integer, name === 'FACTDOUBLE' ? 2 : 1))
  }
  if (name === 'COMBIN' || name === 'COMBINA') {
    if (args.length !== 2) return formulaError('#N/A')
    const numberValue = singleRangeValue(evaluateArgument(args[0]!), resolve)
    const chosenValue = singleRangeValue(evaluateArgument(args[1]!), resolve)
    const number = typeof numberValue === 'number' ? numberValue : numericTextCoercion(numberValue, workbook, evaluationTime)
    const chosen = typeof chosenValue === 'number' ? chosenValue : numericTextCoercion(chosenValue, workbook, evaluationTime)
    if (isFormulaError(number)) return number
    if (isFormulaError(chosen)) return chosen
    const integer = Math.floor(number), chosenInteger = Math.floor(chosen)
    if (integer < 0 || chosenInteger < 0 || (name === 'COMBIN' || integer === 0) && integer < chosenInteger) return formulaError('#NUM!')
    const total = name === 'COMBINA' ? integer + chosenInteger - 1 : integer
    const selected = name === 'COMBINA' ? integer - 1 : chosenInteger
    const count = Math.min(total - selected, selected)
    let result = 1
    for (let index = 1; index <= count && Number.isFinite(result); index++) {
      result *= total - index + 1
      result /= index
    }
    return finiteResult(result)
  }
  if (name === 'GCD' || name === 'LCM' || name === 'MULTINOMIAL') {
    if (args.length < 1 || args.length > 255) return formulaError('#N/A')
    let result = name === 'GCD' ? 0 : 1
    let total = 0
    let hasNumber = false
    for (const argument of args) {
      const value = evaluateArgument(argument)
      const values = isRangeValue(value)
        ? cellsInRange(value).sort((left, right) => left.row - right.row || left.col - right.col)
          .map(cell => resolve(value.sheet, cell.row, cell.col))
        : [value]
      for (const item of values) {
        if (item === null) continue
        if (typeof item === 'boolean') return formulaError('#VALUE!')
        const number = typeof item === 'number' ? item : numericTextCoercion(item, workbook, evaluationTime)
        if (isFormulaError(number)) return number
        const integer = Math.floor(number)
        if (integer < 0 || name === 'GCD' && integer >= 2 ** 53) return formulaError('#NUM!')
        hasNumber = true
        if (name === 'MULTINOMIAL') {
          total += integer
          if (total > 170) return formulaError('#NUM!')
          result *= factorial(integer)
          continue
        }
        let dividend = Math.floor(result)
        let divisor = integer
        while (divisor > 0) {
          const remainder = dividend % divisor
          dividend = divisor
          divisor = remainder
        }
        const gcd = divisor === 0 ? dividend : divisor
        if (name === 'GCD') {
          result = gcd
          continue
        }
        result = gcd === 0 ? 0 : Math.abs(result * integer) / gcd
      }
    }
    if (name === 'MULTINOMIAL') return factorial(total) / result
    if (name === 'GCD') return result
    if (!hasNumber) return 0
    return !Number.isFinite(result) || result >= 2 ** 53 ? formulaError('#VALUE!') : result
  }
  if (name === 'MROUND') {
    if (args.length !== 2) return formulaError('#N/A')
    const numberValue = singleRangeValue(evaluateArgument(args[0]!), resolve)
    if (isFormulaError(numberValue)) return numberValue
    const multipleValue = singleRangeValue(evaluateArgument(args[1]!), resolve)
    if (isFormulaError(multipleValue)) return multipleValue
    if (typeof numberValue === 'boolean' || typeof multipleValue === 'boolean') return formulaError('#VALUE!')
    const number = Number(numberValue)
    const multiple = Number(multipleValue)
    if (Number.isNaN(number) || Number.isNaN(multiple)) return formulaError('#VALUE!')
    if (multiple === 0) return 0
    if (number > 0 && multiple < 0 || number < 0 && multiple > 0) return formulaError('#NUM!')
    const base = number / multiple
    return Math.round(base + Number.EPSILON * Math.max(1, Math.abs(base))) * multiple
  }
  if (name === 'MDETERM') {
    if (args.length !== 1) return formulaError('#N/A')
    const input = formulaArgumentMatrix(evaluateArrayArgument(args[0]!), resolve)
    const matrix = Array.from({ length: input.columnCount }, () => new Array<number>(input.rowCount))
    for (let row = 0; row < input.rowCount; row++) {
      for (let column = 0; column < input.columnCount; column++) {
        const value = input.valueAt(row * input.columnCount + column)
        if (isFormulaError(value)) return value
        if (value === null || typeof value === 'boolean') return formulaError('#VALUE!')
        const number = typeof value === 'string' ? numericTextCoercion(value, workbook, evaluationTime) : value
        if (isFormulaError(number)) return number
        matrix[column]![row] = number
      }
    }
    if (input.rowCount !== input.columnCount) return formulaError('#VALUE!')
    if (matrix.length === 1) return matrix[0]![0]!
    if (matrix.length === 2) return matrix[0]![0]! * matrix[1]![1]! - matrix[0]![1]! * matrix[1]![0]!
    const luMatrix = Array.from({ length: matrix.length }, () => new Array<number>(matrix.length).fill(0))
    let determinant = 1
    for (let column = 0; column < matrix.length; column++) {
      for (let row = 0; row < column; row++) {
        let value = matrix[row]![column]!
        for (let index = 0; index < row; index++) value -= luMatrix[row]![index]! * luMatrix[index]![column]!
        luMatrix[row]![column] = value
      }
      let maxPivot = -Infinity
      let pivotRow = column
      for (let row = column; row < matrix.length; row++) {
        let value = matrix[row]![column]!
        for (let index = 0; index < column; index++) value -= luMatrix[row]![index]! * luMatrix[index]![column]!
        luMatrix[row]![column] = value
        if (Math.abs(value) > maxPivot) {
          maxPivot = Math.abs(value)
          pivotRow = row
        }
      }
      const pivot = Math.abs(luMatrix[pivotRow]![column]!)
      if (!Number.isFinite(pivot) || pivot <= 1e-11) return 0
      if (pivotRow !== column) {
        [luMatrix[column], luMatrix[pivotRow]] = [luMatrix[pivotRow]!, luMatrix[column]!]
        [matrix[column], matrix[pivotRow]] = [matrix[pivotRow]!, matrix[column]!]
        determinant = -determinant
      }
      for (let row = column + 1; row < matrix.length; row++) luMatrix[row]![column]! /= luMatrix[column]![column]!
    }
    for (let index = 0; index < matrix.length; index++) determinant *= luMatrix[index]![index]!
    return determinant === 0 ? 0 : determinant
  }
  if (name === 'SUMSQ') {
    if (args.length < 1 || args.length > 255) return formulaError('#N/A')
    let result: number | FormulaErrorValue = 0
    for (const argument of args) {
      const value = evaluateArrayArgument(argument)
      if (isFormulaError(value)) return value
      const values = isRangeValue(value)
        ? cellsInRange(value).sort((left, right) => left.row - right.row || left.col - right.col)
          .map(cell => formulaReferenceCellValue(value.sheet, cell.row, cell.col, resolve))
        : [value]
      for (const item of values) {
        if (isFormulaError(item)) return item
        if (isRangeValue(value) && (item === null || typeof item === 'string' || typeof item === 'boolean')) continue
        const number = typeof item === 'number' ? item : numericTextCoercion(item, workbook, evaluationTime)
        if (isFormulaError(number)) return number
        const squared = number ** 2
        if (isFormulaError(result) || !Number.isFinite(squared)) {
          result = formulaError('#VALUE!')
          continue
        }
        const total = result + squared
        result = Number.isFinite(total) ? total : formulaError('#NUM!')
      }
      if (isFormulaError(result)) return result
    }
    return result
  }
  if (name === 'SUMX2MY2' || name === 'SUMX2PY2' || name === 'SUMXMY2') {
    if (args.length !== 2) return formulaError('#N/A')
    const inputs = args.map(evaluateArrayArgument)
    for (const input of inputs) if (isFormulaError(input)) return input
    const matrices = inputs.map(input => formulaArgumentMatrix(input, resolve))
    const count = matrices[0]!.rowCount * matrices[0]!.columnCount
    if (count !== matrices[1]!.rowCount * matrices[1]!.columnCount) return formulaError('#N/A')
    const pairResult = (x: number, y: number) => {
      if (name === 'SUMX2MY2') return x ** 2 - y ** 2
      if (name === 'SUMX2PY2') return x ** 2 + y ** 2
      return (x - y) ** 2
    }
    if (count === 1) {
      const x = matrices[0]!.valueAt(0)
      if (isFormulaError(x)) return x
      const y = matrices[1]!.valueAt(0)
      if (isFormulaError(y)) return y
      if (x === null || y === null) return formulaError('#VALUE!')
      if (typeof x === 'boolean' || typeof y === 'boolean'
        || typeof x === 'string' && Number.isNaN(Number(x)) || typeof y === 'string' && Number.isNaN(Number(y))) return formulaError('#DIV/0!')
      return pairResult(Number(x), Number(y))
    }
    let result = 0
    let validPairs = 0
    for (let index = 0; index < count; index += 1) {
      const x = matrices[0]!.valueAt(index)
      if (x === null || typeof x === 'string' || typeof x === 'boolean') continue
      if (isFormulaError(x)) return x
      const y = matrices[1]!.valueAt(index)
      if (y === null || typeof y === 'string' || typeof y === 'boolean') continue
      if (isFormulaError(y)) return y
      result += pairResult(x, y)
      validPairs += 1
    }
    return validPairs ? result : formulaError('#DIV/0!')
  }
  if (name === 'SUMPRODUCT') {
    if (args.length < 1 || args.length > 255) return formulaError('#N/A')
    const inputs = args.map(evaluateArrayArgument)
    if (isFormulaError(inputs[0])) return inputs[0]
    const matrices = inputs.map(input => formulaArgumentMatrix(input, resolve))
    const count = matrices[0]!.rowCount * matrices[0]!.columnCount
    if (inputs.length === 1) {
      let result: number | FormulaErrorValue = 0
      for (let index = 0; index < count; index += 1) {
        const value = matrices[0]!.valueAt(index)
        if (value === null || typeof value === 'string' || typeof value === 'boolean') continue
        if (isFormulaError(value)) return value
        if (isFormulaError(result) || Number.isNaN(value)) {
          result = formulaError('#VALUE!')
          continue
        }
        const total = result + value
        result = Number.isFinite(total) ? total : formulaError('#NUM!')
      }
      return result
    }
    for (let index = 0; index < count; index += 1) {
      const value = matrices[0]!.valueAt(index)
      if (isFormulaError(value)) return value
    }
    for (let index = 1; index < inputs.length; index += 1) {
      if (isFormulaError(inputs[index])) return inputs[index]!
      if (isRangeValue(inputs[index]!) && (matrices[index]!.rowCount !== matrices[0]!.rowCount
        || matrices[index]!.columnCount !== matrices[0]!.columnCount)) return formulaError('#VALUE!')
    }
    let result = 0
    for (let index = 0; index < count; index += 1) {
      const value = matrices[0]!.valueAt(index)
      let product = typeof value === 'number' ? value : 0
      for (let variantIndex = 1; variantIndex < matrices.length; variantIndex += 1) {
        const item = matrices[variantIndex]!.valueAt(index)
        if (isFormulaError(item)) return item
        product = typeof item === 'number' ? product * item : 0
      }
      result += product
    }
    return result
  }
  if (name === 'SERIESSUM') {
    if (args.length !== 4) return formulaError('#N/A')
    const inputs = args.map(evaluateArrayArgument)
    if (inputs.some(input => input === null)) return formulaError('#N/A')
    const scalarValues: FormulaValue[] = []
    for (const input of inputs.slice(0, 3)) {
      const matrix = formulaArgumentMatrix(input, resolve)
      if (matrix.rowCount !== 1 || matrix.columnCount !== 1) return formulaError('#VALUE!')
      const value = matrix.valueAt(0)
      if (isFormulaError(value)) return value
      if (typeof value === 'boolean') return formulaError('#VALUE!')
      scalarValues.push(value)
    }
    const parameters: number[] = []
    for (const value of scalarValues) {
      const number = typeof value === 'string' ? numericTextCoercion(value, workbook, evaluationTime) : Number(value)
      if (isFormulaError(number)) return number
      parameters.push(number)
    }
    const coefficients = formulaArgumentMatrix(inputs[3]!, resolve)
    let result = 0
    for (let index = 0; index < coefficients.rowCount * coefficients.columnCount; index += 1) {
      const value = coefficients.valueAt(index)
      if (isFormulaError(value)) return value
      if (typeof value === 'boolean') return formulaError('#VALUE!')
      const number = typeof value === 'string' ? numericTextCoercion(value, workbook, evaluationTime) : Number(value)
      if (isFormulaError(number)) return number
      result += number * parameters[0]! ** (parameters[1]! + index * parameters[2]!)
    }
    return finiteResult(result)
  }
  if (['CEILING', 'CEILING.MATH', 'CEILING.PRECISE', 'ISO.CEILING', 'FLOOR', 'FLOOR.MATH', 'FLOOR.PRECISE'].includes(name)) {
    const legacy = name === 'CEILING' || name === 'FLOOR'
    const hasMode = name.endsWith('.MATH')
    if (args.length < (legacy ? 2 : 1) || args.length > (hasMode ? 3 : 2)) return formulaError('#N/A')
    const numberValue = singleRangeValue(evaluateArgument(args[0]!), resolve)
    if (isFormulaError(numberValue)) return numberValue
    const significanceValue = args.length > 1 ? singleRangeValue(evaluateArgument(args[1]!), resolve) : 1
    const modeValue = hasMode && args.length === 3 ? singleRangeValue(evaluateArgument(args[2]!), resolve) : 0
    if (name !== 'FLOOR.MATH') {
      if (isFormulaError(significanceValue)) return significanceValue
      if (isFormulaError(modeValue)) return modeValue
    }
    const number = typeof numberValue === 'number' ? numberValue : numericTextCoercion(numberValue, workbook, evaluationTime)
    if (isFormulaError(number)) return number
    const significance = typeof significanceValue === 'number' ? significanceValue : numericTextCoercion(significanceValue, workbook, evaluationTime)
    if (isFormulaError(significance)) return significance
    const mode = typeof modeValue === 'number' ? modeValue : numericTextCoercion(modeValue, workbook, evaluationTime)
    if (isFormulaError(mode)) return mode
    if (legacy && number > 0 && significance < 0) return formulaError('#NUM!')
    if (number === 0) return 0
    if (significance === 0) return name === 'FLOOR' ? formulaError('#DIV/0!') : 0
    const reverseDirection = hasMode && number < 0 && mode !== 0
    const step = legacy ? significance : Math.abs(significance)
    const base = (reverseDirection ? -number : number) / step
    const epsilon = Number.EPSILON * Math.max(1, Math.abs(base))
    const rounded = name.startsWith('FLOOR') ? Math.floor(base + epsilon) : Math.ceil(base - epsilon)
    return (reverseDirection ? -rounded : rounded) * step
  }
  if (['ROUND', 'ROUNDDOWN', 'ROUNDUP'].includes(name)) {
    if (args.length !== 2) return formulaError('#N/A')
    const numberValue = singleRangeValue(evaluateArgument(args[0]!), resolve)
    const number = typeof numberValue === 'number' ? numberValue : numericTextCoercion(numberValue, workbook, evaluationTime)
    if (isFormulaError(number)) return number
    const digitsValue = singleRangeValue(evaluateArgument(args[1]!), resolve)
    const digits = typeof digitsValue === 'number' ? digitsValue : numericTextCoercion(digitsValue, workbook, evaluationTime)
    if (isFormulaError(digits)) return digits
    const booleanDigits = typeof digitsValue === 'boolean'
    if (!booleanDigits && (!Number.isFinite(number) || !Number.isFinite(digits))) return formulaError('#NUM!')
    const base = number < 0 && !booleanDigits ? -number : number
    const factor = 10 ** Math.trunc(digits)
    const epsilon = Number.EPSILON * Math.max(1, Math.abs(base * factor))
    const scaled = base * factor + (name === 'ROUNDUP' ? -epsilon : epsilon)
    const rounded = name === 'ROUND' ? Math.round(scaled) : name === 'ROUNDUP' ? Math.ceil(scaled) : Math.floor(scaled)
    const result = (number < 0 && !booleanDigits ? -rounded : rounded) / factor
    return booleanDigits ? result : finiteResult(result)
  }
  if (name === 'TRUNC') {
    if (args.length < 1 || args.length > 2) return formulaError('#N/A')
    const numberValue = singleRangeValue(evaluateArgument(args[0]!), resolve)
    const number = typeof numberValue === 'number' ? numberValue : numericTextCoercion(numberValue, workbook, evaluationTime)
    if (isFormulaError(number)) return number
    const digitsValue = args.length === 2 ? singleRangeValue(evaluateArgument(args[1]!), resolve) : 0
    const digits = typeof digitsValue === 'number' ? digitsValue : numericTextCoercion(digitsValue, workbook, evaluationTime)
    if (isFormulaError(digits)) return digits
    const factor = 10 ** Math.trunc(digits)
    const epsilon = Number.EPSILON * Math.max(1, Math.abs(number * factor))
    return Math.trunc(number * factor + epsilon) / factor
  }
  if (name === 'ROUNDBANK') {
    if (args.length !== 2) return formulaError('#N/A')
    const numberValue = singleRangeValue(evaluateArgument(args[0]!), resolve)
    if (isFormulaError(numberValue)) return numberValue
    const digitsValue = singleRangeValue(evaluateArgument(args[1]!), resolve)
    if (isFormulaError(digitsValue)) return digitsValue
    const number = typeof numberValue === 'number' ? numberValue : numericTextCoercion(numberValue, workbook, evaluationTime)
    if (isFormulaError(number)) return number
    const digits = typeof digitsValue === 'number' ? digitsValue : numericTextCoercion(digitsValue, workbook, evaluationTime)
    if (isFormulaError(digits)) return digits
    const precision = Math.trunc(digits)
    if (precision > 16) return number
    if (precision < -16) return 0
    const factor = 10 ** precision
    const scaled = Number((number * factor).toFixed(8))
    const lower = Math.floor(scaled)
    if (Math.abs(scaled - lower - 0.5) > 1e-8) return Math.round(scaled) / factor
    return (lower % 2 === 0 ? lower : lower + 1) / factor
  }
  if (['SIN', 'COS', 'TAN', 'COT', 'SEC', 'CSC'].includes(name) && args.length !== 1) return formulaError('#N/A')
  if (name === 'SWITCH') {
    if (args.length === 0) return formulaError('#N/A')
    const expression = evaluateArgument(args[0]!)
    if (isFormulaError(expression)) return expression
    const hasDefault = args.length % 2 === 0
    const pairEnd = hasDefault ? args.length - 1 : args.length
    for (let index = 1; index < pairEnd; index += 2) {
      const switchValue = evaluateArgument(args[index]!)
      if (switchValue === null) continue
      if (isFormulaError(switchValue)) return switchValue
      const expressionText = scalarToText(scalarValue(expression)).toLocaleLowerCase()
      const switchText = scalarToText(scalarValue(switchValue)).toLocaleLowerCase()
      if (expressionText === switchText) {
        const result = evaluateArgument(args[index + 1]!)
        return result === null ? formulaError('#N/A') : result
      }
    }
    if (!hasDefault) return formulaError('#N/A')
    const fallback = evaluateArgument(args.at(-1)!)
    return fallback === null ? formulaError('#N/A') : fallback
  }
  if (name === 'IF') {
    if (args.length < 2 || args.length > 3) return formulaError('#N/A')
    const condition = singleRangeValue(evaluateArgument(args[0]!), resolve)
    if (isFormulaError(condition)) return condition
    const truth = booleanCoercion(condition)
    if (isFormulaError(truth)) return truth
    const selected = truth ? args[1] : args[2]
    return selected ? evaluateArrayArgument(selected) : false
  }
  if (name === 'IFERROR') {
    if (args.length !== 2) return formulaError('#N/A')
    const value = evaluateArrayArgument(args[0]!)
    const scalar = isRangeValue(value) && value.reference.start.row === value.reference.end.row
      && value.reference.start.column === value.reference.end.column ? singleRangeValue(value, resolve) : value
    return isFormulaError(scalar) ? evaluateArrayArgument(args[1]!) : value
  }
  if (name === 'IFNA') {
    if (args.length !== 2) return formulaError('#N/A')
    const value = evaluateArrayArgument(args[0]!)
    const scalar = isRangeValue(value) && value.reference.start.row === value.reference.end.row
      && value.reference.start.column === value.reference.end.column ? singleRangeValue(value, resolve) : value
    return isFormulaError(scalar) && scalar.code === '#N/A' ? evaluateArrayArgument(args[1]!) : value
  }
  if (name === 'IFS') {
    if (args.length < 2 || args.length > 255 || args.length % 2 !== 0) return formulaError('#N/A')
    for (let index = 0; index < args.length; index += 2) {
      const condition = evaluateArgument(args[index]!)
      if (isFormulaError(condition)) return condition
      const truth = booleanCoercion(condition)
      if (isFormulaError(truth)) return truth
      if (truth) return evaluateArgument(args[index + 1]!)
    }
    return formulaError('#N/A')
  }

  if (name === 'SHEETS') return args.length === 0 ? workbook.sheets.length : formulaError('#N/A')
  if (name === 'SHEET') {
    if (args.length > 1) return formulaError('#N/A')
    if (args.length === 0) return workbook.sheets.findIndex(candidate => candidate.id === sheet.id) + 1
    const value = evaluateArrayArgument(args[0]!)
    if (isFormulaError(value)) return value
    if (isRangeValue(value)) return workbook.sheets.findIndex(candidate => candidate.id === value.sheet.id) + 1
    const index = workbook.sheets.findIndex(candidate => candidate.name.toLocaleLowerCase() === String(value).toLocaleLowerCase())
    return index < 0 ? formulaError('#N/A') : index + 1
  }

  if (name === 'ISBETWEEN') {
    if (args.length < 3 || args.length > 5) return formulaError('#N/A')
    const values: CellScalar[] = []
    for (const argument of args) {
      const input = formulaArgumentMatrix(evaluateArrayArgument(argument), resolve)
      if (input.rowCount > 1 || input.columnCount > 1) return formulaError('#VALUE!')
      const value = input.valueAt(0)
      if (isFormulaError(value)) return value
      values.push(value)
    }
    const lowerInclusive = args.length > 3 ? values[3]! : true
    const upperInclusive = args.length > 4 ? values[4]! : true
    if (typeof lowerInclusive === 'string' || typeof upperInclusive === 'string') return formulaError('#VALUE!')

    const compareIntervalValue = (left: CellScalar, right: CellScalar, operator: '>' | '>=' | '<' | '<=') => {
      const leftValue = left === null ? typeof right === 'string' ? '' : typeof right === 'boolean' ? false : 0
        : typeof left === 'boolean' && right === null ? Number(left) : left
      const rightValue = right === null ? 0 : right
      if (typeof leftValue === 'number' && typeof rightValue === 'number') {
        const numbers = [leftValue, rightValue].map(value => {
          if (!Number.isFinite(value)) return value
          const integerDigits = Math.floor(Math.abs(value)).toString().length
          const rounded = integerDigits >= 12 ? value : Number(value.toPrecision(12))
          if (Math.abs(value - rounded) < 1e-10) return rounded
          return integerDigits >= 15 ? value : Number(value.toPrecision(15))
        })
        if (operator === '>') return numbers[0]! > numbers[1]!
        if (operator === '>=') return numbers[0]! >= numbers[1]!
        if (operator === '<') return numbers[0]! < numbers[1]!
        return numbers[0]! <= numbers[1]!
      }
      if (typeof leftValue === 'string' && typeof rightValue === 'string') {
        if (/[*?]/.test(rightValue)) {
          const fold = (text: string) => text.replace(/[\s\S]/g, character => {
            const lower = character.toLowerCase()
            return lower.length === 1 ? lower : character
          })
          const current = fold(leftValue)
          const target = fold(rightValue)
          const pattern = target.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~?[*?]/g, match => {
            if (match.startsWith('~')) return `\\${match.slice(1)}`
            return match === '*' ? '.*' : '.'
          })
          const threshold = target.replace(/~?[*?]/g, match => match.startsWith('~') ? match.slice(1) : ' ')
          if (operator === '>' || operator === '>=') return new RegExp(`^${pattern}$`, 's').test(current) || current > threshold
          return current < threshold
        }
        const texts = [leftValue, rightValue].map(value => value.toLowerCase().normalize('NFC')
          .replace(/[ßæœﬀ-ﬆ]/g, character => FORMULA_TEXT_LIGATURES.get(character)!))
        const ordering = texts[0]!.localeCompare(texts[1]!)
        if (operator === '>') return ordering > 0
        if (operator === '>=') return texts[0] === texts[1] || ordering > 0
        if (operator === '<') return ordering < 0
        return texts[0] === texts[1] || ordering < 0
      }
      if (typeof leftValue === 'boolean' && typeof rightValue === 'boolean') return compareOrdering(Number(leftValue) - Number(rightValue), operator)
      const leftRank = typeof leftValue === 'boolean' ? 2 : typeof leftValue === 'string' ? 1 : 0
      const rightRank = typeof rightValue === 'boolean' ? 2 : typeof rightValue === 'string' ? 1 : 0
      return compareOrdering(leftRank - rightRank, operator)
    }

    if (compareIntervalValue(values[1]!, values[2]!, '>')) return formulaError('#NUM!')
    if (!compareIntervalValue(values[0]!, values[1]!, Number(lowerInclusive) ? '>=' : '>')) return false
    return compareIntervalValue(values[0]!, values[2]!, Number(upperInclusive) ? '<=' : '<')
  }

  if (name === 'ISDATE') {
    if (args.length !== 1) return formulaError('#N/A')
    const input = evaluateArrayArgument(args[0]!)
    const matrix = formulaArgumentMatrix(input, resolve)
    if (matrix.rowCount > 1 || matrix.columnCount > 1) return formulaError('#VALUE!')
    const value = matrix.valueAt(0)
    if (isFormulaError(value)) return value
    if (typeof value === 'number' && isRangeValue(input)) {
      return describeSheetNumberFormat(input.sheet.cells[cellKey(input.reference.start.row, input.reference.start.column)]!.style.numberFormat)?.isDate ?? false
    }
    if (typeof value !== 'string') return false
    const options = { dateSystem: workbook.date1904 ? DateSystem.Date1904 : DateSystem.Date1900, locale: 'en', currentYear: evaluationTime.getFullYear() }
    return Boolean(parseDate(value, options) || parseTime(value, options))
  }

  if (name === 'ISEMAIL' || name === 'ISURL') {
    if (args.length !== 1) return formulaError('#N/A')
    const input = formulaArgumentMatrix(evaluateArrayArgument(args[0]!), resolve)
    if (input.rowCount > 1 || input.columnCount > 1) return formulaError('#VALUE!')
    const value = input.valueAt(0)
    if (isFormulaError(value)) return value
    if (typeof value !== 'string') return false
    const text = name === 'ISURL' ? value.trim() : value
    if (text.length > (name === 'ISURL' ? 1000 : 254)) return false
    return (name === 'ISURL' ? FORMULA_URL_PATTERN : FORMULA_EMAIL_PATTERN).test(text)
  }

  if (name === 'ISEVEN' || name === 'ISODD') {
    if (args.length !== 1) return formulaError('#N/A')
    const input = formulaArgumentMatrix(evaluateArrayArgument(args[0]!), resolve)
    if (input.rowCount > 1 || input.columnCount > 1) return formulaError('#VALUE!')
    const value = input.valueAt(0)
    if (isFormulaError(value)) return value
    if (typeof value === 'boolean') return formulaError('#VALUE!')
    const integer = Math.trunc(Number(value))
    if (Number.isNaN(integer)) return formulaError('#VALUE!')
    return name === 'ISEVEN' ? integer % 2 === 0 : integer % 2 !== 0
  }
  if (name === 'ISFORMULA') {
    if (args.length !== 1) return formulaError('#N/A')
    const input = evaluateArrayArgument(args[0]!)
    if (isFormulaError(input)) return input
    if (!isRangeValue(input)) return formulaError('#N/A')
    if (input.reference.start.row !== input.reference.end.row || input.reference.start.column !== input.reference.end.column) return formulaError('#VALUE!')
    return Boolean(input.sheet.cells[cellKey(input.reference.start.row, input.reference.start.column)]?.formula)
  }

  if (['ERROR.TYPE', 'N', 'TYPE', 'ISREF'].includes(name)) {
    if (args.length !== 1) return formulaError('#N/A')
    const input = evaluateArrayArgument(args[0]!)
    if (name === 'ISREF') return isRangeValue(input)
    const matrix = formulaArgumentMatrix(input, resolve)
    if (name !== 'N' && (matrix.rowCount > 1 || matrix.columnCount > 1)) return name === 'TYPE' ? 64 : formulaError('#VALUE!')
    const value = matrix.valueAt(0)
    if (name === 'TYPE') return isFormulaError(value) ? 16 : typeof value === 'boolean' ? 4 : typeof value === 'string' ? 2 : 1
    if (name === 'N') return isFormulaError(value) ? value : typeof value === 'string' ? 0 : Number(value)
    if (!isFormulaError(value) && typeof value !== 'string') return formulaError('#N/A')
    return ERROR_TYPE_NUMBERS.get(isFormulaError(value) ? value.code : value) ?? formulaError('#N/A')
  }

  if (['ISBLANK', 'ISERR', 'ISERROR', 'ISLOGICAL', 'ISNA', 'ISNONTEXT', 'ISNUMBER', 'ISTEXT'].includes(name)) {
    if (args.length !== 1) return formulaError('#N/A')
    const input = evaluateArrayArgument(args[0]!)
    if (isRangeValue(input) && (input.reference.start.row !== input.reference.end.row
      || input.reference.start.column !== input.reference.end.column)) return formulaError('#VALUE!')
    const value = isRangeValue(input) ? formulaReferenceCellValue(input.sheet, input.reference.start.row, input.reference.start.column, resolve) : input
    if (name === 'ISBLANK' && isRangeValue(input)) {
      const cell = input.sheet.cells[cellKey(input.reference.start.row, input.reference.start.column)]
      return !isFormulaError(value) && (!cell || (cell.value === null || cell.value === '')
        && !cell.formula && cell.richText === undefined && !cell.image)
    }
    if (name === 'ISBLANK') return value === null
    if (name === 'ISERROR') return isFormulaError(value)
    if (name === 'ISERR') return isFormulaError(value) && value.code !== '#N/A'
    if (name === 'ISNA') return (isFormulaError(value) ? value.code : value) === '#N/A'
    if (name === 'ISLOGICAL') return typeof value === 'boolean'
    if (name === 'ISNONTEXT') return typeof value !== 'string'
    if (name === 'ISNUMBER') return typeof value === 'number'
    return typeof value === 'string'
  }

  const evaluated = args.map(argument => {
    const value = evaluateArgument(argument)
    return ['LEFT', 'RIGHT', 'MID', 'WEEKDAY', 'VALUE', 'LEN', 'UPPER', 'LOWER', 'TRIM'].includes(name) ? singleRangeValue(value, resolve) : value
  })
  if (name === 'COUNTBLANK' || name === 'COUNTA') return evaluateCountValue(name, evaluated, args, sheet, resolve)
  if (['COUNTIF', 'COUNTIFS', 'SUMIF', 'SUMIFS', 'AVERAGEIF', 'AVERAGEIFS'].includes(name)) {
    return evaluateConditionalAggregate(name, evaluated, args, sheet, resolve, workbook)
  }
  if (name === 'MATCH') return evaluateMatch(evaluated, args.length, resolve)
  if (name === 'VLOOKUP') return evaluateVlookup(evaluated, args.length, resolve)
  if (name === 'HLOOKUP') return evaluateHlookup(evaluated, args.length, resolve)
  if (name === 'ROWS' || name === 'COLUMNS') return evaluateDimension(name, evaluated)
  if (name === 'PMT' || name === 'PV' || name === 'FV' || name === 'NPER' || name === 'RATE' || name === 'IPMT' || name === 'PPMT') return evaluateFinancialValue(name, evaluated)
  if (name === 'FVSCHEDULE') return evaluateFutureValueSchedule(evaluated, sheet, resolve)
  if (name === 'IRR') return evaluateInternalRateOfReturn(evaluated, sheet, resolve)
  if (name === 'DB' || name === 'DDB' || name === 'DOLLARDE' || name === 'DOLLARFR' || name === 'RRI' || name === 'SLN' || name === 'SYD') return evaluateDepreciationAndDollarValue(name, evaluated)
  if (name === 'NPV') return evaluateNetPresentValue(evaluated, sheet, resolve)
  if (name === 'MEDIAN' || name === 'LARGE' || name === 'SMALL' || name === 'VAR.S' || name === 'VAR.P' || name === 'STDEV.S' || name === 'STDEV.P'
    || name === 'PERCENTILE.INC' || name === 'PERCENTILE.EXC' || name === 'QUARTILE.INC' || name === 'QUARTILE.EXC') {
    return evaluateStatisticalValue(name, evaluated, sheet, resolve, workbook)
  }
  if (name === 'SUM' || name === 'PRODUCT' || name === 'AVERAGE' || name === 'COUNT' || name === 'COUNTA' || name === 'MIN' || name === 'MAX') {
    if (args.length < 1 || args.length > 255) return formulaError('#VALUE!')
    const values: FormulaValue[] = []
    for (let index = 0; index < evaluated.length; index += 1) {
      const value = evaluated[index]!
      if (isRangeValue(value)) {
        for (const cell of cellsInRange(value)) {
          const item = resolve(value.sheet, cell.row, cell.col)
          if (isFormulaError(item)) {
            if (name === 'COUNTA') { values.push(item); continue }
            if (name === 'COUNT') continue
            return item
          }
          if (name === 'COUNTA') { if (item !== null) values.push(item) }
          else if (name === 'COUNT') { if (typeof item === 'number') values.push(item) }
          else if (typeof item === 'number') values.push(item)
        }
      } else if (isFormulaError(value)) {
        if (name === 'COUNTA') values.push(value)
        else if (name !== 'COUNT') return value
      } else if (name === 'COUNTA') { if (value !== null) values.push(value) }
      // A single cell reference has the same aggregation rules as a 1x1 range.
      // Direct literals and scalar expressions keep their own coercion rules.
      else if (args[index]!.kind === 'reference') { if (typeof value === 'number') values.push(value) }
      else if (value === null) continue
      else {
        const number = aggregateLiteralNumber(value, workbook)
        if (isFormulaError(number)) {
          if (name === 'COUNT') continue
          return number
        }
        values.push(number)
      }
    }
    if (name === 'COUNTA') return values.length
    if (name === 'COUNT') return values.length
    const numbers = values.filter((value): value is number => typeof value === 'number')
    if (name === 'SUM') return finiteResult(numbers.reduce((sum, value) => sum + value, 0))
    if (name === 'PRODUCT') return numbers.length ? finiteResult(numbers.reduce((product, value) => product * value, 1)) : 0
    if (name === 'AVERAGE') return numbers.length ? finiteResult(numbers.reduce((sum, value) => sum + value, 0) / numbers.length) : formulaError('#DIV/0!')
    if (!numbers.length) return 0
    return numbers.reduce((selected, value) => name === 'MIN' ? Math.min(selected, value) : Math.max(selected, value), numbers[0]!)
  }
  for (const value of evaluated) if (isFormulaError(value)) return value
  if (name === 'AND' || name === 'OR' || name === 'XOR') {
    const booleans: boolean[] = []
    for (const value of evaluated) {
      for (const item of flattenScalar(value, sheet, resolve)) {
        const truth = booleanCoercion(item)
        if (isFormulaError(truth)) return truth
        booleans.push(truth)
      }
    }
    return name === 'AND' ? booleans.every(Boolean) : name === 'OR' ? booleans.some(Boolean) : booleans.filter(Boolean).length % 2 === 1
  }
  if (name === 'NOT') return invertBoolean(evaluated[0]!)
  if (name === 'DATE') {
    if (evaluated.length !== 3) return formulaError('#VALUE!')
    const parts = evaluated.map(value => numericCoercion(singleRangeValue(value, resolve)))
    const error = parts.find(isFormulaError)
    if (error) return error
    const [rawYear, rawMonth, rawDay] = parts as number[]
    if (![rawYear, rawMonth, rawDay].every(Number.isFinite)) return formulaError('#NUM!')
    let year = Math.floor(rawYear!)
    const month = Math.floor(rawMonth!), day = Math.floor(rawDay!)
    if (year < 0) return formulaError('#NUM!')
    if (year >= 0 && year < 1900) year += 1900
    if (year < 1900 || year > 9999 || Math.abs(month) > 120_000 || Math.abs(day) > 1_000_000) return formulaError('#NUM!')
    return dateArgumentsToSerial(year, month, day, workbook.date1904)
  }
  if (name === 'DATEVALUE') {
    if (evaluated.length !== 1 || typeof evaluated[0] !== 'string') return formulaError('#VALUE!')
    const parts = parseDateText(evaluated[0])
    return parts ? datePartsToSerial(parts.year, parts.month, parts.day, workbook.date1904) : formulaError('#VALUE!')
  }
  if (name === 'TIMEVALUE') {
    if (evaluated.length !== 1) return formulaError('#VALUE!')
    const value = singleRangeValue(evaluated[0]!, resolve)
    if (isFormulaError(value)) return value
    if (typeof value !== 'string') return formulaError('#VALUE!')
    return parseTimeValue(value) ?? formulaError('#VALUE!')
  }
  if (name === 'TO_DATE') {
    if (evaluated.length !== 1) return formulaError('#VALUE!')
    return singleRangeValue(evaluated[0]!, resolve)
  }
  if (name === 'NOW' || name === 'TODAY') {
    const year = evaluationTime.getFullYear()
    const month = evaluationTime.getMonth() + 1
    const day = evaluationTime.getDate()
    const serial = datePartsToSerial(year, month, day, workbook.date1904)
    if (isFormulaError(serial) || typeof serial !== 'number') return formulaError('#NUM!')
    if (name === 'TODAY') return serial
    const seconds = evaluationTime.getHours() * 3600 + evaluationTime.getMinutes() * 60 + evaluationTime.getSeconds()
    return serial + seconds / 86_400
  }
  if (name === 'EPOCHTODATE') {
    if (evaluated.length < 1 || evaluated.length > 2) return formulaError('#VALUE!')
    const timestamp = singleRangeValue(evaluated[0]!, resolve)
    if (isFormulaError(timestamp)) return timestamp
    if (typeof timestamp !== 'number' || !Number.isFinite(timestamp)) return formulaError('#VALUE!')
    const unit = evaluated.length === 2 ? numericCoercion(evaluated[1]!) : 1
    if (isFormulaError(unit)) return unit
    const unitCode = Math.floor(unit)
    if (timestamp < 0 || unitCode < 1 || unitCode > 3) return formulaError('#NUM!')
    const milliseconds = Math.trunc(timestamp * (unitCode === 1 ? 1000 : unitCode === 2 ? 1 : 0.001))
    const epochSerial = workbook.date1904 ? 24_107 : 25_569
    return finiteResult(epochSerial + milliseconds / MS_PER_DAY)
  }
  if (name === 'YEAR' || name === 'MONTH' || name === 'DAY' || name === 'WEEKDAY') {
    if ((name !== 'WEEKDAY' && evaluated.length !== 1) || (name === 'WEEKDAY' && (evaluated.length < 1 || evaluated.length > 2))) return formulaError('#VALUE!')
    const serial = name === 'WEEKDAY' ? numericTextCoercion(evaluated[0]!, workbook, evaluationTime, false) : numericCoercion(evaluated[0]!)
    if (isFormulaError(serial)) return serial
    if (name === 'WEEKDAY' && serial > getDateSystemMaxSerial(workbook.date1904 ? DateSystem.Date1904 : DateSystem.Date1900)) return formulaError('#NUM!')
    const parts = serialToDateParts(serial, workbook.date1904)
    if (!parts) return formulaError('#NUM!')
    if (name === 'YEAR') return parts.year
    if (name === 'MONTH') return parts.month
    if (name === 'DAY') return parts.day
    const returnTypeValue = evaluated.length > 1 ? numericTextCoercion(evaluated[1]!, workbook, evaluationTime) : 1
    if (isFormulaError(returnTypeValue)) return returnTypeValue
    const returnType = Math.floor(returnTypeValue)
    if (![1, 2, 3, 11, 12, 13, 14, 15, 16, 17].includes(returnType)) return formulaError('#NUM!')
    const weekday = serialWeekday(serial, workbook.date1904)
    if (returnType === 1 || returnType === 17) return weekday + 1
    if (returnType === 3) return (weekday + 6) % 7
    const firstWeekday = returnType === 2 ? 1 : returnType - 10
    return (weekday - firstWeekday + 7) % 7 + 1
  }
  if (name === 'ISOWEEKNUM') {
    if (evaluated.length !== 1) return formulaError('#VALUE!')
    const serial = numericCoercion(evaluated[0]!)
    if (isFormulaError(serial)) return serial
    const parts = serialToDateParts(serial, workbook.date1904)
    return parts ? isoWeekNumber(parts) : formulaError('#NUM!')
  }
  if (name === 'WEEKNUM') {
    if (evaluated.length < 1 || evaluated.length > 2) return formulaError('#VALUE!')
    const serial = formulaDateSerial(evaluated[0]!, workbook)
    if (isFormulaError(serial)) return serial
    const returnTypeValue = evaluated.length === 2 ? numericCoercion(evaluated[1]!) : 1
    if (isFormulaError(returnTypeValue)) return returnTypeValue
    const returnType = Math.floor(returnTypeValue)
    const parts = serialToDateParts(serial, workbook.date1904)
    if (!parts) return formulaError('#NUM!')
    if (returnType === 21) return isoWeekNumber(parts)
    const firstWeekday = new Map([[1, 0], [2, 1], [11, 1], [12, 2], [13, 3], [14, 4], [15, 5], [16, 6], [17, 0]])
    const startDay = firstWeekday.get(returnType)
    if (startDay === undefined) return formulaError('#NUM!')
    const yearStart = datePartsToSerial(parts.year, 1, 1, workbook.date1904)
    if (isFormulaError(yearStart)) return yearStart
    if (typeof yearStart !== 'number') return formulaError('#NUM!')
    const weekday = serialWeekday(yearStart, workbook.date1904)
    const weekStart = yearStart - ((weekday - startDay + 7) % 7)
    return Math.floor((Math.floor(serial) - weekStart) / 7) + 1
  }
  if (name === 'YEARFRAC') {
    if (evaluated.length < 2 || evaluated.length > 3) return formulaError('#VALUE!')
    const start = formulaDateSerial(evaluated[0]!, workbook), end = formulaDateSerial(evaluated[1]!, workbook)
    if (isFormulaError(start)) return start
    if (isFormulaError(end)) return end
    const basisValue = evaluated.length === 3 ? numericCoercion(evaluated[2]!) : 0
    if (isFormulaError(basisValue)) return basisValue
    const basis = Math.floor(basisValue)
    if (basis < 0 || basis > 4) return formulaError('#NUM!')
    return finiteResult(yearFraction(start, end, basis, workbook))
  }
  if (name === 'DATEDIF') {
    if (evaluated.length !== 3) return formulaError('#VALUE!')
    const start = formulaDateSerial(evaluated[0]!, workbook)
    const end = formulaDateSerial(evaluated[1]!, workbook)
    if (isFormulaError(start)) return start
    if (isFormulaError(end)) return end
    if (end < start) return formulaError('#NUM!')
    if (typeof evaluated[2] !== 'string') return formulaError('#NUM!')
    const unit = evaluated[2].toLocaleUpperCase()
    const startParts = serialToDateParts(start, workbook.date1904)
    const endParts = serialToDateParts(end, workbook.date1904)
    if (!startParts || !endParts) return formulaError('#NUM!')
    const startYear = !workbook.date1904 && start === 0 ? 1900 : startParts.year
    const startMonth = !workbook.date1904 && start === 0 ? 1 : startParts.month
    const startDay = !workbook.date1904 && start === 0 ? 0 : startParts.day
    if (unit === 'D') return Math.floor(end) - Math.floor(start)
    if (unit === 'Y') {
      return endParts.year - startYear - Number(endParts.month < startMonth || (endParts.month === startMonth && endParts.day < startDay))
    }
    if (unit === 'M') {
      return (endParts.year - startYear) * 12 + endParts.month - startMonth - Number(endParts.day < startDay)
    }
    if (unit === 'YM') {
      let months = endParts.month - startMonth
      if (endParts.month < startMonth || (endParts.month === startMonth && endParts.day < startDay)) months += 12
      if (endParts.day < startDay) months -= 1
      return months
    }
    if (unit === 'MD') {
      let days = endParts.day - startDay
      if (endParts.day < startDay) days += new Date(Date.UTC(endParts.year, endParts.month - 1, 0)).getUTCDate()
      return days
    }
    if (unit === 'YD') {
      const targetYear = endParts.month < startMonth || (endParts.month === startMonth && endParts.day < startDay)
        ? startYear + 1
        : startYear
      const target = datePartsToSerial(targetYear, endParts.month, endParts.day, workbook.date1904)
      if (isFormulaError(target) || typeof target !== 'number') return formulaError('#NUM!')
      return Math.floor(target) - Math.floor(start)
    }
    return formulaError('#NUM!')
  }
  if (name === 'NETWORKDAYS' || name === 'NETWORKDAYS.INTL') {
    const international = name.endsWith('.INTL')
    if ((!international && (evaluated.length < 2 || evaluated.length > 3)) || (international && (evaluated.length < 2 || evaluated.length > 4))) return formulaError('#VALUE!')
    const start = formulaDateSerial(evaluated[0]!, workbook), end = formulaDateSerial(evaluated[1]!, workbook)
    if (isFormulaError(start)) return start
    if (isFormulaError(end)) return end
    const weekend = international && evaluated.length >= 3 ? weekendPattern(evaluated[2]!) : [0, 6]
    if (isFormulaError(weekend)) return weekend
    const holidays = readHolidays(evaluated[international ? 3 : 2], workbook, resolve)
    if (isFormulaError(holidays)) return holidays
    return countWeekdays(start, end, weekend, holidays, workbook.date1904)
  }
  if (name === 'WORKDAY' || name === 'WORKDAY.INTL') {
    const international = name.endsWith('.INTL')
    if ((!international && (evaluated.length < 2 || evaluated.length > 3)) || (international && (evaluated.length < 2 || evaluated.length > 4))) return formulaError('#VALUE!')
    const start = formulaDateSerial(evaluated[0]!, workbook), offset = numericCoercion(evaluated[1]!)
    if (isFormulaError(start)) return start
    if (isFormulaError(offset)) return offset
    if (Math.abs(offset) > 1_000_000) return formulaError('#NUM!')
    const weekend = international && evaluated.length >= 3 ? weekendPattern(evaluated[2]!) : [0, 6]
    if (isFormulaError(weekend)) return weekend
    if (weekend.length === 7) return formulaError('#NUM!')
    const holidays = readHolidays(evaluated[international ? 3 : 2], workbook, resolve)
    if (isFormulaError(holidays)) return holidays
    const target = addWeekdays(start, Math.trunc(offset), weekend, holidays, workbook.date1904)
    if (isFormulaError(target)) return target
    const parts = serialToDateParts(target, workbook.date1904)
    return parts && parts.year <= 9999 ? target : formulaError('#NUM!')
  }
  if (name === 'EDATE' || name === 'EOMONTH') {
    if (evaluated.length !== 2) return formulaError('#VALUE!')
    const serial = numericCoercion(evaluated[0]!), months = numericCoercion(evaluated[1]!)
    if (isFormulaError(serial)) return serial
    if (isFormulaError(months)) return months
    const parts = serialToDateParts(serial, workbook.date1904)
    if (!parts || !Number.isFinite(months)) return formulaError('#NUM!')
    const targetMonth = parts.month - 1 + Math.trunc(months)
    const targetYear = parts.year + Math.floor(targetMonth / 12)
    const monthIndex = ((targetMonth % 12) + 12) % 12
    const day = name === 'EOMONTH' ? new Date(Date.UTC(targetYear, monthIndex + 1, 0)).getUTCDate() : Math.min(parts.day, new Date(Date.UTC(targetYear, monthIndex + 1, 0)).getUTCDate())
    return datePartsToSerial(targetYear, monthIndex + 1, day, workbook.date1904)
  }
  if (name === 'DAYS360') {
    if (evaluated.length < 2 || evaluated.length > 3) return formulaError('#VALUE!')
    const start = numericCoercion(evaluated[0]!), end = numericCoercion(evaluated[1]!)
    if (isFormulaError(start)) return start
    if (isFormulaError(end)) return end
    const method = evaluated.length === 3 ? booleanCoercion(evaluated[2]!) : false
    if (isFormulaError(method)) return method
    const first = serialToDateParts(start, workbook.date1904), last = serialToDateParts(end, workbook.date1904)
    if (!first || !last) return formulaError('#NUM!')
    const days = days360Fraction(first, last, end >= start, !method)
    return end >= start ? days : -days
  }
  if (name === 'DAYS') {
    if (evaluated.length !== 2) return formulaError('#VALUE!')
    const end = numericCoercion(evaluated[0]!), start = numericCoercion(evaluated[1]!)
    if (isFormulaError(end)) return end
    if (isFormulaError(start)) return start
    return end - start
  }
  if (name === 'TIME') {
    if (evaluated.length !== 3) return formulaError('#VALUE!')
    const parts = evaluated.map(numericCoercion)
    const error = parts.find(isFormulaError)
    if (error) return error
    const [hours, minutes, seconds] = parts as number[]
    if (![hours, minutes, seconds].every(Number.isInteger)) return formulaError('#VALUE!')
    if ([hours, minutes, seconds].some((part) => part < 0 || part >= 32_768)) return formulaError('#NUM!')
    return ((hours * 3600 + minutes * 60 + seconds) % 86_400) / 86_400
  }
  if (name === 'HOUR' || name === 'MINUTE' || name === 'SECOND') {
    if (evaluated.length !== 1) return formulaError('#VALUE!')
    const serial = numericCoercion(evaluated[0]!)
    if (isFormulaError(serial)) return serial
    const fraction = ((serial % 1) + 1) % 1
    const totalSeconds = Math.floor(fraction * 86_400 + 1e-8)
    return name === 'HOUR' ? Math.floor(totalSeconds / 3600) : name === 'MINUTE' ? Math.floor(totalSeconds % 3600 / 60) : totalSeconds % 60
  }
  if (['SIN', 'COS', 'TAN', 'COT', 'SEC', 'CSC'].includes(name)) {
    const number = numericTextCoercion(singleRangeValue(evaluated[0]!, resolve), workbook, evaluationTime)
    if (isFormulaError(number)) return number
    if (name === 'SIN') return finiteResult(Math.sin(number))
    if (name === 'COS') return finiteResult(Math.cos(number))
    if (name === 'TAN') return finiteResult(Math.tan(number))
    if (Math.abs(number) >= 2 ** 27) return formulaError('#NUM!')
    const denominator = name === 'COT' ? Math.tan(number) : name === 'SEC' ? Math.cos(number) : Math.sin(number)
    return denominator === 0 ? formulaError('#DIV/0!') : finiteResult(1 / denominator)
  }
  if (name === 'CHAR' || name === 'UNICHAR' || name === 'CODE' || name === 'UNICODE' || name === 'EXACT' || name === 'CLEAN' || name === 'PROPER' || name === 'NUMBERVALUE' || name === 'TEXTAFTER' || name === 'TEXTBEFORE') {
    if (args.length < 1 || args.length > (name === 'NUMBERVALUE' ? 3 : name === 'TEXTAFTER' || name === 'TEXTBEFORE' ? 6 : name === 'EXACT' ? 2 : 1)) return formulaError('#VALUE!')
    const values = evaluated.map(scalarValue)
    const error = values.find(isFormulaError)
    if (error) return error
    const asText = (index: number) => scalarToText(values[index]!)
    if (name === 'CHAR' || name === 'UNICHAR') {
      const code = numericCoercion(values[0]!)
      if (isFormulaError(code)) return code
      const codePoint = Math.floor(code)
      if (codePoint < 1 || codePoint > 0x10ffff) return formulaError('#VALUE!')
      if ((codePoint >= 0xd800 && codePoint <= 0xdfff) || (codePoint >= 0xfdd0 && codePoint <= 0xfdef) || (codePoint & 0xffff) >= 0xfffe) return formulaError('#N/A')
      return String.fromCodePoint(codePoint)
    }
    if (name === 'CODE' || name === 'UNICODE') {
      const chars = Array.from(asText(0))
      if (!chars.length) return formulaError('#VALUE!')
      return chars[0]!.codePointAt(0)!
    }
    if (name === 'EXACT') return asText(0) === asText(1)
    if (name === 'CLEAN') return asText(0).replace(/[\x00-\x1f\x7f-\x9f]/g, '')
    if (name === 'PROPER') return asText(0).toLocaleLowerCase().replace(/(^|[^\p{L}\p{N}])([\p{L}\p{N}])/gu, (_match, prefix: string, initial: string) => prefix + initial.toLocaleUpperCase())
    if (name === 'NUMBERVALUE') {
      let text = asText(0).replace(/\s+/g, '')
      const decimal = args.length >= 2 ? asText(1) : '.'
      const group = args.length >= 3 ? asText(2) : ','
      if (decimal.length > 1 || group.length > 1 || (decimal && decimal === group)) return formulaError('#VALUE!')
      let percentCount = 0
      while (text.endsWith('%')) { percentCount++; text = text.slice(0, -1) }
      const exponentIndex = text.search(/[eE]/)
      const mantissa = exponentIndex < 0 ? text : text.slice(0, exponentIndex)
      const exponent = exponentIndex < 0 ? '' : text.slice(exponentIndex)
      const decimalParts = decimal ? mantissa.split(decimal) : [mantissa]
      if (decimalParts.length > 2) return formulaError('#VALUE!')
      let integerPart = decimalParts[0]!
      if (group && integerPart.includes(group)) {
        const groups = integerPart.split(group)
        if (!/^[+-]?\d{1,3}$/.test(groups[0]!) || groups.slice(1).some(part => !/^\d{3}$/.test(part))) return formulaError('#VALUE!')
        integerPart = groups.join('')
      }
      text = integerPart + (decimalParts.length === 2 ? `.${decimalParts[1]}` : '') + exponent
      if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(text)) return formulaError('#VALUE!')
      const number = Number(text) / (100 ** percentCount)
      return Number.isFinite(number) ? number : formulaError('#VALUE!')
    }
    if (args.length < 2) return formulaError('#VALUE!')
    const source = asText(0), delimiter = asText(1)
    const instance = args.length >= 3 ? numericCoercion(values[2]!) : 1
    const matchMode = args.length >= 4 ? numericCoercion(values[3]!) : 0
    const matchEnd = args.length >= 5 ? numericCoercion(values[4]!) : 0
    const optionError = [instance, matchMode, matchEnd].find(isFormulaError)
    if (optionError) return optionError
    const instanceNumber = Math.floor(instance as number)
    const matchModeNumber = Math.floor(matchMode as number)
    const matchEndNumber = Math.floor(matchEnd as number)
    if (instanceNumber === 0 || ![0, 1].includes(matchModeNumber) || ![0, 1].includes(matchEndNumber)) return formulaError('#VALUE!')
    if (!delimiter) return name === 'TEXTAFTER' ? (instanceNumber > 0 ? source : '') : (instanceNumber > 0 ? '' : source)
    const haystack = matchModeNumber ? source.toLocaleLowerCase() : source
    const needle = matchModeNumber ? delimiter.toLocaleLowerCase() : delimiter
    let found = -1
    let foundCount = 0
    let cursor = instanceNumber > 0 ? 0 : haystack.length
    for (let count = 0; count < Math.abs(instanceNumber); count++) {
      found = instanceNumber > 0 ? haystack.indexOf(needle, cursor) : haystack.lastIndexOf(needle, cursor - needle.length)
      if (found < 0) break
      foundCount++
      cursor = instanceNumber > 0 ? found + needle.length : found
    }
    if (foundCount < Math.abs(instanceNumber)) {
      if (matchEndNumber) return name === 'TEXTAFTER' ? (instanceNumber > 0 ? '' : source) : (instanceNumber > 0 ? source : '')
      if (args.length >= 6) return values[5]!
      return formulaError('#N/A')
    }
    return name === 'TEXTAFTER' ? source.slice(found + needle.length) : source.slice(0, found)
  }
  if (name === 'LEN' || name === 'UPPER' || name === 'LOWER' || name === 'TRIM' || name === 'VALUE' || name === 'LEFT' || name === 'RIGHT' || name === 'MID' || name === 'CONCAT' || name === 'CONCATENATE' || name === 'REPT' || name === 'FIND' || name === 'SEARCH' || name === 'REPLACE' || name === 'SUBSTITUTE' || name === 'TEXTJOIN') {
    if (name === 'CONCAT' || name === 'CONCATENATE') return flattenText(evaluated, sheet, resolve)
    if (name === 'TEXTJOIN') {
      if (args.length < 3 || args.length > 255) return formulaError('#VALUE!')
      const delimiter = scalarValue(evaluated[0]!)
      const ignoreEmpty = booleanCoercion(evaluated[1]!)
      if (isFormulaError(delimiter)) return delimiter
      if (isFormulaError(ignoreEmpty)) return ignoreEmpty
      const segments: string[] = []
      for (const item of evaluated.slice(2)) for (const scalar of flattenScalar(item!, sheet, resolve)) {
        if (isFormulaError(scalar)) return scalar
        if (ignoreEmpty && (scalar === null || scalar === '')) continue
        segments.push(scalarToText(scalar))
      }
      return segments.join(scalarToText(delimiter))
    }
    if (name === 'REPLACE' || name === 'SUBSTITUTE') {
      if ((name === 'REPLACE' && args.length !== 4) || (name === 'SUBSTITUTE' && (args.length < 3 || args.length > 4))) return formulaError('#VALUE!')
      const source = scalarToText(scalarValue(evaluated[0]!))
      const oldText = scalarToText(scalarValue(evaluated[name === 'REPLACE' ? 2 : 1]!))
      const replacement = scalarToText(scalarValue(evaluated[name === 'REPLACE' ? 3 : 2]!))
      if (name === 'REPLACE') {
        const start = numericCoercion(evaluated[1]!), count = numericCoercion(evaluated[2]!)
        if (isFormulaError(start)) return start
        if (isFormulaError(count)) return count
        if (!Number.isInteger(start) || !Number.isInteger(count) || start < 1 || count < 0) return formulaError('#VALUE!')
        const chars = Array.from(source)
        return [...chars.slice(0, start - 1), replacement, ...chars.slice(start - 1 + count)].join('')
      }
      const instance = args.length === 4 ? numericCoercion(evaluated[3]!) : null
      if (isFormulaError(instance)) return instance
      if (instance !== null && (!Number.isInteger(instance) || instance < 1)) return formulaError('#VALUE!')
      if (!oldText) return source
      if (instance === null) return source.split(oldText).join(replacement)
      let occurrence = 0
      return source.replaceAll(oldText, (match) => ++occurrence === instance ? replacement : match)
    }
    if (name === 'REPT') {
      if (args.length !== 2) return formulaError('#VALUE!')
      const text = scalarToText(scalarValue(evaluated[0]!)), count = numericCoercion(evaluated[1]!)
      if (isFormulaError(count)) return count
      if (!Number.isInteger(count) || count < 0 || text.length * count > 32767) return formulaError('#VALUE!')
      return text.repeat(count)
    }
    if (name === 'FIND' || name === 'SEARCH') {
      if (args.length < 2 || args.length > 3) return formulaError('#VALUE!')
      const needle = scalarToText(scalarValue(evaluated[0]!)), haystack = scalarToText(scalarValue(evaluated[1]!))
      const start = args.length === 3 ? numericCoercion(evaluated[2]!) : 1
      if (isFormulaError(start)) return start
      if (!Number.isInteger(start) || start < 1 || start > Array.from(haystack).length + 1) return formulaError('#VALUE!')
      const source = Array.from(haystack).slice(start - 1).join('')
      const offset = name === 'SEARCH' ? source.toLocaleLowerCase().indexOf(needle.toLocaleLowerCase()) : source.indexOf(needle)
      return offset < 0 ? formulaError('#VALUE!') : start + Array.from(source.slice(0, offset)).length
    }
    if ((name === 'LEFT' || name === 'RIGHT') && (args.length < 1 || args.length > 2)) return formulaError('#VALUE!')
    if (name === 'MID' && args.length !== 3) return formulaError('#VALUE!')
    if (['LEN', 'UPPER', 'LOWER', 'TRIM', 'VALUE'].includes(name) && args.length !== 1) return formulaError('#VALUE!')
    if (name === 'VALUE') return typeof evaluated[0] === 'boolean' ? formulaError('#VALUE!') : numericTextCoercion(evaluated[0]!, workbook, evaluationTime)
    const text = scalarToText(scalarValue(evaluated[0]!))
    if (name === 'LEN') return text.length
    if (name === 'UPPER') return text.toLocaleUpperCase()
    if (name === 'LOWER') return text.toLocaleLowerCase()
    if (name === 'TRIM') return text.replace(/^[ \u3000]+|[ \u3000]+$/g, '').replace(/[ \u3000]{2,}/g, run => run.charAt(0))
    const lengthValue = args.length > 1 ? numericTextCoercion(evaluated[1]!, workbook, evaluationTime) : 1
    if (isFormulaError(lengthValue)) return lengthValue
    const lengthArg = Math.floor(lengthValue)
    if (lengthArg < 0) return formulaError('#VALUE!')
    if (name === 'LEFT') return text.slice(0, lengthArg)
    if (name === 'RIGHT') return lengthArg === 0 ? '' : text.slice(-lengthArg)
    const countValue = numericTextCoercion(evaluated[2]!, workbook, evaluationTime)
    if (isFormulaError(countValue)) return countValue
    const count = Math.floor(countValue)
    if (lengthArg < 1 || count < 0) return formulaError('#VALUE!')
    return text.slice(lengthArg - 1, lengthArg - 1 + count)
  }
  return formulaError('#NAME?')
}

function indexReference(values: RuntimeValue[], argumentCount: number): RangeValue | FormulaErrorValue {
  if (argumentCount !== 2 && argumentCount !== 3) return formulaError('#VALUE!')
  const source = values[0]!
  if (isFormulaError(source)) return source
  if (!isRangeValue(source)) return formulaError('#VALUE!')
  const rowNumber = numericCoercion(values[1]!)
  const columnNumber = argumentCount === 3 ? numericCoercion(values[2]!) : 1
  if (isFormulaError(rowNumber)) return rowNumber
  if (isFormulaError(columnNumber)) return columnNumber
  if (!Number.isInteger(rowNumber) || !Number.isInteger(columnNumber) || rowNumber < 0 || columnNumber < 0 || (rowNumber === 0 && columnNumber === 0)) return formulaError('#VALUE!')
  const minRow = Math.min(source.reference.start.row, source.reference.end.row)
  const maxRow = Math.max(source.reference.start.row, source.reference.end.row)
  const minColumn = Math.min(source.reference.start.column, source.reference.end.column)
  const maxColumn = Math.max(source.reference.start.column, source.reference.end.column)
  const height = maxRow - minRow + 1
  const width = maxColumn - minColumn + 1
  if (rowNumber > height || columnNumber > width) return formulaError('#REF!')
  return {
    ...source,
    reference: {
      ...source.reference,
      start: { ...source.reference.start, row: rowNumber ? minRow + rowNumber - 1 : minRow, column: columnNumber ? minColumn + columnNumber - 1 : minColumn },
      end: { ...source.reference.end, row: rowNumber ? minRow + rowNumber - 1 : maxRow, column: columnNumber ? minColumn + columnNumber - 1 : maxColumn },
    },
  }
}


function evaluateMatch(values: RuntimeValue[], argumentCount: number, resolve: (sheet: SpreadsheetSheet, row: number, column: number) => FormulaValue): FormulaValue {
  if (argumentCount < 2 || argumentCount > 3 || !isRangeValue(values[1]!)) return formulaError('#VALUE!')
  const lookup = scalarValue(values[0]!)
  const matchType = argumentCount === 3 ? numericCoercion(values[2]!) : 1
  if (isFormulaError(lookup)) return lookup
  if (isFormulaError(matchType)) return matchType
  if (![-1, 0, 1].includes(matchType)) return formulaError('#VALUE!')
  const range = values[1] as RangeValue
  const minRow = Math.min(range.reference.start.row, range.reference.end.row)
  const maxRow = Math.max(range.reference.start.row, range.reference.end.row)
  const minColumn = Math.min(range.reference.start.column, range.reference.end.column)
  const maxColumn = Math.max(range.reference.start.column, range.reference.end.column)
  if (maxRow > minRow && maxColumn > minColumn) return formulaError('#N/A')
  const length = Math.max(maxRow - minRow + 1, maxColumn - minColumn + 1)
  let best = -1
  for (let index = 0; index < length; index += 1) {
    const row = maxRow === minRow ? minRow : minRow + index
    const column = maxColumn === minColumn ? minColumn : minColumn + index
    const cell = range.sheet.cells[cellKey(row, column)]
    const value = cell ? resolve(range.sheet, row, column) : null
    if (isFormulaError(value)) { if (matchType === 0) return value; continue }
    const ordering = compare(value, lookup, '=') ? 0 : compare(value, lookup, '<') ? -1 : 1
    if (matchType === 0 && ordering === 0) return index + 1
    if (matchType === 1 && ordering <= 0) best = index
    if (matchType === -1 && ordering >= 0 && best < 0) best = index
  }
  return best >= 0 ? best + 1 : formulaError('#N/A')
}

function evaluateVlookup(values: RuntimeValue[], argumentCount: number, resolve: (sheet: SpreadsheetSheet, row: number, column: number) => FormulaValue): FormulaValue {
  if ((argumentCount !== 3 && argumentCount !== 4) || !isRangeValue(values[1]!)) return formulaError('#VALUE!')
  const lookup = scalarValue(values[0]!)
  const columnNumber = numericCoercion(values[2]!)
  const approximate = argumentCount === 3 ? true : booleanCoercion(values[3]!)
  if (isFormulaError(lookup)) return lookup
  if (isFormulaError(columnNumber)) return columnNumber
  if (isFormulaError(approximate)) return approximate
  if (!Number.isInteger(columnNumber) || columnNumber < 1) return formulaError('#VALUE!')
  const range = values[1] as RangeValue
  const minRow = Math.min(range.reference.start.row, range.reference.end.row)
  const maxRow = Math.max(range.reference.start.row, range.reference.end.row)
  const minColumn = Math.min(range.reference.start.column, range.reference.end.column)
  const maxColumn = Math.max(range.reference.start.column, range.reference.end.column)
  const width = maxColumn - minColumn + 1
  if (columnNumber > width) return formulaError('#REF!')
  let bestRow = -1
  for (let row = minRow; row <= maxRow; row += 1) {
    const first = range.sheet.cells[cellKey(row, minColumn)]
    const value = first ? resolve(range.sheet, row, minColumn) : null
    if (isFormulaError(value)) { if (!approximate) return value; continue }
    if (compare(value, lookup, '=')) { bestRow = row; break }
    if (approximate && compare(value, lookup, '<')) bestRow = row
    else if (approximate && compare(value, lookup, '>')) break
  }
  if (bestRow < 0) return formulaError('#N/A')
  const resultCell = range.sheet.cells[cellKey(bestRow, minColumn + columnNumber - 1)]
  return resultCell ? resolve(range.sheet, bestRow, minColumn + columnNumber - 1) : null
}

function evaluateHlookup(values: RuntimeValue[], argumentCount: number, resolve: (sheet: SpreadsheetSheet, row: number, column: number) => FormulaValue): FormulaValue {
  if ((argumentCount !== 3 && argumentCount !== 4) || !isRangeValue(values[1]!)) return formulaError('#VALUE!')
  const lookup = scalarValue(values[0]!)
  const rowNumber = numericCoercion(values[2]!)
  const approximate = argumentCount === 3 ? true : booleanCoercion(values[3]!)
  if (isFormulaError(lookup)) return lookup
  if (isFormulaError(rowNumber)) return rowNumber
  if (isFormulaError(approximate)) return approximate
  if (!Number.isInteger(rowNumber) || rowNumber < 1) return formulaError('#VALUE!')
  const range = values[1] as RangeValue
  const minRow = Math.min(range.reference.start.row, range.reference.end.row)
  const maxRow = Math.max(range.reference.start.row, range.reference.end.row)
  const minColumn = Math.min(range.reference.start.column, range.reference.end.column)
  const maxColumn = Math.max(range.reference.start.column, range.reference.end.column)
  if (rowNumber > maxRow - minRow + 1) return formulaError('#REF!')
  let bestColumn = -1
  for (let column = minColumn; column <= maxColumn; column += 1) {
    const cell = range.sheet.cells[cellKey(minRow, column)]
    const value = cell ? resolve(range.sheet, minRow, column) : null
    if (isFormulaError(value)) { if (!approximate) return value; continue }
    if (compare(value, lookup, '=')) { bestColumn = column; break }
    if (approximate && compare(value, lookup, '<')) bestColumn = column
    else if (approximate && compare(value, lookup, '>')) break
  }
  if (bestColumn < 0) return formulaError('#N/A')
  const resultCell = range.sheet.cells[cellKey(minRow + rowNumber - 1, bestColumn)]
  return resultCell ? resolve(range.sheet, minRow + rowNumber - 1, bestColumn) : null
}


function evaluateAddress(args: FormulaNode[], evaluateArgument: (argument: FormulaNode) => RuntimeValue): FormulaValue {
  if (args.length < 2 || args.length > 5) return formulaError('#VALUE!')
  const values = args.map(evaluateArgument)
  for (const value of values) if (isFormulaError(value)) return value
  if (values.some(isRangeValue)) return formulaError('#VALUE!')
  const rowValue = numericCoercion(values[0]!)
  const columnValue = numericCoercion(values[1]!)
  if (isFormulaError(rowValue)) return rowValue
  if (isFormulaError(columnValue)) return columnValue
  const rowIndex = Number.parseInt(String(rowValue - 1), 10)
  const columnIndex = Number.parseInt(String(columnValue - 1), 10)
  if (rowIndex < 0 || rowIndex >= 1_048_576 || columnIndex < 0 || columnIndex >= 16_384) return formulaError('#REF!')

  const absoluteValue = args.length > 2 && values[2] !== null ? numericCoercion(values[2]!) : 1
  if (isFormulaError(absoluteValue)) return absoluteValue
  const absoluteType = Math.trunc(absoluteValue)
  if (absoluteType < 1 || absoluteType > 4) return formulaError('#VALUE!')
  const a1Value = args.length < 4 || values[3] === null ? true : values[3]
  const useA1 = typeof a1Value === 'string' ? a1Value.length === 0 : a1Value !== false && a1Value !== 0
  let reference: string
  if (useA1) {
    const columnName = columnIndexToName(columnIndex)
    const absoluteColumn = absoluteType === 1 || absoluteType === 3
    const absoluteRow = absoluteType === 1 || absoluteType === 2
    reference = `${absoluteColumn ? '$' : ''}${columnName}${absoluteRow ? '$' : ''}${rowIndex + 1}`
  } else {
    const absoluteRow = absoluteType === 1 || absoluteType === 2
    const absoluteColumn = absoluteType === 1 || absoluteType === 3
    const row = absoluteRow ? String(rowIndex + 1) : `[${rowIndex + 1}]`
    const column = absoluteColumn ? String(columnIndex + 1) : `[${columnIndex + 1}]`
    reference = `R${row}C${column}`
  }
  const sheetValue = args.length === 5 && values[4] !== null ? scalarToText(scalarValue(values[4]!)) : ''
  if (!sheetValue) return reference
  const a1SheetName = /^([A-Z]+)([1-9][0-9]*)$/i.exec(sheetValue)
  const isCellAddress = a1SheetName !== null
    && Number(a1SheetName[2]) <= 1_048_576
    && columnLabelToIndex(a1SheetName[1]!.toUpperCase()) < 16_384
  const isR1C1SheetName = /^(R(-?[0-9]+)?C(-?[0-9]+)?|C(-?[0-9]+)?|R(-?[0-9]+)?)$/i.test(sheetValue)
  const needsQuotes = !/^\p{Letter}/u.test(sheetValue)
    || isCellAddress
    || isR1C1SheetName
    || /[\s!$%^&*()+\-=\[\]{};':"\\|,.<>\/?（）]/.test(sheetValue)
  const sheet = needsQuotes ? `'${sheetValue.replaceAll("'", "''")}'` : sheetValue
  return `${sheet}!${reference}`
}

function evaluateReferenceIndex(
  name: 'ROW' | 'COLUMN',
  args: FormulaNode[],
  evaluateArgument: (argument: FormulaNode) => RuntimeValue,
  formulaCell: Pick<FormulaCellAddress, 'row' | 'column'> | undefined,
  workbook: WorkbookModel,
  currentSheet: SpreadsheetSheet,
  sheetsByName: Map<string, SpreadsheetSheet>,
): FormulaValue {
  if (args.length > 1) return formulaError('#VALUE!')
  if (args.length === 0) return (name === 'ROW' ? formulaCell?.row ?? 0 : formulaCell?.column ?? 0) + 1
  const argument = args[0]!
  if (argument.kind !== 'reference' && argument.kind !== 'range') {
    const value = evaluateArgument(argument)
    return isFormulaError(value) ? value : formulaError('#N/A')
  }
  const reference = argument.kind === 'reference' ? argument.reference : argument.reference
  const sheet = resolveSheet(workbook, reference.sheetName, currentSheet, sheetsByName)
  if (!sheet) return formulaError('#REF!')
  const startRow = Math.min(reference.kind === 'cell' ? reference.row : reference.start.row, reference.kind === 'cell' ? reference.row : reference.end.row)
  const endRow = Math.max(reference.kind === 'cell' ? reference.row : reference.start.row, reference.kind === 'cell' ? reference.row : reference.end.row)
  const startColumn = Math.min(reference.kind === 'cell' ? reference.column : reference.start.column, reference.kind === 'cell' ? reference.column : reference.end.column)
  const endColumn = Math.max(reference.kind === 'cell' ? reference.column : reference.start.column, reference.kind === 'cell' ? reference.column : reference.end.column)
  if (startRow !== endRow || startColumn !== endColumn) return formulaError('#VALUE!')
  return (name === 'ROW' ? startRow : startColumn) + 1
}

function evaluateReferenceAreas(args: FormulaNode[], evaluateArgument: (argument: FormulaNode) => RuntimeValue): FormulaValue {
  if (args.length !== 1) return formulaError('#VALUE!')
  const argument = args[0]!
  if (argument.kind === 'reference' || argument.kind === 'range') return 1
  const value = evaluateArgument(argument)
  return isFormulaError(value) ? value : formulaError('#VALUE!')
}

function columnIndexToName(column: number): string {
  let value = column + 1
  let name = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    name = String.fromCharCode(65 + remainder) + name
    value = Math.floor((value - 1) / 26)
  }
  return name
}

function columnLabelToIndex(label: string): number {
  return Array.from(label).reduce((column, letter) => column * 26 + letter.charCodeAt(0) - 64, 0) - 1
}

function evaluateDimension(name: 'ROWS' | 'COLUMNS', values: RuntimeValue[]): FormulaValue {
  if (values.length !== 1) return formulaError('#VALUE!')
  const value = values[0]
  if (!isRangeValue(value)) return formulaError('#VALUE!')
  return name === 'ROWS'
    ? Math.abs(value.reference.end.row - value.reference.start.row) + 1
    : Math.abs(value.reference.end.column - value.reference.start.column) + 1
}

function formulaDateSerial(value: RuntimeValue, workbook: WorkbookModel): number | FormulaErrorValue {
  if (isFormulaError(value)) return value
  if (isRangeValue(value) || typeof value === 'boolean') return formulaError('#VALUE!')
  if (typeof value === 'string') {
    const date = parseDateText(value)
    if (!date) return formulaError('#VALUE!')
    const serial = datePartsToSerial(date.year, date.month, date.day, workbook.date1904)
    return isFormulaError(serial) ? serial : typeof serial === 'number' ? serial : formulaError('#VALUE!')
  }
  if (value === null || !Number.isFinite(value)) return formulaError('#VALUE!')
  if (value < 0) return formulaError('#NUM!')
  const parts = serialToDateParts(value, workbook.date1904)
  return parts && parts.year <= 9999 ? value : formulaError('#NUM!')
}

function weekendPattern(value: RuntimeValue): number[] | FormulaErrorValue {
  if (isFormulaError(value)) return value
  if (isRangeValue(value)) return formulaError('#VALUE!')
  const weekend = typeof value === 'string' ? value : typeof value === 'boolean' ? Number(value) : value
  if (typeof weekend === 'string') {
    if (!/^[01]{7}$/.test(weekend)) return formulaError('#VALUE!')
    return Array.from(weekend, (bit, index) => bit === '1' ? (index + 1) % 7 : -1).filter(day => day >= 0)
  }
  if (typeof weekend !== 'number' || !Number.isInteger(weekend)) return formulaError('#NUM!')
  const patterns: Record<number, number[]> = {
    1: [0, 6], 2: [0, 1], 3: [1, 2], 4: [2, 3], 5: [3, 4], 6: [4, 5], 7: [5, 6],
    11: [0], 12: [1], 13: [2], 14: [3], 15: [4], 16: [5], 17: [6],
  }
  return patterns[weekend] ?? formulaError('#NUM!')
}

function readHolidays(value: RuntimeValue | undefined, workbook: WorkbookModel, resolve: (sheet: SpreadsheetSheet, row: number, column: number) => FormulaValue): number[] | FormulaErrorValue {
  if (value === undefined) return []
  const inputs = isRangeValue(value) ? cellsInRange(value).map(cell => resolve(value.sheet, cell.row, cell.col)) : [value]
  const holidays: number[] = []
  for (const item of inputs) {
    if (isFormulaError(item)) return item
    if (item === null || item === '') continue
    if (typeof item === 'boolean') return formulaError('#VALUE!')
    const serial = formulaDateSerial(item, workbook)
    if (isFormulaError(serial)) {
      if (typeof item === 'string') continue
      return serial
    }
    holidays.push(Math.floor(serial))
  }
  return holidays
}

function countWeekdays(start: number, end: number, weekend: number[], holidays: number[], date1904: boolean): FormulaValue {
  const first = Math.min(Math.floor(start), Math.floor(end))
  const last = Math.max(Math.floor(start), Math.floor(end))
  if (last - first > 1_000_000) return formulaError('#NUM!')
  const daysOff = new Set(weekend)
  const holidaySet = new Set(holidays)
  let count = 0
  for (let serial = first; serial <= last; serial += 1) {
    if (!daysOff.has(serialWeekday(serial, date1904)) && !holidaySet.has(serial)) count += 1
  }
  return end >= start ? count : -count
}

function addWeekdays(start: number, offset: number, weekend: number[], holidays: number[], date1904: boolean): number | FormulaErrorValue {
  let serial = Math.floor(start)
  if (offset === 0) return serial
  const direction = Math.sign(offset)
  let remaining = Math.abs(offset)
  const weekendSet = new Set(weekend), holidaySet = new Set(holidays)
  while (remaining > 0) {
    serial += direction
    if (serial < 0 || serial > 3_000_000) return formulaError('#NUM!')
    if (weekendSet.has(serialWeekday(serial, date1904)) || holidaySet.has(serial)) continue
    remaining -= 1
  }
  return serial
}

function serialWeekday(serial: number, date1904: boolean): number {
  return ((Math.floor(serial) + (date1904 ? 5 : 6)) % 7 + 7) % 7
}

function yearFraction(start: number, end: number, basis: number, workbook: WorkbookModel): number {
  const first = serialToDateParts(start, workbook.date1904)!
  const last = serialToDateParts(end, workbook.date1904)!
  const days = Math.abs(end - start)
  if (basis === 2) return days / 360
  if (basis === 3) return days / 365
  if (basis === 0 || basis === 4) return days360Fraction(first, last, end >= start, basis === 0) / 360
  if (days === 0) return 0
  if (first.year === last.year) return days / (isLeapYear(first.year) ? 366 : 365)
  if (end >= start && first.month === last.month && first.day === last.day) return last.year - first.year
  const startYear = Math.min(first.year, last.year), endYear = Math.max(first.year, last.year)
  let totalDays = 0
  for (let year = startYear; year <= endYear; year += 1) totalDays += isLeapYear(year) || (!workbook.date1904 && year === 1900) ? 366 : 365
  return days / (totalDays / (endYear - startYear + 1))
}

function days360Fraction(first: DateParts, last: DateParts, forward: boolean, nasd: boolean): number {
  let startYear = forward ? first.year : last.year
  let startMonth = forward ? first.month : last.month
  let startDay = forward ? first.day : last.day
  let endYear = forward ? last.year : first.year
  let endMonth = forward ? last.month : first.month
  let endDay = forward ? last.day : first.day
  const startLastFeb = startMonth === 2 && startDay === new Date(Date.UTC(startYear, 2, 0)).getUTCDate()
  if (nasd && startLastFeb) startDay = 30
  else if (startDay === 31) startDay = 30
  if (nasd && endMonth === 2 && endDay === new Date(Date.UTC(endYear, 2, 0)).getUTCDate() && startLastFeb) endDay = 30
  else if (endDay === 31) {
    if (nasd && startDay < 30) {
      const shifted = new Date(Date.UTC(endYear, endMonth - 1, endDay + 1))
      endYear = shifted.getUTCFullYear(); endMonth = shifted.getUTCMonth() + 1; endDay = shifted.getUTCDate()
    } else endDay = 30
  }
  return Math.abs((endYear - startYear) * 360 + (endMonth - startMonth) * 30 + endDay - startDay)
}

function isLeapYear(year: number): boolean {
  return year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0)
}

function evaluateFinancialValue(name: 'PMT' | 'PV' | 'FV' | 'NPER' | 'RATE' | 'IPMT' | 'PPMT', values: RuntimeValue[]): FormulaValue {
  if (name === 'IPMT' || name === 'PPMT') {
    if (values.length < 4 || values.length > 6) return formulaError('#VALUE!')
    const numbers = [0, 1, 2, 3, 4, 5].map(index => index < values.length ? numericCoercion(values[index]!) : index === 4 ? 0 : 0)
    const error = numbers.find(isFormulaError)
    if (error) return error
    const [rate, period, periods, principal, future, rawType] = numbers as number[]
    const type = rawType ? 1 : 0
    if (period < 1 || Math.floor(period) > Math.ceil(periods) || period - periods >= 1) return formulaError('#NUM!')
    const payment = financialPayment(rate, periods, principal, future, type)
    if (isFormulaError(payment)) return payment
    const interest = period === 1 ? (type ? 0 : -principal * rate) : (financialFutureValue(rate, period - (type ? 2 : 1), payment, principal, type ? 1 : 0) - (type ? payment : 0)) * rate
    const principalPayment = payment - interest
    return finiteResult(name === 'IPMT' ? interest : principalPayment)
  }
  if (name === 'NPER') {
    if (values.length < 3 || values.length > 5) return formulaError('#VALUE!')
    const rate = numericCoercion(values[0]!), payment = numericCoercion(values[1]!), principal = numericCoercion(values[2]!)
    if (isFormulaError(rate)) return rate
    if (isFormulaError(payment)) return payment
    if (isFormulaError(principal)) return principal
    const future = values.length > 3 ? numericCoercion(values[3]!) : 0
    const type = values.length > 4 ? numericCoercion(values[4]!) : 0
    if (isFormulaError(future)) return future
    if (isFormulaError(type)) return type
    if (rate === 0 && payment === 0) return formulaError('#DIV/0!')
    const numerator = payment * (1 + rate * (type ? 1 : 0)) - future * rate
    const denominator = principal * rate + payment * (1 + rate * (type ? 1 : 0))
    return finiteResult(rate === 0 ? -(principal + future) / payment : Math.log(numerator / denominator) / Math.log(1 + rate))
  }
  if (name === 'RATE') {
    if (values.length < 3 || values.length > 6) return formulaError('#VALUE!')
    const periods = numericCoercion(values[0]!), payment = numericCoercion(values[1]!), principal = numericCoercion(values[2]!)
    if (isFormulaError(periods)) return periods
    if (isFormulaError(payment)) return payment
    if (isFormulaError(principal)) return principal
    const future = values.length > 3 ? numericCoercion(values[3]!) : 0
    const type = values.length > 4 ? numericCoercion(values[4]!) : 0
    const guess = values.length > 5 ? numericCoercion(values[5]!) : 0.1
    if (isFormulaError(future)) return future
    if (isFormulaError(type)) return type
    if (isFormulaError(guess)) return guess
    const allNonnegative = payment >= 0 && principal >= 0 && future >= 0
    const allNonpositive = payment <= 0 && principal <= 0 && future <= 0
    if (periods <= 0 || !Number.isFinite(periods) || payment === 0 && principal === 0 || allNonnegative || allNonpositive) return formulaError('#NUM!')
    return solveRate(periods, payment, principal, future, type ? 1 : 0, guess)
  }
  if (values.length < 3 || values.length > 5) return formulaError('#VALUE!')
  const rate = numericCoercion(values[0]!), periods = numericCoercion(values[1]!), amount = numericCoercion(values[2]!)
  for (const value of [rate, periods, amount]) if (isFormulaError(value)) return value
  const future = values.length > 3 ? numericCoercion(values[3]!) : 0
  const type = values.length > 4 ? numericCoercion(values[4]!) : 0
  if (isFormulaError(future)) return future
  if (isFormulaError(type)) return type
  const r = rate as number, n = periods as number, principal = amount as number, fv = future as number, due = type ? 1 : 0
  if (n <= 0 || !Number.isFinite(n) || ![0, 1].includes(due)) return formulaError('#NUM!')
  const growth = (1 + r) ** n
  const factor = r === 0 ? n : (1 - growth ** -1) / r * (1 + r * due)
  if (!Number.isFinite(factor) || factor === 0 || !Number.isFinite(growth)) return formulaError('#NUM!')
  if (name === 'PMT') return finiteResult(-(principal + fv / growth) / factor)
  if (name === 'PV') return finiteResult(-(principal * factor + fv / growth))
  return finiteResult(-(principal * factor * growth + fv * growth))
}

function evaluateDepreciationAndDollarValue(name: 'DB' | 'DDB' | 'DOLLARDE' | 'DOLLARFR' | 'RRI' | 'SLN' | 'SYD', values: RuntimeValue[]): FormulaValue {
  const minArgs = name === 'DB' || name === 'DDB' || name === 'SYD' ? 4 : name === 'SLN' || name === 'RRI' ? 3 : 2
  const maxArgs = name === 'DB' || name === 'DDB' ? 5 : name === 'RRI' ? 6 : minArgs
  if (values.length < minArgs || values.length > maxArgs) return formulaError('#VALUE!')

  const usedValues = name === 'RRI' ? values.slice(0, 3) : values
  const read = (value: RuntimeValue): number | FormulaErrorValue => {
    if (isFormulaError(value)) return value
    if (isRangeValue(value)) return formulaError('#VALUE!')
    if (name.startsWith('DOLLAR') && typeof value === 'boolean') return formulaError('#VALUE!')
    if (typeof value === 'string') {
      const number = Number(value)
      return Number.isFinite(number) ? number : formulaError('#VALUE!')
    }
    return numericCoercion(value)
  }
  const numbers = usedValues.map((value, index) => {
    if ((name === 'DB' && index === 4 && value === null) || (name === 'DDB' && index === 4 && value === null)) return name === 'DB' ? 12 : 2
    return read(value)
  })
  const error = numbers.find(isFormulaError)
  if (error) return error
  const args = numbers as number[]

  if (name === 'DOLLARDE' || name === 'DOLLARFR') {
    const [amount, rawFraction] = args
    const fraction = Math.floor(rawFraction!)
    if (fraction < 0) return formulaError('#NUM!')
    if (fraction < 1) return formulaError('#DIV/0!')
    const integerPart = Number.parseInt(String(amount), 10)
    const decimalPart = amount! % 1
    if (name === 'DOLLARDE') {
      let result = integerPart + (decimalPart * 10 ** Math.ceil(Math.log(fraction) / Math.LN10)) / fraction
      const power = 10 ** (Math.ceil(Math.log(fraction) / Math.LN2) + 1)
      result = Math.round(result * power) / power
      return finiteResult(result)
    }
    return finiteResult(integerPart + decimalPart * 10 ** -Math.ceil(Math.log(fraction) / Math.LN10) * fraction)
  }

  if (name === 'RRI') {
    const [periods, principal, future] = args
    if (periods! <= 0) return formulaError('#NUM!')
    if (principal === 0 && future === 0) return 0
    const ratio = future! / principal!
    const result = ratio ** (1 / periods!) - 1
    return ratio < 0 ? formulaError('#NUM!') : finiteResult(result)
  }

  if (name === 'SLN') {
    const [cost, salvage, life] = args
    if (life === 0) return formulaError('#DIV/0!')
    return finiteResult((cost! - salvage!) / life!)
  }

  if (name === 'SYD') {
    const [cost, salvage, life, period] = args
    if (salvage! < 0 || life! <= 0 || period! > life!) return formulaError('#NUM!')
    return finiteResult((cost! - salvage!) * (life! - period! + 1) * 2 / (life! * (life! + 1)))
  }

  if (name === 'DDB') {
    const [cost, salvage, life, period] = args
    const factor = values.length === 5 ? args[4]! : 2
    if (cost! < 0 || salvage! < 0 || life! <= 0 || period! <= 0 || period! > life! || factor <= 0) return formulaError('#NUM!')
    const rate = factor / life!
    const oldCost = rate >= 1 ? (period === 1 ? cost! : 0) : cost! * (1 - rate) ** (period! - 1)
    const newCost = rate >= 1 ? 0 : cost! * (1 - rate) ** period!
    return finiteResult(Math.max(0, oldCost - Math.max(newCost, salvage!)))
  }

  const [cost, salvage, life, period] = args
  const month = values.length === 5 ? args[4]! : 12
  if (cost! < 0 || salvage! < 0 || life! <= 0 || period! <= 0 || Math.floor(period!) > Math.floor(life!) || month < 1 || month > 12) return formulaError('#NUM!')
  const roundedPeriod = Math.max(1, Math.floor(period!))
  const rate = Number(((1 - (salvage! / cost!) ** (1 / life!)).toFixed(3)))
  const initial = cost! * rate * Math.floor(month) / 12
  let total = initial
  let current = 0
  const ceiling = roundedPeriod === life ? life! - 1 : roundedPeriod
  for (let index = 2; index <= ceiling; index += 1) {
    current = (cost! - total) * rate
    total += current
  }
  const result = roundedPeriod === 1 ? initial : roundedPeriod === life ? (cost! - total) * rate : current
  return finiteResult(result)
}

function solveRate(periods: number, payment: number, principal: number, future: number, type: number, guess: number): FormulaValue {
  const equation = (rate: number) => rate === 0 ? principal + payment * periods + future
    : principal * (1 + rate) ** periods + payment * (1 + rate * type) * (((1 + rate) ** periods - 1) / rate) + future
  let rate = Math.max(-0.999999, guess)
  for (let iteration = 0; iteration < 100; iteration++) {
    const result = equation(rate)
    if (Number.isFinite(result) && Math.abs(result) < 1e-9) return rate
    const delta = Math.max(1e-7, Math.abs(rate) * 1e-6)
    const derivative = (equation(rate + delta) - equation(rate - delta)) / (2 * delta)
    if (!Number.isFinite(result) || !Number.isFinite(derivative) || Math.abs(derivative) < 1e-14) break
    const next = rate - result / derivative
    if (!Number.isFinite(next) || next <= -1 || Math.abs(next - rate) > Math.max(10, Math.abs(rate) * 100)) break
    rate = next
  }
  const intervals: Array<[number, number]> = [[-0.999999999, 0]]
  let edge = 0.01
  while (edge < 1e8) { intervals.push([edge / 2, edge]); edge *= 2 }
  for (const [lowBound, highBound] of intervals) {
    let low = lowBound, high = highBound
    let lowValue = equation(low), highValue = equation(high)
    if (!Number.isFinite(lowValue) || !Number.isFinite(highValue) || lowValue * highValue > 0) continue
    for (let iteration = 0; iteration < 200; iteration++) {
      const middle = (low + high) / 2, middleValue = equation(middle)
      if (Math.abs(middleValue) < 1e-9 || high - low < 1e-12) return middle
      if (lowValue * middleValue <= 0) high = middle
      else { low = middle; lowValue = middleValue }
    }
  }
  return formulaError('#NUM!')
}

function financialPayment(rate: number, periods: number, principal: number, future: number, type: number): number | FormulaErrorValue {
  if (!Number.isFinite(periods) || periods === 0) return formulaError('#NUM!')
  const growth = (1 + rate) ** periods
  const factor = rate === 0 ? periods : (1 - growth ** -1) / rate * (1 + rate * type)
  return factor && Number.isFinite(factor) && Number.isFinite(growth) ? finiteResult(-(principal + future / growth) / factor) as number | FormulaErrorValue : formulaError('#NUM!')
}

function financialFutureValue(rate: number, periods: number, payment: number, principal: number, type: number): number {
  if (rate === 0) return -(principal + payment * periods)
  return -(principal * (1 + rate) ** periods + payment * (type ? 1 + rate : 1) * ((1 + rate) ** periods - 1) / rate)
}

function evaluateFutureValueSchedule(values: RuntimeValue[], sheet: SpreadsheetSheet, resolve: (sheet: SpreadsheetSheet, row: number, column: number) => FormulaValue): FormulaValue {
  if (values.length !== 2) return formulaError('#VALUE!')
  if (typeof values[0] === 'boolean') return formulaError('#VALUE!')
  const principal = numericCoercion(values[0]!)
  if (isFormulaError(principal)) return principal
  const schedule = flattenScalar(values[1]!, sheet, resolve)
  let result = principal
  for (const rateValue of schedule) {
    if (isFormulaError(rateValue)) return rateValue
    if (typeof rateValue === 'boolean') return formulaError('#VALUE!')
    const rate = typeof rateValue === 'number' ? rateValue : typeof rateValue === 'string' && rateValue.trim() ? Number(rateValue) : NaN
    if (!Number.isFinite(rate)) return formulaError('#VALUE!')
    result *= 1 + rate
  }
  return finiteResult(result)
}

function evaluateInternalRateOfReturn(values: RuntimeValue[], sheet: SpreadsheetSheet, resolve: (sheet: SpreadsheetSheet, row: number, column: number) => FormulaValue): FormulaValue {
  if (values.length < 1 || values.length > 2) return formulaError('#VALUE!')
  const cashflows: number[] = []
  for (const item of flattenScalar(values[0]!, sheet, resolve)) {
    if (isFormulaError(item)) return formulaError('#VALUE!')
    if (item === null || typeof item === 'boolean' || item === '') continue
    if (typeof item === 'number') cashflows.push(item)
    else if (item.trim() && Number.isFinite(Number(item))) cashflows.push(Number(item))
  }
  if (!cashflows.some(value => value < 0) || !cashflows.some(value => value > 0)) return formulaError('#NUM!')
  const guessValue = values.length === 2 ? numericCoercion(values[1]!) : 0.1
  if (isFormulaError(guessValue)) return guessValue
  if (guessValue <= -1) return formulaError('#NUM!')
  const npvAt = (rate: number) => cashflows.reduce((sum, value, index) => sum + value / (1 + rate) ** index, 0)
  let rate = guessValue
  for (let iteration = 0; iteration < 100; iteration++) {
    if (rate <= -1 || !Number.isFinite(rate)) break
    const result = npvAt(rate)
    if (Math.abs(result) < 1e-10) return rate
    let derivative = 0
    for (let index = 1; index < cashflows.length; index++) derivative -= index * cashflows[index]! / (1 + rate) ** (index + 1)
    if (!Number.isFinite(derivative) || Math.abs(derivative) < 1e-14) break
    const next = rate - result / derivative
    if (!Number.isFinite(next) || next <= -1 || Math.abs(next - rate) > Math.max(10, Math.abs(rate) * 100)) break
    rate = next
  }
  let low = -0.999999999, lowValue = npvAt(low)
  for (let high = Math.max(0.01, guessValue, 0); high <= 1e10; high = high * 2 + 0.01) {
    const highValue = npvAt(high)
    if (Number.isFinite(lowValue) && Number.isFinite(highValue) && lowValue * highValue <= 0) {
      for (let iteration = 0; iteration < 200; iteration++) {
        const middle = (low + high) / 2
        const middleValue = npvAt(middle)
        if (Math.abs(middleValue) < 1e-10 || high - low < 1e-12) return middle
        if (lowValue * middleValue <= 0) high = middle
        else { low = middle; lowValue = middleValue }
      }
      return formulaError('#NUM!')
    }
  }
  return formulaError('#NUM!')
}

function evaluateNetPresentValue(values: RuntimeValue[], sheet: SpreadsheetSheet, resolve: (sheet: SpreadsheetSheet, row: number, column: number) => FormulaValue): FormulaValue {
  if (values.length < 2) return formulaError('#VALUE!')
  const rate = numericCoercion(values[0]!)
  if (isFormulaError(rate) || rate <= -1) return isFormulaError(rate) ? rate : formulaError('#NUM!')
  let total = 0, period = 1
  for (const value of values.slice(1)) for (const item of flattenScalar(value!, sheet, resolve)) {
    if (isFormulaError(item)) return item
    const number = numericCoercion(item)
    if (isFormulaError(number)) return number
    total += number / (1 + rate) ** period
    period += 1
  }
  return finiteResult(total)
}

interface ConditionalCell { row: number; column: number; value: FormulaValue }

function evaluateCountValue(name: 'COUNTBLANK' | 'COUNTA', values: RuntimeValue[], args: FormulaNode[], sheet: SpreadsheetSheet, resolve: (sheet: SpreadsheetSheet, row: number, column: number) => FormulaValue): FormulaValue {
  if (args.length < 1 || args.length > (name === 'COUNTA' ? 255 : 1)) return formulaError('#VALUE!')
  let count = 0
  for (let index = 0; index < values.length; index++) {
    const value = values[index]!
    if (isFormulaError(value)) {
      if (name === 'COUNTBLANK') return value
      count++
      continue
    }
    if (name === 'COUNTA' && !isRangeValue(value)) {
      if (value !== null) count++
      continue
    }
    const cells = conditionalCells(value, args[index]!, sheet, resolve)
    if (isFormulaError(cells)) return cells
    for (const cell of cells) {
      if (name === 'COUNTA' ? cell.value !== null : cell.value === null || cell.value === '') count++
    }
  }
  return count
}

function conditionalCells(value: RuntimeValue, node: FormulaNode, _sheet: SpreadsheetSheet, resolve: (sheet: SpreadsheetSheet, row: number, column: number) => FormulaValue): ConditionalCell[] | FormulaErrorValue {
  if (isRangeValue(value)) {
    const minRow = Math.min(value.reference.start.row, value.reference.end.row), maxRow = Math.max(value.reference.start.row, value.reference.end.row)
    const minColumn = Math.min(value.reference.start.column, value.reference.end.column), maxColumn = Math.max(value.reference.start.column, value.reference.end.column)
    const cells: ConditionalCell[] = []
    for (let row = minRow; row <= maxRow; row++) for (let column = minColumn; column <= maxColumn; column++) {
      const cell = value.sheet.cells[cellKey(row, column)]
      cells.push({ row, column, value: cell ? resolve(value.sheet, row, column) : null })
    }
    return cells
  }
  if (node.kind === 'reference') return [{ row: node.reference.row, column: node.reference.column, value }]
  return formulaError('#VALUE!')
}

function evaluateConditionalAggregate(name: string, values: RuntimeValue[], args: FormulaNode[], sheet: SpreadsheetSheet, resolve: (sheet: SpreadsheetSheet, row: number, column: number) => FormulaValue, workbook: WorkbookModel): FormulaValue {
  const multi = name.endsWith('IFS')
  const average = name.startsWith('AVERAGE')
  const sum = name.startsWith('SUM')
  const isCount = !sum && !average
  if (sum && (isFormulaError(values[0]!) || (!multi && args.length === 3 && isFormulaError(values[2]!)))) {
    return isFormulaError(values[0]!) ? values[0]! as FormulaErrorValue : values[2]! as FormulaErrorValue
  }
  if ((!multi && (args.length < 2 || args.length > (isCount ? 2 : 3)))
    || (multi && (args.length < (isCount ? 2 : 3) || args.length > 255 || args.length % 2 !== (isCount ? 0 : 1)))) return formulaError('#VALUE!')
  const rangeStart = multi && !isCount ? 1 : 0
  const criteriaEnd = multi ? args.length : 2
  const criteriaPairs: Array<{ cells: ConditionalCell[]; criterion: FormulaValue }> = []
  let base: ConditionalCell[] | undefined
  for (let index = rangeStart; index < criteriaEnd; index += 2) {
    if (index + 1 >= args.length) return formulaError('#VALUE!')
    const cells = conditionalCells(values[index]!, args[index]!, sheet, resolve)
    if (isFormulaError(cells)) return cells
    if (!base) base = cells
    else if (multi && cells.length !== base.length) return formulaError('#VALUE!')
    criteriaPairs.push({ cells, criterion: scalarValue(values[index + 1]!) })
  }
  if (!base || !criteriaPairs.length) return formulaError('#VALUE!')
  let target: ConditionalCell[] | undefined
  if (!isCount) {
    const targetIndex = multi ? 0 : args.length === 3 ? 2 : 0
    const targetCells = conditionalCells(values[targetIndex]!, args[targetIndex]!, sheet, resolve)
    if (isFormulaError(targetCells)) return targetCells
    if (multi && targetCells.length !== base.length) return formulaError('#VALUE!')
    target = !multi && targetCells.length !== base.length ? alignConditionalTarget(values[targetIndex]!, targetCells, base, resolve) : targetCells
  }
  let total = 0, count = 0
  for (let index = 0; index < base.length; index++) {
    let matches = true
    for (const pair of criteriaPairs) {
      const criterion = pair.criterion
      if (isFormulaError(criterion) || !matchesCriterion(pair.cells[index]!.value, criterion, workbook)) { matches = false; break }
    }
    if (!matches) continue
    if (!sum && !average) { count++; continue }
    const targetValue = target?.[index]?.value ?? null
    if (isFormulaError(targetValue)) return targetValue
    if (typeof targetValue === 'number') { total += targetValue; count++ }
  }
  if (!sum && !average) return count
  if (average) return count ? total / count : formulaError('#DIV/0!')
  return finiteResult(total)
}

function matchesCriterion(value: FormulaValue, criterion: FormulaValue, workbook: WorkbookModel): boolean {
  let operator = '='
  let expected = criterion
  if (typeof criterion === 'string') {
    const match = /^(>=|<=|<>|>|<|=)/.exec(criterion)
    if (match) { operator = match[0]; expected = criterion.slice(operator.length) }
    if (typeof expected === 'string' && expected === '' && ['>', '>=', '<', '<='].includes(operator)) return false
    if (typeof expected === 'string') {
      const date = parseDateText(expected)
      if (date) {
        const serial = datePartsToSerial(date.year, date.month, date.day, workbook.date1904)
        if (typeof serial === 'number') expected = serial
      } else if (/^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(expected.trim())) {
        expected = Number(expected)
      }
    }
  }
  if (typeof expected === 'string' && /[*?~]/.test(expected)) {
    const pattern = wildcardPattern(expected)
    const matched = typeof value === 'string' && pattern.test(value)
    return operator === '<>' ? !matched && value !== null : operator === '=' && matched
  }
  if (expected === '' && operator === '<>') return value !== null
  if (value === null || value === '') {
    if (operator === '=' && (expected === '' || expected === 0 || expected === null)) return true
    if (operator === '<>' && expected === '') return value === ''
    return false
  }
  if (isFormulaError(value)) return isFormulaError(expected) ? operator === '=' ? value.code === expected.code : operator === '<>' && value.code !== expected.code : false
  if (typeof value === 'number' && typeof expected !== 'number') return false
  if (typeof value !== 'number' && typeof expected === 'number') return false
  if (typeof value === 'boolean' || typeof expected === 'boolean') {
    const ordering = typeof value === 'boolean' && typeof expected === 'boolean' ? Number(value) - Number(expected) : 1
    return compareOrdering(ordering, operator)
  }
  const ordering = typeof value === 'number' && typeof expected === 'number' ? Math.sign(value - expected)
    : scalarToText(value).localeCompare(scalarToText(expected), 'en-US', { sensitivity: 'base', numeric: true })
  return compareOrdering(ordering, operator)
}

function alignConditionalTarget(value: RuntimeValue, target: ConditionalCell[], base: ConditionalCell[], resolve: (sheet: SpreadsheetSheet, row: number, column: number) => FormulaValue): ConditionalCell[] {
  if (!isRangeValue(value) || !base.length) return base.map((_cell, index) => target[index] ?? { row: 0, column: 0, value: null })
  const first = base[0]!
  const startRow = Math.min(value.reference.start.row, value.reference.end.row)
  const startColumn = Math.min(value.reference.start.column, value.reference.end.column)
  return base.map((cell) => {
    const row = startRow + cell.row - first.row
    const column = startColumn + cell.column - first.column
    const source = value.sheet.cells[cellKey(row, column)]
    return { row, column, value: source ? resolve(value.sheet, row, column) : null }
  })
}

function compareOrdering(ordering: number, operator: string): boolean {
  if (operator === '=') return ordering === 0
  if (operator === '<>') return ordering !== 0
  if (operator === '>') return ordering > 0
  if (operator === '>=') return ordering >= 0
  if (operator === '<') return ordering < 0
  if (operator === '<=') return ordering <= 0
  return false
}

function wildcardPattern(value: string): RegExp {
  let pattern = '^'
  for (let index = 0; index < value.length; index++) {
    const character = value[index]!
    if (character === '~' && index + 1 < value.length) pattern += value[++index]!.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    else if (character === '*') pattern += '.*'
    else if (character === '?') pattern += '.'
    else pattern += character.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  }
  return new RegExp(`${pattern}$`, 'i')
}

function evaluateStatisticalValue(name: 'MEDIAN' | 'LARGE' | 'SMALL' | 'VAR.S' | 'VAR.P' | 'STDEV.S' | 'STDEV.P' | 'PERCENTILE.INC' | 'PERCENTILE.EXC' | 'QUARTILE.INC' | 'QUARTILE.EXC', values: RuntimeValue[], sheet: SpreadsheetSheet, resolve: (sheet: SpreadsheetSheet, row: number, column: number) => FormulaValue, workbook: WorkbookModel): FormulaValue {
  const ranked = ['LARGE', 'SMALL'].includes(name)
  const percentile = name.startsWith('PERCENTILE.')
  const quartile = name.startsWith('QUARTILE.')
  const isVarianceFunction = ['VAR.S', 'VAR.P', 'STDEV.S', 'STDEV.P'].includes(name)
  if (!values.length || ((ranked || percentile || quartile) && values.length !== 2) || (isVarianceFunction && values.length > 255)) return formulaError('#VALUE!')
  const numbers: number[] = []
  const sourceValues = values.slice(0, ranked || percentile || quartile ? 1 : values.length)
  for (const value of sourceValues) {
    for (const item of flattenScalar(value!, sheet, resolve)) {
      if (isFormulaError(item)) return item
      if (typeof item === 'number') numbers.push(item)
      else if (!isRangeValue(value) && item !== null && item !== '') {
        const coerced = aggregateLiteralNumber(item, workbook)
        if (isFormulaError(coerced)) return coerced
        numbers.push(coerced)
      }
    }
  }
  if (name === 'LARGE' || name === 'SMALL') {
    const rank = numericCoercion(values[1]!)
    if (isFormulaError(rank) || !Number.isInteger(rank) || rank < 1 || rank > numbers.length) return formulaError('#NUM!')
    numbers.sort((a, b) => a - b)
    return numbers[name === 'LARGE' ? numbers.length - rank : rank - 1]!
  }
  if (percentile || quartile) {
    const raw = numericCoercion(values[1]!)
    if (isFormulaError(raw)) return raw
    const isExclusive = name.endsWith('.EXC')
    const index = quartile ? raw / 4 : raw
    if ((quartile && (raw < 0 || raw > 4 || (isExclusive && (raw === 0 || raw === 4))))
      || (!quartile && (index < 0 || index > 1))
      || (isExclusive && (index < 1 / (numbers.length + 1) || index > numbers.length / (numbers.length + 1)))) return formulaError('#NUM!')
    if (!numbers.length) return formulaError('#NUM!')
    numbers.sort((a, b) => a - b)
    const position = isExclusive ? (numbers.length + 1) * index - 1 : (numbers.length - 1) * index
    const lower = Math.floor(position), fraction = position - lower
    const left = numbers[Math.max(0, lower)]!, right = numbers[Math.min(numbers.length - 1, lower + 1)]!
    return finiteResult(left + (right - left) * fraction)
  }
  if (!numbers.length) return formulaError('#NUM!')
  numbers.sort((a, b) => a - b)
  if (name === 'MEDIAN') {
    const middle = Math.floor(numbers.length / 2)
    return numbers.length % 2 ? numbers[middle]! : (numbers[middle - 1]! + numbers[middle]!) / 2
  }
  const sample = name === 'VAR.S' || name === 'STDEV.S'
  if (sample && numbers.length < 2) return formulaError('#DIV/0!')
  const mean = numbers.reduce((sum, value) => sum + value, 0) / numbers.length
  const squaredDifference = numbers.reduce((sum, value) => sum + (value - mean) ** 2, 0)
  const variance = squaredDifference / (numbers.length - Number(sample))
  return name === 'STDEV.S' || name === 'STDEV.P' ? finiteResult(Math.sqrt(variance)) : finiteResult(variance)
}

/** Aggregates coerce direct scalar text, while referenced text is ignored above. */
function aggregateLiteralNumber(value: FormulaValue, workbook: WorkbookModel): number | FormulaErrorValue {
  if (typeof value !== 'string') return numericCoercion(value)
  const text = value.trim()
  const date = parseDateText(text)
  if (date) {
    const serial = datePartsToSerial(date.year, date.month, date.day, workbook.date1904)
    return typeof serial === 'number' ? serial : formulaError('#VALUE!')
  }
  // Decimal/scientific notation, conventional thousands groups and percentages.
  // Locale-specific date/currency forms need the shared number-format parser.
  if (!/^[+-]?(?:\d+\.?\d*|\d{1,3}(?:,\d{3})+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?%?$/i.test(text)) return formulaError('#VALUE!')
  const number = Number(text.replaceAll(',', '').replace(/%$/, '')) / (text.endsWith('%') ? 100 : 1)
  return Number.isFinite(number) ? number : formulaError('#NUM!')
}

function evaluateBinary(operator: string, left: FormulaValue, right: FormulaValue): FormulaValue {
  if (isFormulaError(left)) return left
  if (isFormulaError(right)) return right
  if (operator === '&') return scalarToText(left) + scalarToText(right)
  if (['=', '<>', '<', '>', '<=', '>='].includes(operator)) return compare(left, right, operator)
  const leftNumber = numericCoercion(left), rightNumber = numericCoercion(right)
  if (isFormulaError(leftNumber)) return leftNumber
  if (isFormulaError(rightNumber)) return rightNumber
  if (operator === '/' && rightNumber === 0) return formulaError('#DIV/0!')
  switch (operator) {
    case '+': return finiteResult(leftNumber + rightNumber)
    case '-': return finiteResult(leftNumber - rightNumber)
    case '*': return finiteResult(leftNumber * rightNumber)
    case '/': return finiteResult(leftNumber / rightNumber)
    case '^': return finiteResult(leftNumber ** rightNumber)
    default: return formulaError('#VALUE!')
  }
}

function compare(left: FormulaValue, right: FormulaValue, operator: string): boolean {
  let result: number
  if (typeof left === 'number' && typeof right === 'number') result = Math.sign(left - right)
  else if (typeof left === 'boolean' && typeof right === 'boolean') result = Number(left) - Number(right)
  else if ((left === null || left === '') && (right === null || right === '')) result = 0
  else result = scalarToText(left).localeCompare(scalarToText(right), 'en-US', { sensitivity: 'base', numeric: true })
  switch (operator) {
    case '=': return result === 0
    case '<>': return result !== 0
    case '<': return result < 0
    case '>': return result > 0
    case '<=': return result <= 0
    case '>=': return result >= 0
    default: return false
  }
}

function collectReferences(node: FormulaNode): Array<FormulaReference | FormulaRangeReference> {
  const result: Array<FormulaReference | FormulaRangeReference> = [], queue = [node], visited = new Set<FormulaNode>()
  while (queue.length) {
    const current = queue.pop()!
    if (visited.has(current)) continue
    visited.add(current)
    if (current.kind === 'reference' || current.kind === 'range') result.push(current.reference)
    else if (current.kind === 'unary') queue.push(current.operand)
    else if (current.kind === 'binary') queue.push(current.left, current.right)
    else if (current.kind === 'function') queue.push(...current.arguments)
  }
  return result
}
function containsDynamicReference(node: FormulaNode): boolean {
  const queue = [node], visited = new Set<FormulaNode>()
  while (queue.length) {
    const current = queue.pop()!
    if (visited.has(current)) continue
    visited.add(current)
    if (current.kind === 'function') { if (['INDIRECT', 'OFFSET'].includes(current.name)) return true; queue.push(...current.arguments) }
    else if (current.kind === 'unary') queue.push(current.operand)
    else if (current.kind === 'binary') queue.push(current.left, current.right)
  }
  return false
}

function getSparseRangeTargets(sheet: SpreadsheetSheet, reference: FormulaRangeReference): string[] {
  const minRow = Math.min(reference.start.row, reference.end.row), maxRow = Math.max(reference.start.row, reference.end.row)
  const minCol = Math.min(reference.start.column, reference.end.column), maxCol = Math.max(reference.start.column, reference.end.column)
  return Object.values(sheet.cells).filter((cell) => cell.row >= minRow && cell.row <= maxRow && cell.col >= minCol && cell.col <= maxCol)
    .map((cell) => formulaCellKey({ sheetId: sheet.id, row: cell.row, column: cell.col }))
}

function resolveSheet(workbook: WorkbookModel, sheetName: string | undefined, current: SpreadsheetSheet, cache?: Map<string, SpreadsheetSheet>): SpreadsheetSheet | undefined {
  if (!sheetName) return current
  return cache?.get(sheetName.toLocaleLowerCase()) ?? workbook.sheets.find((sheet) => sheet.name.toLocaleLowerCase() === sheetName.toLocaleLowerCase())
}

function cellsInRange(value: RangeValue): SpreadsheetCell[] {
  const minRow = Math.min(value.reference.start.row, value.reference.end.row), maxRow = Math.max(value.reference.start.row, value.reference.end.row)
  const minCol = Math.min(value.reference.start.column, value.reference.end.column), maxCol = Math.max(value.reference.start.column, value.reference.end.column)
  return Object.values(value.sheet.cells).filter((cell) => cell.row >= minRow && cell.row <= maxRow && cell.col >= minCol && cell.col <= maxCol)
}
function referenceBounds(reference: FormulaRangeReference): import('../../domain/model').CellRange {
  return { startRow: Math.min(reference.start.row, reference.end.row), endRow: Math.max(reference.start.row, reference.end.row),
    startCol: Math.min(reference.start.column, reference.end.column), endCol: Math.max(reference.start.column, reference.end.column) }
}
function checkedReference(value: RangeValue, resolve: FormulaValueResolver): RangeValue | FormulaErrorValue {
  return resolve.canReadRange?.(value.sheet, referenceBounds(value.reference)) === false ? formulaError('#REF!') : value
}

function rawCellValue(cell: SpreadsheetCell): FormulaValue {
  const type = normalizeSheetCellValueType(cell.formula ? cell.cachedDataType : cell.dataType)
  if (type === 'error' || type === undefined && typeof cell.value === 'string' && ERROR_CODES.has(cell.value.toUpperCase())) {
    return formulaError(typeof cell.value === 'string' && ERROR_CODES.has(cell.value.toUpperCase()) ? cell.value.toUpperCase() : cell.display || '#VALUE!')
  }
  return cell.value
}

function flattenScalar(value: RuntimeValue, _sheet: SpreadsheetSheet, resolve: (sheet: SpreadsheetSheet, row: number, column: number) => FormulaValue): FormulaValue[] {
  if (!isRangeValue(value)) return [value]
  return cellsInRange(value).map((cell) => resolve(value.sheet, cell.row, cell.col))
}

function flattenText(values: RuntimeValue[], sheet: SpreadsheetSheet, resolve: (sheet: SpreadsheetSheet, row: number, column: number) => FormulaValue): string | FormulaErrorValue {
  const parts: string[] = []
  for (const value of values) for (const item of flattenScalar(value, sheet, resolve)) {
    if (isFormulaError(item)) return item
    parts.push(scalarToText(item))
  }
  return parts.join('')
}

function numericTextCoercion(value: RuntimeValue, workbook: WorkbookModel, evaluationTime: Date, formatted = true): number | FormulaErrorValue {
  if (typeof value !== 'string') return numericCoercion(value)
  const options = { dateSystem: workbook.date1904 ? DateSystem.Date1904 : DateSystem.Date1900, locale: 'en', currentYear: evaluationTime.getFullYear() }
  const parsed = parseDate(value, options) ?? parseTime(value, options) ?? (formatted ? parseNumber(value, options) : null)
  const number = parsed?.v ?? Number(value)
  if (!value.trim() || Number.isNaN(number)) return formulaError('#VALUE!')
  return number
}

function numericCoercion(value: RuntimeValue): number | FormulaErrorValue {
  if (isFormulaError(value)) return value
  if (isRangeValue(value)) return formulaError('#VALUE!')
  if (value === null || value === '') return 0
  if (typeof value === 'number') return Number.isFinite(value) ? value : formulaError('#NUM!')
  if (typeof value === 'boolean') return value ? 1 : 0
  return formulaError('#VALUE!')
}

function booleanCoercion(value: RuntimeValue): boolean | FormulaErrorValue {
  if (isFormulaError(value)) return value
  if (isRangeValue(value)) return formulaError('#VALUE!')
  if (value === null || value === '') return false
  if (typeof value === 'boolean') return value
  if (typeof value === 'number') return value !== 0
  if (typeof value === 'string') return value.length > 0
  return formulaError('#VALUE!')
}

function invertBoolean(value: RuntimeValue): boolean | FormulaErrorValue {
  const result = booleanCoercion(value)
  return isFormulaError(result) ? result : !result
}

function scalarToText(value: FormulaValue): string {
  if (isFormulaError(value)) return value.code
  if (value === null) return ''
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  return String(value)
}

function scalarValue(value: RuntimeValue): FormulaValue { return isRangeValue(value) ? formulaError('#VALUE!') : value }
function singleRangeValue(value: RuntimeValue, resolve: (sheet: SpreadsheetSheet, row: number, column: number) => FormulaValue): FormulaValue {
  if (!isRangeValue(value)) return value
  const { start, end } = value.reference
  if (start.row !== end.row || start.column !== end.column) return formulaError('#VALUE!')
  return resolve(value.sheet, start.row, start.column)
}
function formulaReferenceCellValue(sheet: SpreadsheetSheet, row: number, column: number, resolve: (sheet: SpreadsheetSheet, row: number, column: number) => FormulaValue): FormulaValue {
  const value = resolve(sheet, row, column)
  const cell = sheet.cells[cellKey(row, column)]
  if (!cell) return value
  if (typeof value !== 'number') return value
  if (normalizeSheetCellValueType(cell.formula ? cell.cachedDataType : cell.dataType) === 'number'
    && isSheetTextNumberFormat(cell.style.numberFormat)) return String(value)
  return Number.isFinite(value) ? value : formulaError('#NUM!')
}
function formulaArgumentMatrix(value: RuntimeValue, resolve: (sheet: SpreadsheetSheet, row: number, column: number) => FormulaValue) {
  if (!isRangeValue(value)) return { rowCount: 1, columnCount: 1, valueAt: (_index: number) => value }
  const columnCount = Math.abs(value.reference.end.column - value.reference.start.column) + 1
  const startRow = Math.min(value.reference.start.row, value.reference.end.row)
  const startColumn = Math.min(value.reference.start.column, value.reference.end.column)
  return {
    rowCount: Math.abs(value.reference.end.row - value.reference.start.row) + 1,
    columnCount,
    valueAt: (index: number) => formulaReferenceCellValue(value.sheet, startRow + Math.floor(index / columnCount), startColumn + index % columnCount, resolve),
  }
}
function normalizeResult(value: RuntimeValue, resolve: (sheet: SpreadsheetSheet, row: number, column: number) => FormulaValue): FormulaValue { return singleRangeValue(value, resolve) }
function isRangeValue(value: RuntimeValue): value is RangeValue { return Boolean(value && typeof value === 'object' && 'kind' in value && value.kind === 'rangeValue') }
function factorial(integer: number, step = 1) {
  let result = 1
  for (let factor = integer; factor > 1 && Number.isFinite(result); factor -= step) {
    result *= factor
  }
  return result
}
function finiteResult(value: number): FormulaValue { return Number.isFinite(value) ? value : formulaError('#NUM!') }
function displayFormulaValue(value: FormulaValue): string { return isFormulaError(value) ? value.code : typeof value === 'boolean' ? (value ? 'TRUE' : 'FALSE') : value === null ? '' : String(value) }
function valueType(value: FormulaValue): string { return isFormulaError(value) ? 'error' : value === null ? 'empty' : typeof value }
function hasCachedValue(cell: SpreadsheetCell): boolean {
  const type = normalizeSheetCellValueType(cell.cachedDataType)
  return cell.value !== null && (type !== undefined && type !== 'formula' || cell.value !== '' || cell.display !== '')
}
function diagnostic(cell: FormulaCellAddress, formula: string, error: unknown): FormulaEngineDiagnostic {
  if (error instanceof FormulaSyntaxError) return { code: error.code, message: error.message, cell, formula }
  return { code: 'FORMULA_ERROR', message: error instanceof Error ? error.message : String(error), cell, formula }
}
function toFormulaError(error: unknown): FormulaErrorValue {
  if (error instanceof FormulaSyntaxError) return formulaError(error.code === 'FORMULA_NAME' ? '#NAME?' : error.code === 'FORMULA_REFERENCE' ? '#REF!' : '#VALUE!')
  return formulaError('#VALUE!')
}
function cellKey(row: number, column: number): string { return `${row}:${column}` }

function validEvaluationTime(value: Date | undefined): Date | undefined {
  return value instanceof Date && Number.isFinite(value.getTime()) ? value : undefined
}

interface DateParts { year: number; month: number; day: number }
const MS_PER_DAY = 86_400_000

function datePartsToSerial(year: number, month: number, day: number, date1904: boolean): FormulaValue {
  return excelDateTimePartsToSerial({ year, month, day, hours: 0, minutes: 0, seconds: 0, fractionalSecond: 0 },
    { dateSystem: date1904 ? DateSystem.Date1904 : DateSystem.Date1900 }) ?? formulaError('#NUM!')
}

function dateArgumentsToSerial(year: number, month: number, day: number, date1904: boolean): FormulaValue {
  const normalizedYear = year + Math.floor((month - 1) / 12)
  const normalizedMonth = ((month - 1) % 12 + 12) % 12
  if (normalizedYear < 0 || normalizedYear > 9999) return formulaError('#NUM!')
  const normalizedMonthNumber = normalizedMonth + 1
  const first = new Date(0)
  first.setUTCFullYear(normalizedYear, normalizedMonthNumber - 1, 1)
  const dateSystem = date1904 ? DateSystem.Date1904 : DateSystem.Date1900
  const serial = excelDateSerial(first, dateSystem) + Math.trunc(day) - 1
  return Number.isFinite(serial) && serial >= 0 && serial < getDateSystemMaxSerial(dateSystem) + 1 ? serial : formulaError('#NUM!')
}

function serialToDateParts(serial: number, date1904: boolean): DateParts | undefined {
  const parts = excelSerialToDateTimeParts(serial, { dateSystem: date1904 ? DateSystem.Date1904 : DateSystem.Date1900 })
  return parts ? { year: parts.year, month: parts.month, day: parts.day } : undefined
}

function parseDateText(value: string): DateParts | undefined {
  const match = /^\s*(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})\s*$/.exec(value)
  if (!match) return undefined
  const year = Number(match[1]), month = Number(match[2]), day = Number(match[3])
  const date = new Date(Date.UTC(year, month - 1, day))
  return date.getUTCFullYear() === year && date.getUTCMonth() + 1 === month && date.getUTCDate() === day ? { year, month, day } : undefined
}

function parseTimeValue(value: string): number | undefined {
  const text = value.trim()
  if (!text) return undefined
  if (parseDateText(text)) return 0
  let timeText = text
  const dateTime = /^(\d{4}[-/.]\d{1,2}[-/.]\d{1,2})[ T](.+)$/.exec(text)
  if (dateTime) {
    if (!parseDateText(dateTime[1]!)) return undefined
    timeText = dateTime[2]!
  }
  const match = /^(\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?\s*(AM|PM)?$/i.exec(timeText)
  if (!match) return undefined
  let hours = Number(match[1]), minutes = Number(match[2]), seconds = Number(match[3] ?? 0)
  if (minutes > 59 || seconds > 59) return undefined
  const meridiem = match[5]?.toUpperCase()
  if (meridiem) {
    if (hours < 1 || hours > 12) return undefined
    hours = hours % 12 + (meridiem === 'PM' ? 12 : 0)
  } else if (hours > 23) return undefined
  const fractionalSecond = match[4] ? Number(`0.${match[4]}`) : 0
  return (hours * 3600 + minutes * 60 + seconds + fractionalSecond) / 86_400
}

function isoWeekNumber(parts: DateParts): number {
  const date = new Date(Date.UTC(parts.year, parts.month - 1, parts.day))
  const weekday = date.getUTCDay() || 7
  date.setUTCDate(date.getUTCDate() + 4 - weekday)
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1))
  return Math.ceil(((date.getTime() - yearStart.getTime()) / MS_PER_DAY + 1) / 7)
}

