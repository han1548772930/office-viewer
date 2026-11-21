import type { CellScalar, SpreadsheetCell, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import { formulaCellKey, formulaError, isFormulaError } from './formula-types.ts'
import type {
  FormulaCellAddress, FormulaDependencyGraph, FormulaEngineDiagnostic, FormulaErrorValue, FormulaNode,
  FormulaEvaluationOptions, FormulaRangeReference, FormulaRecalculationResult, FormulaReference, FormulaValue,
} from './formula-types.ts'
import { FormulaSyntaxError, parseFormula } from './formula-parser.ts'

interface RangeValue { kind: 'rangeValue'; sheet: SpreadsheetSheet; reference: FormulaRangeReference }
type RuntimeValue = FormulaValue | RangeValue

const ERROR_CODES = new Set(['#DIV/0!', '#N/A', '#NAME?', '#NULL!', '#NUM!', '#REF!', '#VALUE!', '#CIRCULAR!'])

/** Parses references once and builds a sparse range-aware dependency index. */
export function createFormulaDependencyGraph(workbook: WorkbookModel): FormulaDependencyGraph {
  const dependencies = new Map<string, Array<FormulaReference | FormulaRangeReference>>()
  const dependents = new Map<string, Set<string>>()
  const diagnostics: FormulaEngineDiagnostic[] = []
  const sheetByName = new Map(workbook.sheets.map((sheet) => [sheet.name.toLocaleLowerCase(), sheet]))
  for (const sheet of workbook.sheets) for (const cell of Object.values(sheet.cells)) {
    if (!cell.formula) continue
    const owner: FormulaCellAddress = { sheetId: sheet.id, row: cell.row, column: cell.col }
    const ownerKey = formulaCellKey(owner)
    try {
      const node = parseFormula(cell.formula)
      const refs = collectReferences(node)
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
  const active: string[] = []
  const recalculatedCells: FormulaCellAddress[] = []
  const diagnostics = [...graph.diagnostics]
  const cells = new Map<string, { sheet: SpreadsheetSheet; cell: SpreadsheetCell }>()
  for (const sheet of workbook.sheets) for (const cell of Object.values(sheet.cells)) {
    cells.set(formulaCellKey({ sheetId: sheet.id, row: cell.row, column: cell.col }), { sheet, cell })
  }

  const evaluateCell = (sheet: SpreadsheetSheet, row: number, column: number): FormulaValue => {
    const address = { sheetId: sheet.id, row, column }
    const key = formulaCellKey(address)
    const cached = memo.get(key)
    if (cached !== undefined) return cached
    const current = sheet.cells[cellKey(row, column)]
    if (!current) return null
    if (!current.formula) {
      if (current.dataType === 'error' || current.cachedDataType === 'error' || (typeof current.value === 'string' && ERROR_CODES.has(current.value.toUpperCase()))) {
        return formulaError(typeof current.value === 'string' && ERROR_CODES.has(current.value.toUpperCase()) ? current.value.toUpperCase() : current.display || '#VALUE!')
      }
      return current.value
    }
    const cycleStart = active.indexOf(key)
    if (cycleStart >= 0) {
      for (const cycleKey of active.slice(cycleStart)) memo.set(cycleKey, formulaError('#CIRCULAR!'))
      memo.set(key, formulaError('#CIRCULAR!'))
      return formulaError('#CIRCULAR!')
    }
    active.push(key)
    let result: FormulaValue
    const previousCell = previous?.sheets.find((candidate) => candidate.id === sheet.id)?.cells[cellKey(row, column)]
    try {
      result = normalizeResult(evaluateNode(parseFormula(current.formula), sheet, evaluateCell, workbook, sheetsByName, { row, column }, evaluationTime))
    } catch (error) {
      const failure = toFormulaError(error)
      // Keep a cached result for an unchanged imported formula outside our grammar/catalog.
      if (previousCell?.formula === current.formula && hasCachedValue(previousCell)) {
        result = previousCell.value
        diagnostics.push({ code: 'FORMULA_UNSUPPORTED_CACHE_PRESERVED', message: `${current.formula} is not implemented; the imported cached value was retained`, cell: address, formula: current.formula })
      } else {
        result = failure
      }
    }
    if (isFormulaError(result) && result.code === '#NAME?' && previousCell?.formula === current.formula && hasCachedValue(previousCell)) {
      result = previousCell.value
      diagnostics.push({ code: 'FORMULA_UNSUPPORTED_CACHE_PRESERVED', message: `${current.formula} is not implemented; the imported cached value was retained`, cell: address, formula: current.formula })
    }
    if (!memo.has(key)) memo.set(key, result)
    const stored = memo.get(key)!
    if (active.at(-1) === key) active.pop()
    if (stored !== null || current.value !== null || current.display !== '') recalculatedCells.push(address)
    return stored
  }

  for (const { sheet, cell } of cells.values()) if (cell.formula) evaluateCell(sheet, cell.row, cell.col)

  const updatedSheets = workbook.sheets.map((sheet) => {
    let changed = false
    const updatedCells: Record<string, SpreadsheetCell> = { ...sheet.cells }
    for (const cell of Object.values(sheet.cells)) {
      if (!cell.formula) continue
      const value = memo.get(formulaCellKey({ sheetId: sheet.id, row: cell.row, column: cell.col }))
      if (value === undefined) continue
      const scalar: CellScalar = isFormulaError(value) ? value.code : value
      const display = displayFormulaValue(value)
      const errorCode = isFormulaError(value) ? value.code : undefined
      if (cell.value === scalar && cell.display === display && cell.dataType === 'formula' && cell.cachedDataType === valueType(value)
        && cell.metadata?.formulaError === errorCode) continue
      const metadata = { ...cell.metadata }
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
  const evaluationTime = validEvaluationTime(options.now) ?? new Date()
  const sheet = workbook.sheets.find((candidate) => candidate.id === currentSheetId)
  if (!sheet) return formulaError('#REF!')
  const byName = new Map(workbook.sheets.map((item) => [item.name.toLocaleLowerCase(), item]))
  const memo = new Map<string, FormulaValue>(), active: string[] = []
  const resolve = (targetSheet: SpreadsheetSheet, row: number, column: number): FormulaValue => {
    const key = formulaCellKey({ sheetId: targetSheet.id, row, column })
    if (memo.has(key)) return memo.get(key)!
    const cell = targetSheet.cells[cellKey(row, column)]
    if (!cell) return null
    if (!cell.formula) return rawCellValue(cell)
    if (active.includes(key)) return formulaError('#CIRCULAR!')
    active.push(key)
    let value: FormulaValue
    try { value = normalizeResult(evaluateNode(parseFormula(cell.formula), targetSheet, resolve, workbook, byName, { row, column }, evaluationTime)) }
    catch (error) { value = toFormulaError(error) }
    active.pop()
    memo.set(key, value)
    return value
  }
  try { return normalizeResult(evaluateNode(parseFormula(formula), sheet, resolve, workbook, byName, undefined, evaluationTime)) }
  catch (error) { return toFormulaError(error) }
}

export function formulaValueToDisplay(value: FormulaValue): string { return displayFormulaValue(value) }
export function formulaErrorCode(error: FormulaErrorValue): string { return error.code }

function evaluateNode(
  node: FormulaNode,
  currentSheet: SpreadsheetSheet,
  resolveCell: (sheet: SpreadsheetSheet, row: number, column: number) => FormulaValue,
  workbook: WorkbookModel,
  sheetsByName: Map<string, SpreadsheetSheet>,
  formulaCell?: Pick<FormulaCellAddress, 'row' | 'column'>,
  evaluationTime: Date = new Date(),
): RuntimeValue {
  switch (node.kind) {
    case 'literal': return typeof node.value === 'string' && ERROR_CODES.has(node.value.toUpperCase()) ? formulaError(node.value.toUpperCase()) : node.value
    case 'reference': {
      const targetSheet = resolveSheet(workbook, node.reference.sheetName, currentSheet, sheetsByName)
      return targetSheet ? resolveCell(targetSheet, node.reference.row, node.reference.column) : formulaError('#REF!')
    }
    case 'range': {
      const targetSheet = resolveSheet(workbook, node.reference.sheetName, currentSheet, sheetsByName)
      return targetSheet ? { kind: 'rangeValue', sheet: targetSheet, reference: node.reference } : formulaError('#REF!')
    }
    case 'unary': {
      const value = evaluateNode(node.operand, currentSheet, resolveCell, workbook, sheetsByName, formulaCell, evaluationTime)
      if (isFormulaError(value)) return value
      const number = numericCoercion(value)
      if (isFormulaError(number)) return number
      return finiteResult(node.operator === '-' ? -number : number)
    }
    case 'binary': {
      const left = evaluateNode(node.left, currentSheet, resolveCell, workbook, sheetsByName, formulaCell, evaluationTime)
      if (isFormulaError(left)) return left
      const right = evaluateNode(node.right, currentSheet, resolveCell, workbook, sheetsByName, formulaCell, evaluationTime)
      if (isFormulaError(right)) return right
      return evaluateBinary(node.operator, scalarValue(left), scalarValue(right))
    }
    case 'function': return evaluateFunction(node.name, node.arguments, currentSheet, resolveCell, workbook, sheetsByName, formulaCell, evaluationTime)
    case 'error': return formulaError(node.code)
  }
  return formulaError('#VALUE!')
}

function evaluateFunction(
  name: string,
  args: FormulaNode[],
  sheet: SpreadsheetSheet,
  resolve: (sheet: SpreadsheetSheet, row: number, column: number) => FormulaValue,
  workbook: WorkbookModel,
  byName: Map<string, SpreadsheetSheet>,
  formulaCell?: Pick<FormulaCellAddress, 'row' | 'column'>,
  evaluationTime: Date = new Date(),
): RuntimeValue {
  const evaluateArgument = (argument: FormulaNode): RuntimeValue => evaluateNode(argument, sheet, resolve, workbook, byName, formulaCell, evaluationTime)
  if (name === 'ADDRESS') return evaluateAddress(args, evaluateArgument)
  if (name === 'ROW' || name === 'COLUMN') return evaluateReferenceIndex(name, args, evaluateArgument, formulaCell, workbook, sheet, byName)
  if (name === 'AREAS') return evaluateReferenceAreas(args, evaluateArgument)
  if (name === 'TRUE' || name === 'FALSE') return args.length === 0 ? name === 'TRUE' : formulaError('#VALUE!')
  if (name === 'SWITCH') {
    if (args.length < 3 || args.length > 255) return formulaError('#VALUE!')
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
    if (args.length < 2 || args.length > 3) return formulaError('#VALUE!')
    const condition = evaluateArgument(args[0]!)
    if (isFormulaError(condition)) return condition
    const truth = booleanCoercion(condition)
    if (isFormulaError(truth)) return truth
    const selected = truth ? args[1] : args[2]
    return selected ? evaluateArgument(selected) : false
  }
  if (name === 'IFERROR') {
    if (args.length !== 2) return formulaError('#VALUE!')
    const value = evaluateArgument(args[0]!)
    return isFormulaError(value) ? evaluateArgument(args[1]!) : value
  }
  if (name === 'IFNA') {
    if (args.length !== 2) return formulaError('#VALUE!')
    const value = evaluateArgument(args[0]!)
    return isFormulaError(value) && value.code === '#N/A' ? evaluateArgument(args[1]!) : value
  }
  if (name === 'IFS') {
    if (args.length < 2 || args.length % 2 !== 0) return formulaError('#VALUE!')
    for (let index = 0; index < args.length; index += 2) {
      const condition = evaluateArgument(args[index]!)
      if (isFormulaError(condition)) return condition
      const truth = booleanCoercion(condition)
      if (isFormulaError(truth)) return truth
      if (truth) return evaluateArgument(args[index + 1]!)
    }
    return formulaError('#N/A')
  }

  const evaluated = args.map(evaluateArgument)
  if (name === 'COUNTBLANK' || name === 'COUNTA') return evaluateCountValue(name, evaluated, args, sheet, resolve)
  if (['COUNTIF', 'COUNTIFS', 'SUMIF', 'SUMIFS', 'AVERAGEIF', 'AVERAGEIFS'].includes(name)) {
    return evaluateConditionalAggregate(name, evaluated, args, sheet, resolve, workbook)
  }
  if (name === 'INDEX') return evaluateIndex(evaluated, args.length, resolve)
  if (name === 'MATCH') return evaluateMatch(evaluated, args.length, resolve)
  if (name === 'VLOOKUP') return evaluateVlookup(evaluated, args.length, resolve)
  if (name === 'HLOOKUP') return evaluateHlookup(evaluated, args.length, resolve)
  if (name === 'CHOOSE') return evaluateChoose(evaluated)
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
          const item = cell.formula ? resolve(value.sheet, cell.row, cell.col) : rawCellValue(cell)
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
    if (evaluated.length === 0) return formulaError('#VALUE!')
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
  if (name === 'NOT') return args.length === 1 ? invertBoolean(evaluated[0]!) : formulaError('#VALUE!')
  if (name === 'DATE') {
    if (evaluated.length !== 3) return formulaError('#VALUE!')
    const parts = evaluated.map(numericCoercion)
    const error = parts.find(isFormulaError)
    if (error) return error
    const [rawYear, month, day] = parts as number[]
    if (![rawYear, month, day].every(Number.isFinite) || ![rawYear, month, day].every(Number.isInteger)) return formulaError('#VALUE!')
    let year = rawYear
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
    if (evaluated.length !== 0) return formulaError('#VALUE!')
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
    const serial = numericCoercion(evaluated[0]!)
    if (isFormulaError(serial)) return serial
    const parts = serialToDateParts(serial, workbook.date1904)
    if (!parts) return formulaError('#NUM!')
    if (name === 'YEAR') return parts.year
    if (name === 'MONTH') return parts.month
    if (name === 'DAY') return parts.day
    const returnTypeValue = evaluated.length > 1 ? numericCoercion(evaluated[1]!) : 1
    if (isFormulaError(returnTypeValue) || ![1, 2, 3].includes(returnTypeValue)) return formulaError('#NUM!')
    const weekday = excelWeekday(parts, serial, workbook.date1904)
    return returnTypeValue === 1 ? weekday : returnTypeValue === 2 ? (weekday + 5) % 7 + 1 : (weekday + 5) % 7
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
  if (name === 'ABS' || name === 'SQRT' || name === 'INT' || name === 'ROUND' || name === 'MOD' || name === 'POWER') {
    const numbers = evaluated.map(numericCoercion)
    const error = numbers.find(isFormulaError)
    if (error) return error
    const values = numbers as number[]
    if ((name === 'ABS' || name === 'SQRT' || name === 'INT') && values.length !== 1) return formulaError('#VALUE!')
    if (name === 'ABS') return Math.abs(values[0]!)
    if (name === 'SQRT') return values[0]! < 0 ? formulaError('#NUM!') : Math.sqrt(values[0]!)
    if (name === 'INT') return Math.floor(values[0]!)
    if (name === 'ROUND') {
      if (values.length !== 2 || Math.abs(values[1]!) > 308) return formulaError('#NUM!')
      const scale = 10 ** values[1]!
      return finiteResult(Math.sign(values[0]!) * Math.round(Math.abs(values[0]!) * scale) / scale)
    }
    if (name === 'MOD') return values.length === 2 ? values[1] === 0 ? formulaError('#DIV/0!') : values[0]! - values[1]! * Math.floor(values[0]! / values[1]!) : formulaError('#VALUE!')
    if (name === 'POWER') return values.length === 2 ? finiteResult(values[0]! ** values[1]!) : formulaError('#VALUE!')
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
    const text = scalarToText(scalarValue(evaluated[0]!))
    if (name === 'LEN') return text.length
    if (name === 'UPPER') return text.toLocaleUpperCase()
    if (name === 'LOWER') return text.toLocaleLowerCase()
    if (name === 'TRIM') return text.trim().replace(/\s+/g, ' ')
    if (name === 'VALUE') { const number = Number(text); return Number.isFinite(number) ? number : formulaError('#VALUE!') }
    const lengthArg = args.length > 1 ? numericCoercion(evaluated[1]!) : 1
    if (isFormulaError(lengthArg) || !Number.isInteger(lengthArg) || lengthArg < 0) return formulaError('#VALUE!')
    if (name === 'LEFT') return Array.from(text).slice(0, lengthArg).join('')
    if (name === 'RIGHT') return Array.from(text).slice(-lengthArg || undefined).join('')
    const start = numericCoercion(evaluated[1]!)
    const count = numericCoercion(evaluated[2]!)
    if (isFormulaError(start) || isFormulaError(count) || start < 1 || !Number.isInteger(start) || count < 0 || !Number.isInteger(count)) return formulaError('#VALUE!')
    return Array.from(text).slice(start - 1, start - 1 + count).join('')
  }
  if (name === 'ISNUMBER' || name === 'ISTEXT' || name === 'ISBLANK' || name === 'ISERROR') {
    if (args.length !== 1) return formulaError('#VALUE!')
    const value = evaluated[0]!
    if (name === 'ISERROR') return isFormulaError(value)
    if (isFormulaError(value)) return false
    if (name === 'ISNUMBER') return typeof value === 'number'
    if (name === 'ISTEXT') return typeof value === 'string' && !ERROR_CODES.has(value.toUpperCase())
    return value === null || value === ''
  }
  if (name === 'NA' && args.length === 0) return formulaError('#N/A')
  return formulaError('#NAME?')
}

function evaluateIndex(values: RuntimeValue[], argumentCount: number, resolve: (sheet: SpreadsheetSheet, row: number, column: number) => FormulaValue): RuntimeValue {
  if ((argumentCount !== 2 && argumentCount !== 3) || !isRangeValue(values[0]!)) return formulaError('#VALUE!')
  const source = values[0] as RangeValue
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
  if (rowNumber === 0 || columnNumber === 0) {
    return {
      ...source,
      reference: {
        ...source.reference,
        start: { ...source.reference.start, row: rowNumber ? minRow + rowNumber - 1 : minRow, column: columnNumber ? minColumn + columnNumber - 1 : minColumn },
        end: { ...source.reference.end, row: rowNumber ? minRow + rowNumber - 1 : maxRow, column: columnNumber ? minColumn + columnNumber - 1 : maxColumn },
      },
    }
  }
  const cell = source.sheet.cells[cellKey(minRow + rowNumber - 1, minColumn + columnNumber - 1)]
  return cell ? cell.formula ? resolve(source.sheet, minRow + rowNumber - 1, minColumn + columnNumber - 1) : rawCellValue(cell) : null
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
    const value = cell ? cell.formula ? resolve(range.sheet, row, column) : rawCellValue(cell) : null
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
    const value = first ? first.formula ? resolve(range.sheet, row, minColumn) : rawCellValue(first) : null
    if (isFormulaError(value)) { if (!approximate) return value; continue }
    if (compare(value, lookup, '=')) { bestRow = row; break }
    if (approximate && compare(value, lookup, '<')) bestRow = row
    else if (approximate && compare(value, lookup, '>')) break
  }
  if (bestRow < 0) return formulaError('#N/A')
  const resultCell = range.sheet.cells[cellKey(bestRow, minColumn + columnNumber - 1)]
  return resultCell ? resultCell.formula ? resolve(range.sheet, bestRow, minColumn + columnNumber - 1) : rawCellValue(resultCell) : null
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
    const value = cell ? cell.formula ? resolve(range.sheet, minRow, column) : rawCellValue(cell) : null
    if (isFormulaError(value)) { if (!approximate) return value; continue }
    if (compare(value, lookup, '=')) { bestColumn = column; break }
    if (approximate && compare(value, lookup, '<')) bestColumn = column
    else if (approximate && compare(value, lookup, '>')) break
  }
  if (bestColumn < 0) return formulaError('#N/A')
  const resultCell = range.sheet.cells[cellKey(minRow + rowNumber - 1, bestColumn)]
  return resultCell ? resultCell.formula ? resolve(range.sheet, minRow + rowNumber - 1, bestColumn) : rawCellValue(resultCell) : null
}

function evaluateChoose(values: RuntimeValue[]): RuntimeValue {
  if (values.length < 2 || values.length > 255) return formulaError('#VALUE!')
  const index = numericCoercion(values[0]!)
  if (isFormulaError(index)) return index
  if (!Number.isInteger(index) || index < 1 || index >= values.length) return formulaError('#VALUE!')
  return values[index]!
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
  const inputs = isRangeValue(value) ? cellsInRange(value).map(cell => cell.formula ? resolve(value.sheet, cell.row, cell.col) : rawCellValue(cell)) : [value]
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
      cells.push({ row, column, value: cell ? cell.formula ? resolve(value.sheet, row, column) : rawCellValue(cell) : null })
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
    return { row, column, value: source ? source.formula ? resolve(value.sheet, row, column) : rawCellValue(source) : null }
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
  switch (node.kind) {
    case 'reference': return [node.reference]
    case 'range': return [node.reference]
    case 'unary': return collectReferences(node.operand)
    case 'binary': return [...collectReferences(node.left), ...collectReferences(node.right)]
    case 'function': return node.arguments.flatMap(collectReferences)
    case 'literal': return []
    case 'error': return []
  }
  return []
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

function rawCellValue(cell: SpreadsheetCell): FormulaValue {
  if (cell.dataType === 'error' || cell.cachedDataType === 'error' || (typeof cell.value === 'string' && ERROR_CODES.has(cell.value.toUpperCase()))) {
    return formulaError(typeof cell.value === 'string' && ERROR_CODES.has(cell.value.toUpperCase()) ? cell.value.toUpperCase() : cell.display || '#VALUE!')
  }
  return cell.value
}

function flattenScalar(value: RuntimeValue, _sheet: SpreadsheetSheet, resolve: (sheet: SpreadsheetSheet, row: number, column: number) => FormulaValue): FormulaValue[] {
  if (!isRangeValue(value)) return [value]
  return cellsInRange(value).map((cell) => cell.formula ? resolve(value.sheet, cell.row, cell.col) : rawCellValue(cell))
}

function flattenText(values: RuntimeValue[], sheet: SpreadsheetSheet, resolve: (sheet: SpreadsheetSheet, row: number, column: number) => FormulaValue): string | FormulaErrorValue {
  const parts: string[] = []
  for (const value of values) for (const item of flattenScalar(value, sheet, resolve)) {
    if (isFormulaError(item)) return item
    parts.push(scalarToText(item))
  }
  return parts.join('')
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
  const cell = value.sheet.cells[cellKey(start.row, start.column)]
  return cell ? cell.formula ? resolve(value.sheet, cell.row, cell.col) : rawCellValue(cell) : null
}
function normalizeResult(value: RuntimeValue): FormulaValue { return isRangeValue(value) ? formulaError('#VALUE!') : value }
function isRangeValue(value: RuntimeValue): value is RangeValue { return Boolean(value && typeof value === 'object' && 'kind' in value && value.kind === 'rangeValue') }
function finiteResult(value: number): FormulaValue { return Number.isFinite(value) ? value : formulaError('#NUM!') }
function displayFormulaValue(value: FormulaValue): string { return isFormulaError(value) ? value.code : typeof value === 'boolean' ? (value ? 'TRUE' : 'FALSE') : value === null ? '' : String(value) }
function valueType(value: FormulaValue): string { return isFormulaError(value) ? 'error' : value === null ? 'empty' : typeof value }
function hasCachedValue(cell: SpreadsheetCell): boolean { return cell.value !== null && cell.value !== '' && cell.display !== '' }
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
  if (!Number.isInteger(year) || !Number.isInteger(month) || !Number.isInteger(day) || month < 1 || month > 12 || day < 1 || day > 31) return formulaError('#NUM!')
  if (date1904) {
    const value = Date.UTC(year, month - 1, day)
    const epoch = Date.UTC(1904, 0, 1)
    const normalized = new Date(value)
    if (normalized.getUTCFullYear() !== year || normalized.getUTCMonth() !== month - 1 || normalized.getUTCDate() !== day) return formulaError('#NUM!')
    return (value - epoch) / MS_PER_DAY
  }
  if (year === 1900 && month === 2 && day === 29) return 60
  const value = Date.UTC(year, month - 1, day)
  const epoch = Date.UTC(1899, 11, 31)
  const normalized = new Date(value)
  if (normalized.getUTCFullYear() !== year || normalized.getUTCMonth() !== month - 1 || normalized.getUTCDate() !== day) return formulaError('#NUM!')
  const days = (value - epoch) / MS_PER_DAY
  return days >= 60 ? days + 1 : days
}

function dateArgumentsToSerial(year: number, month: number, day: number, date1904: boolean): FormulaValue {
  const normalizedYear = year + Math.floor((month - 1) / 12)
  const normalizedMonth = ((month - 1) % 12 + 12) % 12
  if (normalizedYear < 0 || normalizedYear > 9999) return formulaError('#NUM!')
  const normalizedMonthNumber = normalizedMonth + 1
  const normalizedDay = Math.trunc(day)
  if (!date1904 && normalizedYear === 1900 && normalizedMonthNumber === 2 && normalizedDay === 29) return 60
  const timestamp = Date.UTC(normalizedYear, normalizedMonth, normalizedDay)
  const epoch = Date.UTC(date1904 ? 1904 : 1899, date1904 ? 0 : 11, date1904 ? 1 : 31)
  const serial = (timestamp - epoch) / MS_PER_DAY
  return Number.isFinite(serial) ? (!date1904 && serial >= 60 ? serial + 1 : serial) : formulaError('#NUM!')
}

function serialToDateParts(serial: number, date1904: boolean): DateParts | undefined {
  if (!Number.isFinite(serial) || serial < 0) return undefined
  if (!date1904 && serial === 60) return { year: 1900, month: 2, day: 29 }
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31)
  const days = date1904 || serial < 60 ? serial : serial - 1
  const date = new Date(epoch + Math.floor(days) * MS_PER_DAY)
  if (!Number.isFinite(date.getTime())) return undefined
  return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate() }
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

function excelWeekday(parts: DateParts, serial: number, date1904: boolean): number {
  if (!date1904) return ((Math.floor(serial) % 7) + 7) % 7 + 1
  return new Date(Date.UTC(parts.year, parts.month - 1, parts.day)).getUTCDay() + 1
}

function isoWeekNumber(parts: DateParts): number {
  const date = new Date(Date.UTC(parts.year, parts.month - 1, parts.day))
  const weekday = date.getUTCDay() || 7
  date.setUTCDate(date.getUTCDate() + 4 - weekday)
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1))
  return Math.ceil(((date.getTime() - yearStart.getTime()) / MS_PER_DAY + 1) / 7)
}

