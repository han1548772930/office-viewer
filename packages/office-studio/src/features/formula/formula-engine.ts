import type { CellScalar, SpreadsheetCell, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import { formulaCellKey, formulaError, isFormulaError } from './formula-types.ts'
import type {
  FormulaCellAddress, FormulaDependencyGraph, FormulaEngineDiagnostic, FormulaErrorValue, FormulaNode,
  FormulaRangeReference, FormulaRecalculationResult, FormulaReference, FormulaValue,
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
export function recalculateWorkbookFormulas(workbook: WorkbookModel, previous?: WorkbookModel): FormulaRecalculationResult {
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
      result = normalizeResult(evaluateNode(parseFormula(current.formula), sheet, evaluateCell, workbook, sheetsByName))
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

export function evaluateFormula(formula: string, workbook: WorkbookModel, currentSheetId: string): FormulaValue {
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
    try { value = normalizeResult(evaluateNode(parseFormula(cell.formula), targetSheet, resolve, workbook, byName)) }
    catch (error) { value = toFormulaError(error) }
    active.pop()
    memo.set(key, value)
    return value
  }
  try { return normalizeResult(evaluateNode(parseFormula(formula), sheet, resolve, workbook, byName)) }
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
      const value = evaluateNode(node.operand, currentSheet, resolveCell, workbook, sheetsByName)
      if (isFormulaError(value)) return value
      const number = numericCoercion(value)
      if (isFormulaError(number)) return number
      return finiteResult(node.operator === '-' ? -number : number)
    }
    case 'binary': {
      const left = evaluateNode(node.left, currentSheet, resolveCell, workbook, sheetsByName)
      if (isFormulaError(left)) return left
      const right = evaluateNode(node.right, currentSheet, resolveCell, workbook, sheetsByName)
      if (isFormulaError(right)) return right
      return evaluateBinary(node.operator, scalarValue(left), scalarValue(right))
    }
    case 'function': return evaluateFunction(node.name, node.arguments, currentSheet, resolveCell, workbook, sheetsByName)
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
): RuntimeValue {
  if (name === 'IF') {
    if (args.length < 2 || args.length > 3) return formulaError('#VALUE!')
    const condition = evaluateNode(args[0]!, sheet, resolve, workbook, byName)
    if (isFormulaError(condition)) return condition
    const truth = booleanCoercion(condition)
    if (isFormulaError(truth)) return truth
    const selected = truth ? args[1] : args[2]
    return selected ? evaluateNode(selected, sheet, resolve, workbook, byName) : false
  }
  if (name === 'IFERROR') {
    if (args.length !== 2) return formulaError('#VALUE!')
    const value = evaluateNode(args[0]!, sheet, resolve, workbook, byName)
    return isFormulaError(value) ? evaluateNode(args[1]!, sheet, resolve, workbook, byName) : value
  }

  const evaluated = args.map((argument) => evaluateNode(argument, sheet, resolve, workbook, byName))
  if (name === 'SUM' || name === 'AVERAGE' || name === 'COUNT' || name === 'COUNTA' || name === 'MIN' || name === 'MAX') {
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
    if (name === 'AVERAGE') return numbers.length ? finiteResult(numbers.reduce((sum, value) => sum + value, 0) / numbers.length) : formulaError('#DIV/0!')
    if (!numbers.length) return 0
    return numbers.reduce((selected, value) => name === 'MIN' ? Math.min(selected, value) : Math.max(selected, value), numbers[0]!)
  }
  for (const value of evaluated) if (isFormulaError(value)) return value
  if (name === 'AND' || name === 'OR') {
    if (evaluated.length === 0) return formulaError('#VALUE!')
    const booleans: boolean[] = []
    for (const value of evaluated) {
      for (const item of flattenScalar(value, sheet, resolve)) {
        const truth = booleanCoercion(item)
        if (isFormulaError(truth)) return truth
        booleans.push(truth)
      }
    }
    return name === 'AND' ? booleans.every(Boolean) : booleans.some(Boolean)
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
  if (name === 'LEN' || name === 'UPPER' || name === 'LOWER' || name === 'TRIM' || name === 'VALUE' || name === 'LEFT' || name === 'RIGHT' || name === 'MID' || name === 'CONCAT' || name === 'CONCATENATE') {
    if (name === 'CONCAT' || name === 'CONCATENATE') return flattenText(evaluated, sheet, resolve)
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

function excelWeekday(parts: DateParts, serial: number, date1904: boolean): number {
  if (!date1904) return ((Math.floor(serial) % 7) + 7) % 7 + 1
  return new Date(Date.UTC(parts.year, parts.month - 1, parts.day)).getUTCDay() + 1
}

