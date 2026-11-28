import { readFormulaReferenceArea, readFormulaSheetQualifier } from './formula-reference-areas'
import { readFormulaStructuredSpecifier } from './formula-structured-references'

export interface FormulaIdentifier {
  name: string
  kind: 'name' | 'table' | 'function'
  sheetName?: string
  external: boolean
  threeDimensional: boolean
}
export interface FormulaIdentifierReplacement { name: string; sheetName: string }
interface IdentifierSpan { start: number; end: number; qualifierStart: number; identifier: FormulaIdentifier }
const identifierCharacter = (value: string | undefined): boolean => value !== undefined && /[\p{L}\p{M}\p{N}_.\\]/u.test(value)

/** A lexical transform, not a second parser/evaluator. Strings, A1 areas and
 * structured column captions retain their exact authored text. */
function identifierSpans(formula: string): IdentifierSpan[] {
  const spans: IdentifierSpan[] = []
  let index = 0
  while (index < formula.length) {
    if (formula[index] === '#') {
      const error = /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|SPILL!|CALC!|CONNECT!|BLOCKED!|UNKNOWN!|FIELD!|BUSY!)/i.exec(formula.slice(index))
      if (error) { index += error[0].length; continue }
    }
    if (formula[index] === '"') {
      index++
      while (index < formula.length) {
        if (formula[index++] !== '"') continue
        if (formula[index] === '"') { index++; continue }
        break
      }
      continue
    }
    if (identifierCharacter(formula[index - 1])) { index++; continue }
    const area = readFormulaReferenceArea(formula, index)
    if (area) { index = area.end; continue }
    const qualifier = readFormulaSheetQualifier(formula, index)
    if (!qualifier && formula[index] === '[') {
      index = readFormulaStructuredSpecifier(formula, index)?.end ?? formula.length
      continue
    }
    const start = qualifier?.end ?? index
    const name = /^[\p{L}_\\][\p{L}\p{M}\p{N}_.\\]*/u.exec(formula.slice(start))?.[0]
    if (!name) { index = qualifier?.end ?? index + 1; continue }
    const end = start + name.length, following = formula.slice(end).trimStart()[0]
    spans.push({ start, end, qualifierStart: index, identifier: { name, kind: following === '(' ? 'function' : following === '[' ? 'table' : 'name',
      ...(qualifier ? { sheetName: qualifier.sheetName } : {}), external: qualifier?.external ?? false, threeDimensional: qualifier?.threeDimensional ?? false } })
    index = end
  }
  return spans
}

export function visitFormulaIdentifiers(formula: string, visit: (identifier: Readonly<FormulaIdentifier>) => void): void {
  if (typeof formula !== 'string') throw new TypeError('公式必须是字符串。')
  for (const span of identifierSpans(formula)) visit({ ...span.identifier })
}

export function mapFormulaIdentifiers(formula: string,
  transform: (identifier: Readonly<FormulaIdentifier>) => string | FormulaIdentifierReplacement | undefined): string {
  if (typeof formula !== 'string') throw new TypeError('公式必须是字符串。')
  let result = formula
  for (const span of identifierSpans(formula).reverse()) {
    const replacement = transform({ ...span.identifier })
    if (replacement === undefined || replacement === span.identifier.name) continue
    const name = typeof replacement === 'string' ? replacement : replacement.name
    if (!/^[\p{L}_\\][\p{L}\p{M}\p{N}_.\\]*$/u.test(name)) throw new TypeError('变换后的公式标识无效。')
    const qualified = typeof replacement !== 'string'
    if (qualified && (typeof replacement.sheetName !== 'string' || !replacement.sheetName.trim()
      || /[\r\n]/.test(replacement.sheetName))) throw new TypeError('变换后的公式工作表无效。')
    result = result.slice(0, qualified ? span.qualifierStart : span.start)
      + (qualified ? `'${replacement.sheetName.replace(/'/g, "''")}'!` : '') + name + result.slice(span.end)
  }
  return result
}
