import { parseFormula } from './formula-parser'
import { readFormulaSheetQualifier, mapFormulaSheetQualifiers } from './formula-reference-areas'
import { mapFormulaIdentifiers } from './formula-identifiers'
import type { FormulaImportAliases } from './formula-import-context'
import { formulaRangeReference } from './formula-reference-binding'
import type { FormulaCellAddress, FormulaNode } from './formula-types'

export function remapDynamicFormulaReference(text: string, aliases: FormulaImportAliases | undefined, a1 = true): string {
  if (!aliases) return text
  const mappedNames = new Set(Object.values(aliases.names).map(name => name.toLowerCase()))
  const mappedTables = new Set(Object.values(aliases.tables).map(name => name.toLowerCase()))
  const qualifier = readFormulaSheetQualifier(text, 0), body = text.slice(qualifier?.end ?? 0)
  const r1c1 = !a1 && (/^R(?:\d+|\[[+-]?\d+\])?C(?:\d+|\[[+-]?\d+\])?(?:\s*:\s*R(?:\d+|\[[+-]?\d+\])?C(?:\d+|\[[+-]?\d+\])?)?$/i.test(body)
    || /^(R|C)(?:\d+|\[[+-]?\d+\])?\s*:\s*\1(?:\d+|\[[+-]?\d+\])?$/i.test(body))
  let unresolved = false
  const names = r1c1 ? text : mapFormulaIdentifiers(text, identifier => {
    if (identifier.external || identifier.kind === 'function') return undefined
    const key = identifier.name.toLowerCase()
    if (identifier.kind === 'table') {
      const mapped = aliases.tables[key] ?? (mappedTables.has(key) ? identifier.name : undefined)
      unresolved ||= mapped === undefined; return mapped
    }
    const scopeKey = identifier.sheetName ? `${identifier.sheetName.toLowerCase()}!${key}` : key
    const scoped = aliases.names[scopeKey]
    const mapped = scoped ?? aliases.names[key] ?? aliases.tables[key]
      ?? (mappedNames.has(key) || mappedTables.has(key) ? identifier.name : undefined)
    unresolved ||= mapped === undefined
    const sheetName = aliases.nameSheets?.[scopeKey] ?? aliases.nameSheets?.[key]
    return mapped && sheetName ? { name: mapped, sheetName } : mapped
  })
  if (unresolved) return '#REF!'
  const mappedSheets = new Set(Object.values(aliases.sheets).map(name => name.toLowerCase()))
  return mapFormulaSheetQualifiers(names, qualifier => qualifier.external ? qualifier.sheetName
    : aliases.sheets[qualifier.sheetName.toLowerCase()] ?? (mappedSheets.has(qualifier.sheetName.toLowerCase()) ? qualifier.sheetName : null))
}
function coordinate(text: string, base: number | undefined, maximum: number): number | undefined {
  let result: number
  if (!text) { if (base === undefined) return undefined; result = base }
  else if (text[0] === '[') { if (base === undefined) return undefined; result = base + Number(text.slice(1, -1)) }
  else result = Number(text) - 1
  return Number.isSafeInteger(result) && result >= 0 && result <= maximum ? result : undefined
}

/** Resolve authored A1 and absolute/relative R1C1 text through the same AST
 * binder. Arbitrary expressions in a ref_text value are never executed. */
export function parseDynamicFormulaReference(text: string, a1: boolean, cell?: Pick<FormulaCellAddress, 'row' | 'column'>): FormulaNode {
  const error = /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|CIRCULAR!|CALC!)$/i.exec(text)
  if (error) return { kind: 'error', code: error[0].toUpperCase() }
  if (a1) {
    const node = parseFormula(text)
    return ['reference', 'range', 'name', 'structured', 'error'].includes(node.kind) ? node : { kind: 'error', code: '#REF!' }
  }
  const qualifier = readFormulaSheetQualifier(text, 0)
  if (qualifier?.external || qualifier?.threeDimensional) return { kind: 'error', code: '#REF!' }
  const body = text.slice(qualifier?.end ?? 0).trim(), sheetName = qualifier?.sheetName
  const endpoint = (part: string): { row: number; column: number } | undefined => {
    const match = /^R(\d+|\[[+-]?\d+\])?C(\d+|\[[+-]?\d+\])?$/i.exec(part)
    if (!match) return undefined
    const row = coordinate(match[1] ?? '', cell?.row, 1_048_575), column = coordinate(match[2] ?? '', cell?.column, 16_383)
    return row === undefined || column === undefined ? undefined : { row, column }
  }
  const parts = body.split(/\s*:\s*/), first = endpoint(parts[0] ?? ''), last = endpoint(parts[1] ?? parts[0] ?? '')
  if (parts.length <= 2 && first && last) {
    const reference = formulaRangeReference({ startRow: Math.min(first.row, last.row), endRow: Math.max(first.row, last.row),
      startCol: Math.min(first.column, last.column), endCol: Math.max(first.column, last.column) }, sheetName ?? '')
    if (!sheetName) { delete reference.sheetName; delete reference.start.sheetName; delete reference.end.sheetName }
    return parts.length === 1 ? { kind: 'reference', reference: reference.start } : { kind: 'range', reference }
  }
  const axis = /^(R|C)(\d+|\[[+-]?\d+\])?\s*:\s*\1(\d+|\[[+-]?\d+\])?$/i.exec(body)
  if (axis) {
    const row = axis[1]!.toUpperCase() === 'R', start = coordinate(axis[2] ?? '', row ? cell?.row : cell?.column, row ? 1_048_575 : 16_383),
      end = coordinate(axis[3] ?? '', row ? cell?.row : cell?.column, row ? 1_048_575 : 16_383)
    if (start !== undefined && end !== undefined) {
      const reference = formulaRangeReference({ startRow: row ? Math.min(start, end) : 0, endRow: row ? Math.max(start, end) : 1_048_575,
        startCol: row ? 0 : Math.min(start, end), endCol: row ? 16_383 : Math.max(start, end) }, sheetName ?? '')
      if (!sheetName) { delete reference.sheetName; delete reference.start.sheetName; delete reference.end.sheetName }
      return { kind: 'range', reference }
    }
  }
  // Named ranges are independent of reference notation.
  try { const node = parseFormula(text); if (node.kind === 'name' || node.kind === 'structured') return node } catch { /* invalid R1C1 text */ }
  return { kind: 'error', code: '#REF!' }
}
