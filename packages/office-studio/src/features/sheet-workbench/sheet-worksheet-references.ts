import type { SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import { mapConditionalFormatFormulas } from '../conditional-formatting/conditional-format-formulas.ts'
import { mapFormulaReferenceAreas, mapFormulaSheetQualifiers } from '../formula/formula-reference-areas.ts'
import { normalizeSheetHyperlinkTarget, parseSheetHyperlinkTarget } from './sheet-hyperlink.ts'
import { renamePivotSourceWorksheet } from '../pivot/services/pivot-source-lifecycle.ts'

export interface WorksheetReferenceChange {
  sheetId: string
  oldName: string
  newName?: string
}

const equalName = (first: string, second: string): boolean => first.toLowerCase() === second.toLowerCase()
const quote = (name: string): string => `'${name.replace(/'/g, "''")}'!`

function changedQualifier(name: string, original: WorkbookModel, change: WorksheetReferenceChange): string | null {
  const endpoints = name.split(':')
  if (change.newName !== undefined) return endpoints.map(endpoint => equalName(endpoint, change.oldName) ? change.newName! : endpoint).join(':')
  if (endpoints.length === 1) return equalName(name, change.oldName) ? null : name
  if (endpoints.length !== 2 || !endpoints.some(endpoint => equalName(endpoint, change.oldName))) return name
  const first = original.sheets.findIndex(sheet => equalName(sheet.name, endpoints[0]!))
  const last = original.sheets.findIndex(sheet => equalName(sheet.name, endpoints[1]!))
  if (first < 0 || last < 0 || first === last) return null
  const direction = first < last ? 1 : -1
  const start = equalName(endpoints[0]!, change.oldName) ? first + direction : first
  const end = equalName(endpoints[1]!, change.oldName) ? last - direction : last
  if (direction * (end - start) < 0) return null
  const startName = original.sheets[start]!.name, endName = original.sheets[end]!.name
  return start === end ? startName : `${startName}:${endName}`
}

/** Reference areas and qualified names share the formula engine's lexical rules. */
export function rewriteWorksheetFormulaReferences(formula: string, original: WorkbookModel, change: WorksheetReferenceChange): string {
  const areas = mapFormulaReferenceAreas(formula, area => {
    const start = area.sheetName === undefined || area.qualifier.includes('[') ? undefined : changedQualifier(area.sheetName, original, change)
    const end = area.endSheetName === undefined || area.endQualifier.includes('[') ? undefined : changedQualifier(area.endSheetName, original, change)
    if (start === null || end === null) return null
    return { ...area,
      ...(start !== undefined && start !== area.sheetName ? { qualifier: quote(start), sheetName: start, threeDimensional: start.includes(':') } : {}),
      ...(end !== undefined && end !== area.endSheetName ? { endQualifier: quote(end), endSheetName: end } : {}),
    }
  }, { invalidReference: '#REF!' })
  return mapFormulaSheetQualifiers(areas, qualifier => qualifier.external ? qualifier.sheetName
    : changedQualifier(qualifier.sheetName, original, change))
}

/** A removed table invalidates its entire structured reference, without touching strings or external tables. */
export function invalidateRemovedTableReferences(formula: string, names: ReadonlySet<string>): string {
  if (!names.size) return formula
  let result = '', cursor = 0, index = 0
  while (index < formula.length) {
    if (formula[index] === '"') {
      index++
      while (index < formula.length) {
        if (formula[index] !== '"') { index++; continue }
        if (formula[index + 1] === '"') { index += 2; continue }
        index++; break
      }
      continue
    }
    const identifier = /^[\p{L}_\\][\p{L}\p{N}_.\\]*/u.exec(formula.slice(index))
    if (identifier && !/[\p{L}\p{N}_.\\]/u.test(formula[index - 1] ?? '')) {
      const end = index + identifier[0].length
      if (formula[end] === '[') {
        let close = end + 1, depth = 1
        while (close < formula.length && depth) {
          if (formula[close] === "'" && /[\[\]#@']/u.test(formula[close + 1] ?? '')) { close += 2; continue }
          if (formula[close] === '[') depth++
          else if (formula[close] === ']') depth--
          close++
        }
        const previous = formula.slice(0, index).trimEnd().at(-1)
        if (!depth && previous !== '!' && previous !== ']' && names.has(identifier[0].toLowerCase())) {
          result += formula.slice(cursor, index) + '#REF!'; cursor = close
        }
        index = close; continue
      }
      index = end; continue
    }
    if (formula[index] === '[' || formula[index] === "'") {
      const bracket = formula[index] === '['
      index++
      while (index < formula.length) {
        if (bracket && formula[index] === ']') { index++; break }
        if (!bracket && formula[index] === "'") {
          if (formula[index + 1] === "'") { index += 2; continue }
          index++; break
        }
        index++
      }
    } else index++
  }
  return result + formula.slice(cursor)
}

/** Positional localSheetId is rebound by stable sheet ID when sheets are inserted, removed or reordered. */
export function rebindWorksheetScopedNames(metadata: Record<string, unknown> | undefined, before: readonly SpreadsheetSheet[],
  after: readonly SpreadsheetSheet[], rewrite: (formula: string) => string = formula => formula,
  change?: WorksheetReferenceChange): void {
  if (!metadata) return
  for (const key of ['definedNames', 'localDefinedNames']) {
    const entries = metadata[key]
    if (!Array.isArray(entries)) continue
    metadata[key] = entries.flatMap((entry: unknown) => {
      if (!entry || typeof entry !== 'object' || Array.isArray(entry)) return [entry]
      const record = entry as Record<string, unknown>
      const indexedSheet = Number.isSafeInteger(record.localSheetId) ? before[Number(record.localSheetId)] : undefined
      const owns = change && (record.sheetId === change.sheetId || record.subUnitId === change.sheetId
        || indexedSheet?.id === change.sheetId || typeof record.sheet === 'string' && equalName(record.sheet, change.oldName)
        || typeof record.sheetName === 'string' && equalName(record.sheetName, change.oldName))
      if (owns && change.newName === undefined) return []
      return [{ ...record,
        ...(indexedSheet ? { localSheetId: after.findIndex(sheet => sheet.id === indexedSheet.id) } : {}),
        ...(owns && typeof record.sheet === 'string' ? { sheet: change.newName } : {}),
        ...(owns && typeof record.sheetName === 'string' ? { sheetName: change.newName } : {}),
        ...(typeof record.refersTo === 'string' ? { refersTo: rewrite(record.refersTo) } : {}),
      }]
    })
  }
}

/** Updates all modeled reference owners on a detached workbook candidate. */
export function rewriteWorkbookWorksheetReferences(candidate: WorkbookModel, original: WorkbookModel, change: WorksheetReferenceChange,
  removedTableNames: ReadonlySet<string> = new Set()): void {
  const rewrite = (formula: string) => invalidateRemovedTableReferences(rewriteWorksheetFormulaReferences(formula, original, change), removedTableNames)
  const rewriteLink = (link: string): string | undefined => {
    const target = parseSheetHyperlinkTarget(link, original)
    if (target?.type !== 'internal' || target.sheetId !== change.sheetId) return link
    if (change.newName === undefined) return undefined
    return normalizeSheetHyperlinkTarget(link, original) ?? link
  }
  for (const sheet of candidate.sheets) {
    for (const cell of Object.values(sheet.cells)) {
      if (cell.formula) cell.formula = rewrite(cell.formula)
      if (cell.sparkline) cell.sparkline.sourceRange = rewrite(cell.sparkline.sourceRange)
      if (cell.hyperlink) cell.hyperlink = rewriteLink(cell.hyperlink) ?? null
      if (cell.richText) cell.richText = cell.richText.map(run => {
        if (!run.hyperlink) return run
        const hyperlink = rewriteLink(run.hyperlink)
        if (hyperlink !== undefined) return { ...run, hyperlink }
        const { hyperlink: _removed, ...plain } = run
        return plain
      })
    }
    sheet.dataValidations = sheet.dataValidations?.map(rule => ({ ...rule,
      ...(rule.formula1 && (!(rule.type === 'list' || rule.type === 'listMultiple') || rule.formula1.startsWith('=')) ? { formula1: rewrite(rule.formula1) } : {}),
      ...(rule.formula2 ? { formula2: rewrite(rule.formula2) } : {}),
    }))
    sheet.conditionalFormats = sheet.conditionalFormats?.map(rule => mapConditionalFormatFormulas(rule, rewrite))
    rebindWorksheetScopedNames(sheet.metadata, original.sheets, candidate.sheets, rewrite, change)
  }
  rebindWorksheetScopedNames(candidate.metadata, original.sheets, candidate.sheets, rewrite, change)
  if (change.newName !== undefined) renamePivotSourceWorksheet(original, candidate, change.sheetId, change.newName)
}
