import type { SpreadsheetCell, SpreadsheetSheet, WorkbookModel } from '../../domain/model'
import { classifySheetNumberFormat } from '../sheet-workbench/sheet-number-format'
import { DateSystem } from '../sheet-workbench/number-format-engine/date-system'
import { translateExcelCalendarSerial } from '../sheet-workbench/number-format-engine/serial-date'
import type { FormulaValueResolver } from './formula-types'

export interface FormulaImportSource { version: 1; formula: string; date1904: boolean; dateResult: boolean; aliasSheetId?: string }
export interface FormulaImportAliases {
  sheets: Record<string, string>
  names: Record<string, string>
  tables: Record<string, string>
  /** Global source names are isolated in a hidden destination sheet scope. */
  nameSheets?: Record<string, string>
}
const object = (value: unknown): value is Record<string, unknown> => !!value && typeof value === 'object' && !Array.isArray(value)

export function formulaCalendarDelta(sourceDate1904: boolean, targetDate1904: boolean): number {
  return sourceDate1904 === targetDate1904 ? 0 : sourceDate1904 ? 1462 : -1462
}
export function formulaCalendarSerial(value: number, sourceDate1904: boolean, targetDate1904: boolean): number {
  return translateExcelCalendarSerial(value, sourceDate1904 ? DateSystem.Date1904 : DateSystem.Date1900,
    targetDate1904 ? DateSystem.Date1904 : DateSystem.Date1900)
}
export interface FormulaCalendarValue { date1904: boolean; value: number; targetValue: number }
/** Exact originals disambiguate serial 60 from its physical-date carrier and
 * avoid fraction drift after repeated calendar projections. A changed value
 * or manually changed formula cannot inherit an obsolete original. */
export function readFormulaCalendarValue(cell: SpreadsheetCell, value: number, date1904: boolean): FormulaCalendarValue | undefined {
  const retained = cell.metadata?.chartDependencyDateSource
  if (!object(retained) || typeof retained.date1904 !== 'boolean'
    || retained.targetDate1904 !== undefined && retained.targetDate1904 !== date1904
    || retained.targetFormula !== undefined && retained.targetFormula !== cell.formula) return undefined
  const original = typeof retained.value === 'number' ? retained.value
    : typeof retained.value === 'string' && retained.value.trim() ? Number(retained.value) : Number.NaN
  if (!Number.isFinite(original)) return undefined
  const target = retained.targetValue === undefined ? formulaCalendarSerial(original, retained.date1904, date1904) : Number(retained.targetValue)
  return value === target ? { date1904: retained.date1904, value: original, targetValue: value } : undefined
}
export function formulaCellUsesCalendar(cell: SpreadsheetCell): boolean {
  const format = classifySheetNumberFormat(cell.style.numberFormat)
  const type = (cell.formula ? cell.cachedDataType ?? cell.dataType : cell.dataType)?.toLowerCase()
  return format === 'date' || format === 'datetime' || type === 'date' || type === 'd' || type === 'datetime'
    || object(cell.metadata?.chartDependencyFormulaSource) && cell.metadata.chartDependencyFormulaSource.dateResult === true
    || object(cell.metadata?.chartDependencyDateSource)
}
export function readFormulaImportSource(cell: SpreadsheetCell): FormulaImportSource | undefined {
  const value = cell.metadata?.chartDependencyFormulaSource
  if (!object(value) || value.version !== 1 || value.formula !== cell.formula
    || typeof value.formula !== 'string' || typeof value.date1904 !== 'boolean' || typeof value.dateResult !== 'boolean') return undefined
  return { version: 1, formula: value.formula, date1904: value.date1904, dateResult: value.dateResult,
    ...(typeof value.aliasSheetId === 'string' ? { aliasSheetId: value.aliasSheetId } : {}) }
}

/** Automatic reference rewrites preserve the source calendar only when the
 * marker still owns the old expression. A manually changed expression cannot
 * acquire an obsolete marker through a later rename, move or fill. */
export function rewriteFormulaImportCell<T extends { formula?: string | null; metadata?: Record<string, unknown> | null }>(cell: T, formula: string): void {
  const marker = cell.metadata?.chartDependencyFormulaSource
  if (object(marker) && marker.version === 1 && marker.formula === cell.formula
    && typeof marker.date1904 === 'boolean' && typeof marker.dateResult === 'boolean') {
    cell.metadata = { ...cell.metadata, chartDependencyFormulaSource: { ...marker, formula } }
  }
  const date = cell.metadata?.chartDependencyDateSource
  if (object(date) && date.targetFormula === cell.formula) {
    cell.metadata = { ...cell.metadata, chartDependencyDateSource: { ...date, targetFormula: formula } }
  }
  cell.formula = formula
}

/** Authored INDIRECT strings keep their original aliases while renamed or
 * deleted destination sheets update the alias values. Deleted aliases remain
 * invalid instead of falling through to an unrelated same-name worksheet. */
export function rewriteFormulaImportSheetAliases(sheet: SpreadsheetSheet, oldName: string, newName?: string): void {
  const provenance = sheet.metadata?.chartDependencyImport
  const aliases = readFormulaImportAliases(sheet)
  if (!object(provenance) || !aliases) return
  const changed = Object.entries(aliases.sheets).filter(([, value]) => value.toLowerCase() === oldName.toLowerCase())
  const nameScopes = Object.entries(aliases.nameSheets ?? {}).filter(([, value]) => value.toLowerCase() === oldName.toLowerCase())
  if (!changed.length && !nameScopes.length) return
  for (const [key] of changed) aliases.sheets[key] = newName ?? '#REF!'
  for (const [key] of nameScopes) aliases.nameSheets![key] = newName ?? '#REF!'
  sheet.metadata = { ...sheet.metadata, chartDependencyImport: { ...provenance, formulaAliases: aliases } }
}
export function readFormulaImportAliases(sheet: SpreadsheetSheet): FormulaImportAliases | undefined {
  const provenance = sheet.metadata?.chartDependencyImport
  const value = object(provenance) ? provenance.formulaAliases : undefined
  if (!object(value) || !object(value.sheets) || !object(value.names) || !object(value.tables)) return undefined
  const read = (values: Record<string, unknown>): Record<string, string> => Object.fromEntries(Object.entries(values)
    .filter((entry): entry is [string, string] => typeof entry[1] === 'string'))
  return { sheets: read(value.sheets), names: read(value.names), tables: read(value.tables),
    ...(object(value.nameSheets) ? { nameSheets: read(value.nameSheets) } : {}) }
}
export function composeFormulaImportAliases(current: FormulaImportAliases, previous?: FormulaImportAliases): FormulaImportAliases {
  const result: FormulaImportAliases = { sheets: { ...current.sheets }, names: { ...current.names }, tables: { ...current.tables },
    ...(current.nameSheets || previous?.nameSheets ? { nameSheets: { ...current.nameSheets } } : {}) }
  for (const kind of ['sheets', 'names', 'tables'] as const) for (const [original, value] of Object.entries(previous?.[kind] ?? {})) {
    result[kind][original] = current[kind][value.toLowerCase()] ?? value
  }
  for (const [original, value] of Object.entries(previous?.nameSheets ?? {})) {
    const nextName = previous?.names[original]?.toLowerCase()
    result.nameSheets![original] = nextName && current.nameSheets?.[nextName] || current.sheets[value.toLowerCase()] || value
  }
  return result
}
/** One evaluator runs in the formula's retained calendar. Its resolver converts
 * typed date references from the workbook calendar, and its caller converts a
 * date result back. Numeric quantities and duration-only cells are unchanged. */
export function formulaImportEnvironment(workbook: WorkbookModel, sheet: SpreadsheetSheet, cell: SpreadsheetCell,
  resolve: FormulaValueResolver,
  calculatedDate?: (sheet: SpreadsheetSheet, row: number, column: number) => FormulaCalendarValue | undefined) {
  const source = readFormulaImportSource(cell)
  if (!source) return { workbook, resolve, source: undefined, aliases: undefined }
  const aliasOwner = source.aliasSheetId ? workbook.sheets.find(sheet => sheet.id === source.aliasSheetId) : sheet
  const mappedResolve: FormulaValueResolver = (targetSheet, row, column) => {
      const value = resolve(targetSheet, row, column), targetCell = targetSheet.cells[`${row}:${column}`]
      if (typeof value !== 'number' || !targetCell || !formulaCellUsesCalendar(targetCell)) return value
      const retained = calculatedDate?.(targetSheet, row, column) ?? readFormulaCalendarValue(targetCell, value, workbook.date1904)
      return retained && retained.targetValue === value
        ? formulaCalendarSerial(retained.value, retained.date1904, source.date1904)
        : formulaCalendarSerial(value, workbook.date1904, source.date1904)
  }
  mappedResolve.canReadRange = resolve.canReadRange
  return { workbook: source.date1904 === workbook.date1904 ? workbook : { ...workbook, date1904: source.date1904 }, source,
    aliases: aliasOwner ? readFormulaImportAliases(aliasOwner) : { sheets: {}, names: {}, tables: {} }, resolve: mappedResolve }
}
