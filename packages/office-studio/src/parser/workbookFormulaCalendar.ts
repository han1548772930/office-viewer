import type { SpreadsheetCell, SpreadsheetSheet, WorkbookModel } from '../domain/model.ts'
import { cloneModel } from '../domain/serialization.ts'
import { conditionalFormatSerializedSize, conditionalFormatUtf8Size } from '../features/conditional-formatting/conditional-format-source-budget.ts'
import { formulaCalendarSerial } from '../features/formula/formula-import-context.ts'
import type { FormulaCalendarSourceDto } from './contracts.ts'

const byteLimit = 16 * 1024 * 1024
const cellLimit = 250_000
const object = (value: unknown): value is Record<string, unknown> => value !== null && typeof value === 'object' && !Array.isArray(value)
const invalid = (detail: string): never => { throw new TypeError(`FORMULA_CALENDAR_SOURCE_INVALID: ${detail}`) }
const limit = (detail: string): never => { throw new RangeError(`FORMULA_CALENDAR_SOURCE_LIMIT: ${detail}`) }
const own = (value: object, key: string) => Object.prototype.hasOwnProperty.call(value, key)
const formula = (value: string | null | undefined) => value == null ? null : `=${value.replace(/^=+/, '')}`

function plainObject(value: unknown, detail: string): Record<string, unknown> {
  if (!object(value) || ![Object.prototype, null].includes(Object.getPrototypeOf(value))) invalid(detail)
  return value
}
function string(value: unknown, maxBytes: number, detail: string): string {
  if (typeof value !== 'string' || !value || conditionalFormatUtf8Size(value, maxBytes) > maxBytes) invalid(detail)
  return value
}
function uint(value: unknown, max: number, detail: string, min = 0): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < min || value > max) invalid(detail)
  return value
}
function boolean(value: unknown, detail: string): boolean {
  if (typeof value !== 'boolean') invalid(detail)
  return value
}
function list(value: unknown, max: number, detail: string): unknown[] {
  if (!Array.isArray(value)) invalid(detail)
  if (value.length > max) limit(detail)
  return value
}
function fields(value: Record<string, unknown>, allowed: readonly string[]): void {
  for (const key of Object.keys(value)) if (!allowed.includes(key)) invalid(`unknown field ${key}`)
}
function finiteValue(value: unknown): number | undefined {
  if (typeof value === 'number') return Number.isFinite(value) ? value : undefined
  if (typeof value !== 'string' || !/^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(value)) return undefined
  const result = Number(value)
  return Number.isFinite(result) ? result : undefined
}

/** Stop at the byte budget while inspecting JSON, before enumerating extension
 * values or cloning. Accessors and toJSON cannot replace the admitted source. */
function assertJson(value: unknown): void {
  let bytes = 0
  const ancestors = new Set<object>()
  const add = (count: number) => { bytes += count; if (bytes > byteLimit) limit('source exceeds 16 MiB') }
  const textSize = (value: string) => {
    try { return conditionalFormatSerializedSize(value, byteLimit - bytes) }
    catch (error) { if (error instanceof RangeError) limit('source exceeds 16 MiB'); invalid('invalid JSON string') }
  }
  const visit = (input: unknown, depth: number): void => {
    if (depth > 128) invalid('JSON depth exceeds 128')
    if (input === null) { add(4); return }
    if (typeof input === 'string') { add(textSize(input)); return }
    if (typeof input === 'boolean') { add(input ? 4 : 5); return }
    if (typeof input === 'number') {
      if (!Number.isFinite(input)) invalid('non-finite JSON number')
      add(String(input).length); return
    }
    if (typeof input !== 'object' || ancestors.has(input)) invalid('source is not acyclic JSON')
    if (!Array.isArray(input)) plainObject(input, 'source has a non-JSON prototype')
    else if (Object.getPrototypeOf(input) !== Array.prototype) invalid('source has a non-JSON array prototype')
    if (Object.getOwnPropertySymbols(input).length) invalid('source has symbol keys')
    const conversion = Object.getOwnPropertyDescriptor(input, 'toJSON')
    if (conversion && (!('value' in conversion) || typeof conversion.value === 'function')) invalid('source has a JSON conversion')
    ancestors.add(input); add(2)
    if (Array.isArray(input)) {
      for (let index = 0; index < input.length; index++) {
        if (index) add(1)
        const property = Object.getOwnPropertyDescriptor(input, String(index))
        if (!property || !('value' in property)) invalid('source has sparse or accessor array entries')
        visit(property.value, depth + 1)
      }
      for (const key in input) if (own(input, key)
        && (!/^(?:0|[1-9]\d*)$/.test(key) || Number(key) >= input.length)) invalid('source has extra array properties')
    } else {
      let entries = 0
      for (const key in input) {
        const property = Object.getOwnPropertyDescriptor(input, key)
        if (!property) continue
        if (!('value' in property)) invalid('source has accessor fields')
        if (entries++) add(1)
        add(textSize(key) + 1); visit(property.value, depth + 1)
      }
    }
    ancestors.delete(input)
  }
  visit(value, 0)
}

function validateImport(value: unknown): void {
  const input = plainObject(value, 'invalid sheet import')
  if (input.version !== 1) invalid('invalid sheet import version')
  if (!own(input, 'formulaAliases')) return
  const aliases = plainObject(input.formulaAliases, 'invalid alias catalog')
  let entries = 0
  for (const key of ['sheets', 'names', 'tables', 'nameSheets']) {
    if (!own(aliases, key)) continue
    for (const [name, target] of Object.entries(plainObject(aliases[key], 'invalid alias map'))) {
      if (++entries > 32_768) limit('alias catalog exceeds 32768 entries')
      string(name, 2048, 'invalid alias key')
      if (typeof target !== 'string' || conditionalFormatUtf8Size(target, 2048) > 2048) invalid('invalid alias target')
    }
  }
}

function validateSource(value: unknown): asserts value is FormulaCalendarSourceDto {
  assertJson(value)
  const root = plainObject(value, 'missing source object')
  fields(root, ['version', 'date1904', 'sheets', 'helpers', 'tables'])
  if (root.version !== 1) invalid('unsupported source version')
  const date1904 = boolean(root.date1904, 'missing workbook calendar')
  const sheets = list(root.sheets, 1024, 'too many author worksheets')
  const helpers = list(root.helpers, 2049, 'too many calculation worksheets')
  const tables = list(root.tables, 24_576, 'too many table identities')
  const names = new Set<string>(), ids = new Set<string>(), nativeIds = new Set<number>()
  const authorIds = new Set<string>()
  const helperByName = new Map<string, { owner?: string; date1904: boolean }>()
  const identity = (entry: Record<string, unknown>) => {
    const name = string(entry.name, 124, 'invalid sheet name')
    if ([...name].length > 31) invalid('sheet name exceeds 31 characters')
    const id = string(entry.sheetId, 512, 'invalid sheet ID')
    const nativeId = uint(entry.officeSheetId, 0xffff_ffff, 'invalid native sheet ID', 1)
    if (names.has(name.toLowerCase()) || ids.has(id) || nativeIds.has(nativeId)) invalid('duplicate worksheet identity')
    names.add(name.toLowerCase()); ids.add(id); nativeIds.add(nativeId)
    return { name, id }
  }
  let cells = 0
  for (const value of sheets) {
    const sheet = plainObject(value, 'invalid author worksheet')
    fields(sheet, ['name', 'sheetId', 'officeSheetId', 'import', 'cells'])
    const { id } = identity(sheet); authorIds.add(id)
    const coordinates = new Set<string>()
    for (const value of list(sheet.cells, cellLimit, 'too many source cells')) {
      if (++cells > cellLimit) limit('source exceeds 250000 authored cells')
      const cell = plainObject(value, 'invalid source cell')
      fields(cell, ['row', 'col', 'formula', 'projectedFormula', 'formulaSource', 'dateSource', 'calculationSheet', 'calculationFormula'])
      const row = uint(cell.row, 1_048_575, 'invalid source row'), col = uint(cell.col, 16_383, 'invalid source column')
      if (coordinates.has(`${row}:${col}`)) invalid('duplicate source coordinate')
      coordinates.add(`${row}:${col}`)
      for (const key of ['formula', 'projectedFormula', 'calculationFormula']) if (own(cell, key)) string(cell[key], 32_768, 'invalid formula text')
      if (own(cell, 'calculationSheet')) string(cell.calculationSheet, 124, 'invalid calculation sheet name')
      if (own(cell, 'formulaSource')) {
        const marker = plainObject(cell.formulaSource, 'invalid formula provenance')
        if (marker.version !== 1 || marker.formula !== cell.formula || typeof cell.formula !== 'string'
          || typeof cell.projectedFormula !== 'string' || typeof cell.calculationFormula !== 'string' || typeof cell.calculationSheet !== 'string') invalid('incomplete formula provenance')
        boolean(marker.date1904, 'invalid formula calendar'); boolean(marker.dateResult, 'invalid formula result type')
        if (own(marker, 'aliasSheetId')) string(marker.aliasSheetId, 512, 'invalid alias owner')
      }
      if (own(cell, 'dateSource')) {
        const marker = plainObject(cell.dateSource, 'invalid exact date provenance')
        if (finiteValue(marker.value) === undefined) invalid('invalid original date serial')
        boolean(marker.date1904, 'invalid original date calendar')
        if (own(marker, 'targetValue') && finiteValue(marker.targetValue) === undefined) invalid('invalid target date serial')
        if (own(marker, 'targetDate1904') && marker.targetDate1904 !== date1904) invalid('target date calendar differs from workbook')
        if (own(marker, 'targetFormula') && marker.targetFormula !== null && marker.targetFormula !== cell.formula) invalid('target formula differs from author formula')
      }
    }
    if (own(sheet, 'import')) validateImport(sheet.import)
  }
  for (const value of helpers) {
    const helper = plainObject(value, 'invalid calculation worksheet')
    fields(helper, ['name', 'sheetId', 'officeSheetId', 'date1904', 'sourceSheetId', 'fingerprint', 'reclaim'])
    const { name } = identity(helper)
    const calendar = boolean(helper.date1904, 'invalid calculation calendar')
    const fingerprint = string(helper.fingerprint, 16, 'invalid calculation fingerprint')
    if (!/^[\da-f]{16}$/i.test(fingerprint)) invalid('invalid calculation fingerprint')
    const owner = own(helper, 'sourceSheetId') ? string(helper.sourceSheetId, 512, 'invalid calculation owner') : undefined
    if (owner !== undefined && !authorIds.has(owner)) invalid('missing calculation owner')
    if (own(helper, 'reclaim')) boolean(helper.reclaim, 'invalid reclaim approval')
    helperByName.set(name, { owner, date1904: calendar })
  }
  for (const value of sheets) {
    const sheet = plainObject(value, 'invalid author worksheet')
    for (const value of list(sheet.cells, cellLimit, 'too many source cells')) {
      const cell = plainObject(value, 'invalid source cell')
      if (own(cell, 'calculationSheet')) {
        const helper = helperByName.get(string(cell.calculationSheet, 124, 'invalid calculation sheet'))
        const marker = plainObject(cell.formulaSource, 'missing calculation provenance')
        if (!helper || helper.owner !== sheet.sheetId || helper.date1904 !== marker.date1904) invalid('calculation owner or calendar differs')
        if (own(marker, 'aliasSheetId') && !ids.has(string(marker.aliasSheetId, 512, 'invalid alias owner'))) invalid('alias owner is missing')
      }
    }
  }
  const tableNames = new Set<string>(), tableIds = new Set<string>(), nativeTableIds = new Set<number>()
  for (const value of tables) {
    const table = plainObject(value, 'invalid table identity')
    fields(table, ['name', 'tableId', 'officeTableId', 'sheetId'])
    const name = string(table.name, 255, 'invalid table name'), id = string(table.tableId, 512, 'invalid table ID')
    const nativeId = uint(table.officeTableId, 0xffff_ffff, 'invalid native table ID', 1)
    if (!ids.has(string(table.sheetId, 512, 'invalid table owner')) || tableNames.has(name.toLowerCase()) || tableIds.has(id) || nativeTableIds.has(nativeId)) invalid('duplicate or missing table identity')
    tableNames.add(name.toLowerCase()); tableIds.add(id); nativeTableIds.add(nativeId)
  }
}

/** The native parser has already inspected relationships and fingerprints.
 * This boundary repeats structural and byte guards before any source clone. */
export function normalizeWorkbookFormulaCalendarSource(value: unknown): FormulaCalendarSourceDto | undefined {
  if (value === undefined || value === null) return undefined
  validateSource(value)
  return cloneModel(value)
}

function rawCellIndex(value: unknown): Map<string, Record<string, unknown>> {
  const result = new Map<string, Record<string, unknown>>()
  if (!object(value) || !Array.isArray(value.cells)) return result
  for (const cell of value.cells) {
    if (!object(cell) || !Number.isSafeInteger(cell.row) || !Number.isSafeInteger(cell.col)) continue
    const key = `${cell.row}:${cell.col}`
    if (result.has(key)) invalid('duplicate native cell coordinate')
    result.set(key, cell)
  }
  return result
}

/** Remap import IDs into the newly created domain model. Changes are prepared
 * first; no calculation sheet is removed after a failed native cell match. */
export function applyWorkbookFormulaCalendarSource(workbook: WorkbookModel, rawSheets: unknown, rawTables: unknown, source: FormulaCalendarSourceDto | undefined): void {
  if (!source) return
  if (source.date1904 !== workbook.date1904) invalid('native workbook calendar changed during normalization')
  const nativeSheets = Array.isArray(rawSheets) ? rawSheets : []
  const byName = new Map<string, { sheet: SpreadsheetSheet; cells: Map<string, Record<string, unknown>> }>()
  for (const [index, sheet] of workbook.sheets.entries()) {
    if (byName.has(sheet.name.toLowerCase())) invalid('duplicate native worksheet name')
    byName.set(sheet.name.toLowerCase(), { sheet, cells: rawCellIndex(nativeSheets[index]) })
  }
  const idMap = new Map<string, string>()
  for (const entry of [...source.sheets, ...source.helpers]) {
    const native = byName.get(entry.name.toLowerCase())
    if (native) idMap.set(entry.sheetId, native.sheet.id)
  }
  const staged: Array<{ sheet: SpreadsheetSheet; cell: SpreadsheetCell; next: SpreadsheetCell }> = []
  let unmatched = 0, missingSheets = 0
  for (const entry of source.sheets) {
    const native = byName.get(entry.name.toLowerCase())
    if (!native) { unmatched += entry.cells.length; missingSheets++; continue }
    for (const saved of entry.cells) {
      const key = `${saved.row}:${saved.col}`, cell = native.sheet.cells[key], raw = native.cells.get(key)
      if (!cell || !raw) { unmatched++; continue }
      const rawFormula = typeof raw.formula === 'string' ? formula(raw.formula) : null
      if (saved.formulaSource) {
        const calculation = saved.calculationSheet ? byName.get(saved.calculationSheet.toLowerCase())?.cells.get(key) : undefined
        const calculationFormula = typeof calculation?.formula === 'string' ? formula(calculation.formula) : null
        if (rawFormula !== formula(saved.projectedFormula) || calculationFormula !== formula(saved.calculationFormula)
          || saved.formulaSource.aliasSheetId && !idMap.has(saved.formulaSource.aliasSheetId)) { unmatched++; continue }
      } else if (rawFormula !== formula(saved.formula)) { unmatched++; continue }
      const restoredFormula = formula(saved.formula)
      const metadata = { ...cell.metadata }
      if (saved.formulaSource) metadata.chartDependencyFormulaSource = { ...saved.formulaSource, formula: restoredFormula,
        ...(saved.formulaSource.aliasSheetId ? { aliasSheetId: idMap.get(saved.formulaSource.aliasSheetId) } : {}) }
      if (saved.dateSource) {
        const original = finiteValue(saved.dateSource.value)
        const target = saved.dateSource.targetValue === undefined && original !== undefined
          ? formulaCalendarSerial(original, saved.dateSource.date1904, workbook.date1904) : finiteValue(saved.dateSource.targetValue)
        if (typeof cell.value !== 'number' || target !== cell.value) { unmatched++; continue }
        metadata.chartDependencyDateSource = { ...saved.dateSource, targetValue: cell.value, targetDate1904: workbook.date1904, targetFormula: restoredFormula }
      }
      staged.push({ sheet: native.sheet, cell, next: { ...cell, formula: restoredFormula, metadata } })
    }
  }
  const removable = new Set<string>()
  const helperIds = new Set(source.helpers.map(helper => idMap.get(helper.sheetId)).filter((id): id is string => id !== undefined))
  const retainedAliasOwner = staged.some(({ next }) => {
    const marker = next.metadata?.chartDependencyFormulaSource
    return object(marker) && typeof marker.aliasSheetId === 'string' && helperIds.has(marker.aliasSheetId)
  })
  // A retained alias owner is part of the live formula environment. Keep its
  // calculation family, whose members can depend on one another.
  if (!unmatched && !missingSheets && !retainedAliasOwner) for (const helper of source.helpers) {
    const native = byName.get(helper.name.toLowerCase())?.sheet
    if (helper.reclaim === true && native && !native.visible && native.metadata?.hiddenState === 'veryHidden') removable.add(native.id)
  }
  const remaining = workbook.sheets.filter(sheet => !removable.has(sheet.id))
  if (!remaining.length || !remaining.some(sheet => sheet.visible)) removable.clear()
  for (const { sheet, cell, next } of staged) sheet.cells[cell.id] = next
  for (const entry of source.sheets) {
    const sheet = byName.get(entry.name.toLowerCase())?.sheet
    if (sheet && entry.import) sheet.metadata = { ...sheet.metadata, chartDependencyImport: entry.import }
  }
  for (const helper of source.helpers) {
    const sheet = byName.get(helper.name.toLowerCase())?.sheet
    if (!sheet) continue
    sheet.metadata = { ...sheet.metadata, formulaCalendarHelper: { version: 1, date1904: helper.date1904,
      officeSheetId: helper.officeSheetId, fingerprint: helper.fingerprint,
      ...(helper.sourceSheetId ? { sourceSheetId: idMap.get(helper.sourceSheetId) } : {}) } }
  }
  const nativeTables = Array.isArray(rawTables) ? rawTables : []
  const tableSources = new Map(source.tables.map(table => [table.officeTableId, table]))
  const tableIds = new Map<string, string>()
  for (const [index, table] of (workbook.tables ?? []).entries()) {
    const raw = nativeTables[index]
    const nativeId = table.officeTableId ?? (object(raw) && typeof raw.officeTableId === 'number' ? raw.officeTableId : undefined)
    const saved = nativeId !== undefined ? tableSources.get(nativeId) : undefined
    if (!saved || saved.name.toLowerCase() !== table.name.toLowerCase() || idMap.get(saved.sheetId) !== table.sheetId) continue
    tableIds.set(table.id, saved.tableId); table.id = saved.tableId
  }
  for (const slicer of workbook.slicers ?? []) if (slicer.tableId) slicer.tableId = tableIds.get(slicer.tableId) ?? slicer.tableId
  for (const cache of workbook.slicerCaches ?? []) if (cache.source.kind === 'table') cache.source.tableId = tableIds.get(cache.source.tableId) ?? cache.source.tableId
  const before = workbook.sheets
  workbook.sheets = before.filter(sheet => !removable.has(sheet.id))
  workbook.tables = workbook.tables?.filter(table => !removable.has(table.sheetId))
  const metadata = { ...workbook.metadata }
  for (const key of ['definedNames', 'localDefinedNames']) {
    if (!Array.isArray(metadata[key])) continue
    metadata[key] = metadata[key].flatMap(value => {
      if (!object(value)) return [value]
      const declaration = [value.sheetId, value.subUnitId, value.sheet, value.localSheetId].find(item => typeof item === 'string')
      const scoped = typeof declaration === 'string' ? idMap.get(declaration) ?? byName.get(declaration.toLowerCase())?.sheet.id : undefined
      const indexed = typeof value.localSheetId === 'number' ? before[value.localSheetId]?.id : undefined
      if (scoped && removable.has(scoped) || indexed && removable.has(indexed)) return []
      if (indexed) return [{ ...value, localSheetId: workbook.sheets.findIndex(sheet => sheet.id === indexed) }]
      return [value]
    })
  }
  metadata.formulaCalendarImport = { version: 1, restoredCells: staged.length, unmatchedCells: unmatched, missingSheets,
    reclaimedSheets: removable.size, retainedCalculationSheets: source.helpers.length - removable.size }
  workbook.metadata = metadata
  if (!workbook.sheets.some(sheet => sheet.id === workbook.activeSheetId)) workbook.activeSheetId = (workbook.sheets.find(sheet => sheet.visible) ?? workbook.sheets[0])!.id
}
