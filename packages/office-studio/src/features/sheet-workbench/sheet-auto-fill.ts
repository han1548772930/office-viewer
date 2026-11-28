import type { CellRange, SpreadsheetCell, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import { DEFAULT_SHEET_COLUMN_COUNT, DEFAULT_SHEET_ROW_COUNT } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { validateSheetDataValidationChanges, type SheetValidationIssue } from '../sheet-validation/validation-engine.ts'
import { sheetValidationSourceReader } from '../sheet-validation/validation-value-read.ts'
import { canViewProtectedRanges, canViewWorksheetProtection } from '../sheet-protection/range-protection.ts'
import { createSheetClipboardPastedCell } from './sheet-clipboard-cell.ts'
import { sheetClipboardCellAt } from './sheet-clipboard-data.ts'
import { pasteSheetClipboardRangeResources } from './sheet-clipboard-resources.ts'
import { sheetClipboardRowsToRanges, type SheetClipboardTargetGeometry } from './sheet-clipboard-target.ts'
import type { SheetClipboardSnapshot } from './sheet-clipboard-types.ts'
import { clearSpreadsheetCellContent } from './sheet-clear-command.ts'
import { createSheetFillSeriesPattern, sheetFillSeriesValueAt, type SheetFillSeriesPattern } from './sheet-fill-series.ts'
import { resolveSheetFillTarget, type SheetFillTarget } from './sheet-fill-target.ts'
import type { SheetFillApplyType } from './sheet-fill.ts'
import { classifySheetNumberFormat, formatSheetValue } from './sheet-number-format.ts'
import { assertSheetRangesEditable, sheetRangeContainsCell, sheetRangesEqual, sheetRangesIntersect } from './sheet-range-command.ts'
import { hasInheritedSheetStyle, resolveSheetCellStyle } from './sheet-style-inheritance.ts'
import { captureSheetStyleRegion } from './sheet-style-region.ts'

export const SHEET_AUTO_FILL_COMMAND = 'sheet.command.auto-fill'
export const SHEET_REFILL_COMMAND = 'sheet.command.refill'
export const SHEET_COPY_DOWN_COMMAND = 'sheet.command.copy-down'
export const SHEET_COPY_RIGHT_COMMAND = 'sheet.command.copy-right'
export const SHEET_AUTO_CLEAR_COMMAND = 'sheet.command.auto-clear-content'
export const SHEET_FILL_COMMAND_IDS = [SHEET_AUTO_FILL_COMMAND, SHEET_REFILL_COMMAND, SHEET_COPY_DOWN_COMMAND,
  SHEET_COPY_RIGHT_COMMAND, SHEET_AUTO_CLEAR_COMMAND] as const
export const SHEET_FILL_SYNC_CELL_BUDGET = 100_000

export interface SheetAutoFillResult {
  accepted: boolean
  changed: boolean
  selection?: CellRange
  ranges: CellRange[]
  applyType: SheetFillApplyType
  supportsSeries: boolean
  issues: SheetValidationIssue[]
  sessionId?: string
  reason?: 'empty-source' | 'same-range' | 'validation' | 'stale-target' | 'stale-fill'
}
export interface SheetAutoFillPlan {
  sheet?: SpreadsheetSheet
  result: SheetAutoFillResult
  target?: SheetFillTarget
}
const mod = (value: number, length: number): number => (value % length + length) % length

function filteredRows(workbook: WorkbookModel, sheet: SpreadsheetSheet): Set<number> {
  return new Set([...(sheet.autoFilter?.filteredRows ?? []), ...(workbook.tables ?? []).filter(table => table.sheetId === sheet.id)
    .flatMap(table => table.autoFilter?.filteredRows ?? [])])
}
function rowsFor(range: CellRange, filtered: ReadonlySet<number>): number[] {
  const rows: number[] = []
  for (let row = range.startRow; row <= range.endRow; row++) if (!filtered.has(row)) rows.push(row)
  return rows
}
function assertCompleteMerges(sheet: SpreadsheetSheet, range: CellRange): void {
  for (const merge of sheet.mergedRanges) if (sheetRangesIntersect(merge, range)
    && (!sheetRangeContainsCell(range, merge.startRow, merge.startCol) || !sheetRangeContainsCell(range, merge.endRow, merge.endCol))) {
    throw new RangeError('填充范围不能截断合并单元格。')
  }
}
function cellIndex(sheet: SpreadsheetSheet): Map<string, SpreadsheetCell> {
  return new Map(Object.values(sheet.cells).map(cell => [`${cell.row}:${cell.col}`, cell]))
}
function seedPattern(cells: readonly (SpreadsheetCell | undefined)[]): SheetFillSeriesPattern {
  return createSheetFillSeriesPattern(cells.map(cell => ({ value: cell?.value ?? null, formula: cell?.formula, image: Boolean(cell?.image),
    date: cell !== undefined && ['date', 'time', 'datetime', 'duration'].includes(classifySheetNumberFormat(cell.style.numberFormat)) })))
}
function setSeriesValue(cell: SpreadsheetCell, value: string | number | boolean | null, workbook: WorkbookModel, region?: string): void {
  cell.value = value; cell.formula = null; cell.cachedDataType = null
  cell.dataType = value === null ? 'blank' : typeof value === 'number' ? 'number' : typeof value === 'boolean' ? 'boolean' : 'text'
  cell.display = formatSheetValue(value, cell.style.numberFormat, workbook.date1904, region) ?? String(value ?? '')
  delete cell.richText; delete cell.image
  if (cell.metadata) delete cell.metadata.formulaError
}

function sourceSnapshot(workbook: WorkbookModel, sheet: SpreadsheetSheet, source: CellRange, rows: number[], columns: number[]): SheetClipboardSnapshot {
  const merges = sheet.mergedRanges.filter(merge => sheetRangesIntersect(merge, source)).map(merge => {
    const startRow = rows.indexOf(merge.startRow), endRow = rows.indexOf(merge.endRow)
    if (startRow < 0 || endRow < 0 || endRow - startRow !== merge.endRow - merge.startRow) {
      throw new RangeError('源合并单元格不能跨过筛选隐藏行。')
    }
    return { startRow, endRow, startCol: merge.startCol - source.startCol, endCol: merge.endCol - source.startCol }
  })
  const rowOffsets = new Map(rows.map((row, offset) => [row, offset])), cells: Record<string, SpreadsheetCell> = {}
  for (const cell of Object.values(sheet.cells)) {
    const row = rowOffsets.get(cell.row), col = cell.col - source.startCol
    if (row === undefined || col < 0 || col >= columns.length) continue
    cells[`${row}:${col}`] = { ...cloneModel(cell), id: `${row}:${col}`, row, col,
      style: resolveSheetCellStyle(sheet, cell.row, cell.col, cell, workbook) }
  }
  return { version: 1, id: 'auto-fill-source', copyType: 'copy', formulaOnly: false, rows, columns, cells, merges,
    styleRegion: captureSheetStyleRegion(sheet, rows, columns, workbook),
    columnWidths: [], conditionalFormats: sheet.conditionalFormats ? cloneModel(sheet.conditionalFormats) : undefined,
    dataValidations: sheet.dataValidations ? cloneModel(sheet.dataValidations) : undefined }
}

interface FillTile { source: number[]; target: number[]; offset: number }
function tiles(source: readonly number[], target: readonly number[], backwards: boolean): FillTile[] {
  const result: FillTile[] = []
  let offset = backwards ? mod(-target.length, source.length) : 0
  for (let index = 0; index < target.length;) {
    const count = Math.min(source.length - offset, target.length - index)
    result.push({ source: source.slice(offset, offset + count), target: target.slice(index, index + count), offset })
    index += count; offset = 0
  }
  return result
}

/** Each partial/backward tile has an exact source projection; clipboard resource hooks remain the shared implementation. */
function applyRangeResources(sheet: SpreadsheetSheet, snapshot: SheetClipboardSnapshot, targetRows: number[], targetColumns: number[],
  target: SheetFillTarget, applyType: SheetFillApplyType): void {
  const vertical = target.direction === 'up' || target.direction === 'down'
  const rowTiles = vertical ? tiles(snapshot.rows, targetRows, target.direction === 'up') : [{ source: snapshot.rows, target: targetRows, offset: 0 }]
  const columnTiles = vertical ? [{ source: snapshot.columns, target: targetColumns, offset: 0 }]
    : tiles(snapshot.columns, targetColumns, target.direction === 'left')
  const conditional = applyType === 'NO_FORMAT' && sheet.conditionalFormats ? cloneModel(sheet.conditionalFormats) : undefined
  for (const rowTile of rowTiles) for (const columnTile of columnTiles) {
    const local = { ...snapshot, rows: rowTile.source, columns: columnTile.source }
    const geometry: SheetClipboardTargetGeometry = { rows: rowTile.target, columns: columnTile.target,
      ranges: sheetClipboardRowsToRanges(rowTile.target, columnTile.target[0]!, columnTile.target.at(-1)!), merges: [] }
    pasteSheetClipboardRangeResources(sheet, local, geometry, applyType === 'ONLY_FORMAT' ? 'formats' : 'all')
    if (applyType === 'NO_FORMAT') continue
    for (const merge of snapshot.merges) {
      const top = snapshot.rows[merge.startRow]!, bottom = snapshot.rows[merge.endRow]!
      const left = snapshot.columns[merge.startCol]!, right = snapshot.columns[merge.endCol]!
      const sourceRange = { startRow: rowTile.source[0]!, endRow: rowTile.source.at(-1)!, startCol: columnTile.source[0]!, endCol: columnTile.source.at(-1)! }
      if (!sheetRangesIntersect(mergeRange(snapshot, merge), sourceRange)) continue
      const startRow = rowTile.source.indexOf(top), endRow = rowTile.source.indexOf(bottom)
      const startCol = columnTile.source.indexOf(left), endCol = columnTile.source.indexOf(right)
      if ([startRow, endRow, startCol, endCol].some(index => index < 0)) throw new RangeError('填充目标不能截断源合并单元格。')
      const projected = { startRow: rowTile.target[startRow]!, endRow: rowTile.target[endRow]!,
        startCol: columnTile.target[startCol]!, endCol: columnTile.target[endCol]! }
      if (projected.endRow - projected.startRow !== bottom - top) throw new RangeError('目标合并单元格不能跨过筛选隐藏行。')
      if (!sheet.mergedRanges.some(value => sheetRangesEqual(value, projected))) sheet.mergedRanges.push(projected)
    }
  }
  if (applyType === 'NO_FORMAT') sheet.conditionalFormats = conditional
}
function mergeRange(snapshot: SheetClipboardSnapshot, merge: CellRange): CellRange {
  return { startRow: snapshot.rows[merge.startRow]!, endRow: snapshot.rows[merge.endRow]!,
    startCol: snapshot.columns[merge.startCol]!, endCol: snapshot.columns[merge.endCol]! }
}

/** A detached candidate includes content, styles, validation, CF and merges before any publication. */
export function createSheetAutoFillPlan(workbook: WorkbookModel, sheetId: string, source: CellRange, selection: CellRange,
  options: { applyType?: SheetFillApplyType; userId?: string; region?: string; allowClear?: boolean } = {}): SheetAutoFillPlan {
  const applyType = options.applyType ?? 'COPY', result: SheetAutoFillResult = { accepted: false, changed: false,
    ranges: [], issues: [], applyType, supportsSeries: false }
  const target = resolveSheetFillTarget(source, selection, options.allowClear ?? true)
  if (!target) return { result: { ...result, reason: 'same-range' } }
  const original = workbook.sheets.find(sheet => sheet.id === sheetId)
  if (!original) return { result: { ...result, reason: 'stale-target' } }
  const work = (source.endRow - source.startRow + 1) * (source.endCol - source.startCol + 1)
    + (target.destination.endRow - target.destination.startRow + 1) * (target.destination.endCol - target.destination.startCol + 1)
  if (work > SHEET_FILL_SYNC_CELL_BUDGET) throw new RangeError('填充范围超出当前同步预算，需要分块 Worker 填充。')
  const userId = options.userId ?? 'local-user'
  if (!canViewWorksheetProtection(original, userId) || !canViewProtectedRanges(original, [source], userId)) throw new Error('没有读取填充源区域的权限。')
  assertCompleteMerges(original, source); assertCompleteMerges(original, target.destination)
  const filtered = filteredRows(workbook, original), sourceRows = rowsFor(source, filtered), destinationRows = rowsFor(target.destination, filtered)
  const destinationColumns = Array.from({ length: target.destination.endCol - target.destination.startCol + 1 }, (_, index) => target.destination.startCol + index)
  const ranges = sheetClipboardRowsToRanges(destinationRows, target.destination.startCol, target.destination.endCol)
  if (!sourceRows.length || !destinationRows.length) return { result: { ...result, reason: 'empty-source' }, target }
  for (const merge of original.mergedRanges) if (ranges.some(range => sheetRangesIntersect(range, merge))
    && !ranges.some(range => sheetRangeContainsCell(range, merge.startRow, merge.startCol) && sheetRangeContainsCell(range, merge.endRow, merge.endCol))) {
    throw new RangeError('目标合并单元格不能跨过筛选隐藏行。')
  }
  assertSheetRangesEditable(original, ranges, userId)
  const sheet = cloneModel(original), index = cellIndex(original)
  if (target.clear) {
    for (const [key, cell] of Object.entries(sheet.cells)) if (ranges.some(range => sheetRangeContainsCell(range, cell.row, cell.col))) {
      sheet.cells[key] = clearSpreadsheetCellContent(cell)
    }
    return { sheet, target, result: { ...result, accepted: true, changed: JSON.stringify(sheet) !== JSON.stringify(original), selection: target.selection, ranges } }
  }
  sheet.rowCount = Math.max(sheet.rowCount ?? DEFAULT_SHEET_ROW_COUNT, target.selection.endRow + 1)
  sheet.columnCount = Math.max(sheet.columnCount ?? DEFAULT_SHEET_COLUMN_COUNT, target.selection.endCol + 1)
  const sourceColumns = Array.from({ length: source.endCol - source.startCol + 1 }, (_, col) => source.startCol + col)
  const vertical = target.direction === 'up' || target.direction === 'down'
  const snapshot = sourceSnapshot(workbook, original, source, sourceRows, sourceColumns)
  const patterns = vertical ? sourceColumns.map((_, col) => seedPattern(sourceRows.map((_, row) => sheetClipboardCellAt(snapshot, row, col))))
    : sourceRows.map((_, row) => seedPattern(sourceColumns.map((_, col) => sheetClipboardCellAt(snapshot, row, col))))
  const effectiveType = options.applyType ?? (patterns.every(pattern => pattern.preferredApplyType === 'COPY') ? 'COPY' : 'SERIES')
  if (effectiveType !== 'NO_FORMAT') sheet.mergedRanges = sheet.mergedRanges.filter(merge => !ranges.some(range => sheetRangesIntersect(range, merge)))
  const changedPositions: Array<{ row: number; column: number }> = []
  for (let r = 0; r < destinationRows.length; r++) for (let c = 0; c < destinationColumns.length; c++) {
    const row = destinationRows[r]!, col = destinationColumns[c]!
    const ordinal = vertical ? target.direction === 'up' ? r - destinationRows.length : sourceRows.length + r
      : target.direction === 'left' ? c - destinationColumns.length : sourceColumns.length + c
    const sourceRowOffset = vertical ? mod(ordinal, sourceRows.length) : r
    const sourceColOffset = vertical ? c : mod(ordinal, sourceColumns.length)
    const sourceRow = sourceRows[sourceRowOffset]!, sourceCol = sourceColumns[sourceColOffset]!
    const seed = sheetClipboardCellAt(snapshot, sourceRowOffset, sourceColOffset), existing = index.get(`${row}:${col}`)
    const sourceIsVirtual = !snapshot.cells[`${sourceRowOffset}:${sourceColOffset}`]
    const targetStyle = resolveSheetCellStyle(original, row, col, existing, workbook)
    let cell = createSheetClipboardPastedCell(seed, existing, { row, col }, effectiveType === 'ONLY_FORMAT' ? 'formats' : 'all', {
      date1904: workbook.date1904, region: options.region, rowDelta: row - sourceRow, columnDelta: col - sourceCol,
      sourceIsVirtual: effectiveType !== 'NO_FORMAT' && sourceIsVirtual, targetStyle,
      targetHasInheritedStyle: hasInheritedSheetStyle(original, row, col, workbook, existing?.themeStyle) })
    const key = `${row}:${col}`
    if (cell && effectiveType !== 'ONLY_FORMAT' && effectiveType !== 'COPY' && !seed?.formula && !seed?.image) {
      const generated = sheetFillSeriesValueAt(patterns[vertical ? c : r]!, ordinal)
      if (generated !== undefined) setSeriesValue(cell, generated, workbook, options.region)
    }
    if (effectiveType === 'NO_FORMAT') {
      if (cell && sourceIsVirtual && !existing && cell.value === null && !cell.formula && !cell.richText?.length && !cell.sparkline && !cell.image) cell = undefined
      if (cell) {
        cell.style = cloneModel(targetStyle)
        cell.authoredStyle = existing?.authoredStyle === undefined ? existing ? undefined : {} : cloneModel(existing.authoredStyle)
        cell.authoredStyleReference = existing?.authoredStyleReference
        cell.clearedStyleProperties = existing?.clearedStyleProperties ? [...existing.clearedStyleProperties] : undefined
        cell.themeStyle = existing?.themeStyle === undefined ? undefined : cloneModel(existing.themeStyle)
        if (!cell.formula) cell.display = formatSheetValue(cell.value, cell.style.numberFormat, workbook.date1904, options.region) ?? String(cell.value ?? '') }
      else if (existing) cell = { ...clearSpreadsheetCellContent(existing), id: key, row, col }
    }
    if (existing && existing.id !== key) delete sheet.cells[existing.id]
    if (cell) sheet.cells[key] = cell
    else delete sheet.cells[key]
    changedPositions.push({ row, column: col })
  }
  applyRangeResources(sheet, snapshot, destinationRows, destinationColumns, target, effectiveType)
  if (effectiveType !== 'ONLY_FORMAT') {
    const validationPositions = changedPositions.filter(({ row, column }) => {
      const merge = sheet.mergedRanges.find(item => sheetRangeContainsCell(item, row, column))
      if (!merge || merge.startRow === row && merge.startCol === column) return true
      const cell = sheet.cells[`${row}:${column}`]
      if (cell) sheet.cells[`${row}:${column}`] = clearSpreadsheetCellContent(cell)
      return false
    })
    const validation = validateSheetDataValidationChanges({ ...workbook, sheets: workbook.sheets.map(item => item.id === sheetId ? sheet : item) },
      sheetId, validationPositions, sheetValidationSourceReader(userId), options.region)
    result.issues = validation.issues
    if (validation.blocked) return { target, result: { ...result, applyType: effectiveType, issues: validation.issues,
      supportsSeries: patterns.some(pattern => pattern.supportsSeries), reason: 'validation' } }
  }
  return { sheet, target, result: { ...result, applyType: effectiveType, accepted: true, selection: target.selection, ranges,
    supportsSeries: patterns.some(pattern => pattern.supportsSeries), changed: JSON.stringify(sheet) !== JSON.stringify(original) } }
}
