import type { CellRange, CellScalar, SpreadsheetCell, SpreadsheetSheet, SheetAutoFilter, SheetFilterCriteria, SheetFilterConditionOperator, WorkbookModel,
  SheetSortCriterion } from '../../domain/model'
import { sheetRowHeight, sheetAxisSpanAt } from '../../domain/sheet-axis.ts'
import { compilePivotRecordSelection, normalizePivotRecordSelection } from '../pivot/services/pivot-filter-predicate.ts'
import { sheetCellRecordValue } from '../sheet-workbench/sheet-record-value.ts'
import { getSheetOutlineHiddenSpans } from '../sheet-outline/sheet-outline.ts'
import { canonicalSheetFilterColor, createSheetFilterColorReader } from './sheet-filter-colors.ts'
import { sheetFilterColumnEntries } from './sheet-filter-column.ts'
export { canonicalSheetFilterColor } from './sheet-filter-colors.ts'
export type { SheetSortDirection, SheetSortCriterion, SheetFilterConditionOperator, SheetValueFilter, SheetConditionFilter,
  SheetFilterCriteria, SheetAutoFilter } from '../../domain/model.ts'
type FilteredSpreadsheetSheet = SpreadsheetSheet & { autoFilter?: SheetAutoFilter }

export interface SortRangeOptions { hasHeader?: boolean; locale?: string }
export type FormulaReferenceTranslator = (formula: string, rowDelta: number, columnDelta: number) => string
export interface NumericSortColumnData {
  rowCount: number
  columns: Array<Array<number | null | undefined>>
  directions: Array<'ascending' | 'descending'>
  destinationOffsets: number[]
}

export function normalizeAutoFilterRange(range: CellRange, sheet: FilteredSpreadsheetSheet): CellRange {
  validateRange(range)
  if (range.endRow >= getRowCount(sheet) || range.endCol >= getColumnCount(sheet)) throw new RangeError('Auto-filter range exceeds worksheet bounds')
  if (range.startRow === range.endRow) throw new RangeError('Auto-filter requires a header and at least one data row')
  if (range.startCol < 0 || range.startRow < 0) throw new RangeError('Auto-filter range must start at worksheet origin or later')
  return { ...range }
}

export function setAutoFilter(sheet: FilteredSpreadsheetSheet, range: CellRange | null, workbook?: WorkbookModel): void {
  if (range === null) { sheet.autoFilter = undefined; return }
  const normalized = normalizeAutoFilterRange(range, sheet)
  const previous = sheet.autoFilter
  const columns = previous && sameRange(previous.range, normalized) ? previous.columns : {}
  sheet.autoFilter = { range: normalized, columns: cloneCriteria(columns), filteredRows: [] }
  refreshAutoFilter(sheet, workbook)
}

export function setFilterCriteria(sheet: FilteredSpreadsheetSheet, column: number, criteria: SheetFilterCriteria | null, workbook?: WorkbookModel): void {
  const filter = sheet.autoFilter
  if (!filter) throw new Error('Create an auto-filter before setting column criteria')
  if (!Number.isInteger(column) || column < filter.range.startCol || column > filter.range.endCol) throw new RangeError('Filter column must be inside the auto-filter range')
  if (criteria === null) delete filter.columns[column]
  else filter.columns[column] = normalizeCriteria(criteria)
  refreshAutoFilter(sheet, workbook)
}

/** Derive source UI state with the same native cell/error and condition semantics as row filtering. */
export function compileSheetCellFilterCriteria(criteria: SheetFilterCriteria, context?: { sheet: SpreadsheetSheet; workbook?: WorkbookModel }): (cell: SpreadsheetCell | undefined, row?: number, column?: number) => boolean {
  const normalized = normalizeCriteria(criteria)
  if (normalized.kind === 'color') {
    const fills = normalized.cellFillColors && new Set(normalized.cellFillColors.map(color => color === null ? null : canonicalSheetFilterColor(color)))
    const text = normalized.cellTextColors && new Set(normalized.cellTextColors.map(canonicalSheetFilterColor))
    const colors = context && createSheetFilterColorReader(context.sheet, context.workbook)
    return (cell, row, column) => {
      const resolved = colors?.(cell?.row ?? row ?? 0, cell?.col ?? column ?? 0, cell)
      if (fills) return fills.has(resolved ? resolved.fill : cell?.style.fill == null ? null : canonicalSheetFilterColor(cell.style.fill))
      return text?.has(resolved?.text ?? canonicalSheetFilterColor(cell?.style.font.color ?? '#000000')) ?? false
    }
  }
  if (normalized.kind === 'values' && normalized.recordSelection) {
    const matches = compilePivotRecordSelection(normalized.recordSelection)
    return cell => matches(cell?.value ?? null, sheetCellRecordValue(cell, normalized.recordDate1904))
  }
  return cell => matchesCriteria(readCellValue(cell), normalized)
}

export function clearFilter(sheet: FilteredSpreadsheetSheet, column?: number, workbook?: WorkbookModel): void {
  const filter = sheet.autoFilter
  if (!filter) return
  if (column === undefined) filter.columns = {}
  else {
    if (!Number.isInteger(column) || column < filter.range.startCol || column > filter.range.endCol) throw new RangeError('Filter column must be inside the auto-filter range')
    delete filter.columns[column]
  }
  refreshAutoFilter(sheet, workbook)
}

export function refreshAutoFilter(sheet: FilteredSpreadsheetSheet, workbook?: WorkbookModel): number[] {
  const filter = sheet.autoFilter
  if (!filter) return []
  normalizeAutoFilterRange(filter.range, sheet)
  const hidden = new Set<number>()
  for (const [column, criteria] of Object.entries(filter.columns)) {
    for (const row of sheetFilterColumnHiddenRows(sheet, filter.range, Number(column), criteria, workbook)) hidden.add(row)
  }
  filter.filteredRows = [...hidden].sort((left, right) => left - right)
  return [...filter.filteredRows]
}

/** Compute one column independently; unioning columns preserves merged group failures and other hide reasons. */
export function sheetFilterColumnHiddenRows(sheet: SpreadsheetSheet, range: CellRange, column: number, criteria: SheetFilterCriteria, workbook?: WorkbookModel): Set<number> {
  const matches = compileSheetCellFilterCriteria(criteria, { sheet, workbook }), hidden = new Set<number>()
  for (const entry of sheetFilterColumnEntries(sheet, range, column)) {
    if (matches(entry.cell, entry.row, entry.column)) continue
    for (let offset = 0; offset < entry.rowSpan; offset++) hidden.add(entry.row + offset)
  }
  return hidden
}

export function sortRange(sheet: FilteredSpreadsheetSheet, range: CellRange, criteria: readonly SheetSortCriterion[], order?: readonly number[], options: SortRangeOptions = {}, translateFormula?: FormulaReferenceTranslator, workbook?: WorkbookModel): void {
  validateRange(range)
  if (range.endRow >= getRowCount(sheet) || range.endCol >= getColumnCount(sheet)) throw new RangeError('Sort range exceeds worksheet bounds')
  if (!criteria.length) throw new TypeError('At least one sort criterion is required')
  for (const criterion of criteria) {
    if (!Number.isInteger(criterion.column) || criterion.column < range.startCol || criterion.column > range.endCol) throw new RangeError('Sort key must be inside the sorted range')
    if (criterion.direction !== 'ascending' && criterion.direction !== 'descending') throw new TypeError('Invalid sort direction')
  }
  const firstDataRow = range.startRow + (options.hasHeader ? 1 : 0)
  const rowCount = range.endRow - firstDataRow + 1
  if (rowCount < 0) throw new RangeError('The sort range does not include data rows')
  if (order && !isPermutation(order, rowCount)) throw new RangeError('Sort result indices must be a permutation of the data rows')
  if (sheet.mergedRanges.some((merged) => rangesIntersect(range, merged))) throw new Error('Sort range cannot intersect merged cells')
  const sourceRows = Array.from({ length: rowCount }, (_, index) => {
    const row = firstDataRow + index
    const cells: Array<[number, SpreadsheetCell]> = []
    for (let column = range.startCol; column <= range.endCol; column += 1) {
      const cell = sheet.cells[cellKey(row, column)]
      if (cell) cells.push([column - range.startCol, cell])
    }
    return cells
  })
  const orderedIndices = order ? [...order] : getSortedRowOrder(sheet, range, criteria, options, workbook)
  // Resolve every moved formula before mutating the sheet. A failed reference
  // translation must not leave an emptied or partially reordered range.
  const sortedCells: Array<[string, SpreadsheetCell]> = []
  for (let destinationOffset = 0; destinationOffset < orderedIndices.length; destinationOffset += 1) {
    const destinationRow = firstDataRow + destinationOffset
    const sourceRow = firstDataRow + orderedIndices[destinationOffset]!
    for (const [columnOffset, source] of sourceRows[orderedIndices[destinationOffset]!]!) {
      const column = range.startCol + columnOffset
      const formula = source.formula && translateFormula ? translateFormula(source.formula, destinationRow - sourceRow, 0) : source.formula
      sortedCells.push([cellKey(destinationRow, column), { ...source, row: destinationRow, col: column, id: `${destinationRow}:${column}`, formula }])
    }
  }
  // Formulas, styles, rich text and metadata travel with complete cell records.
  // Clearing destinations first also removes cells moved into sparse positions.
  for (let row = firstDataRow; row <= range.endRow; row += 1) {
    for (let column = range.startCol; column <= range.endCol; column += 1) delete sheet.cells[cellKey(row, column)]
  }
  for (const [key, cell] of sortedCells) sheet.cells[key] = cell
  if (sheet.autoFilter && rangesIntersect(sheet.autoFilter.range, range)) refreshAutoFilter(sheet, workbook)
}

/** Returns stable data-row offsets suitable for a Worker-produced order. */
export function getSortedRowOrder(sheet: FilteredSpreadsheetSheet, range: CellRange, criteria: readonly SheetSortCriterion[], options: SortRangeOptions = {}, workbook?: WorkbookModel): number[] {
  validateRange(range)
  if (sheet.autoFilter) refreshAutoFilter(sheet, workbook)
  if (!criteria.length) throw new TypeError('At least one sort criterion is required')
  for (const criterion of criteria) {
    if (!Number.isInteger(criterion.column) || criterion.column < range.startCol || criterion.column > range.endCol) throw new RangeError('Sort key must be inside the sorted range')
    if (criterion.direction !== 'ascending' && criterion.direction !== 'descending') throw new TypeError('Invalid sort direction')
  }
  const firstDataRow = range.startRow + (options.hasHeader ? 1 : 0)
  const count = Math.max(0, range.endRow - firstDataRow + 1)
  const destinationOffsets = getSortableRowOffsets(sheet, range, options, workbook)
  const visibleOrder = [...destinationOffsets]
  const collator = new Intl.Collator(options.locale ?? 'zh-CN', { numeric: true, sensitivity: 'base' })
  visibleOrder.sort((left, right) => {
    for (const criterion of criteria) {
      const leftValue = readCellValue(sheet.cells[cellKey(firstDataRow + left, criterion.column)])
      const rightValue = readCellValue(sheet.cells[cellKey(firstDataRow + right, criterion.column)])
      const compared = compareValues(leftValue, rightValue, collator)
      if (compared !== 0) {
        // Blank placement is independent of direction; only nonblank values
        // reverse their comparison in a descending sort.
        if (rank(leftValue) === 4 || rank(rightValue) === 4) return compared
        return criterion.direction === 'ascending' ? compared : -compared
      }
    }
    return left - right
  })
  const order = Array.from({ length: count }, (_, index) => index)
  destinationOffsets.forEach((destinationOffset, index) => { order[destinationOffset] = visibleOrder[index]! })
  return order
}

export function getNumericSortColumns(sheet: FilteredSpreadsheetSheet, range: CellRange, criteria: readonly SheetSortCriterion[], options: SortRangeOptions = {}, workbook?: WorkbookModel): NumericSortColumnData | null {
  if (!criteria.length || criteria.some((criterion) => criterion.column < range.startCol || criterion.column > range.endCol)) return null
  const destinationOffsets = getSortableRowOffsets(sheet, range, options, workbook)
  const rowCount = destinationOffsets.length
  const first = range.startRow + (options.hasHeader ? 1 : 0)
  const columns = criteria.map((criterion) => destinationOffsets.map((offset) => {
    const value = readCellValue(sheet.cells[cellKey(first + offset, criterion.column)])
    return value === null || value === undefined || typeof value === 'number' ? value as number | null | undefined : Number.NaN
  }))
  if (columns.some((column) => column.some((value) => typeof value === 'number' && Number.isNaN(value)))) return null
  return { rowCount, columns, directions: criteria.map((criterion) => criterion.direction), destinationOffsets }
}

/** Expands a Worker-sorted visible-row index list into a full range permutation.
 * Filtered and manually hidden rows retain their exact positions. */
export function expandVisibleSortOrder(rowCount: number, destinationOffsets: readonly number[], sortedIndices: readonly number[]): number[] {
  if (destinationOffsets.some((offset) => !Number.isSafeInteger(offset) || offset < 0 || offset >= rowCount)
    || new Set(destinationOffsets).size !== destinationOffsets.length) throw new RangeError('Visible row offsets must be a unique subset of the range')
  if (!isPermutation(sortedIndices, destinationOffsets.length)) throw new RangeError('Worker sort result is not a permutation of visible rows')
  const order = Array.from({ length: rowCount }, (_, index) => index)
  destinationOffsets.forEach((destinationOffset, index) => { order[destinationOffset] = destinationOffsets[sortedIndices[index]!]! })
  return order
}

export function getSortableRowOffsets(sheet: FilteredSpreadsheetSheet, range: CellRange, options: SortRangeOptions = {}, workbook?: WorkbookModel): number[] {
  const first = range.startRow + (options.hasHeader ? 1 : 0)
  const filtered = new Set(sheet.autoFilter ? refreshAutoFilter(sheet, workbook) : [])
  const manuallyHidden = new Set(sheet.hiddenRows ?? [])
  const outlineHidden = sheet.outline ? getSheetOutlineHiddenSpans(sheet.outline, 'row', sheet.rowCount ?? 1_048_576) : []
  const result: number[] = []
  for (let row = first; row <= range.endRow; row += 1) {
    if (filtered.has(row) || manuallyHidden.has(row) || sheetAxisSpanAt(sheet.rowHiddenSpans, row)?.value
      || sheetAxisSpanAt(outlineHidden, row)?.value) continue
    if (sheetRowHeight(sheet, row) > 0) result.push(row - first)
  }
  return result
}

export function getFilteredRows(sheet: FilteredSpreadsheetSheet, workbook?: WorkbookModel): number[] {
  return sheet.autoFilter ? refreshAutoFilter(sheet, workbook) : []
}

export function cellKey(row: number, column: number): string { return `${row}:${column}` }

export function readCellValue(cell: SpreadsheetCell | undefined): CellScalar | { error: string } | null {
  if (!cell) return null
  if (cell.metadata?.error && typeof cell.metadata.error === 'string') return { error: cell.metadata.error }
  if (cell.dataType === 'e' || cell.dataType === 'error') return { error: typeof cell.value === 'string' ? cell.value : cell.display }
  if (cell.formula && cell.value == null && cell.display) {
    const cached = Number(cell.display)
    return Number.isFinite(cached) && cell.display.trim() !== '' ? cached : cell.display
  }
  return cell.value
}

function matchesCriteria(value: CellScalar | { error: string } | null, criteria: SheetFilterCriteria): boolean {
  if (criteria.kind === 'values') return criteria.selectedItems.some((item) => sameFilterValue(value, item)) !== (criteria.selectionMode === 'exclude')
  if (criteria.kind === 'color') throw new TypeError('Color filtering requires a cell style context')
  return matchesCondition(value, criteria)
}

function matchesCondition(value: CellScalar | { error: string } | null, criteria: Extract<SheetFilterCriteria, { kind: 'condition' }>): boolean {
  const { operator, value: expected, secondValue } = criteria
  if (operator === 'isError') return isError(value)
  if (operator === 'isBlank') return value === null || value === ''
  if (operator === 'isNotBlank') return value !== null && value !== ''
  if (value === null || isError(value) || expected === undefined) return false
  const leftText = String(value), rightText = String(expected)
  const collator = new Intl.Collator('zh-CN', { sensitivity: criteria.caseSensitive ? 'variant' : 'base', numeric: true })
  const compare = compareValues(value, expected, collator)
  switch (operator) {
    case 'equals': return criteria.caseSensitive ? leftText === rightText : compare === 0
    case 'doesNotEqual': return criteria.caseSensitive ? leftText !== rightText : compare !== 0
    case 'greaterThan': return compare > 0
    case 'greaterThanOrEqual': return compare >= 0
    case 'lessThan': return compare < 0
    case 'lessThanOrEqual': return compare <= 0
    case 'between': return secondValue !== undefined && compare >= 0 && compareValues(value, secondValue, collator) <= 0
    case 'notBetween': return secondValue !== undefined && (compare < 0 || compareValues(value, secondValue, collator) > 0)
    case 'contains': return includes(leftText, rightText, criteria.caseSensitive)
    case 'doesNotContain': return !includes(leftText, rightText, criteria.caseSensitive)
    case 'beginsWith': return criteria.caseSensitive ? leftText.startsWith(rightText) : leftText.toLocaleLowerCase().startsWith(rightText.toLocaleLowerCase())
    case 'endsWith': return criteria.caseSensitive ? leftText.endsWith(rightText) : leftText.toLocaleLowerCase().endsWith(rightText.toLocaleLowerCase())
  }
}

function normalizeCriteria(criteria: SheetFilterCriteria): SheetFilterCriteria {
  if (!criteria || (criteria.kind !== 'values' && criteria.kind !== 'condition' && criteria.kind !== 'color')) throw new TypeError('Invalid filter criteria')
  if (criteria.kind === 'color') {
    if (criteria.cellFillColors === undefined && criteria.cellTextColors === undefined) throw new TypeError('Color filter needs fill or text colors')
    for (const key of ['cellFillColors', 'cellTextColors'] as const) {
      const values = criteria[key]
      if (values !== undefined && (!Array.isArray(values) || values.length > 4096 || values.some(color => key === 'cellFillColors' && color === null ? false
        : typeof color !== 'string' || !color.trim() || color.length > 256))) throw new TypeError('Invalid filter colors')
    }
    return { kind: 'color', ...(criteria.cellFillColors === undefined ? {} : { cellFillColors: [...criteria.cellFillColors] }),
      ...(criteria.cellTextColors === undefined ? {} : { cellTextColors: [...criteria.cellTextColors] }) }
  }
  if (criteria.kind === 'values') {
    if (!Array.isArray(criteria.selectedItems)) throw new TypeError('Value filter needs a selectedItems array')
    if (criteria.selectionMode !== undefined && criteria.selectionMode !== 'include' && criteria.selectionMode !== 'exclude') throw new TypeError('Invalid value filter selection mode')
    if (criteria.recordDate1904 !== undefined && typeof criteria.recordDate1904 !== 'boolean') throw new TypeError('Invalid record calendar context')
    return { kind: 'values', selectedItems: [...criteria.selectedItems], ...(criteria.selectionMode === 'exclude' ? { selectionMode: 'exclude' as const } : {}),
      ...(criteria.recordSelection ? { recordSelection: normalizePivotRecordSelection(criteria.recordSelection) } : {}),
      ...(criteria.recordDate1904 !== undefined ? { recordDate1904: criteria.recordDate1904 } : {}) }
  }
  const operators: SheetFilterConditionOperator[] = ['equals', 'doesNotEqual', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual', 'between', 'notBetween', 'contains', 'doesNotContain', 'beginsWith', 'endsWith', 'isBlank', 'isNotBlank', 'isError']
  if (!operators.includes(criteria.operator)) throw new TypeError(`Unsupported filter operator: ${criteria.operator}`)
  if (!['isBlank', 'isNotBlank', 'isError'].includes(criteria.operator) && criteria.value === undefined) throw new TypeError('Filter condition requires a comparison value')
  if ((criteria.operator === 'between' || criteria.operator === 'notBetween') && criteria.secondValue === undefined) throw new TypeError('Between filters require two comparison values')
  return { ...criteria }
}

function cloneCriteria(value: Record<number, SheetFilterCriteria>): Record<number, SheetFilterCriteria> {
  return Object.fromEntries(Object.entries(value).map(([column, criteria]) => [column, normalizeCriteria(criteria)]))
}
function sameFilterValue(left: CellScalar | { error: string } | null, right: CellScalar): boolean {
  if (isError(left)) return false
  if (left === null || right === null) return left === right
  return typeof left === typeof right && Object.is(left, right)
}
function compareValues(left: CellScalar | { error: string } | null, right: CellScalar | { error: string } | null, collator: Intl.Collator): number {
  const leftRank = rank(left), rightRank = rank(right)
  const leftBlank = leftRank === 4, rightBlank = rightRank === 4
  if (leftBlank || rightBlank) return leftBlank === rightBlank ? 0 : leftBlank ? 1 : -1
  if (leftRank !== rightRank) return leftRank - rightRank
  if (left === null || right === null || isError(left) || isError(right)) return 0
  if (typeof left === 'number' && typeof right === 'number') return left < right ? -1 : left > right ? 1 : 0
  if (typeof left === 'boolean' && typeof right === 'boolean') return Number(left) - Number(right)
  return collator.compare(String(left), String(right))
}
function rank(value: CellScalar | { error: string } | null): number {
  if (typeof value === 'number') return 0
  if (typeof value === 'string' && value !== '') return 1
  if (typeof value === 'boolean') return 2
  if (isError(value)) return 3
  return 4
}
function includes(text: string, search: string, caseSensitive = false): boolean { return caseSensitive ? text.includes(search) : text.toLocaleLowerCase().includes(search.toLocaleLowerCase()) }
function isError(value: unknown): value is { error: string } { return Boolean(value && typeof value === 'object' && 'error' in value) }
function validateRange(range: CellRange): void {
  if (!range || ![range.startRow, range.startCol, range.endRow, range.endCol].every((value) => Number.isSafeInteger(value) && value >= 0) || range.endRow < range.startRow || range.endCol < range.startCol) throw new RangeError('Invalid worksheet range')
}
function getRowCount(sheet: SpreadsheetSheet): number { let count = Math.max(sheet.rowCount ?? 0, 1); for (const cell of Object.values(sheet.cells)) count = Math.max(count, cell.row + 1); return count }
function getColumnCount(sheet: SpreadsheetSheet): number { let count = Math.max(sheet.columnCount ?? 0, 1); for (const cell of Object.values(sheet.cells)) count = Math.max(count, cell.col + 1); return count }
function sameRange(left: CellRange, right: CellRange): boolean { return left.startRow === right.startRow && left.startCol === right.startCol && left.endRow === right.endRow && left.endCol === right.endCol }
function rangesIntersect(left: CellRange, right: CellRange): boolean { return left.startRow <= right.endRow && left.endRow >= right.startRow && left.startCol <= right.endCol && left.endCol >= right.startCol }
function isPermutation(order: readonly number[], rowCount: number): boolean {
  if (order.length !== rowCount) return false
  const seen = new Uint8Array(rowCount)
  for (const index of order) { if (!Number.isSafeInteger(index) || index < 0 || index >= rowCount || seen[index]) return false; seen[index] = 1 }
  return true
}
