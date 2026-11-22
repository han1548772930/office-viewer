import { defaultCellStyle, type CellRange, type SpreadsheetCell, type SpreadsheetSheet } from '../../domain/model.ts'
import { parseCellKey, type ExcelCellView, type ExcelSheetView } from '../../services/excelViewAdapter.ts'
import { clearSheetCellFormat } from './sheet-toolbar-format.ts'
import { formatSheetValue } from './sheet-number-format.ts'
import { translateFormulaReferences } from '../formula/formula-references.ts'
import type { SheetDataValidationRule, SheetConditionalFormatRule, SheetConditionalFormatValueConfig } from '../../domain/model.ts'
import { assertSheetRangesEditable, sheetCommandRangesWithoutFilteredRows, sheetRangeContainsCell, sheetRangesIntersect,
  subtractSheetCommandRange, type SheetRangeCommandInput } from './sheet-range-command.ts'

export const SHEET_CLEAR_COMMANDS = Object.freeze({
  content: 'sheet.command.clear-content', format: 'sheet.command.clear-format', all: 'sheet.command.clear-all',
  range: 'sheet.command.clear-range',
})
export const SHEET_CLEAR_COMMAND_ALIASES = Object.freeze({
  content: 'sheet.command.clear-selection-content', format: 'sheet.command.clear-selection-format', all: 'sheet.command.clear-selection-all',
})
export const SHEET_CLEAR_COMMAND_IDS: readonly string[] = [...Object.values(SHEET_CLEAR_COMMANDS), ...Object.values(SHEET_CLEAR_COMMAND_ALIASES)]
export type SheetClearMode = 'content' | 'format' | 'all' | 'none'
export interface SheetClearOptions {
  contentsOnly?: boolean
  formatOnly?: boolean
  /** Compatibility for the project's earlier clear API. */
  contents?: boolean
  format?: boolean
}
export interface SheetClearCommandInput extends SheetRangeCommandInput { options?: SheetClearOptions }

export function sheetClearModeForOptions(options: SheetClearOptions = {}): SheetClearMode {
  if (!options || typeof options !== 'object' || Array.isArray(options)) throw new TypeError('清除选项无效。')
  for (const flag of [options.contentsOnly, options.formatOnly, options.contents, options.format]) {
    if (flag !== undefined && typeof flag !== 'boolean') throw new TypeError('清除选项必须是布尔值。')
  }
  if (options.contentsOnly !== undefined || options.formatOnly !== undefined) {
    return options.contentsOnly && !options.formatOnly ? 'content' : options.formatOnly && !options.contentsOnly ? 'format' : 'all'
  }
  if (options.contents !== undefined || options.format !== undefined) {
    const contents = options.contents ?? true, format = options.format ?? false
    return contents ? format ? 'all' : 'content' : format ? 'format' : 'none'
  }
  return 'all'
}
export function clearModeForCommand(id: string, input?: SheetClearCommandInput): SheetClearMode | undefined {
  if (id === SHEET_CLEAR_COMMANDS.range) return sheetClearModeForOptions(input?.options)
  if (id === SHEET_CLEAR_COMMANDS.content || id === SHEET_CLEAR_COMMAND_ALIASES.content) return 'content'
  if (id === SHEET_CLEAR_COMMANDS.format || id === SHEET_CLEAR_COMMAND_ALIASES.format) return 'format'
  if (id === SHEET_CLEAR_COMMANDS.all || id === SHEET_CLEAR_COMMAND_ALIASES.all) return 'all'
  return undefined
}

/** Clearing a value invalidates every cached representation while retaining cell identity and formatting. */
export function clearSpreadsheetCellContent(cell: SpreadsheetCell): SpreadsheetCell {
  const metadata = cell.metadata ? { ...cell.metadata } : undefined
  if (metadata) delete metadata.formulaError
  return { ...cell, value: null, formula: null, display: '', dataType: 'blank', cachedDataType: null,
    richText: undefined, hyperlink: null, sparkline: undefined, metadata }
}
export function clearSheetViewCellContent(cell: ExcelCellView): ExcelCellView {
  const metadata = cell.metadata ? { ...cell.metadata } : cell.metadata
  if (metadata) delete metadata.formulaError
  return { ...cell, value: '', rawValue: null, formula: undefined, cachedValue: undefined, displayValue: '',
    dataType: 'blank', cachedDataType: null, richText: null, hyperlink: null, sparkline: undefined, metadata }
}
function trimResourceRanges<T extends { ranges: CellRange[] }>(rules: T[] | undefined, ranges: readonly CellRange[], rebase?: (rule: T, ranges: CellRange[]) => T): T[] | undefined {
  return rules?.flatMap(rule => {
    let remaining = rule.ranges.map(range => ({ ...range }))
    for (const range of ranges) remaining = remaining.flatMap(source => subtractSheetCommandRange(source, range))
    return remaining.length ? [rebase ? rebase(rule, remaining) : { ...rule, ranges: remaining }] : []
  })
}
function formulaFromNewAnchor(formula: string | undefined, before: CellRange | undefined, after: CellRange | undefined): string | undefined {
  if (!formula || !before || !after || before.startRow === after.startRow && before.startCol === after.startCol) return formula
  return translateFormulaReferences(formula, after.startRow - before.startRow, after.startCol - before.startCol)
}
function rebaseValidation(rule: SheetDataValidationRule, ranges: CellRange[]): SheetDataValidationRule {
  // List sources and checkbox labels are evaluated without a per-cell relative offset.
  if (rule.type === 'list' || rule.type === 'listMultiple' || rule.type === 'checkbox') return { ...rule, ranges }
  return { ...rule, ranges, formula1: formulaFromNewAnchor(rule.formula1, rule.ranges[0], ranges[0]),
    formula2: formulaFromNewAnchor(rule.formula2, rule.ranges[0], ranges[0]) }
}
function rebaseConditional(rule: SheetConditionalFormatRule, ranges: CellRange[]): SheetConditionalFormatRule {
  // Conditional rendering anchors to the first range by position, while validation preserves serialized range order.
  const anchor = (items: readonly CellRange[]) => items.reduce<CellRange | undefined>((first, range) =>
    !first || range.startRow < first.startRow || range.startRow === first.startRow && range.startCol < first.startCol ? range : first, undefined)
  const before = anchor(rule.ranges), after = anchor(ranges)
  const rebase = (config: SheetConditionalFormatValueConfig) => config.type === 'formula' && typeof config.value === 'string'
    ? { ...config, value: formulaFromNewAnchor(config.value, before, after) } : { ...config }
  const expression = rule.type === 'expression' || rule.subType === 'formula'
  return { ...rule, ranges,
    ...(expression ? { formula1: formulaFromNewAnchor(rule.formula1, before, after) ?? '',
      ...(typeof rule.value === 'string' ? { value: formulaFromNewAnchor(rule.value, before, after) } : {}) } : {}),
    ...(rule.config ? { config: { ...rule.config, min: rule.config.min ? rebase(rule.config.min) : undefined,
      max: rule.config.max ? rebase(rule.config.max) : undefined,
      thresholds: rule.config.thresholds?.map(item => ({ ...item, value: rebase(item.value) })),
      icons: rule.config.icons?.map(item => ({ ...item, value: rebase(item.value) })) } } : {}) }
}
function clearViewFormat(cell: ExcelCellView): ExcelCellView {
  const display = cell.rawValue === undefined ? undefined : cell.formula && cell.rawValue === null ? cell.cachedValue ?? cell.displayValue
    : formatSheetValue(cell.rawValue, 'General')
  return { ...clearSheetCellFormat(cell), hyperlink: null, styleMetadata: null,
    displayValue: display, ...(cell.formula && display !== undefined ? { cachedValue: display } : {}) }
}
function parseViewMerges(sheet: ExcelSheetView): CellRange[] {
  return (sheet.merges ?? []).map(value => {
    const [start, end] = value.split(':'), first = parseCellKey(start ?? ''), last = parseCellKey(end ?? start ?? '')
    if (!first || !last) throw new RangeError('合并区域坐标无效。')
    return { startRow: first.row, startCol: first.col, endRow: last.row, endCol: last.col }
  })
}
function clearScope(
  sheet: SpreadsheetSheet | ExcelSheetView, ranges: readonly CellRange[], merges: readonly CellRange[],
  mode: SheetClearMode, userId: string, filteredRows: readonly number[],
) {
  const removedMerges = mode === 'format' || mode === 'all' ? merges.filter(merge => ranges.some(range => sheetRangesIntersect(range, merge))) : []
  // Even a one-cell API request must be authorized for the entire topology it removes.
  assertSheetRangesEditable(sheet, [...ranges, ...removedMerges], userId)
  return { removedMerges, visibleRanges: sheetCommandRangesWithoutFilteredRows(ranges, filteredRows) }
}

/** Sparse cells and resource rectangles are changed once, regardless of selection area or overlap. */
export function applySheetClearCommand(
  sheet: SpreadsheetSheet, ranges: readonly CellRange[], mode: SheetClearMode, userId = 'local-user',
  filteredRows: readonly number[] = sheet.autoFilter?.filteredRows ?? [],
): boolean {
  if (!ranges.length || mode === 'none') return false
  const scope = clearScope(sheet, ranges, sheet.mergedRanges, mode, userId, filteredRows)
  let changed = false
  for (const [key, cell] of Object.entries(sheet.cells)) {
    if (!scope.visibleRanges.some(range => sheetRangeContainsCell(range, cell.row, cell.col))) continue
    if (mode === 'all') { delete sheet.cells[key]; changed = true; continue }
    const next = mode === 'content' ? clearSpreadsheetCellContent(cell)
      : { ...cell, style: defaultCellStyle(), hyperlink: null,
        display: cell.value === null ? cell.display : formatSheetValue(cell.value, 'General') ?? String(cell.value) }
    if (JSON.stringify(next) !== JSON.stringify(cell)) { sheet.cells[key] = next; changed = true }
  }
  if (scope.removedMerges.length) {
    const removed = new Set(scope.removedMerges)
    sheet.mergedRanges = sheet.mergedRanges.filter(merge => !removed.has(merge))
    changed = true
  }
  if (mode === 'all') {
    const validation = trimResourceRanges(sheet.dataValidations, scope.visibleRanges, rebaseValidation)
    const conditional = trimResourceRanges(sheet.conditionalFormats, scope.visibleRanges, rebaseConditional)
    if (JSON.stringify(validation) !== JSON.stringify(sheet.dataValidations)) { sheet.dataValidations = validation; changed = true }
    if (JSON.stringify(conditional) !== JSON.stringify(sheet.conditionalFormats)) { sheet.conditionalFormats = conditional; changed = true }
  }
  return changed
}
export function applySheetViewClearCommand(
  sheet: ExcelSheetView, ranges: readonly CellRange[], mode: SheetClearMode, userId = 'local-user',
): boolean {
  if (!ranges.length || mode === 'none') return false
  const merges = parseViewMerges(sheet), scope = clearScope(sheet, ranges, merges, mode, userId, sheet.filteredRows ?? sheet.autoFilter?.filteredRows ?? [])
  let changed = false
  for (const [key, cell] of Object.entries(sheet.cells)) {
    const position = parseCellKey(key)
    if (!position || !scope.visibleRanges.some(range => sheetRangeContainsCell(range, position.row, position.col))) continue
    if (mode === 'all') { delete sheet.cells[key]; changed = true; continue }
    const next = mode === 'content' ? clearSheetViewCellContent(cell) : clearViewFormat(cell)
    if (JSON.stringify(next) !== JSON.stringify(cell)) { sheet.cells[key] = next; changed = true }
  }
  if (scope.removedMerges.length) {
    const removed = new Set(scope.removedMerges)
    sheet.merges = (sheet.merges ?? []).filter((_, index) => !removed.has(merges[index]!))
    changed = true
  }
  if (mode === 'all') {
    const validation = trimResourceRanges(sheet.dataValidations, scope.visibleRanges, rebaseValidation), conditional = trimResourceRanges(sheet.conditionalFormats, scope.visibleRanges, rebaseConditional)
    if (JSON.stringify(validation) !== JSON.stringify(sheet.dataValidations)) { sheet.dataValidations = validation; changed = true }
    if (JSON.stringify(conditional) !== JSON.stringify(sheet.conditionalFormats)) { sheet.conditionalFormats = conditional; changed = true }
  }
  return changed
}
