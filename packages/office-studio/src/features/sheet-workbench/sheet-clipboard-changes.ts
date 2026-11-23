import type { CellRange, SheetConditionalFormatRule, SheetDataValidationRule, SpreadsheetSheet } from '../../domain/model.ts'
import { sheetColumnWidth } from '../../domain/sheet-axis.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { conditionalFormatFormulaAnchor, mapConditionalFormatFormulas } from '../conditional-formatting/conditional-format-formulas.ts'
import { formulaRelativeReferenceSignature } from '../formula/formula-references.ts'
import { sheetValidationFormulaAnchor, sheetValidationUsesRelativeAnchor } from '../sheet-validation/validation-formulas.ts'
import { subtractSheetCommandRange } from './sheet-range-command.ts'

const MAX_RANGE_WORK = 5_000_000, MAX_RANGES = 100_000

/** Coalesce adjacent points without replacing sparse changes by a bounding rectangle. */
export function compactSheetClipboardRanges(input: readonly CellRange[]): CellRange[] {
  const rows = new Map<string, CellRange[]>()
  for (const range of input) {
    const key = `${range.startRow}:${range.endRow}`, entries = rows.get(key) ?? []
    entries.push({ ...range }); rows.set(key, entries)
  }
  const horizontal: CellRange[] = []
  for (const entries of rows.values()) {
    entries.sort((a, b) => a.startCol - b.startCol || a.endCol - b.endCol)
    let previous: CellRange | undefined
    for (const range of entries) {
      if (previous && range.startCol <= previous.endCol + 1) previous.endCol = Math.max(previous.endCol, range.endCol)
      else { previous = range; horizontal.push(previous) }
    }
  }
  horizontal.sort((a, b) => a.startRow - b.startRow || a.startCol - b.startCol || a.endRow - b.endRow)
  const result: CellRange[] = [], previousByColumns = new Map<string, CellRange>()
  for (const range of horizontal) {
    const key = `${range.startCol}:${range.endCol}`, previous = previousByColumns.get(key)
    if (previous && range.startRow <= previous.endRow + 1 && range.startRow >= previous.startRow) previous.endRow = Math.max(previous.endRow, range.endRow)
    else { result.push(range); previousByColumns.set(key, range) }
  }
  return result
}

function conditionalDefinition(rule: SheetConditionalFormatRule): string {
  const anchor = conditionalFormatFormulaAnchor(rule)
  const normalized = anchor ? mapConditionalFormatFormulas(rule, formula => formulaRelativeReferenceSignature(formula, anchor.row, anchor.col)) : cloneModel(rule)
  const { id: _id, ranges: _ranges, formulaAnchor: _anchor, officeSource: _source, metadata: _metadata, ...definition } = normalized
  return JSON.stringify(definition)
}
function validationDefinition(rule: SheetDataValidationRule): string {
  const anchor = sheetValidationFormulaAnchor(rule), normalized = cloneModel(rule)
  if (anchor && sheetValidationUsesRelativeAnchor(rule)) for (const key of ['formula1', 'formula2'] as const) {
    if (normalized[key]) normalized[key] = formulaRelativeReferenceSignature(normalized[key]!, anchor.row, anchor.col)
  }
  const { id: _id, ranges: _ranges, formulaAnchor: _anchor, metadata: _metadata, ...definition } = normalized
  return JSON.stringify(definition)
}
function rangeDifference(original: readonly CellRange[], removed: readonly CellRange[], work: { count: number }): CellRange[] {
  if (JSON.stringify(original) === JSON.stringify(removed)) return []
  let result = original.map(range => ({ ...range }))
  for (const cut of removed) {
    work.count += result.length
    if (work.count > MAX_RANGE_WORK) throw new RangeError('剪贴板资源差异超出当前同步计算预算。')
    result = result.flatMap(range => subtractSheetCommandRange(range, cut))
    if (result.length > MAX_RANGES) throw new RangeError('剪贴板资源差异范围过多。')
  }
  return result
}
function changedRules<R extends { id: string; ranges: CellRange[] }>(before: readonly R[], after: readonly R[], definition: (rule: R) => string): CellRange[] {
  const originals = new Map(before.map(rule => [rule.id, rule])), updated = new Map(after.map(rule => [rule.id, rule]))
  const result: CellRange[] = [], work = { count: 0 }
  for (const id of new Set([...originals.keys(), ...updated.keys()])) {
    const first = originals.get(id), next = updated.get(id)
    if (!first) result.push(...next!.ranges)
    else if (!next) result.push(...first.ranges)
    else if (definition(first) !== definition(next)) result.push(...first.ranges, ...next.ranges)
    else result.push(...rangeDifference(first.ranges, next.ranges, work), ...rangeDifference(next.ranges, first.ranges, work))
    if (result.length > MAX_RANGES) throw new RangeError('剪贴板资源差异范围过多。')
  }
  return result
}

/** Reanchoring an unchanged predicate affects only its added/removed membership, not every retained cell. */
export function sheetClipboardChangedRuleRanges(before: SpreadsheetSheet, after: SpreadsheetSheet): CellRange[] {
  return compactSheetClipboardRanges([
    ...changedRules(before.conditionalFormats ?? [], after.conditionalFormats ?? [], conditionalDefinition),
    ...changedRules(before.dataValidations ?? [], after.dataValidations ?? [], validationDefinition),
  ])
}

export function sheetClipboardChangedSheetRanges(before: SpreadsheetSheet, after: SpreadsheetSheet): CellRange[] {
  const result: CellRange[] = []
  for (const key of new Set([...Object.keys(before.cells), ...Object.keys(after.cells)])) {
    if (JSON.stringify(before.cells[key]) === JSON.stringify(after.cells[key])) continue
    const cell = after.cells[key] ?? before.cells[key]!
    result.push({ startRow: cell.row, endRow: cell.row, startCol: cell.col, endCol: cell.col })
  }
  result.push(...sheetClipboardChangedRuleRanges(before, after))
  const oldMerges = new Set(before.mergedRanges.map(range => JSON.stringify(range))), newMerges = new Set(after.mergedRanges.map(range => JSON.stringify(range)))
  result.push(...before.mergedRanges.filter(range => !newMerges.has(JSON.stringify(range))), ...after.mergedRanges.filter(range => !oldMerges.has(JSON.stringify(range))))
  result.push(...sheetClipboardChangedColumnWidthRanges(before, after))
  return compactSheetClipboardRanges(result)
}

/** Column geometry is independent of which optional paste mode caused the replacement. */
export function sheetClipboardChangedColumnWidthRanges(before: SpreadsheetSheet, after: SpreadsheetSheet): CellRange[] {
  const result: CellRange[] = []
  const columns = new Set<number>([0, Math.min(16_384, Math.max(before.columnCount ?? 0, after.columnCount ?? 0))])
  for (const sheet of [before, after]) {
    for (const key of Object.keys(sheet.columnWidths)) { columns.add(Number(key)); columns.add(Number(key) + 1) }
    for (const span of sheet.columnWidthSpans ?? []) { columns.add(span.start); columns.add(span.end + 1) }
  }
  const boundaries = [...columns].filter(column => column >= 0 && column <= 16_384).sort((a, b) => a - b)
  boundaries.forEach((column, index) => {
    const end = boundaries[index + 1]
    if (end !== undefined && end > column && sheetColumnWidth(before, column) !== sheetColumnWidth(after, column)) {
      result.push({ startRow: 0, endRow: Math.max(1, before.rowCount ?? 1_048_576, after.rowCount ?? 1_048_576) - 1, startCol: column, endCol: end - 1 })
    }
  })
  return compactSheetClipboardRanges(result)
}
