import type { CellRange, SheetConditionalFormatRule, SheetDataValidationRule, SpreadsheetSheet } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { conditionalFormatFormulaAnchor, mapConditionalFormatFormulas, reanchorConditionalFormatRule } from '../conditional-formatting/conditional-format-formulas.ts'
import { translateFormulaReferenceAreas } from '../formula/formula-reference-areas.ts'
import { reanchorSheetDataValidationRule, sheetValidationFormulaAnchor, sheetValidationUsesRelativeAnchor } from '../sheet-validation/validation-formulas.ts'
import { moveSheetClipboardFormulaReferences, type SheetClipboardMove } from './sheet-clipboard-references.ts'
import { sheetClipboardRowsToRanges, type SheetClipboardTargetGeometry } from './sheet-clipboard-target.ts'
import { assertSheetCommandRanges, subtractSheetCommandRange } from './sheet-range-command.ts'
import type { SheetClipboardSnapshot } from './sheet-clipboard-types.ts'
import type { SheetPasteMode } from './sheet-paste.ts'

const RANGE_BUDGET = 100_000, RANGE_WORK_BUDGET = 5_000_000
interface RangeWork { count: number }
function countWork(work: RangeWork, count = 1): void {
  work.count += count
  if (work.count > RANGE_WORK_BUDGET) throw new RangeError('剪贴板范围资源超出当前同步计算预算。')
}
function assertRangeBudget(ranges: readonly CellRange[]): void {
  if (ranges.length > RANGE_BUDGET) throw new RangeError('剪贴板范围资源数量超出当前同步预算。')
}
function lowerBound(values: readonly number[], value: number): number {
  let low = 0, high = values.length
  while (low < high) { const middle = (low + high) >>> 1; if (values[middle]! < value) low = middle + 1; else high = middle }
  return low
}

function projectedRuleRanges(
  original: readonly CellRange[], snapshot: SheetClipboardSnapshot, geometry: SheetClipboardTargetGeometry, work: RangeWork,
  rowTile?: number, columnTile?: number,
): CellRange[] {
  const ranges: CellRange[] = []
  for (let row = rowTile ?? 0; row < (rowTile === undefined ? geometry.rows.length : rowTile + snapshot.rows.length); row += snapshot.rows.length) {
    for (let col = columnTile ?? 0; col < (columnTile === undefined ? geometry.columns.length : columnTile + snapshot.columns.length); col += snapshot.columns.length) {
      for (const range of original) {
        countWork(work)
        const left = lowerBound(snapshot.columns, range.startCol), right = lowerBound(snapshot.columns, range.endCol + 1) - 1
        const top = lowerBound(snapshot.rows, range.startRow), bottom = lowerBound(snapshot.rows, range.endRow + 1)
        if (left > right || top >= bottom || left >= snapshot.columns.length || top >= snapshot.rows.length) continue
        const rows = geometry.rows.slice(row + top, row + bottom)
        countWork(work, rows.length)
        const startCol = geometry.columns[col + left], endCol = geometry.columns[col + right]
        if (startCol === undefined || endCol === undefined || rows.length !== bottom - top) throw new RangeError('剪贴板范围资源不能被截断。')
        ranges.push(...sheetClipboardRowsToRanges(rows, startCol, endCol)); assertRangeBudget(ranges)
      }
    }
  }
  return ranges
}

function subtractRanges(ranges: readonly CellRange[], removed: readonly CellRange[], work: RangeWork): CellRange[] {
  let next = ranges.map(range => ({ ...range }))
  for (const cut of removed) {
    countWork(work, next.length)
    next = next.flatMap(range => subtractSheetCommandRange(range, cut)); assertRangeBudget(next)
  }
  return next
}

export function trimSheetClipboardRangeResources(sheet: SpreadsheetSheet, removed: readonly CellRange[], conditional: boolean, validation: boolean): void {
  const work: RangeWork = { count: 0 }
  if (conditional) sheet.conditionalFormats = sheet.conditionalFormats?.flatMap(rule => {
    const ranges = subtractRanges(rule.ranges, removed, work)
    return ranges.length ? [reanchorConditionalFormatRule(rule, ranges)] : []
  })
  if (validation) sheet.dataValidations = sheet.dataValidations?.flatMap(rule => {
    const ranges = subtractRanges(rule.ranges, removed, work)
    return ranges.length ? [reanchorSheetDataValidationRule(rule, ranges)] : []
  })
}

function moveAnchor(anchor: { row: number; col: number }, move: SheetClipboardMove): { row: number; col: number } {
  const next = { row: anchor.row + move.targetRange.startRow - move.sourceRange.startRow,
    col: anchor.col + move.targetRange.startCol - move.sourceRange.startCol }
  assertSheetCommandRanges([{ startRow: next.row, endRow: next.row, startCol: next.col, endCol: next.col }])
  return next
}

function sourceRuleRanges(ranges: readonly CellRange[], snapshot: SheetClipboardSnapshot): CellRange[] {
  const geometry: SheetClipboardTargetGeometry = { rows: snapshot.rows, columns: snapshot.columns, ranges: [], merges: [] }
  return projectedRuleRanges(ranges, snapshot, geometry, { count: 0 })
}
function copiedConditionalRule(rule: SheetConditionalFormatRule, ranges: CellRange[], snapshot: SheetClipboardSnapshot, move?: SheetClipboardMove): SheetConditionalFormatRule {
  if (!move) return reanchorConditionalFormatRule(rule, ranges)
  const source = reanchorConditionalFormatRule(rule, sourceRuleRanges(rule.ranges, snapshot)), origin = conditionalFormatFormulaAnchor(source)
  const next = mapConditionalFormatFormulas(source, formula => moveSheetClipboardFormulaReferences(formula, move.sourceSheetName, move.targetSheetName, move))
  next.ranges = ranges
  if (origin) next.formulaAnchor = moveAnchor(origin, move)
  return next
}

function copiedValidationRule(
  rule: SheetDataValidationRule, ranges: CellRange[], snapshot: SheetClipboardSnapshot, geometry: SheetClipboardTargetGeometry,
  rowTile: number, columnTile: number, move?: SheetClipboardMove,
): SheetDataValidationRule {
  if (move) {
    const next = reanchorSheetDataValidationRule(rule, sourceRuleRanges(rule.ranges, snapshot)), origin = sheetValidationFormulaAnchor(next)
    next.ranges = ranges
    if (sheetValidationUsesRelativeAnchor(rule) && origin) next.formulaAnchor = moveAnchor(origin, move)
    if (rule.type !== 'checkbox') for (const key of ['formula1', 'formula2'] as const) if (next[key]) {
      next[key] = moveSheetClipboardFormulaReferences(next[key]!, move.sourceSheetName, move.targetSheetName, move)
    }
    return next
  }
  const next = reanchorSheetDataValidationRule(rule, ranges)
  if (rule.type === 'list' || rule.type === 'listMultiple') {
    const rowDelta = geometry.rows[rowTile]! - snapshot.rows[0]!, columnDelta = geometry.columns[columnTile]! - snapshot.columns[0]!
    for (const key of ['formula1', 'formula2'] as const) if (next[key] && !next[key]!.trim().startsWith('"')) {
      next[key] = translateFormulaReferenceAreas(next[key]!, rowDelta, columnDelta)
    }
  }
  return next
}

/** CF follows format paste; validation follows all/except-borders, as in the fixed resource hooks. */
export function pasteSheetClipboardRangeResources(
  sheet: SpreadsheetSheet, snapshot: SheetClipboardSnapshot, geometry: SheetClipboardTargetGeometry, mode: SheetPasteMode, move?: SheetClipboardMove,
): void {
  const conditional = mode === 'all' || mode === 'formats' || mode === 'exceptBorders'
  const validation = mode === 'all' || mode === 'exceptBorders'
  if (!conditional && !validation || geometry.contentOnlyMerge || snapshot.externalFormat) return
  trimSheetClipboardRangeResources(sheet, geometry.ranges, conditional, validation)
  const work: RangeWork = { count: 0 }
  if (conditional && snapshot.conditionalFormats?.length) {
    let priority = (sheet.conditionalFormats ?? []).reduce((maximum, rule) => Math.max(maximum, rule.priority), 0)
    const additions = [...snapshot.conditionalFormats].sort((first, second) => first.priority - second.priority).flatMap(rule => {
      const ranges = projectedRuleRanges(rule.ranges, snapshot, geometry, work)
      return ranges.length ? [{ ...copiedConditionalRule(rule, ranges, snapshot, move),
        id: `clipboard-cf-${crypto.randomUUID()}`, priority: ++priority }] : []
    })
    sheet.conditionalFormats = [...(sheet.conditionalFormats ?? []), ...additions]
  }
  if (validation && snapshot.dataValidations?.length) {
    const additions: SheetDataValidationRule[] = []
    for (let row = 0; row < geometry.rows.length; row += snapshot.rows.length) for (let col = 0; col < geometry.columns.length; col += snapshot.columns.length) {
      for (const rule of snapshot.dataValidations) {
        const ranges = projectedRuleRanges(rule.ranges, snapshot, geometry, work, row, col)
        if (ranges.length) additions.push({ ...copiedValidationRule(rule, ranges, snapshot, geometry, row, col, move), id: `clipboard-dv-${crypto.randomUUID()}` })
        if (additions.length > RANGE_BUDGET) throw new RangeError('剪贴板数据验证规则数量超出当前同步预算。')
      }
    }
    sheet.dataValidations = [...(sheet.dataValidations ?? []), ...additions]
  }
}

export function moveSheetClipboardLinkedRuleReferences(sheet: SpreadsheetSheet, move: SheetClipboardMove): void {
  sheet.conditionalFormats = sheet.conditionalFormats?.map(rule => mapConditionalFormatFormulas(rule,
    formula => moveSheetClipboardFormulaReferences(formula, sheet.name, sheet.name, move)))
  sheet.dataValidations = sheet.dataValidations?.map(rule => {
    const next = cloneModel(rule)
    if (rule.type !== 'checkbox') for (const key of ['formula1', 'formula2'] as const) if (next[key]) {
      next[key] = moveSheetClipboardFormulaReferences(next[key]!, sheet.name, sheet.name, move)
    }
    return next
  })
}
