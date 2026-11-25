import type { CellRange, SheetConditionalFormatRule, SheetDataValidationRule, SpreadsheetSheet } from '../../domain/model.ts'
import { sheetColumnWidth } from '../../domain/sheet-axis.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { conditionalFormatFormulaAnchor, mapConditionalFormatFormulas } from '../conditional-formatting/conditional-format-formulas.ts'
import { formulaRelativeReferenceSignature } from '../formula/formula-references.ts'
import { sheetValidationFormulaAnchor, sheetValidationUsesRelativeAnchor } from '../sheet-validation/validation-formulas.ts'
import { sheetRangesIntersect, subtractSheetCommandRange } from './sheet-range-command.ts'
import { sheetRangeThemeChangedRanges } from './sheet-range-theme-structure.ts'

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
  // Fragment ID changes do not change a retained predicate. Group by its
  // per-cell definition so split resources do not widen range permissions.
  const ordered = (rules: readonly R[]) => rules.map(rule => ({ rule, key: definition(rule) }))
  const firstEntries = ordered(before), nextEntries = ordered(after)
  const group = (entries: typeof firstEntries) => {
    const result = new Map<string, CellRange[]>()
    for (const { rule, key } of entries) {
      const ranges = result.get(key) ?? []
      ranges.push(...rule.ranges); result.set(key, ranges)
    }
    return result
  }
  const originals = group(firstEntries), updated = group(nextEntries)
  const result: CellRange[] = [], work = { count: 0 }
  for (const key of new Set([...originals.keys(), ...updated.keys()])) {
    const first = originals.get(key) ?? [], next = updated.get(key) ?? []
    result.push(...rangeDifference(first, next, work), ...rangeDifference(next, first, work))
    if (result.length > MAX_RANGES) throw new RangeError('剪贴板资源差异范围过多。')
  }
  // Compare the actual per-cell predicate sequence when precedence can change.
  // Repeated definitions may move across another predicate without changing
  // the global definition list, so retain their order inside the coverage map.
  const order = (entries: typeof firstEntries) => entries.map(entry => entry.key).filter((key, index, keys) => index === 0 || key !== keys[index - 1])
  const firstOrder = order(firstEntries), nextOrder = order(nextEntries)
  if (JSON.stringify(firstOrder) !== JSON.stringify(nextOrder)
    || new Set(firstOrder).size !== firstOrder.length || new Set(nextOrder).size !== nextOrder.length) {
    const coverage = (entries: typeof firstEntries) => {
      let regions: Array<{ range: CellRange; keys: string[] }> = []
      for (const { rule, key } of entries) {
        const masks: CellRange[] = []
        for (const range of rule.ranges) {
          masks.push(...rangeDifference([range], masks, work))
          if (masks.length > MAX_RANGES) throw new RangeError('剪贴板规则优先级差异范围过多。')
        }
        const next: typeof regions = []
        for (const region of regions) {
          next.push(...rangeDifference([region.range], masks, work).map(range => ({ range, keys: region.keys })))
          for (const mask of masks) {
            if (++work.count > MAX_RANGE_WORK) throw new RangeError('剪贴板规则优先级差异超出当前同步计算预算。')
            if (!sheetRangesIntersect(region.range, mask)) continue
            next.push({ range: { startRow: Math.max(region.range.startRow, mask.startRow), endRow: Math.min(region.range.endRow, mask.endRow),
              startCol: Math.max(region.range.startCol, mask.startCol), endCol: Math.min(region.range.endCol, mask.endCol) },
              keys: region.keys.at(-1) === key ? region.keys : [...region.keys, key] })
          }
          if (next.length > MAX_RANGES) throw new RangeError('剪贴板规则优先级差异范围过多。')
        }
        next.push(...rangeDifference(masks, regions.map(region => region.range), work).map(range => ({ range, keys: [key] })))
        if (next.length > MAX_RANGES) throw new RangeError('剪贴板规则优先级差异范围过多。')
        const groups = new Map<string, { keys: string[]; ranges: CellRange[] }>()
        for (const region of next) {
          const id = JSON.stringify(region.keys), group = groups.get(id) ?? { keys: region.keys, ranges: [] }
          group.ranges.push(region.range); groups.set(id, group)
        }
        regions = [...groups.values()].flatMap(group => compactSheetClipboardRanges(group.ranges).map(range => ({ range, keys: group.keys })))
      }
      const result = new Map<string, CellRange[]>()
      for (const region of regions) {
        const id = JSON.stringify(region.keys), ranges = result.get(id) ?? []
        ranges.push(region.range); result.set(id, ranges)
      }
      return result
    }
    const oldCoverage = coverage(firstEntries), newCoverage = coverage(nextEntries)
    for (const key of new Set([...oldCoverage.keys(), ...newCoverage.keys()])) {
      const first = oldCoverage.get(key) ?? [], next = newCoverage.get(key) ?? []
      result.push(...rangeDifference(first, next, work), ...rangeDifference(next, first, work))
      if (result.length > MAX_RANGES) throw new RangeError('剪贴板资源差异范围过多。')
    }
  }
  return result
}

/** Reanchoring an unchanged predicate affects only its added/removed membership, not every retained cell. */
export function sheetClipboardChangedRuleRanges(before: SpreadsheetSheet, after: SpreadsheetSheet): CellRange[] {
  return compactSheetClipboardRanges([
    ...changedRules(before.conditionalFormats ?? [], after.conditionalFormats ?? [], conditionalDefinition),
    ...changedRules(before.dataValidations ?? [], after.dataValidations ?? [], validationDefinition),
    ...sheetRangeThemeChangedRanges(before.rangeThemeRules ?? [], after.rangeThemeRules ?? []),
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
