import type { CellRange, CellScalar, SheetDataValidationRule, SpreadsheetCell, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import { sheetModelCellViewAt } from '../../services/excelViewAdapter.ts'
import { sheetRangeContainsCell, sheetRangesIntersect, subtractSheetCommandRange } from '../sheet-workbench/sheet-range-command.ts'
import { normalizeSheetCheckboxValue, sheetCheckboxRuleValues, type SheetCheckboxRuleValues } from './validation-checkbox.ts'
import { sheetValidationFormulaValue, sheetValidationSourceReader } from './validation-value-read.ts'
import { createSheetValidationScalarCell } from './validation-widget-input.ts'

export interface SheetCheckboxRuleCellPlan { cells: SpreadsheetCell[]; ranges: CellRange[] }

/** Rule ranges can overlap or describe whole axes; initialization visits their union within the worksheet dimensions. */
export function sheetCheckboxWorksheetRanges(sheet: SpreadsheetSheet, ranges: readonly CellRange[]): CellRange[] {
  const rows = sheet.rowCount ?? 1_048_576, columns = sheet.columnCount ?? 16_384, union: CellRange[] = []
  for (const source of ranges) {
    const clipped = { ...source, endRow: Math.min(source.endRow, rows - 1), endCol: Math.min(source.endCol, columns - 1) }
    if (clipped.startRow > clipped.endRow || clipped.startCol > clipped.endCol) continue
    let remaining = [clipped]
    for (const previous of union) remaining = remaining.flatMap(range => subtractSheetCommandRange(range, previous))
    union.push(...remaining)
  }
  return union
}

/** Observer ranges carry only the actual scalar changes, coalescing contiguous row runs and equivalent adjoining rows. */
function changedRanges(cells: readonly SpreadsheetCell[]): CellRange[] {
  const rows: CellRange[] = []
  for (const cell of cells) {
    const previous = rows.at(-1)
    if (previous && previous.startRow === cell.row && previous.endCol + 1 === cell.col) previous.endCol = cell.col
    else rows.push({ startRow: cell.row, endRow: cell.row, startCol: cell.col, endCol: cell.col })
  }
  const result: CellRange[] = [], lastByColumns = new Map<string, CellRange>()
  for (const range of rows) {
    const key = `${range.startCol}:${range.endCol}`, previous = lastByColumns.get(key)
    if (previous && previous.endRow + 1 === range.startRow) previous.endRow = range.endRow
    else { const detached = { ...range }; result.push(detached); lastByColumns.set(key, detached) }
  }
  return result
}

/** A prepared rule edit materializes missing unchecked scalars and migrates matching old states in the same content transaction. */
export interface SheetCheckboxRuleCellPreparation {
  ranges: CellRange[]; values: SheetCheckboxRuleValues; oldValues?: SheetCheckboxRuleValues
  previous?: SheetDataValidationRule; migrate: boolean
}
export function prepareSheetCheckboxRuleCells(workbook: WorkbookModel, sheet: SpreadsheetSheet, rule: SheetDataValidationRule,
  previous: SheetDataValidationRule | undefined, userId: string): SheetCheckboxRuleCellPreparation | undefined {
  if (rule.type !== 'checkbox') return undefined
  const oldCheckbox = previous?.type === 'checkbox' ? previous : undefined
  const rangesChanged = !oldCheckbox || JSON.stringify(oldCheckbox.ranges) !== JSON.stringify(rule.ranges)
  const definitionChanged = !oldCheckbox || oldCheckbox.formula1 !== rule.formula1 || oldCheckbox.formula2 !== rule.formula2
    || JSON.stringify(oldCheckbox.formulaAnchor) !== JSON.stringify(rule.formulaAnchor)
  if (!rangesChanged && !definitionChanged) return undefined
  const canRead = sheetValidationSourceReader(userId), ranges = sheetCheckboxWorksheetRanges(sheet, rule.ranges)
  const readableRanges = [...ranges, ...sheet.mergedRanges.filter(merge => ranges.some(range => sheetRangesIntersect(range, merge)))]
  if (canRead && readableRanges.some(range => !canRead(sheet, range))) throw new Error('当前用户没有查看复选框初始化区域的权限。')
  const read = (formula: string, anchor: { row: number; column: number }) =>
    sheetValidationFormulaValue(formula, workbook, sheet.id, anchor, canRead)
  const values = sheetCheckboxRuleValues(rule, read)
  if (!values) throw new Error('复选框的勾选或未勾选值无法读取，请检查规则公式和来源权限。')
  const oldValues = oldCheckbox ? sheetCheckboxRuleValues(oldCheckbox, read) : undefined
  const migrate = !!oldValues && (definitionChanged || !Object.is(oldValues.checked, values.checked) || !Object.is(oldValues.unchecked, values.unchecked))
  return { ranges, values, oldValues, previous: oldCheckbox, migrate }
}

/** Scalar classification is shared by the small planner and the bulk Worker adapter. Rust receives only these typed actions. */
export function sheetCheckboxRuleCellAction(prepared: SheetCheckboxRuleCellPreparation, cell: SpreadsheetCell | undefined, row: number, column: number): 0 | 1 | 2 {
  if (cell?.formula || cell?.richText?.length || cell?.image) return 0
  const actual = cell?.value, blank = actual === undefined || actual === null || actual === ''
  if (blank) return 1
  if (!prepared.migrate || !prepared.previous!.ranges.some(source => sheetRangeContainsCell(source, row, column))) return 0
  const state = normalizeSheetCheckboxValue(actual)
  const action = state === normalizeSheetCheckboxValue(prepared.oldValues!.unchecked) ? 1
    : state === normalizeSheetCheckboxValue(prepared.oldValues!.checked) ? 2 : 0
  const next = action === 1 ? prepared.values.unchecked : prepared.values.checked
  return action && !Object.is(next, actual) ? action : 0
}

export function createSheetCheckboxRuleCell(workbook: WorkbookModel, sheet: SpreadsheetSheet, row: number, column: number,
  value: CellScalar, region: string): SpreadsheetCell {
  return createSheetValidationScalarCell({ workbook, sheet, position: { row, column }, view: sheetModelCellViewAt(workbook, sheet, row, column),
    date1904: workbook.date1904, region }, value, typeof value === 'number' ? 'number' : typeof value === 'boolean' ? 'boolean' : 'text')
}

export function planSheetCheckboxRuleCells(workbook: WorkbookModel, sheet: SpreadsheetSheet, rule: SheetDataValidationRule,
  previous: SheetDataValidationRule | undefined, userId: string, region: string): SheetCheckboxRuleCellPlan {
  const prepared = prepareSheetCheckboxRuleCells(workbook, sheet, rule, previous, userId)
  if (!prepared) return { cells: [], ranges: [] }
  const cells: SpreadsheetCell[] = []
  for (const range of prepared.ranges) for (let row = range.startRow; row <= range.endRow; row++) {
    for (let column = range.startCol; column <= range.endCol; column++) {
      const merge = sheet.mergedRanges.find(item => sheetRangeContainsCell(item, row, column))
      if (merge && (merge.startRow !== row || merge.startCol !== column)) { column = Math.min(merge.endCol, range.endCol); continue }
      const cell = sheet.cells[`${row}:${column}`]
      const action = sheetCheckboxRuleCellAction(prepared, cell, row, column)
      if (action) cells.push(createSheetCheckboxRuleCell(workbook, sheet, row, column, action === 1 ? prepared.values.unchecked : prepared.values.checked, region))
    }
  }
  cells.sort((left, right) => left.row - right.row || left.col - right.col)
  return { cells, ranges: changedRanges(cells) }
}
