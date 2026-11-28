import type { CellRange, SheetConditionalFormatRule, SheetDataValidationRule, WorkbookModel } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { inverseSheetCellShiftPosition, mapSheetCellShiftPosition, sheetCellShiftAffectedRange, sheetCellShiftContainsBand,
  transformSheetCellShiftRange, type SheetCellShiftEdit } from '../../domain/sheet-cell-shift.ts'
import { conditionalFormatFormulaAnchor, mapConditionalFormatFormulas } from '../conditional-formatting/conditional-format-formulas.ts'
import { sheetValidationFormulaAnchor, sheetValidationUsesRelativeAnchor } from '../sheet-validation/validation-formulas.ts'
import { readFormulaReferenceArea, transformFormulaReferenceAreasForCellShift, translateFormulaReferenceAreas } from '../formula/formula-reference-areas.ts'
import { rewriteFormulaImportCell } from '../formula/formula-import-context.ts'
import { rewriteWorkbookFormulaDefinedNames } from '../formula/formula-defined-names.ts'
import { compactSheetClipboardRanges } from './sheet-clipboard-changes.ts'
import { sheetClipboardObjectAxes } from './sheet-clipboard-object-geometry.ts'
import { resolveSlicerAnchorBounds } from '../slicer/slicer-drawing.ts'
import { formatProtectionRange } from '../sheet-protection/range-protection.ts'
import { parseSheetHyperlinkTarget } from './sheet-hyperlink.ts'
import { sheetRangesIntersect } from './sheet-range-command.ts'
import { transformSheetImagesForCellShift } from '../sheet-images/sheet-image-structure.ts'
import { transformSheetAutoFilterForAxisEdit, transformSheetChartForAxisEdit, transformSheetTableForAxisEdit } from './sheet-axis-structure-resources.ts'
import { reconcileSlicerConnectionResources } from '../slicer/slicer-connections-lifecycle.ts'
import { refreshPivotSourceWorksheetOutputs, transformWorkbookPivotSourceReports } from '../pivot/services/pivot-source-lifecycle.ts'
import { synchronizeLinkedPivotChartReports } from '../pivot/services/pivot-chart-report-lifecycle.ts'

const MAX_RULE_FRAGMENTS = 100_000
function assertBudget(length: number): void {
  if (length > MAX_RULE_FRAGMENTS) throw new RangeError('单元格移位资源片段超出当前同步计算预算。')
}
const reference = (formula: string, owner: string | undefined, target: string, edit: SheetCellShiftEdit) =>
  transformFormulaReferenceAreasForCellShift(formula, edit, owner, target)
function oneRange(range: CellRange, edit: SheetCellShiftEdit, label: string): CellRange | undefined {
  const pieces = compactSheetClipboardRanges(transformSheetCellShiftRange(range, edit))
  if (pieces.length > 1) throw new RangeError(`${label}需要连续的矩形范围，当前局部移位会将它拆开。请连同整个范围一起移动。`)
  return pieces[0]
}
function referenceRange(range: CellRange, edit: SheetCellShiftEdit, owner: string): CellRange | undefined {
  const formula = reference(formatProtectionRange(range), owner, owner, edit), parsed = readFormulaReferenceArea(formula, 0)
  return parsed && parsed.end === formula.length ? {
    startRow: Math.min(parsed.area.startRow, parsed.area.endRow), endRow: Math.max(parsed.area.startRow, parsed.area.endRow),
    startCol: Math.min(parsed.area.startColumn, parsed.area.endColumn), endCol: Math.max(parsed.area.startColumn, parsed.area.endColumn),
  } : undefined
}
function rebaseFormula(formula: string, origin: { row: number; col: number } | undefined, range: CellRange,
  owner: string, target: string, edit: SheetCellShiftEdit): string {
  const source = inverseSheetCellShiftPosition(range.startRow, range.startCol, edit)
  if (!source) throw new RangeError('无法确定移位资源的公式起点。')
  return reference(origin ? translateFormulaReferenceAreas(formula, source.row - origin.row, source.col - origin.col) : formula, owner, target, edit)
}
function conditionalRules(rules: readonly SheetConditionalFormatRule[], local: boolean, owner: string, target: string,
  edit: SheetCellShiftEdit): SheetConditionalFormatRule[] {
  const result: SheetConditionalFormatRule[] = []
  for (const rule of rules) {
    if (!local) { result.push(mapConditionalFormatFormulas(rule, formula => reference(formula, owner, target, edit))); continue }
    const origin = conditionalFormatFormulaAnchor(rule)
    let index = 0
    for (const original of rule.ranges) for (const range of transformSheetCellShiftRange(original, edit)) {
      const next = mapConditionalFormatFormulas(rule, formula => rebaseFormula(formula, origin, range, owner, target, edit))
      next.id = index++ ? `cf-${crypto.randomUUID()}` : rule.id
      next.ranges = [{ ...range }]
      if (origin) next.formulaAnchor = { row: range.startRow, col: range.startCol }
      result.push(next); assertBudget(result.length)
    }
  }
  assertBudget(result.length); return result
}
function validationFormula(formula: string | undefined, rule: SheetDataValidationRule, owner: string, target: string,
  edit: SheetCellShiftEdit, range?: CellRange): string | undefined {
  if (!formula || rule.type === 'checkbox' || (rule.type === 'list' || rule.type === 'listMultiple') && !formula.trimStart().startsWith('=')) return formula
  return range && sheetValidationUsesRelativeAnchor(rule)
    ? rebaseFormula(formula, sheetValidationFormulaAnchor(rule), range, owner, target, edit) : reference(formula, owner, target, edit)
}
function validationRules(rules: readonly SheetDataValidationRule[], local: boolean, owner: string, target: string,
  edit: SheetCellShiftEdit): SheetDataValidationRule[] {
  const result: SheetDataValidationRule[] = []
  for (const rule of rules) {
    if (!local) { result.push({ ...cloneModel(rule), formula1: validationFormula(rule.formula1, rule, owner, target, edit),
      formula2: validationFormula(rule.formula2, rule, owner, target, edit) }); continue }
    let index = 0
    for (const original of rule.ranges) for (const range of transformSheetCellShiftRange(original, edit)) {
      const next = { ...cloneModel(rule), id: index++ ? `dv-${crypto.randomUUID()}` : rule.id, ranges: [{ ...range }],
        formula1: validationFormula(rule.formula1, rule, owner, target, edit, range), formula2: validationFormula(rule.formula2, rule, owner, target, edit, range) }
      if (sheetValidationUsesRelativeAnchor(rule)) next.formulaAnchor = { row: range.startRow, col: range.startCol }
      result.push(next); assertBudget(result.length)
    }
  }
  assertBudget(result.length); return result
}
function internalHyperlink(value: string, original: WorkbookModel, targetId: string, targetName: string, edit: SheetCellShiftEdit): string {
  const link = parseSheetHyperlinkTarget(value, original)
  if (link?.type !== 'internal' || link.sheetId !== targetId) return value
  const next = referenceRange(link.range, edit, targetName), address = next ? formatProtectionRange(next) : '#REF!'
  if (value.slice(1).includes('=')) { const params = new URLSearchParams(value.slice(1)); params.set('range', address); return `#${params}` }
  return `${value.slice(0, value.lastIndexOf('!') + 1)}${address}`
}
/** All resources read original coordinates and write only a detached workbook candidate. */
export function transformWorkbookCellShiftResources(workbook: WorkbookModel, original: WorkbookModel, targetId: string, edit: SheetCellShiftEdit): void {
  const target = original.sheets.find(sheet => sheet.id === targetId)
  if (!target) throw new Error('找不到移位目标工作表。')
  const affected = sheetCellShiftAffectedRange(edit)
  for (const sheet of workbook.sheets) {
    const before = original.sheets.find(item => item.id === sheet.id)!, local = sheet.id === targetId
    if (before.conditionalFormats) sheet.conditionalFormats = conditionalRules(before.conditionalFormats, local, sheet.name, target.name, edit)
    if (before.dataValidations) sheet.dataValidations = validationRules(before.dataValidations, local, sheet.name, target.name, edit)
    for (const cell of Object.values(sheet.cells)) {
      if (cell.formula) rewriteFormulaImportCell(cell, reference(cell.formula, sheet.name, target.name, edit))
      if (cell.sparkline) cell.sparkline.sourceRange = reference(cell.sparkline.sourceRange, sheet.name, target.name, edit)
      if (cell.hyperlink) cell.hyperlink = internalHyperlink(cell.hyperlink, original, targetId, target.name, edit)
      if (cell.richText) cell.richText = cell.richText.map(run => typeof run.hyperlink === 'string'
        ? { ...run, hyperlink: internalHyperlink(run.hyperlink, original, targetId, target.name, edit) } : run)
    }
    if (!local) continue
    if (before.rangeProtectionRules) sheet.rangeProtectionRules = before.rangeProtectionRules.flatMap(rule => {
      const ranges = compactSheetClipboardRanges(rule.ranges.flatMap(range => transformSheetCellShiftRange(range, edit)))
      assertBudget(ranges.length); return ranges.length ? [{ ...cloneModel(rule), ranges }] : []
    })
    if (before.autoFilter) {
      const range = oneRange(before.autoFilter.range, edit, '自动筛选')
      sheet.autoFilter = !range ? undefined : sheetCellShiftContainsBand(before.autoFilter.range, edit)
        ? transformSheetAutoFilterForAxisEdit(before.autoFilter, edit) : { ...cloneModel(before.autoFilter), range }
    }
  }
  if (original.tables) workbook.tables = original.tables.flatMap(table => {
    if (table.sheetId !== targetId) return [cloneModel(table)]
    const range = oneRange(table.range, edit, `表格“${table.name}”`)
    if (!range) return []
    return sheetCellShiftContainsBand(table.range, edit) ? transformSheetTableForAxisEdit(table, edit) : [{ ...cloneModel(table), range }]
  })
  if (original.charts) workbook.charts = original.charts.flatMap(chart => {
    if ((chart.source.sheetId ?? chart.sheetId) !== targetId || chart.pivotTableId || chart.ownedPivot) return [cloneModel(chart)]
    if (!sheetRangesIntersect(chart.source.range, affected)) return [cloneModel(chart)]
    const range = referenceRange(chart.source.range, edit, target.name)
    if (!range) throw new RangeError('移位会删除图表的完整数据源，请先调整图表数据绑定。')
    if (sheetCellShiftContainsBand(chart.source.range, edit)) {
      const next = transformSheetChartForAxisEdit(chart, edit)
      return next ? [next] : []
    }
    return [{ ...cloneModel(chart), source: { ...cloneModel(chart.source), range },
      revision: chart.revision + (sheetRangesIntersect(chart.source.range, affected) ? 1 : 0) }]
  })
  const pivotTransform = {
    sheetId: targetId,
    range: (range: CellRange) => referenceRange(range, edit, target.name),
    outputRange: (range: CellRange) => oneRange(range, edit, '透视输出'),
    point: (row: number, col: number) => mapSheetCellShiftPosition(row, col, edit),
    preserveGeneratedCells: true,
    sourceDirty: (range: CellRange) => sheetRangesIntersect(range, affected),
    ...(edit.axis === 'column' ? { column: (column: number, sourceRange: CellRange) => sheetCellShiftContainsBand(sourceRange, edit)
      ? mapSheetCellShiftPosition(sourceRange.startRow, column, edit)?.col : column } : {}),
  }
  transformWorkbookPivotSourceReports(original, workbook, pivotTransform)
  const tableIds = new Set((workbook.tables ?? []).map(table => table.id))
  let slicerAxes: ReturnType<typeof sheetClipboardObjectAxes> | undefined
  if (original.slicers) workbook.slicers = original.slicers.flatMap(slicer => {
    if (slicer.tableId && original.tables?.some(table => table.id === slicer.tableId) && !tableIds.has(slicer.tableId)) return []
    const next = cloneModel(slicer)
    if (slicer.hostSheetId === targetId && (!original.id || slicer.unitId === original.id) && slicer.drawing?.anchorType !== 'none') {
      const mapped = mapSheetCellShiftPosition(slicer.position.row, slicer.position.column, edit)
      if (mapped) next.position = { row: mapped.row, column: mapped.col }
      else next.position = edit.axis === 'row' ? { ...next.position, row: edit.start } : { ...next.position, column: edit.start }
      if (next.drawing) for (const point of [next.drawing.from, ...(next.drawing.anchorType === 'both' ? [next.drawing.to] : [])]) {
        const mapped = mapSheetCellShiftPosition(point.row, point.column, edit)
        if (mapped) { point.row = mapped.row; point.column = mapped.col }
        else if (edit.axis === 'row') point.row = edit.start; else point.column = edit.start
      }
      slicerAxes ??= sheetClipboardObjectAxes(workbook, workbook.sheets.find(sheet => sheet.id === targetId)!)
      if (next.drawing) {
        const bounds = resolveSlicerAnchorBounds(next, slicerAxes)
        // Hidden/deleted Both markers may collapse; retain positive author layout for restoration and settings.
        if (bounds.width > 0 && bounds.height > 0) next.layout = bounds
      }
      else if (next.layout) {
        next.layout.left += slicerAxes.columns.offset(next.position.column) - slicerAxes.columns.offset(slicer.position.column)
        next.layout.top += slicerAxes.rows.offset(next.position.row) - slicerAxes.rows.offset(slicer.position.row)
      }
    }
    if (next.tableId) {
      const table = workbook.tables?.find(table => table.id === next.tableId)
      if (table && !table.columns.some(column => column.id === next.sourceFieldId)) return []
    }
    if (JSON.stringify(next) !== JSON.stringify(slicer)) next.revision++
    return [next]
  })
  rewriteWorkbookFormulaDefinedNames(workbook, (formula, owner) => reference(formula, owner, target.name, edit))
  const after = workbook.sheets.find(sheet => sheet.id === targetId)!
  if (target.images) after.images = transformSheetImagesForCellShift(target.images, edit, after,
    { before: sheetClipboardObjectAxes(original, target), after: sheetClipboardObjectAxes(workbook, after) })
  reconcileSlicerConnectionResources(original, workbook, pivotTransform)
  refreshPivotSourceWorksheetOutputs(original, workbook)
  synchronizeLinkedPivotChartReports(original, workbook)
  let overlapWork = 0
  const budget = () => { if (++overlapWork > 5_000_000) throw new RangeError('单元格移位对象重叠检查超出当前同步计算预算。') }
  const tables = workbook.tables ?? [], pivots = workbook.pivotTables ?? []
  for (const [index, table] of tables.entries()) for (const other of tables.slice(index + 1)) {
    budget()
    if (other.sheetId === table.sheetId && sheetRangesIntersect(other.range, table.range)) throw new RangeError('移位后结构化表格范围重叠。')
  }
  for (const [index, pivot] of pivots.entries()) {
    if (!pivot.outputRange) continue
    if (pivot.source.subUnitId === pivot.anchor.subUnitId && sheetRangesIntersect(pivot.outputRange, pivot.source.range)) throw new RangeError('移位后透视输出覆盖其数据源。')
    for (const table of tables) {
      budget()
      if (table.sheetId === pivot.anchor.subUnitId && sheetRangesIntersect(table.range, pivot.outputRange)) throw new RangeError('移位后透视输出与结构化表格重叠。')
    }
    for (const other of pivots.slice(index + 1)) {
      budget()
      if (other.outputRange && other.anchor.subUnitId === pivot.anchor.subUnitId && sheetRangesIntersect(other.outputRange, pivot.outputRange)) throw new RangeError('移位后透视输出相互重叠。')
    }
  }
}
