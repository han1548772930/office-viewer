import type { CellRange, ChartModel, SheetAutoFilter, SheetConditionalFormatRule, SheetDataValidationRule, SheetTableModel, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import { sheetAxisEditSegments, sheetAxisInsertionEdit, sheetRangeAxisInterval, sheetRangeWithAxisInterval,
  transformSheetAxisRange, transformSheetAxisReferenceInterval, type SheetAxisEdit } from '../../domain/sheet-axis-structure.ts'
import { conditionalFormatFormulaAnchor, mapConditionalFormatFormulas } from '../conditional-formatting/conditional-format-formulas.ts'
import { sheetValidationFormulaAnchor, sheetValidationUsesRelativeAnchor } from '../sheet-validation/validation-formulas.ts'
import { transformFormulaReferenceAreasForAxisEdit, translateFormulaReferenceAreas } from '../formula/formula-reference-areas.ts'
import { parseSheetHyperlinkTarget } from './sheet-hyperlink.ts'
import { formatProtectionRange } from '../sheet-protection/range-protection.ts'
import { transformSheetImagesForAxisEdit } from '../sheet-images/sheet-image-structure.ts'
import { sheetClipboardObjectAxes } from './sheet-clipboard-object-geometry.ts'
import { reconcileSlicerConnectionResources } from '../slicer/slicer-connections-lifecycle.ts'
import { refreshPivotSourceWorksheetOutputs, transformWorkbookPivotSourceReports } from '../pivot/services/pivot-source-lifecycle.ts'
import { synchronizeLinkedPivotChartReports } from '../pivot/services/pivot-chart-report-lifecycle.ts'
import { transformSheetOutlineForAxisEdit } from '../sheet-outline/sheet-outline-structure.ts'

const object = (value: unknown): Record<string, unknown> | undefined => value !== null && typeof value === 'object' && !Array.isArray(value)
  ? value as Record<string, unknown> : undefined
const referenceFormula = (formula: string, owner: string | undefined, target: string, edit: SheetAxisEdit): string => transformFormulaReferenceAreasForAxisEdit(formula, edit, owner, target)

function axisMapper(edit: SheetAxisEdit) {
  const segments = sheetAxisEditSegments(edit)
  return (index: number): number | undefined => {
    const segment = segments.find(item => index >= item.start && index <= item.end)
    return segment ? index + segment.offset : undefined
  }
}
function inverseAxisIndex(index: number, edit: SheetAxisEdit): number | undefined {
  const segment = sheetAxisEditSegments(edit).find(item => index >= item.start + item.offset && index <= item.end + item.offset)
  return segment ? index - segment.offset : undefined
}
function copiedRanges(ranges: readonly CellRange[], edit: SheetAxisEdit): CellRange[] {
  if (edit.kind !== 'copy') return []
  return ranges.flatMap(range => {
    const interval = sheetRangeAxisInterval(range, edit.axis), start = Math.max(interval.start, edit.start), end = Math.min(interval.end, edit.start + edit.count - 1)
    return start <= end ? [sheetRangeWithAxisInterval(range, edit.axis, { start: edit.destination + start - edit.start, end: edit.destination + end - edit.start })] : []
  })
}
function singleResourceRange(range: CellRange, edit: SheetAxisEdit, label: string): CellRange | undefined {
  const ranges = transformSheetAxisRange(range, edit)
  if (ranges.length > 1) throw new Error(`${label}需要连续范围，当前移动会将它拆开。请移动完整范围或在范围内部调整。`)
  return ranges[0]
}
function mapOrigin(origin: { row: number; col: number }, ranges: readonly CellRange[], edit: SheetAxisEdit) {
  const originalIndex = edit.axis === 'row' ? origin.row : origin.col, mapped = axisMapper(edit)(originalIndex)
  if (mapped !== undefined) return { source: { ...origin }, target: edit.axis === 'row' ? { ...origin, row: mapped } : { ...origin, col: mapped } }
  const first = ranges[0]!
  const target = { row: first.startRow, col: first.startCol }, sourceIndex = inverseAxisIndex(edit.axis === 'row' ? target.row : target.col, edit)
  if (sourceIndex === undefined) throw new Error('无法确定范围资源删除后的公式起点。')
  return { source: edit.axis === 'row' ? { ...target, row: sourceIndex } : { ...target, col: sourceIndex }, target }
}
function conditionalForEdit(rule: SheetConditionalFormatRule, owner: string, target: string, edit: SheetAxisEdit, local: boolean): SheetConditionalFormatRule[] {
  const ranges = local ? rule.ranges.flatMap(range => transformSheetAxisRange(range, edit)) : rule.ranges.map(range => ({ ...range }))
  if (!ranges.length) return []
  const origin = conditionalFormatFormulaAnchor(rule), rebased = origin && local ? mapOrigin(origin, ranges, edit) : undefined
  const next = mapConditionalFormatFormulas(rule, formula => referenceFormula(rebased && origin
    ? translateFormulaReferenceAreas(formula, rebased.source.row - origin.row, rebased.source.col - origin.col) : formula, owner, target, edit))
  next.ranges = ranges
  if (rebased) next.formulaAnchor = { ...rebased.target }
  const copy = local ? copiedRanges(rule.ranges, edit) : []
  if (!copy.length || !origin || edit.kind !== 'copy') return [next]
  const first = copy[0]!, source = edit.axis === 'row'
    ? { row: edit.start + first.startRow - edit.destination, col: first.startCol }
    : { row: first.startRow, col: edit.start + first.startCol - edit.destination }
  const movedSource = mapOrigin(source, ranges, edit).target, copiedOrigin = { row: first.startRow, col: first.startCol }
  const duplicate = mapConditionalFormatFormulas(rule, formula => translateFormulaReferenceAreas(referenceFormula(
    translateFormulaReferenceAreas(formula, source.row - origin.row, source.col - origin.col), owner, target, edit),
    copiedOrigin.row - movedSource.row, copiedOrigin.col - movedSource.col))
  duplicate.id = `cf-${crypto.randomUUID()}`
  duplicate.ranges = copy
  duplicate.formulaAnchor = copiedOrigin
  return [next, duplicate]
}
function validationForEdit(rule: SheetDataValidationRule, owner: string, target: string, edit: SheetAxisEdit, local: boolean): SheetDataValidationRule[] {
  const ranges = local ? rule.ranges.flatMap(range => transformSheetAxisRange(range, edit)) : rule.ranges.map(range => ({ ...range }))
  if (!ranges.length) return []
  const origin = sheetValidationFormulaAnchor(rule)
  // Literal list entries and checkbox labels are data. Their range/formula
  // references still follow structural edits, without a per-cell rebase.
  const relative = sheetValidationUsesRelativeAnchor(rule)
  const rebased = relative && local && origin ? mapOrigin(origin, ranges, edit) : undefined
  const transform = (formula: string | undefined) => formula === undefined ? undefined : rule.type === 'checkbox' ? formula
    : referenceFormula(rebased && origin ? translateFormulaReferenceAreas(formula, rebased.source.row - origin.row, rebased.source.col - origin.col) : formula, owner, target, edit)
  const formula1 = transform(rule.formula1), formula2 = transform(rule.formula2)
  const next = { ...rule, ranges, formula1, formula2, ...(rebased ? { formulaAnchor: { ...rebased.target } } : {}) }, copy = local ? copiedRanges(rule.ranges, edit) : []
  if (!copy.length || edit.kind !== 'copy') return [next]
  const oldCopyIndex = edit.start + (edit.axis === 'row' ? copy[0]!.startRow : copy[0]!.startCol) - edit.destination
  const movedCopyIndex = axisMapper(edit)(oldCopyIndex)!
  const delta = (edit.axis === 'row' ? copy[0]!.startRow : copy[0]!.startCol) - movedCopyIndex
  const copyFormula = (formula: string | undefined) => formula === undefined || rule.type === 'checkbox' ? formula
    : translateFormulaReferenceAreas(formula, edit.axis === 'row' ? delta : 0, edit.axis === 'column' ? delta : 0)
  const copiedAnchor = relative && rebased ? { row: rebased.target.row + (edit.axis === 'row' ? delta : 0), col: rebased.target.col + (edit.axis === 'column' ? delta : 0) } : undefined
  return [next, { ...next, id: `dv-${crypto.randomUUID()}`, ranges: copy, formula1: copyFormula(formula1), formula2: copyFormula(formula2),
    ...(copiedAnchor ? { formulaAnchor: copiedAnchor } : {}) }]
}

export function transformSheetAutoFilterForAxisEdit(filter: SheetAutoFilter | undefined, edit: SheetAxisEdit): SheetAutoFilter | undefined {
  if (!filter) return undefined
  const range = singleResourceRange(filter.range, edit, '自动筛选')
  if (!range) return undefined
  const map = axisMapper(edit), columns: SheetAutoFilter['columns'] = {}
  for (const [column, criterion] of Object.entries(filter.columns)) {
    const mapped = edit.axis === 'column' ? map(Number(column)) : Number(column)
    if (mapped !== undefined && mapped >= range.startCol && mapped <= range.endCol) columns[mapped] = structuredClone(criterion)
  }
  return { ...filter, range, columns, filteredRows: edit.axis === 'row'
    ? filter.filteredRows.flatMap(row => { const mapped = map(row); return mapped === undefined ? [] : [mapped] }) : [...filter.filteredRows] }
}
function transformedTableColumns(table: SheetTableModel, range: CellRange, edit: SheetAxisEdit): SheetTableModel['columns'] {
  if (edit.axis !== 'column') return table.columns.map(column => ({ ...column }))
  const map = axisMapper(edit), columns = new Map<number, SheetTableModel['columns'][number]>()
  table.columns.forEach((column, index) => { const position = map(table.range.startCol + index); if (position !== undefined) columns.set(position, { ...column }) })
  const names = new Set(table.columns.map(column => column.name.toLowerCase()))
  return Array.from({ length: range.endCol - range.startCol + 1 }, (_value, index) => {
    const existing = columns.get(range.startCol + index)
    if (existing) return existing
    let sequence = index + 1, name = `Column${sequence}`
    while (names.has(name.toLowerCase())) { sequence++; name = `Column${sequence}` }
    names.add(name.toLowerCase())
    return { id: `table-column-${crypto.randomUUID()}`, name }
  })
}
function tableForEdit(table: SheetTableModel, edit: SheetAxisEdit): SheetTableModel[] {
  const range = singleResourceRange(table.range, edit, `表格“${table.name}”`)
  if (!range) return []
  if (edit.kind === 'move' && edit.axis === 'row' && table.showHeaderRow && table.range.startRow >= edit.start
    && table.range.startRow < edit.start + edit.count && table.range.endRow >= edit.start + edit.count) {
    throw new Error(`移动会将表格“${table.name}”的表头与数据拆开。请一起移动表格。`)
  }
  const next = { ...table, range, columns: transformedTableColumns(table, range, edit), autoFilter: transformSheetAutoFilterForAxisEdit(table.autoFilter, edit) }
  const interval = sheetRangeAxisInterval(table.range, edit.axis), copies = copiedRanges([table.range], edit)
  if (edit.kind !== 'copy' || interval.start < edit.start || interval.end >= edit.start + edit.count || !copies[0]) return [next]
  const copied = copies[0]
  if (copied.startRow <= range.endRow && copied.endRow >= range.startRow && copied.startCol <= range.endCol && copied.endCol >= range.startCol) return [next]
  const duplicate: SheetTableModel = { ...structuredClone(table), id: `table-${crypto.randomUUID()}`, name: table.name,
    range: copied, columns: table.columns.map(column => ({ ...column, id: `table-column-${crypto.randomUUID()}` })) }
  if (duplicate.autoFilter) {
    const delta = edit.destination - edit.start
    duplicate.autoFilter.range = sheetRangeWithAxisInterval(duplicate.autoFilter.range, edit.axis,
      { start: sheetRangeAxisInterval(duplicate.autoFilter.range, edit.axis).start + delta, end: sheetRangeAxisInterval(duplicate.autoFilter.range, edit.axis).end + delta })
    if (edit.axis === 'column') duplicate.autoFilter.columns = Object.fromEntries(Object.entries(duplicate.autoFilter.columns).map(([column, criterion]) => [Number(column) + delta, criterion]))
    else duplicate.autoFilter.filteredRows = duplicate.autoFilter.filteredRows.map(row => row + delta)
  }
  return [next, duplicate]
}

function chartForEdit(chart: ChartModel, edit: SheetAxisEdit): ChartModel | undefined {
  const range = singleResourceRange(chart.source.range, edit, '图表数据源')
  if (!range) return undefined
  const fieldAxis = chart.source.orientation === 'rows' ? 'row' : 'column', map = axisMapper(edit)
  const before = sheetRangeAxisInterval(chart.source.range, fieldAxis).start, after = sheetRangeAxisInterval(range, fieldAxis).start
  const field = (value: number): number | undefined => { const position = fieldAxis === edit.axis ? map(before + value) : before + value; return position === undefined ? undefined : position - after }
  const categoryField = field(chart.categoryField)
  if (categoryField === undefined || categoryField < 0) throw new Error('删除会移除图表分类字段，请先调整图表的数据绑定。')
  const valueFields = chart.valueFields.flatMap(value => { const mapped = field(value); return mapped === undefined ? [] : [mapped] })
  const seriesToCategory = chart.seriesToCategory ? Object.fromEntries(Object.entries(chart.seriesToCategory).flatMap(([series, categories]) => {
    const mapped = field(Number(series))
    return mapped === undefined ? [] : [[mapped, categories.flatMap(category => { const next = field(category); return next === undefined ? [] : [next] })]]
  })) : undefined
  return { ...chart, source: { ...chart.source, range }, categoryField, valueFields, seriesToCategory, revision: chart.revision + 1 }
}
// Cell-band edits reuse these continuous-resource semantics when the whole resource lies in the band.
export { tableForEdit as transformSheetTableForAxisEdit, chartForEdit as transformSheetChartForAxisEdit }
function internalHyperlink(value: string, original: WorkbookModel, targetId: string, edit: SheetAxisEdit): string {
  const target = parseSheetHyperlinkTarget(value, original)
  if (target?.type !== 'internal' || target.sheetId !== targetId) return value
  const interval = transformSheetAxisReferenceInterval(sheetRangeAxisInterval(target.range, edit.axis), edit)
  const address = interval ? formatProtectionRange(sheetRangeWithAxisInterval(target.range, edit.axis, interval)) : '#REF!'
  if (value.slice(1).includes('=')) { const params = new URLSearchParams(value.slice(1)); params.set('range', address); return `#${params}` }
  return `${value.slice(0, value.lastIndexOf('!') + 1)}${address}`
}
function updateDefinedNames(value: unknown, original: WorkbookModel, targetName: string, edit: SheetAxisEdit, local: boolean): unknown {
  if (!Array.isArray(value)) return value
  const active = original.sheets.find(sheet => sheet.id === original.activeSheetId)?.name
  return value.map((entry: unknown) => {
    const record = object(entry)
    if (!record || typeof record.refersTo !== 'string') return entry
    const owner = typeof record.sheet === 'string' ? record.sheet : typeof record.sheetId === 'string'
      ? original.sheets.find(sheet => sheet.id === record.sheetId)?.name : local ? undefined : active
    return { ...record, refersTo: referenceFormula(record.refersTo, owner, targetName, edit) }
  })
}

/** Update typed resources and formula-bearing fields against the same original coordinates. */
export function transformWorkbookAxisResources(workbook: WorkbookModel, original: WorkbookModel, target: SpreadsheetSheet, edit: SheetAxisEdit): void {
  for (const sheet of workbook.sheets) {
    const before = original.sheets.find(item => item.id === sheet.id)!, local = sheet.id === target.id
    sheet.conditionalFormats = before.conditionalFormats?.flatMap(rule => conditionalForEdit(rule, sheet.name, target.name, edit, local))
    sheet.dataValidations = before.dataValidations?.flatMap(rule => validationForEdit(rule, sheet.name, target.name, edit, local))
    for (const cell of Object.values(sheet.cells)) {
      if (cell.formula) cell.formula = referenceFormula(cell.formula, sheet.name, target.name, edit)
      if (cell.sparkline) cell.sparkline = { ...cell.sparkline, sourceRange: referenceFormula(cell.sparkline.sourceRange, sheet.name, target.name, edit) }
      if (cell.hyperlink) cell.hyperlink = internalHyperlink(cell.hyperlink, original, target.id, edit)
      if (cell.richText) cell.richText = cell.richText.map(run => typeof run.hyperlink === 'string' ? { ...run, hyperlink: internalHyperlink(run.hyperlink, original, target.id, edit) } : run)
    }
    if (!local) continue
    sheet.autoFilter = transformSheetAutoFilterForAxisEdit(before.autoFilter, edit)
    sheet.rangeProtectionRules = before.rangeProtectionRules?.flatMap(rule => {
      const ranges = rule.ranges.flatMap(range => transformSheetAxisRange(range, edit))
      return ranges.length ? [{ ...rule, ranges }] : []
    })
    sheet.outline = transformSheetOutlineForAxisEdit(before.outline, edit, before, sheet)
    const insertion = sheetAxisInsertionEdit(edit), count = edit.axis === 'row' ? before.frozenRows : before.frozenCols
    if (insertion && count > insertion.start) {
      const frozen = insertion.kind === 'insert' ? count + insertion.count : count - Math.min(insertion.count, count - insertion.start)
      if (edit.axis === 'row') sheet.frozenRows = Math.min(sheet.rowCount ?? 1_048_576, frozen)
      else sheet.frozenCols = Math.min(sheet.columnCount ?? 16_384, frozen)
    }
  }
  const tableNames = new Set((original.tables ?? []).map(table => table.name.toLowerCase()))
  workbook.tables = original.tables?.flatMap(table => {
    if (table.sheetId !== target.id) return [structuredClone(table)]
    const tables = tableForEdit(table, edit)
    for (const duplicate of tables.slice(1)) {
      let sequence = 2, name = `${table.name}_${sequence}`
      while (tableNames.has(name.toLowerCase())) { sequence++; name = `${table.name}_${sequence}` }
      duplicate.name = name; tableNames.add(name.toLowerCase())
    }
    return tables
  })
  workbook.charts = original.charts?.flatMap(chart => {
    if ((chart.source.sheetId ?? chart.sheetId) !== target.id || chart.pivotTableId || chart.ownedPivot) return [structuredClone(chart)]
    const next = chartForEdit(chart, edit)
    return next ? [next] : []
  })
  const map = axisMapper(edit)
  const pivotTransform = {
    sheetId: target.id,
    range: (range: CellRange) => singleResourceRange(range, edit, '透视数据源'),
    outputRange: (range: CellRange) => singleResourceRange(range, edit, '透视输出'),
    point: (row: number, col: number) => {
      const index = map(edit.axis === 'row' ? row : col)
      return index === undefined ? undefined : edit.axis === 'row' ? { row: index, col } : { row, col: index }
    },
    sourceDirty: (range: CellRange) => {
      const interval = sheetRangeAxisInterval(range, edit.axis)
      if (edit.kind !== 'move') return interval.end >= (edit.kind === 'copy' ? edit.destination : edit.start)
      if (edit.destination === edit.start || edit.destination === edit.start + edit.count) return false
      const start = Math.min(edit.start, edit.destination), end = Math.max(edit.start + edit.count - 1, edit.destination - 1)
      return interval.start <= end && interval.end >= start
    },
    ...(edit.axis === 'column' ? { column: map } : {}),
  }
  transformWorkbookPivotSourceReports(original, workbook, pivotTransform)
  const removedTableIds = new Set((original.tables ?? []).filter(table => !workbook.tables?.some(next => next.id === table.id)).map(table => table.id))
  workbook.slicers = original.slicers?.flatMap(slicer => {
    if (slicer.tableId && removedTableIds.has(slicer.tableId)) return []
    const next = structuredClone(slicer)
    if (slicer.hostSheetId === target.id && slicer.drawing?.anchorType !== 'none') {
      const position = map(edit.axis === 'row' ? slicer.position.row : slicer.position.column)
      if (position !== undefined) { if (edit.axis === 'row') next.position.row = position; else next.position.column = position }
      else { if (edit.axis === 'row') next.position.row = edit.start; else next.position.column = edit.start }
      if (next.drawing && next.drawing.anchorType !== 'none') {
        for (const point of [next.drawing.from, ...(next.drawing.anchorType === 'both' ? [next.drawing.to] : [])]) {
          const index = map(edit.axis === 'row' ? point.row : point.column) ?? edit.start
          if (edit.axis === 'row') point.row = index; else point.column = index
        }
      }
    }
    if (JSON.stringify(next) !== JSON.stringify(slicer)) next.revision++
    return [next]
  })
  if (workbook.metadata) {
    workbook.metadata.definedNames = updateDefinedNames(workbook.metadata.definedNames, original, target.name, edit, false)
    workbook.metadata.localDefinedNames = updateDefinedNames(workbook.metadata.localDefinedNames, original, target.name, edit, true)
  }
  const before = original.sheets.find(sheet => sheet.id === target.id)!, after = workbook.sheets.find(sheet => sheet.id === target.id)!
  if (before.images) after.images = transformSheetImagesForAxisEdit(before.images, edit, after,
    { before: sheetClipboardObjectAxes(original, before), after: sheetClipboardObjectAxes(workbook, after) })
  reconcileSlicerConnectionResources(original, workbook, pivotTransform)
  refreshPivotSourceWorksheetOutputs(original, workbook)
  synchronizeLinkedPivotChartReports(original, workbook)
}
