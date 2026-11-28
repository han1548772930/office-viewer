import type { CellScalar, ChartModel, PivotFieldModel, PivotRecordValue, PivotTableModel, SheetFilterCriteria, SpreadsheetCell } from '../../domain/model'
import { classifySheetNumberFormat, formatSheetValue } from '../sheet-workbench/sheet-number-format'
import { formulaCalendarDelta, formulaCalendarSerial, readFormulaCalendarValue } from '../formula/formula-import-context'

/** Office calendars use a 1462-day modern offset and a 1461-day early offset. Numeric quantities
 * and time/duration-only values do not become dates during data import. */
export function chartSourceDateDelta(sourceDate1904: boolean, targetDate1904: boolean): number {
  return formulaCalendarDelta(sourceDate1904, targetDate1904)
}
const numericSerial = (value: CellScalar): number | undefined => typeof value === 'number' && Number.isFinite(value) ? value
  : typeof value === 'string' && /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(value.trim()) ? Number(value) : undefined
export function remapChartSourceDateValue(value: CellScalar, delta: number, numericText = false): CellScalar {
  const serial = typeof value === 'number' ? value : numericText ? numericSerial(value) : undefined
  if (serial === undefined || !Number.isFinite(serial)) return value
  const next = delta === 1462 || delta === -1462 ? formulaCalendarSerial(serial, delta > 0, delta < 0) : serial + delta
  if (!Number.isFinite(next)) throw new RangeError('图表日期源超出数值范围。')
  return typeof value === 'string' ? String(next) : next
}
export function remapChartSourceDateCell(cell: SpreadsheetCell, delta: number, targetDate1904: boolean, forced = false): void {
  if (!delta) return
  const kind = classifySheetNumberFormat(cell.style.numberFormat), type = (cell.formula ? cell.cachedDataType ?? cell.dataType : cell.dataType)?.toLowerCase()
  const date = forced || kind === 'date' || kind === 'datetime' || type === 'date' || type === 'd' || type === 'datetime'
  if (!date) return
  const previous = cell.value
  const numeric = typeof cell.value === 'number' ? cell.value : forced ? numericSerial(cell.value) : undefined
  const retained = numeric === undefined ? undefined : readFormulaCalendarValue(cell, numeric, !targetDate1904)
  cell.value = retained ? formulaCalendarSerial(retained.value, retained.date1904, targetDate1904) : remapChartSourceDateValue(cell.value, delta, forced)
  if (retained && typeof previous === 'string') cell.value = String(cell.value)
  if (cell.value !== previous) cell.display = formatSheetValue(cell.value, cell.style.numberFormat, targetDate1904) ?? String(cell.value ?? '')
}
function remapRecordDate(value: PivotRecordValue, delta: number, dateColumn: boolean): PivotRecordValue {
  if (value !== null && typeof value === 'object') return value.type === 'date'
    ? { ...value, value: Number(remapChartSourceDateValue(value.value, delta)) } : value
  // Old snapshots use scalar selections. Typed number/text rules keep their
  // exact type identity; only their date-column scalar predecessor is shifted.
  return dateColumn ? remapChartSourceDateValue(value, delta, true) : value
}
export function remapChartSourceDateFilter(criteria: SheetFilterCriteria, delta: number, targetDate1904: boolean, dateColumn: boolean): SheetFilterCriteria {
  if (!delta) return criteria
  if (criteria.kind === 'values') {
    return { ...criteria,
      selectedItems: dateColumn ? criteria.selectedItems.map(value => remapChartSourceDateValue(value, delta, true)) : criteria.selectedItems,
      ...(criteria.recordDate1904 === undefined ? {} : { recordDate1904: targetDate1904 }),
      ...(criteria.recordSelection && criteria.recordSelection.mode !== 'all' ? { recordSelection: { ...criteria.recordSelection,
        values: criteria.recordSelection.values.map(value => remapRecordDate(value, delta, dateColumn)) } } : {}),
    }
  }
  return dateColumn ? { ...criteria, ...(criteria.value === undefined ? {} : { value: remapChartSourceDateValue(criteria.value, delta, true) }),
    ...(criteria.secondValue === undefined ? {} : { secondValue: remapChartSourceDateValue(criteria.secondValue, delta, true) }) } : criteria
}
export function pivotFieldUsesDates(field: PivotFieldModel): boolean {
  return !!field.dateGroupType || !!field.dateFilter || !!field.timelineDateRange || !!field.timelineDateCondition
    || field.recordSelection?.mode !== 'all' && field.recordSelection?.values.some(value => value !== null && typeof value === 'object' && value.type === 'date') === true
}
function remapGroupedDateKey(field: PivotFieldModel, key: string, delta: number): string {
  if (field.manualGroups?.some(group => group.name === key) || field.dateGroupType) return key
  if (field.numberGroup) {
    const format = (value: number): string => Number.isInteger(value) ? String(value) : String(Number(value.toFixed(12)))
    const bound = /^(<|>) ([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?)$/.exec(key)
    if (bound) return `${bound[1]} ${format(Number(remapChartSourceDateValue(Number(bound[2]), delta)))}`
    const interval = /^([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?)-([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?)$/.exec(key)
    if (interval) return `${format(Number(remapChartSourceDateValue(Number(interval[1]), delta)))}-${format(Number(remapChartSourceDateValue(Number(interval[2]), delta)))}`
  }
  return String(remapChartSourceDateValue(key, delta, true))
}
export function remapChartSourcePivotDates(report: PivotTableModel, delta: number, dateColumns: ReadonlySet<number> = new Set()): void {
  if (!delta) return
  const datedFields = new Set(report.fields.filter(field => dateColumns.has(field.sourceColumn) || pivotFieldUsesDates(field)).map(field => field.id))
  for (const field of report.fields) {
    if (field.dateFilter) field.dateFilter = { ...field.dateFilter,
      expected: Array.isArray(field.dateFilter.expected) ? [Number(remapChartSourceDateValue(field.dateFilter.expected[0], delta)), Number(remapChartSourceDateValue(field.dateFilter.expected[1], delta))] : Number(remapChartSourceDateValue(field.dateFilter.expected, delta)) }
    if (field.recordSelection && field.recordSelection.mode !== 'all') field.recordSelection = { ...field.recordSelection,
      values: field.recordSelection.values.map(value => remapRecordDate(value, delta, datedFields.has(field.id))) }
    if (datedFields.has(field.id)) {
      if (field.selectedItems) field.selectedItems = field.selectedItems.map(value => String(remapChartSourceDateValue(value, delta, true)))
      if (field.partialItems) field.partialItems = field.partialItems.map(value => remapGroupedDateKey(field, value, delta))
      if (field.groupedItems) field.groupedItems = field.groupedItems.map(value => remapGroupedDateKey(field, value, delta))
      const state = report.collapseInfo?.[field.id]
      if (state && typeof state === 'object') report.collapseInfo![field.id] = Object.fromEntries(Object.entries(state)
        .map(([key, value]) => [remapGroupedDateKey(field, key, delta), value]))
      if (field.numberGroup) field.numberGroup = { ...field.numberGroup, start: Number(remapChartSourceDateValue(field.numberGroup.start, delta)), end: Number(remapChartSourceDateValue(field.numberGroup.end, delta)) }
      if (field.manualGroups) field.manualGroups = field.manualGroups.map(group => ({ ...group,
        members: group.members.map(value => remapChartSourceDateValue(value, delta, true)) }))
    }
    const base = field.showDataAs && report.fields.find(candidate => candidate.id === field.showDataAs!.baseFieldId)
    if (field.showDataAs && (field.showDataAs.baseItemType ?? 'item') === 'item' && base && datedFields.has(base.id)) field.showDataAs = { ...field.showDataAs,
      baseItem: remapGroupedDateKey(base, field.showDataAs.baseItem, delta) }
    // Timeline rules contain calendar timestamps, not Excel serial days.
  }
}
export function remapChartSourceAxisDates(chart: ChartModel, delta: number): void {
  if (!delta || !chart.options?.mapping?.useDateAxis) return
  chart.xAxis = { ...chart.xAxis,
    ...(chart.xAxis.min === undefined ? {} : { min: Number(remapChartSourceDateValue(chart.xAxis.min, delta)) }),
    ...(chart.xAxis.max === undefined ? {} : { max: Number(remapChartSourceDateValue(chart.xAxis.max, delta)) }) }
}
