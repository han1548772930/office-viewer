import type { CellScalar, PivotAxisOutputEntry, PivotTableModel, WorkbookModel } from '../../domain/model'
import { calculatePivotReport, pivotOutputValues } from '../pivot/services/pivot-engine'
import { formatSheetValue } from '../sheet-workbench/sheet-number-format'
import type { ChartRecords } from './chart-data'
import { equivalentModelData } from '../../platform/commands/model-equivalence.ts'
import { pivotChartAxisMetadata } from './chart-pivot-axis.ts'

export interface PivotChartSourceData extends ChartRecords {
  categoryFields: number[]
  valueFields: number[]
  syntheticValues?: boolean
}

/** Pivot output has multiple header rows and can put value fields on rows.
 * Normalize it once into the same field/record contract as other chart data. */
export function readPivotChartSource(workbook: WorkbookModel, pivotTableId: string, retainFormatting = false): PivotChartSourceData {
  const pivot = workbook.pivotTables?.find(item => item.id === pivotTableId)
  if (!pivot) throw new Error(`Pivot table not found: ${pivotTableId}`)
  if (!pivot.outputRange) return readOwnedPivotChartSource(workbook, pivot, retainFormatting)
  const sheet = workbook.sheets.find(item => item.id === pivot.anchor.subUnitId)
  const format = retainFormatting ? (row: number, column: number, value: CellScalar): string | null => {
    if (value === null) return null
    const cell = sheet?.cells[`${pivot.outputRange!.startRow + row}:${pivot.outputRange!.startCol + column}`]
    return cell?.style.numberFormat ? formatSheetValue(value, cell.style.numberFormat, workbook.date1904) ?? String(value) : String(value)
  } : undefined
  const matrix = pivotOutputValues(workbook, pivot)
  let axes = validPivotAxes(pivot, matrix, pivot.axisOutput) ? pivot.axisOutput : undefined
  if (!axes) {
    // Legacy output has no retained metadata. Use the authoritative calculation
    // only when it describes the exact stored matrix, preserving unrefreshed imports.
    const result = calculatePivotReport(workbook, pivot)
    if (equivalentModelData(result.values, matrix)) axes = { rowAxis: result.rowAxis, columnAxis: result.columnAxis }
  }
  return normalizePivotChartReport(pivot, matrix, format, axes)
}

export function readOwnedPivotChartSource(workbook: WorkbookModel, pivot: PivotTableModel, retainFormatting = false): PivotChartSourceData {
  const result = calculatePivotReport(workbook, pivot)
  const format = retainFormatting ? (row: number, column: number, value: CellScalar): string | null => {
    if (value === null) return null
    const pattern = result.numberFormats[`${row}:${column}`]
    return pattern ? formatSheetValue(value, pattern, workbook.date1904) ?? String(value) : String(value)
  } : undefined
  return normalizePivotChartReport(pivot, result.values, format, { rowAxis: result.rowAxis, columnAxis: result.columnAxis })
}

function normalizePivotChartReport(pivot: PivotTableModel, matrix: CellScalar[][],
  format?: (row: number, column: number, value: CellScalar) => string | null,
  axes?: PivotTableModel['axisOutput']): PivotChartSourceData {
  const rowFields = pivot.fields.filter(field => field.area === 'row').sort((a, b) => a.position - b.position)
  const columnFields = pivot.fields.filter(field => field.area === 'column').sort((a, b) => a.position - b.position)
  const valueFields = pivot.fields.filter(field => field.area === 'value').sort((a, b) => a.position - b.position)
  const emptyReport = (): PivotChartSourceData => {
    const fields = [...rowFields.map(field => field.displayName ?? field.name), ...valueFields.map(field => field.displayName ?? field.name)]
    return { fields, records: [], categoryFields: rowFields.map((_field, index) => index),
      valueFields: valueFields.map((_field, index) => rowFields.length + index), visibleFields: fields.map((_field, index) => index) }
  }
  // An independently owned report can be created before fields are placed.
  // Retain its empty field contract so authoring can continue in the field list.
  if (!rowFields.length || !valueFields.length) return emptyReport()
  if (matrix.length < 2 || (matrix[0]?.length ?? 0) < 2) return emptyReport()
  const valueRows = pivot.options.valuePosition === 'row'
  const headerRows = Math.max(1, columnFields.length + (valueRows ? 0 : 1))
  const rowHeaderCount = rowFields.length + (valueRows ? 1 : 0)
  const rowsByIndex = new Map<number, PivotAxisOutputEntry>(axes?.rowAxis.map(entry => [entry.index, entry]))
  const columnsByIndex = new Map<number, PivotAxisOutputEntry>(axes?.columnAxis.map(entry => [entry.index, entry]))
  const sourceRows = matrix.slice(headerRows).map((row, index) => ({ values: row, index: headerRows + index, axis: rowsByIndex.get(headerRows + index) }))
    .filter(row => row.axis ? plottedAxisEntry(row.axis) : row.values.some(value => value != null && value !== '') && !isTotalRow(row.values, rowFields.length))
  const rows = sourceRows.map(row => row.values)
  if (!rows.length) return emptyReport()
  const previousLabels = Array.from({ length: rowFields.length }, () => '')
  const categoryRows = sourceRows.map(sourceRow => rowsByIndex.has(sourceRow.index)
    ? rowFields.map((_field, index) => scalarLabel(sourceRow.axis!.labels[index]))
    : sourceRow.values.slice(0, rowFields.length).map((value, index) => {
    if (value != null && value !== '') previousLabels[index] = scalarLabel(value)
    return previousLabels[index]!
  }))
  if (!valueRows) {
    const plottedValueFields = matrix[0]!.map((_value, index) => index)
      .filter(index => index >= rowHeaderCount && (columnsByIndex.has(index)
        ? plottedAxisEntry(columnsByIndex.get(index)!) : !isTotalColumn(matrix, index, headerRows)))
    return {
      fields: matrix[0]!.map((_value, field) => {
        const axis = columnsByIndex.get(field), value = valueFields.find(value => value.id === axis?.valueFieldId)
        const labels = axis && plottedAxisEntry(axis) ? [...axis.labels, ...(value ? [value.displayName ?? value.name] : [])]
          : matrix.slice(0, headerRows).map(row => row[field])
        return labels.filter(label => label != null && label !== '').map(scalarLabel).join(' · ') || `字段 ${field + 1}`
      }),
      records: rows.map((row, index) => row.map((value, field) => field < rowFields.length ? categoryRows[index]![field]! : value)),
      ...(format ? { formattedRecords: rows.map((row, index) => row.map((value, field) => field < rowFields.length
        ? categoryRows[index]![field]! : format(sourceRows[index]!.index, field, value))) } : {}),
      categoryFields: rowFields.map((_field, index) => index), valueFields: plottedValueFields,
      ...(axes ? { pivotAxis: pivotChartAxisMetadata(rowFields, columnFields, sourceRows.map(row => row.axis!), plottedValueFields.map(index => columnsByIndex.get(index)!)) } : {}),
    }
  }

  // Preserve the existing synthetic value-field indices (1..N). Independent
  // hierarchy fields follow them, so legacy snapshots keep their bindings.
  const categories = categoryRows.map(labels => labels.filter(Boolean).join(' / '))
  const categoryKeys = sourceRows.map((row, index) => JSON.stringify(row.axis?.labels ?? categoryRows[index]))
  const categoryNames = new Map<string, string>()
  const categoryLabels = new Map<string, string[]>()
  const categoryAxes = new Map<string, PivotAxisOutputEntry>()
  const seriesByKey = new Map<string, { name: string; values: Map<string, CellScalar>; formatted?: Map<string, string | null> }>()
  for (const [rowIndex, row] of rows.entries()) {
    const category = categoryKeys[rowIndex]!
    categoryNames.set(category, categories[rowIndex]!)
    categoryLabels.set(category, categoryRows[rowIndex]!)
    if (sourceRows[rowIndex]!.axis && !categoryAxes.has(category)) categoryAxes.set(category, sourceRows[rowIndex]!.axis!)
    const valueName = scalarLabel(row[rowFields.length])
    for (let column = rowHeaderCount; column < matrix[0]!.length; column++) {
      const columnAxis = columnsByIndex.get(column)
      if (columnAxis ? !plottedAxisEntry(columnAxis) : isTotalColumn(matrix, column, headerRows)
        || columnFields.length > 0 && pivot.options.showColGrandTotal && column === matrix[0]!.length - 1) continue
      const raw = row[column] ?? null
      const value = numericValue(raw) ?? raw
      const columnName = (columnAxis ? columnAxis.labels : matrix.slice(0, headerRows).map(header => header[column]))
        .filter(label => label != null && label !== '').map(scalarLabel).join(' · ')
      const name = [columnName, valueName].filter(Boolean).join(' · ')
      const seriesKey = JSON.stringify([columnAxis?.labels ?? columnName, sourceRows[rowIndex]!.axis?.valueFieldId ?? valueName])
      const series = seriesByKey.get(seriesKey) ?? { name, values: new Map<string, CellScalar>(), ...(format ? { formatted: new Map<string, string | null>() } : {}) }
      series.values.set(category, value); seriesByKey.set(seriesKey, series)
      if (format) series.formatted!.set(category, format(sourceRows[rowIndex]!.index, column, raw))
    }
  }
  const allSeries = [...seriesByKey.values()]
  const uniqueCategories = [...new Set(categoryKeys)]
  return {
    fields: ['分类', ...allSeries.map(series => series.name), ...rowFields.map(field => field.name)],
    records: uniqueCategories.map(category => [categoryNames.get(category)!, ...allSeries.map(series => series.values.get(category) ?? null), ...categoryLabels.get(category)!]),
    ...(format ? { formattedRecords: uniqueCategories.map(category => [categoryNames.get(category)!,
      ...allSeries.map(series => series.formatted!.get(category) ?? null), ...categoryLabels.get(category)!]) } : {}),
    categoryFields: [0], valueFields: allSeries.map((_series, index) => index + 1), syntheticValues: true,
    ...(axes ? { pivotAxis: pivotChartAxisMetadata(rowFields, columnFields, uniqueCategories.map(key => categoryAxes.get(key)!),
      axes.columnAxis.filter(plottedAxisEntry), rowFields.map((_field, index) => allSeries.length + 1 + index)) } : {}),
  }
}

function plottedAxisEntry(entry: PivotAxisOutputEntry): boolean { return entry.kind === 'detail' || entry.kind === 'collapsed' }

function validPivotAxes(pivot: PivotTableModel, matrix: CellScalar[][], axes: PivotTableModel['axisOutput']): boolean {
  if (!axes || !Array.isArray(axes.rowAxis) || !Array.isArray(axes.columnAxis)) return false
  const headerRows = Math.max(1, pivot.fields.filter(field => field.area === 'column').length + (pivot.options.valuePosition === 'row' ? 0 : 1))
  const headerColumns = pivot.fields.filter(field => field.area === 'row').length + (pivot.options.valuePosition === 'row' ? 1 : 0)
  const validEntry = (entry: PivotAxisOutputEntry, maximum: number, minimum: number): boolean => !!entry
    && Number.isSafeInteger(entry.index) && entry.index >= minimum && entry.index < maximum
    && ['detail', 'collapsed', 'subtotal', 'grand-total'].includes(entry.kind) && Array.isArray(entry.labels) && entry.labels.every(isScalar)
    && Number.isSafeInteger(entry.level) && entry.level >= -1
    && !!entry.fieldFilters && typeof entry.fieldFilters === 'object' && !Array.isArray(entry.fieldFilters)
    && Object.values(entry.fieldFilters).every(isScalar)
    && !!entry.fieldMembers && typeof entry.fieldMembers === 'object' && !Array.isArray(entry.fieldMembers)
    && Object.values(entry.fieldMembers).every(members => Array.isArray(members) && members.every(isScalar))
  return axes.rowAxis.length === Math.max(0, matrix.length - headerRows)
    && axes.columnAxis.length === Math.max(0, (matrix[0]?.length ?? 0) - headerColumns)
    && axes.rowAxis.every(entry => validEntry(entry, matrix.length, headerRows))
    && axes.columnAxis.every(entry => validEntry(entry, matrix[0]?.length ?? 0, headerColumns))
    && new Set(axes.rowAxis.map(entry => entry.index)).size === axes.rowAxis.length
    && new Set(axes.columnAxis.map(entry => entry.index)).size === axes.columnAxis.length
}

function isScalar(value: unknown): value is CellScalar {
  return value === null || typeof value === 'string' || typeof value === 'boolean' || typeof value === 'number' && Number.isFinite(value)
}

function isTotalColumn(matrix: CellScalar[][], column: number, headerRows: number): boolean {
  return matrix.slice(0, headerRows).some(row => typeof row[column] === 'string' && /^grand\s+total(?:\s+.+)?$/i.test(row[column]!.trim()))
}
function isTotalRow(row: CellScalar[], rowFieldCount: number): boolean {
  return row.slice(0, rowFieldCount).some(value => typeof value === 'string' && /^grand\s+total$/i.test(value.trim()))
}
function numericValue(value: CellScalar | undefined): number | null {
  return typeof value === 'number' && Number.isFinite(value) ? value
    : typeof value === 'string' && value.trim() !== '' && Number.isFinite(Number(value)) ? Number(value) : null
}
function scalarLabel(value: CellScalar | undefined): string { return value == null ? '' : String(value) }
