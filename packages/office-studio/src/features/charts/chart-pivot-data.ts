import type { CellScalar, WorkbookModel } from '../../domain/model'
import { pivotOutputValues } from '../pivot/services/pivot-engine'
import type { ChartRecords } from './chart-data'

export interface PivotChartSourceData extends ChartRecords {
  categoryFields: number[]
  valueFields: number[]
  syntheticValues?: boolean
}

/** Pivot output has multiple header rows and can put value fields on rows.
 * Normalize it once into the same field/record contract as other chart data. */
export function readPivotChartSource(workbook: WorkbookModel, pivotTableId: string): PivotChartSourceData {
  const pivot = workbook.pivotTables?.find(item => item.id === pivotTableId)
  if (!pivot) throw new Error(`Pivot table not found: ${pivotTableId}`)
  const matrix = pivotOutputValues(workbook, pivot)
  if (matrix.length < 2 || matrix[0]!.length < 2) throw new RangeError('Pivot chart requires a pivot table with row and value fields')
  const rowFields = pivot.fields.filter(field => field.area === 'row').sort((a, b) => a.position - b.position)
  const columnFields = pivot.fields.filter(field => field.area === 'column').sort((a, b) => a.position - b.position)
  if (!rowFields.length) throw new RangeError('Pivot chart requires at least one row field to define categories')
  const valueRows = pivot.options.valuePosition === 'row'
  const headerRows = Math.max(1, columnFields.length + (valueRows ? 0 : 1))
  const rowHeaderCount = rowFields.length + (valueRows ? 1 : 0)
  const rows = matrix.slice(headerRows).filter(row => row.some(value => typeof value === 'number') && !isTotalRow(row, rowFields.length))
  if (!rows.length) throw new RangeError('Pivot chart has no numeric categories to display')
  const previousLabels = Array.from({ length: rowFields.length }, () => '')
  const categoryRows = rows.map(row => row.slice(0, rowFields.length).map((value, index) => {
    if (value != null && value !== '') previousLabels[index] = scalarLabel(value)
    return previousLabels[index]!
  }))
  if (!valueRows) {
    const valueFields = matrix[0]!.map((_value, index) => index)
      .filter(index => index >= rowHeaderCount && rows.some(row => typeof row[index] === 'number') && !isTotalColumn(matrix, index, headerRows))
    return {
      fields: matrix[0]!.map((_value, field) => matrix.slice(0, headerRows).map(row => row[field])
        .filter(value => value != null && value !== '').map(scalarLabel).join(' · ') || `字段 ${field + 1}`),
      records: rows.map((row, index) => row.map((value, field) => field < rowFields.length ? categoryRows[index]![field]! : value)),
      categoryFields: rowFields.map((_field, index) => index), valueFields,
    }
  }

  // Preserve the existing synthetic value-field indices (1..N). Independent
  // hierarchy fields follow them, so legacy snapshots keep their bindings.
  const categories = categoryRows.map(labels => labels.filter(Boolean).join(' / '))
  const categoryLabels = new Map<string, string[]>()
  const seriesByName = new Map<string, Map<string, number>>()
  for (const [rowIndex, row] of rows.entries()) {
    const category = categories[rowIndex]!
    categoryLabels.set(category, categoryRows[rowIndex]!)
    const valueName = scalarLabel(row[rowFields.length])
    for (let column = rowHeaderCount; column < matrix[0]!.length; column++) {
      if (isTotalColumn(matrix, column, headerRows) || columnFields.length > 0 && pivot.options.showColGrandTotal && column === matrix[0]!.length - 1) continue
      const value = numericValue(row[column])
      if (value === null) continue
      const columnName = matrix.slice(0, headerRows).map(header => header[column])
        .filter(label => label != null && label !== '').map(scalarLabel).join(' · ')
      const name = [columnName, valueName].filter(Boolean).join(' · ')
      const series = seriesByName.get(name) ?? new Map<string, number>()
      series.set(category, value); seriesByName.set(name, series)
    }
  }
  const allSeries = [...seriesByName.entries()]
  const uniqueCategories = [...new Set(categories)]
  return {
    fields: ['分类', ...allSeries.map(([name]) => name), ...rowFields.map(field => field.name)],
    records: uniqueCategories.map(category => [category, ...allSeries.map(([, values]) => values.get(category) ?? null), ...categoryLabels.get(category)!]),
    categoryFields: [0], valueFields: allSeries.map((_series, index) => index + 1), syntheticValues: true,
  }
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
