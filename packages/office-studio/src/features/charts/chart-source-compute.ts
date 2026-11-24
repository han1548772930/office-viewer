import type { CellScalar, ChartModel, SpreadsheetCell, WorkbookModel } from '../../domain/model'
import { chartNumber, type ChartRecords } from './chart-data'
import { readPivotChartSource } from './chart-pivot-data'
import { chartSourceVisibility } from './chart-visibility'
import { formatSheetValue } from '../sheet-workbench/sheet-number-format'
import { COMPUTE_LIMITS } from '../../compute/protocol.ts'
import { ComputeError } from '../../compute/errors.ts'
import { assertChartPreparationMemory, CHART_PREPARATION_CHUNK, consumeChartPreparation,
  type ChartComputePreparationContext } from './chart-compute-lifecycle'

export type ChartRecordSource = Pick<ChartModel, 'sheetId' | 'source' | 'pivotTableId' | 'options'> & Partial<Pick<ChartModel, 'type'>>

export function readChartRecordsAsync(workbook: WorkbookModel, chart: ChartRecordSource, context: ChartComputePreparationContext = {}): Promise<ChartRecords> {
  return consumeChartPreparation(readChartRecordSteps(workbook, chart, context), context)
}

/** Read directly from sparse cells in the authored orientation. No dense
 * worksheet matrix or additional transpose is retained during preparation. */
export function* readChartRecordSteps(workbook: WorkbookModel, chart: ChartRecordSource, context: ChartComputePreparationContext = {}): Generator<void, ChartRecords, void> {
  if (chart.pivotTableId) {
    const source = readPivotChartSource(workbook, chart.pivotTableId)
    const numericFields = new Set<number>()
    let work = 0
    for (const record of source.records) for (let field = 0; field < source.fields.length; field++) {
      if (!numericFields.has(field) && chartNumber(record[field]) !== null) numericFields.add(field)
      if (++work % CHART_PREPARATION_CHUNK === 0) yield
    }
    return { fields: source.fields, records: source.records, date1904: workbook.date1904, sourceCategoryFields: source.categoryFields,
      sourceValueFields: source.valueFields, sourceNumericFields: [...numericFields].sort((a, b) => a - b), ...(source.syntheticValues ? { syntheticValues: true } : {}) }
  }
  const sheet = workbook.sheets.find(candidate => candidate.id === (chart.source.sheetId ?? chart.sheetId))
  if (!sheet) throw new ComputeError('COMPUTE_INVALID_INPUT', '图表源工作表不存在')
  const range = chart.source.range
  for (const key of ['startRow', 'endRow', 'startCol', 'endCol'] as const) {
    const value = range[key], maximum = key.endsWith('Row') ? 1_048_575 : 16_383
    if (!Number.isSafeInteger(value) || value < 0 || value > maximum) throw new ComputeError('COMPUTE_INVALID_INPUT', '图表数据源坐标无效')
  }
  if (range.startRow > range.endRow || range.startCol > range.endCol || range.endRow >= (sheet.rowCount ?? 1_048_576)
    || range.endCol >= (sheet.columnCount ?? 16_384) || !['columns', 'rows'].includes(chart.source.orientation)) {
    throw new ComputeError('COMPUTE_INVALID_INPUT', '图表数据源区域或方向无效')
  }
  const byColumns = chart.source.orientation === 'columns', height = range.endRow - range.startRow + 1, width = range.endCol - range.startCol + 1
  const fieldCount = byColumns ? width : height, sampleCount = (byColumns ? height : width) - 1
  if (sampleCount < 1 || fieldCount < 1) throw new ComputeError('COMPUTE_INVALID_INPUT', '图表数据区域须包含字段名和至少一条数据')
  if (sampleCount > COMPUTE_LIMITS.maxRows || fieldCount * (sampleCount + 1) > COMPUTE_LIMITS.maxElements) {
    throw new ComputeError('COMPUTE_MEMORY_LIMIT', '图表数据源的完整区域超出样本预算')
  }
  // Conservative JS reference/row/string allowance, separate from exact
  // TypedArray boundary bytes. Actual heap usage remains a benchmark metric.
  const retainFormatting = chart.type === 'heatmap' || chart.type === 'treemap' || chart.type === 'sunburst'
  let estimatedBytes = fieldCount * sampleCount * (retainFormatting ? 40 : 24) + sampleCount * (retainFormatting ? 96 : 48) + fieldCount * 64 + 8 * 1024 * 1024
  assertChartPreparationMemory(estimatedBytes, context)
  // Fixed v1.6.0 IChartContext declares true as the historical default.
  const visibility = chart.options?.plotVisibleOnly !== false ? chartSourceVisibility(workbook, sheet, range) : undefined
  const fields: string[] = [], records: CellScalar[][] = [], visibleFields: number[] = []
  const numericFields = new Set<number>()
  const formattedRecords: Array<Array<string | null>> | undefined = retainFormatting ? [] : undefined
  let work = 0
  for (let field = 0; field < fieldCount; field++) {
    const row = range.startRow + (byColumns ? 0 : field), column = range.startCol + (byColumns ? field : 0)
    const value = scalarValue(sheet.cells[`${row}:${column}`]), label = value == null ? '' : String(value)
    fields.push(label); estimatedBytes += label.length * 2
    if (!visibility || (byColumns ? visibility.columnVisible(column) : visibility.rowVisible(row))) visibleFields.push(field)
    if (++work % CHART_PREPARATION_CHUNK === 0) { assertChartPreparationMemory(estimatedBytes, context); yield }
  }
  for (let sample = 0; sample < sampleCount; sample++) {
    if (++work % CHART_PREPARATION_CHUNK === 0) { assertChartPreparationMemory(estimatedBytes, context); yield }
    const position = (byColumns ? range.startRow : range.startCol) + sample + 1
    if (visibility && !(byColumns ? visibility.rowVisible(position) : visibility.columnVisible(position))) continue
    const record: CellScalar[] = []
    const formattedRecord: Array<string | null> | undefined = formattedRecords ? [] : undefined
    for (let field = 0; field < fieldCount; field++) {
      const row = byColumns ? position : range.startRow + field, column = byColumns ? range.startCol + field : position
      const cell = sheet.cells[`${row}:${column}`], value = scalarValue(cell)
      record.push(value)
      if (!numericFields.has(field) && chartNumber(value) !== null) numericFields.add(field)
      if (typeof value === 'string') estimatedBytes += value.length * 2
      if (formattedRecord) {
        const display = value === null ? null : formatSheetValue(value, cell?.style.numberFormat ?? 'General', workbook.date1904)
          ?? cell?.display ?? String(value)
        formattedRecord.push(display)
        if (display !== null) estimatedBytes += 32 + display.length * 2
      }
      if (++work % CHART_PREPARATION_CHUNK === 0) { assertChartPreparationMemory(estimatedBytes, context); yield }
    }
    records.push(record)
    if (formattedRecord) formattedRecords!.push(formattedRecord)
    if (++work % CHART_PREPARATION_CHUNK === 0) { assertChartPreparationMemory(estimatedBytes, context); yield }
  }
  assertChartPreparationMemory(estimatedBytes, context)
  return { fields, records, date1904: workbook.date1904, sourceNumericFields: [...numericFields].sort((a, b) => a - b),
    ...(formattedRecords ? { formattedRecords } : {}), ...(visibility ? { visibleFields } : {}) }
}

function scalarValue(cell: SpreadsheetCell | undefined): CellScalar {
  if (!cell) return null
  // Formula recalculation and import put the authoritative cache in value.
  // A formatted/rounded display must never become a statistics input.
  return cell.value
}
