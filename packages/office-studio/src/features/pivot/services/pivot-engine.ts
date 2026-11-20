import {
  createCell,
  defaultCellStyle,
  type CellRange,
  type CellScalar,
  type PivotAggregation,
  type PivotDateGroupType,
  type PivotDateGroupingMode,
  type PivotFieldArea,
  type PivotFieldFilter,
  type PivotFieldModel,
  type PivotManualGroup,
  type PivotNumberGroup,
  type PivotTableModel,
  type PivotTableOptions,
  type PivotTableSource,
  type SpreadsheetSheet,
  type WorkbookModel,
} from '../../../domain/model'
import { randomId } from '../../../api/model-utils'

export interface PivotSourceFieldInfo {
  fieldId: string
  fieldName: string
  sourceColumn: number
}

export interface PivotDrillDownEntry {
  fieldFilters: Record<string, CellScalar>
  fieldMembers?: Record<string, CellScalar[]>
  valueFieldId: string
}

export interface PivotTableModelWithDrillDown extends PivotTableModel {
  drillDownMap?: Record<string, PivotDrillDownEntry>
}

export interface PivotTableSourceInput {
  unitId: string
  subUnitId: string
  sheetName: string
  range: { startRow: number; startColumn: number; endRow: number; endColumn: number }
}

export interface PivotTableAnchorInput {
  unitId: string
  subUnitId: string
  row: number
  col: number
}

export interface PivotRenderResult {
  outputRange?: CellRange
  generatedCellKeys: string[]
  drillDownMap: Record<string, PivotDrillDownEntry>
}

interface SourceRecord {
  values: CellScalar[]
}

interface Group {
  values: CellScalar[]
  key: string
  memberValues: Record<string, CellScalar[]>
}

export type PivotDateGroup = PivotDateGroupType
export type PivotDateMode = PivotDateGroupingMode
export type PivotNumberGroupSettings = PivotNumberGroup
export type PivotManualGroupSettings = PivotManualGroup

interface OutputRow {
  group: Group | null
  valueFieldId?: string
  total: boolean
}

interface OutputColumn {
  group: Group | null
  valueFieldId?: string
  total: boolean
}

const defaultOptions = (): PivotTableOptions => ({
  repeatRowLabels: false,
  showRowGrandTotal: true,
  showColGrandTotal: true,
  showRowSubtotals: true,
  showColSubtotals: true,
  valuePosition: 'column',
})

export function normalizePivotSource(input: PivotTableSourceInput, workbook: WorkbookModel): PivotTableSource {
  const sheet = workbook.sheets.find((candidate) => candidate.id === input.subUnitId)
  if (!sheet || workbook.id !== input.unitId) throw new Error('Pivot source worksheet does not belong to the workbook')
  const range = {
    startRow: input.range.startRow,
    startCol: input.range.startColumn,
    endRow: input.range.endRow,
    endCol: input.range.endColumn,
  }
  validateRange(range)
  if (range.endRow <= range.startRow) throw new RangeError('Pivot source must contain a header and at least one data row')
  if (range.endCol < range.startCol) throw new RangeError('Pivot source must contain at least one column')
  return { unitId: input.unitId, subUnitId: input.subUnitId, sheetName: input.sheetName, range }
}

export function createPivotTable(
  workbook: WorkbookModel,
  sourceInput: PivotTableSourceInput,
  anchorInput: PivotTableAnchorInput,
  id = randomId('pivot')
): PivotTableModelWithDrillDown {
  const source = normalizePivotSource(sourceInput, workbook)
  const anchorSheet = workbook.sheets.find((sheet) => sheet.id === anchorInput.subUnitId)
  if (!anchorSheet || anchorInput.unitId !== workbook.id) throw new Error('Pivot anchor worksheet does not belong to the workbook')
  validateIndex(anchorInput.row, 'anchor row')
  validateIndex(anchorInput.col, 'anchor column')
  const model: PivotTableModelWithDrillDown = {
    id,
    source,
    anchor: { unitId: anchorInput.unitId, subUnitId: anchorInput.subUnitId, row: anchorInput.row, col: anchorInput.col },
    fields: [],
    options: defaultOptions(),
    revision: 0,
    generatedCellKeys: [],
    drillDownMap: {},
  }
  // Reserve the requested anchor as part of creation. Without this preflight,
  // a pivot with no value field has no output range yet and a later field
  // mutation could fail after leaving a partially configured pivot in history.
  ensureNoOutputConflict(
    workbook,
    model,
    { startRow: anchorInput.row, startCol: anchorInput.col, endRow: anchorInput.row, endCol: anchorInput.col },
    anchorSheet,
  )
  workbook.pivotTables ??= []
  workbook.pivotTables.push(model)
  renderPivotTable(workbook, model)
  return model
}

export function getPivotSourceFields(workbook: WorkbookModel, pivot: PivotTableModel): PivotSourceFieldInfo[] {
  const source = sourceSheet(workbook, pivot)
  const result: PivotSourceFieldInfo[] = []
  for (let column = pivot.source.range.startCol; column <= pivot.source.range.endCol; column++) {
    const header = cellValue(source, pivot.source.range.startRow, column)
    result.push({ fieldId: String(column - pivot.source.range.startCol), fieldName: header == null || header === '' ? `Column ${column - pivot.source.range.startCol + 1}` : String(header), sourceColumn: column })
  }
  return result
}

export function setPivotField(
  workbook: WorkbookModel,
  pivot: PivotTableModel,
  fieldIdOrIndex: string | number,
  area: PivotFieldArea,
  position?: number
): void {
  const info = resolveField(workbook, pivot, fieldIdOrIndex)
  const existing = pivot.fields.find((field) => field.id === info.fieldId)
  if (existing) {
    const nextPosition = normalizePosition(position, areaFields(pivot, area, existing.id).length)
    if (existing.area === area && existing.position === nextPosition) return
    existing.area = area
    existing.position = nextPosition
  } else {
    const field: PivotFieldModel = {
      id: info.fieldId,
      sourceColumn: info.sourceColumn,
      name: info.fieldName,
      area,
      position: normalizePosition(position, areaFields(pivot, area).length),
      ...(area === 'value' ? { aggregation: 'sum' as const } : {}),
    }
    pivot.fields.push(field)
  }
  reindexFields(pivot)
  refreshPivotTable(workbook, pivot)
}

export function removePivotFields(workbook: WorkbookModel, pivot: PivotTableModel, fieldIds: string | string[]): void {
  const ids = new Set(Array.isArray(fieldIds) ? fieldIds.map(String) : [String(fieldIds)])
  pivot.fields = pivot.fields.filter((field) => !ids.has(field.id))
  reindexFields(pivot)
  refreshPivotTable(workbook, pivot)
}

export function setPivotOptions(workbook: WorkbookModel, pivot: PivotTableModel, options: Partial<PivotTableOptions>): void {
  if (options.valuePosition && !['row', 'column'].includes(options.valuePosition)) throw new TypeError('Invalid pivot value position')
  pivot.options = { ...pivot.options, ...options }
  refreshPivotTable(workbook, pivot)
}

export function setPivotFieldSetting(
  workbook: WorkbookModel,
  pivot: PivotTableModel,
  fieldId: string,
  setting: { displayName?: string; format?: string; subtotalType?: PivotAggregation }
): void {
  const field = requirePivotField(pivot, fieldId)
  if (setting.displayName !== undefined) field.displayName = setting.displayName
  if (setting.format !== undefined) field.format = setting.format
  if (setting.subtotalType !== undefined) field.aggregation = setting.subtotalType
  refreshPivotTable(workbook, pivot)
}

export function setPivotSort(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string, direction: 'ascending' | 'descending'): void {
  const field = requirePivotField(pivot, fieldId)
  if (field.area !== 'row' && field.area !== 'column') throw new Error('Only row and column fields can be label-sorted')
  field.sort = direction
  refreshPivotTable(workbook, pivot)
}

export function setPivotLabelFilter(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string, selectedItems: string[] | null): void {
  const field = requirePivotField(pivot, fieldId)
  if (field.area === 'value') throw new Error('Value fields cannot use label filters')
  field.selectedItems = selectedItems == null ? undefined : selectedItems.map(String)
  refreshPivotTable(workbook, pivot)
}

export function setPivotValueFilter(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string, filter: PivotFieldFilter | null): void {
  const field = requirePivotField(pivot, fieldId)
  if (field.area === 'value') throw new Error('Value filters require a row, column, or filter field')
  if (filter && !pivot.fields.some((candidate) => candidate.id === filter.valueFieldId && candidate.area === 'value')) {
    throw new Error(`Value field not found: ${filter.valueFieldId}`)
  }
  field.valueFilter = filter ? { ...filter } : undefined
  refreshPivotTable(workbook, pivot)
}

/** Configure numeric buckets for a row/column/filter label field. Passing null removes grouping. */
export function setPivotNumberGroup(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string, settings: PivotNumberGroup | null): void {
  const field = requirePivotField(pivot, fieldId)
  if (field.area === 'value') throw new Error('Numeric grouping requires a label field')
  if (settings != null) validateNumberGroup(settings)
  field.numberGroup = settings == null ? undefined : { ...settings }
  if (settings != null) { field.manualGroups = undefined; field.dateGroupType = undefined }
  refreshPivotTable(workbook, pivot)
}

/** Select the calendar grouping mode used by date label fields. */
export function setPivotDateGroupingMode(workbook: WorkbookModel, pivot: PivotTableModel, mode: PivotDateGroupingMode): void {
  if (mode !== 'derived' && mode !== 'excelLike') throw new TypeError(`Invalid pivot date grouping mode: ${mode}`)
  pivot.dateGroupingMode = mode
  refreshPivotTable(workbook, pivot)
}

/** Configure a date bucket for a row/column/filter label field. */
export function setPivotDateGroupType(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string, type: PivotDateGroupType | null): void {
  const field = requirePivotField(pivot, fieldId)
  if (field.area === 'value') throw new Error('Date grouping requires a label field')
  if (type != null && !['year', 'quarter', 'month', 'day'].includes(type)) throw new TypeError(`Invalid pivot date group type: ${type}`)
  field.dateGroupType = type == null ? undefined : type
  if (type != null) { field.manualGroups = undefined; field.numberGroup = undefined }
  refreshPivotTable(workbook, pivot)
}

/** Excel-like calendar grouping entry point; requires the explicit mode to avoid accidental text parsing. */
export function setPivotExcelDateGroup(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string, type: PivotDateGroupType): void {
  if (pivot.dateGroupingMode !== 'excelLike') throw new Error('Excel-like date grouping requires date grouping mode excelLike')
  setPivotDateGroupType(workbook, pivot, fieldId, type)
}

/** Set or clear manual member groups. Every non-empty group must contain at least two distinct members. */
export function setPivotManualGroups(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string, groups: PivotManualGroup[]): void {
  const field = requirePivotField(pivot, fieldId)
  if (field.area === 'value') throw new Error('Manual grouping requires a label field')
  const normalized = normalizeManualGroups(groups)
  field.manualGroups = normalized.length ? normalized : undefined
  if (normalized.length) { field.numberGroup = undefined; field.dateGroupType = undefined }
  refreshPivotTable(workbook, pivot)
}

export function refreshPivotTable(workbook: WorkbookModel, pivot: PivotTableModel): void {
  pivot.revision += 1
  renderPivotTable(workbook, pivot)
}

export function renderPivotTable(workbook: WorkbookModel, pivot: PivotTableModel): PivotRenderResult {
  const target = targetSheet(workbook, pivot)
  removeGeneratedCells(target, pivot)
  const result = buildPivotOutput(workbook, pivot)
  if (!result.values.length) {
    pivot.outputRange = undefined
    pivot.generatedCellKeys = []
    pivot.drillDownMap = {}
    return { generatedCellKeys: [], drillDownMap: {} }
  }
  const output = {
    startRow: pivot.anchor.row,
    startCol: pivot.anchor.col,
    endRow: pivot.anchor.row + result.values.length - 1,
    endCol: pivot.anchor.col + result.values[0]!.length - 1,
  }
  ensureNoOutputConflict(workbook, pivot, output, target)
  const generatedCellKeys: string[] = []
  for (let row = 0; row < result.values.length; row++) {
    for (let col = 0; col < result.values[row]!.length; col++) {
      const targetRow = output.startRow + row
      const targetCol = output.startCol + col
      const key = `${targetRow}:${targetCol}`
      const cell = createCell(targetRow, targetCol, result.values[row]![col] ?? null)
      cell.style = { ...defaultCellStyle(), font: { ...defaultCellStyle().font, bold: row < result.headerRows } }
      cell.metadata = { pivotTableId: pivot.id, pivotGenerated: true }
      target.cells[key] = cell
      generatedCellKeys.push(key)
    }
  }
  pivot.outputRange = output
  pivot.generatedCellKeys = generatedCellKeys
  pivot.drillDownMap = result.drillDownMap
  return { outputRange: output, generatedCellKeys, drillDownMap: result.drillDownMap }
}

export function getPivotTableAtCell(workbook: WorkbookModel, sheetId: string, row: number, col: number): PivotTableModel | undefined {
  return workbook.pivotTables?.find((pivot) => pivot.anchor.subUnitId === sheetId && pivot.outputRange && contains(pivot.outputRange, row, col))
}

export function pivotOutputValues(workbook: WorkbookModel, pivot: PivotTableModel): CellScalar[][] {
  const range = pivot.outputRange
  if (!range) return []
  const sheet = targetSheet(workbook, pivot)
  const result: CellScalar[][] = []
  for (let row = range.startRow; row <= range.endRow; row++) {
    const values: CellScalar[] = []
    for (let col = range.startCol; col <= range.endCol; col++) values.push(cellValue(sheet, row, col))
    result.push(values)
  }
  return result
}

export function pivotDrillDown(
  workbook: WorkbookModel,
  pivot: PivotTableModel,
  row: number,
  col: number
): { headers: string[]; rows: CellScalar[][] } {
  validateIndex(row, 'pivot row offset')
  validateIndex(col, 'pivot column offset')
  const map = (pivot as PivotTableModelWithDrillDown).drillDownMap ?? {}
  const entry = map[`${row}:${col}`] ?? map[`${pivot.anchor.row + row}:${pivot.anchor.col + col}`]
  if (!entry) throw new RangeError('The selected pivot cell does not contain drill-down detail')
  const source = sourceSheet(workbook, pivot)
  const fields = getPivotSourceFields(workbook, pivot)
  const records = readSourceRecords(source, pivot)
  const filtered = records.filter((record) => matchesFilters(record, fields, entry.fieldFilters, entry.fieldMembers))
  return { headers: fields.map((field) => field.fieldName), rows: filtered.map((record) => record.values) }
}

export function pivotFieldIds(pivot: PivotTableModel, area: PivotFieldArea): string[] {
  return pivot.fields.filter((field) => field.area === area).sort((a, b) => a.position - b.position).map((field) => field.id)
}

function buildPivotOutput(workbook: WorkbookModel, pivot: PivotTableModel): { values: CellScalar[][]; headerRows: number; drillDownMap: Record<string, PivotDrillDownEntry> } {
  const fields = getPivotSourceFields(workbook, pivot)
  const byId = new Map(pivot.fields.map((field) => [field.id, field]))
  const rows = sortedFields(pivot, 'row')
  const columns = sortedFields(pivot, 'column')
  const filters = sortedFields(pivot, 'filter')
  const values = sortedFields(pivot, 'value')
  if (!values.length) return { values: [], headerRows: 0, drillDownMap: {} }
  const allRecords = readSourceRecords(sourceSheet(workbook, pivot), pivot)
  let records = allRecords.filter((record) => filters.every((field) => field.selectedItems == null || field.selectedItems.includes(String(record.values[field.sourceColumn - pivot.source.range.startCol] ?? ''))))
  records = records.filter((record) => [...rows, ...columns].every((field) => field.selectedItems == null || field.selectedItems.includes(String(record.values[field.sourceColumn - pivot.source.range.startCol] ?? ''))))
  records = applyValueFilters(records, rows.concat(columns), values, byId, pivot)

  const rowGroups = groupsFor(records, rows, pivot)
  const colGroups = groupsFor(records, columns, pivot)
  const valueRows = pivot.options.valuePosition === 'row'
  const emptyGroup = (): Group => ({ values: [], key: '[]', memberValues: {} })
  const rowAxis = rows.length ? rowGroups : [emptyGroup()]
  const colAxis = columns.length ? colGroups : [emptyGroup()]
  const outputRows: OutputRow[] = []
  for (const group of rowAxis) for (const field of valueRows ? values : [undefined]) outputRows.push({ group, valueFieldId: field?.id, total: false })
  if (pivot.options.showRowGrandTotal && rows.length) {
    for (const field of valueRows ? values : [undefined]) outputRows.push({ group: null, valueFieldId: field?.id, total: true })
  }
  const outputColumns: OutputColumn[] = []
  for (const group of colAxis) for (const field of valueRows ? [undefined] : values) outputColumns.push({ group, valueFieldId: field?.id, total: false })
  if (pivot.options.showColGrandTotal && columns.length) {
    for (const field of valueRows ? [undefined] : values) outputColumns.push({ group: null, valueFieldId: field?.id, total: true })
  }

  const headerRows = Math.max(1, columns.length + (valueRows ? 0 : 1))
  const width = rows.length + (valueRows ? 1 : 0) + outputColumns.length
  // Allocate both the header and detail rows before populating the sparse output.
  // Pivot output is a rectangular range even when a field has no matching records.
  const matrix: CellScalar[][] = Array.from({ length: headerRows + outputRows.length }, () => Array.from({ length: width }, () => null))
  const drillDownMap: Record<string, PivotDrillDownEntry> = {}
  const rowHeaderCount = rows.length + (valueRows ? 1 : 0)
  rows.forEach((field, index) => { matrix[headerRows - 1]![index] = field.displayName ?? field.name })
  if (valueRows) matrix[headerRows - 1]![rows.length] = 'Values'
  for (let colIndex = 0; colIndex < outputColumns.length; colIndex++) {
    const outputColumn = outputColumns[colIndex]!
    const targetColumn = rowHeaderCount + colIndex
    columns.forEach((_field, level) => { matrix[level]![targetColumn] = outputColumn.total ? 'Grand total' : outputColumn.group?.values[level] ?? null })
    if (!valueRows) matrix[headerRows - 1]![targetColumn] = outputColumn.total
      ? (outputColumn.valueFieldId ? displayValueField(byId.get(outputColumn.valueFieldId)!, values.length) : 'Grand total')
      : outputColumn.valueFieldId ? displayValueField(byId.get(outputColumn.valueFieldId)!, values.length) : 'Value'
    else if (headerRows === 1) matrix[0]![targetColumn] = columns.length ? outputColumn.group?.values[0] ?? null : 'Value'
  }
  if (columns.length && !valueRows && values.length > 1) {
    for (let i = 0; i < outputColumns.length; i++) {
      const outputColumn = outputColumns[i]!
      if (outputColumn.total) matrix[headerRows - 1]![rowHeaderCount + i] = `Grand total ${displayValueField(byId.get(outputColumn.valueFieldId!)!, values.length)}`
    }
  }

  for (let rowIndex = 0; rowIndex < outputRows.length; rowIndex++) {
    const outputRow = outputRows[rowIndex]!
    const targetRow = headerRows + rowIndex
    const rowFieldFilters: Record<string, CellScalar> = {}
    const rowFieldMembers: Record<string, CellScalar[]> = {}
    if (outputRow.group) rows.forEach((field, index) => {
      rowFieldFilters[field.id] = outputRow.group!.values[index] ?? null
      rowFieldMembers[field.id] = outputRow.group!.memberValues[field.id] ?? [outputRow.group!.values[index] ?? null]
    })
    if (outputRow.valueFieldId && valueRows) {
      const valueField = byId.get(outputRow.valueFieldId)!
      matrix[targetRow]![rows.length] = outputRow.total ? 'Grand total' : valueField.displayName ?? valueField.name
    } else {
      rows.forEach((_field, index) => { matrix[targetRow]![index] = outputRow.total ? 'Grand total' : outputRow.group?.values[index] ?? null })
    }
    for (let colIndex = 0; colIndex < outputColumns.length; colIndex++) {
      const outputColumn = outputColumns[colIndex]!
      const valueFieldId = valueRows ? outputRow.valueFieldId! : outputColumn.valueFieldId!
      const colFieldFilters: Record<string, CellScalar> = {}
      const colFieldMembers: Record<string, CellScalar[]> = {}
      if (outputColumn.group) columns.forEach((field, index) => {
        colFieldFilters[field.id] = outputColumn.group!.values[index] ?? null
        colFieldMembers[field.id] = outputColumn.group!.memberValues[field.id] ?? [outputColumn.group!.values[index] ?? null]
      })
      const groupFilter = { ...rowFieldFilters, ...colFieldFilters }
      const groupMembers = { ...rowFieldMembers, ...colFieldMembers }
      const aggregateRecords = records.filter((record) => matchesFilters(record, fields, groupFilter, groupMembers))
      const valueField = byId.get(valueFieldId)!
      const aggregate = aggregateValue(aggregateRecords, valueField, pivot)
      const cellColumn = rowHeaderCount + colIndex
      matrix[targetRow]![cellColumn] = aggregate
      if (!outputRow.total && !outputColumn.total) drillDownMap[`${targetRow}:${cellColumn}`] = { fieldFilters: groupFilter, fieldMembers: groupMembers, valueFieldId }
    }
  }
  return { values: matrix, headerRows, drillDownMap }
}

function applyValueFilters(records: SourceRecord[], dimensionFields: PivotFieldModel[], _values: PivotFieldModel[], byId: Map<string, PivotFieldModel>, pivot: PivotTableModel): SourceRecord[] {
  const active = dimensionFields.filter((field) => field.valueFilter)
  if (!active.length) return records
  return records.filter((record) => active.every((field) => {
    const filter = field.valueFilter!
    const valueField = byId.get(filter.valueFieldId)
    if (!valueField) return false
    const label = String(record.values[field.sourceColumn - pivot.source.range.startCol] ?? '')
    const matching = records.filter((candidate) => String(candidate.values[field.sourceColumn - pivot.source.range.startCol] ?? '') === label)
    const value = aggregateValue(matching, valueField, pivot)
    return typeof value === 'number' && testFilter(value, filter)
  }))
}

function testFilter(value: number, filter: PivotFieldFilter): boolean {
  switch (filter.operator) {
    case 'valueGreaterThan': return value > filter.expected
    case 'valueGreaterThanOrEqual': return value >= filter.expected
    case 'valueLessThan': return value < filter.expected
    case 'valueLessThanOrEqual': return value <= filter.expected
    case 'valueEqual': return value === filter.expected
    case 'valueNotEqual': return value !== filter.expected
    case 'valueBetween': return value >= Math.min(filter.expected, filter.expected2 ?? filter.expected) && value <= Math.max(filter.expected, filter.expected2 ?? filter.expected)
    case 'valueNotBetween': return value < Math.min(filter.expected, filter.expected2 ?? filter.expected) || value > Math.max(filter.expected, filter.expected2 ?? filter.expected)
  }
}

function groupsFor(records: SourceRecord[], fields: PivotFieldModel[], pivot: PivotTableModel): Group[] {
  if (!fields.length) return []
  const result = new Map<string, Group>()
  for (const record of records) {
    const values = fields.map((field) => groupedLabel(field, record.values[field.sourceColumn - pivot.source.range.startCol] ?? null))
    const key = JSON.stringify(values)
    if (!result.has(key)) {
      const memberValues: Record<string, CellScalar[]> = {}
      fields.forEach((field) => { memberValues[field.id] = [record.values[field.sourceColumn - pivot.source.range.startCol] ?? null] })
      result.set(key, { values, key, memberValues })
    } else {
      const group = result.get(key)!
      fields.forEach((field) => {
        const sourceValue = record.values[field.sourceColumn - pivot.source.range.startCol] ?? null
        const members = group.memberValues[field.id] ?? (group.memberValues[field.id] = [])
        if (!members.some((member) => Object.is(member, sourceValue))) members.push(sourceValue)
      })
    }
  }
  const groups = [...result.values()]
  for (let index = fields.length - 1; index >= 0; index--) {
    const field = fields[index]!
    if (!field.sort) continue
    groups.sort((a, b) => {
      const order = compareLabels(a.values[index] ?? null, b.values[index] ?? null)
      return field.sort === 'ascending' ? order : -order
    })
  }
  return groups
}

function aggregateValue(records: SourceRecord[], field: PivotFieldModel, pivot: PivotTableModel): CellScalar {
  const index = field.sourceColumn - pivot.source.range.startCol
  const values = records.map((record) => record.values[index]).filter((value) => value !== null && value !== '')
  const aggregation = field.aggregation ?? 'sum'
  if (aggregation === 'count') return values.length
  const numbers = values.map(Number).filter(Number.isFinite)
  if (!numbers.length) return null
  if (aggregation === 'min') return Math.min(...numbers)
  if (aggregation === 'max') return Math.max(...numbers)
  const sum = numbers.reduce((total, value) => total + value, 0)
  if (aggregation === 'average') return sum / numbers.length
  return sum
}

function readSourceRecords(sheet: SpreadsheetSheet, pivot: PivotTableModel): SourceRecord[] {
  const records: SourceRecord[] = []
  for (let row = pivot.source.range.startRow + 1; row <= pivot.source.range.endRow; row++) {
    records.push({ values: Array.from({ length: pivot.source.range.endCol - pivot.source.range.startCol + 1 }, (_, index) => cellValue(sheet, row, pivot.source.range.startCol + index)) })
  }
  return records
}

function resolveField(workbook: WorkbookModel, pivot: PivotTableModel, value: string | number): PivotSourceFieldInfo {
  const fields = getPivotSourceFields(workbook, pivot)
  const normalized = String(value)
  const field = fields.find((candidate) => candidate.fieldId === normalized)
  if (!field) throw new Error(`Pivot source field not found: ${normalized}`)
  return field
}

function sourceSheet(workbook: WorkbookModel, pivot: PivotTableModel): SpreadsheetSheet {
  if (workbook.id !== pivot.source.unitId) throw new Error('Pivot source workbook is not available')
  const sheet = workbook.sheets.find((candidate) => candidate.id === pivot.source.subUnitId)
  if (!sheet) throw new Error(`Pivot source worksheet not found: ${pivot.source.subUnitId}`)
  return sheet
}

function targetSheet(workbook: WorkbookModel, pivot: PivotTableModel): SpreadsheetSheet {
  if (workbook.id !== pivot.anchor.unitId) throw new Error('Pivot target workbook is not available')
  const sheet = workbook.sheets.find((candidate) => candidate.id === pivot.anchor.subUnitId)
  if (!sheet) throw new Error(`Pivot target worksheet not found: ${pivot.anchor.subUnitId}`)
  return sheet
}

function cellValue(sheet: SpreadsheetSheet, row: number, col: number): CellScalar {
  const cell = sheet.cells[`${row}:${col}`]
  return cell?.value ?? null
}

function removeGeneratedCells(sheet: SpreadsheetSheet, pivot: PivotTableModel): void {
  for (const key of pivot.generatedCellKeys ?? []) {
    if (sheet.cells[key]?.metadata?.pivotTableId === pivot.id) delete sheet.cells[key]
  }
}

function ensureNoOutputConflict(workbook: WorkbookModel, pivot: PivotTableModel, range: CellRange, target: SpreadsheetSheet): void {
  if (pivot.source.subUnitId === pivot.anchor.subUnitId && overlaps(range, pivot.source.range)) throw new Error('Pivot output cannot overlap its source range')
  const allowed = new Set(pivot.generatedCellKeys ?? [])
  for (let row = range.startRow; row <= range.endRow; row++) for (let col = range.startCol; col <= range.endCol; col++) {
    const key = `${row}:${col}`
    const existing = target.cells[key]
    // Clearing a range leaves a sparse cell record with null content in the
    // public API. It must not keep an old pivot anchor blocked forever, while
    // formulas, values, and displayed text still count as real content.
    const hasContent = existing != null && (existing.value !== null || existing.formula !== null || existing.display !== '')
    if (hasContent && !allowed.has(key)) throw new Error(`Pivot output overlaps existing content at ${key}`)
  }
  for (const other of workbook.pivotTables ?? []) {
    if (other.id !== pivot.id && other.anchor.subUnitId === pivot.anchor.subUnitId && other.outputRange && overlaps(range, other.outputRange)) throw new Error(`Pivot output overlaps pivot table ${other.id}`)
  }
}

function matchesFilters(record: SourceRecord, fields: PivotSourceFieldInfo[], fieldFilters: Record<string, CellScalar>, fieldMembers: Record<string, CellScalar[]> = {}): boolean {
  return Object.entries(fieldFilters).every(([fieldId, value]) => {
    const field = fields.find((candidate) => candidate.fieldId === fieldId)
    if (!field) return false
    const sourceValue = record.values[field.sourceColumn - fields[0]!.sourceColumn] ?? null
    return (fieldMembers[fieldId] ?? [value]).some((member) => Object.is(member, sourceValue))
  })
}

function areaFields(pivot: PivotTableModel, area: PivotFieldArea, omitId?: string): PivotFieldModel[] {
  return pivot.fields.filter((field) => field.area === area && field.id !== omitId).sort((a, b) => a.position - b.position)
}

function sortedFields(pivot: PivotTableModel, area: PivotFieldArea): PivotFieldModel[] {
  return areaFields(pivot, area)
}

function requirePivotField(pivot: PivotTableModel, id: string): PivotFieldModel {
  const field = pivot.fields.find((candidate) => candidate.id === String(id))
  if (!field) throw new Error(`Pivot field not configured: ${id}`)
  return field
}

function reindexFields(pivot: PivotTableModel): void {
  for (const area of ['filter', 'row', 'column', 'value'] as const) areaFields(pivot, area).forEach((field, index) => { field.position = index })
}

function normalizePosition(position: number | undefined, length: number): number {
  if (position == null) return length
  if (!Number.isInteger(position) || position < 0 || position > length) throw new RangeError(`Invalid pivot field position: ${position}`)
  return position
}

function displayValueField(field: PivotFieldModel, count: number): string {
  return field.displayName ?? (count > 1 ? field.name : field.aggregation === 'count' ? `Count of ${field.name}` : field.name)
}

function compareLabels(a: CellScalar, b: CellScalar): number {
  if (typeof a === 'number' && typeof b === 'number') return a - b
  return String(a ?? '').localeCompare(String(b ?? ''), undefined, { numeric: true, sensitivity: 'base' })
}

function groupedLabel(field: PivotFieldModel, value: CellScalar): CellScalar {
  if (field.manualGroups?.length) {
    const group = field.manualGroups.find((candidate) => candidate.members.some((member) => Object.is(member, value)))
    return group?.name ?? value
  }
  if (field.numberGroup && typeof value === 'number' && Number.isFinite(value)) {
    const { start, end, interval } = field.numberGroup
    if (value < start) return `< ${start}`
    if (value > end) return `> ${end}`
    const bucket = Math.min(Math.floor((value - start) / interval), Math.max(0, Math.ceil((end - start) / interval) - 1))
    const lower = start + bucket * interval
    const upper = Math.min(end, lower + interval)
    return `${formatGroupNumber(lower)}-${formatGroupNumber(upper)}`
  }
  if (field.dateGroupType) return dateGroupLabel(value, field.dateGroupType)
  return value
}

function dateGroupLabel(value: CellScalar, type: PivotDateGroupType): CellScalar {
  const date = parsePivotDate(value)
  if (!date) return value
  const year = date.getUTCFullYear()
  const month = date.getUTCMonth() + 1
  if (type === 'year') return String(year)
  if (type === 'quarter') return `${year} Q${Math.floor((month - 1) / 3) + 1}`
  if (type === 'month') return `${year}-${String(month).padStart(2, '0')}`
  return `${year}-${String(month).padStart(2, '0')}-${String(date.getUTCDate()).padStart(2, '0')}`
}

function parsePivotDate(value: CellScalar): Date | undefined {
  if (typeof value === 'string') {
    const timestamp = Date.parse(value)
    if (!Number.isNaN(timestamp)) return new Date(timestamp)
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    // Numeric spreadsheet dates use the workbook's 1900/1904 date system.
    const epoch = Date.UTC(1899, 11, 30)
    const timestamp = epoch + value * 86400000
    const date = new Date(timestamp)
    if (!Number.isNaN(date.getTime())) return date
  }
  return undefined
}

function formatGroupNumber(value: number): string {
  return Number.isInteger(value) ? String(value) : String(Number(value.toFixed(12)))
}

function validateNumberGroup(settings: PivotNumberGroup): void {
  if (![settings.start, settings.end, settings.interval].every(Number.isFinite)) throw new TypeError('Pivot number group bounds must be finite')
  if (settings.end <= settings.start) throw new RangeError('Pivot number group end must be greater than start')
  if (settings.interval <= 0) throw new RangeError('Pivot number group interval must be positive')
}

function normalizeManualGroups(groups: PivotManualGroup[]): PivotManualGroup[] {
  if (!Array.isArray(groups)) throw new TypeError('Pivot manual groups must be an array')
  return groups.map((group, index) => {
    if (!group || typeof group.name !== 'string' || !group.name.trim()) throw new TypeError(`Pivot manual group ${index} requires a name`)
    if (!Array.isArray(group.members) || group.members.length < 2) throw new RangeError(`Pivot manual group ${index} must contain at least two members`)
    const members = group.members.filter((member, memberIndex, all) => all.findIndex((candidate) => Object.is(candidate, member)) === memberIndex)
    if (members.length < 2) throw new RangeError(`Pivot manual group ${index} must contain at least two distinct members`)
    return { name: group.name.trim(), members: [...members] }
  })
}

function overlaps(a: CellRange, b: CellRange): boolean {
  return !(a.endRow < b.startRow || a.startRow > b.endRow || a.endCol < b.startCol || a.startCol > b.endCol)
}

function contains(range: CellRange, row: number, col: number): boolean {
  return row >= range.startRow && row <= range.endRow && col >= range.startCol && col <= range.endCol
}

function validateRange(range: CellRange): void {
  validateIndex(range.startRow, 'source start row')
  validateIndex(range.startCol, 'source start column')
  validateIndex(range.endRow, 'source end row')
  validateIndex(range.endCol, 'source end column')
  if (range.endRow < range.startRow || range.endCol < range.startCol) throw new RangeError('Invalid pivot source range')
}

function validateIndex(value: number, name: string): void {
  if (!Number.isInteger(value) || value < 0) throw new RangeError(`${name} must be a non-negative integer`)
}
